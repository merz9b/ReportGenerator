//! Word (docx) 渲染器
//!
//! 使用 rdocx 处理 docx 文件，支持完整的行迭代功能

use crate::data::{DataContext, DataRow};
use crate::error::{ReportError, Result};
use crate::engine::document_ext::DocumentExt;
use crate::template::{AggFunc, AggTarget, AstNode, Expression, FormatSpec, FormatType, Modifier, Parser, Operator};
use rdocx::Document;
use serde_json::Value;
use std::collections::HashMap;
use std::path::Path;

/// Word 渲染器
pub struct WordRenderer<'a> {
    context: &'a DataContext,
}

impl<'a> WordRenderer<'a> {
    pub fn new(context: &'a DataContext) -> Self {
        Self { context }
    }

    /// 加载并渲染 docx 文件
    pub fn render_file(&self, template_path: &Path, output_path: &Path) -> Result<()> {
        let mut doc = Document::open(template_path)
            .map_err(|e| ReportError::WordTemplateRead(format!("读取失败: {}", e)))?;

        self.render_document(&mut doc)?;

        doc.save(output_path)
            .map_err(|e| ReportError::WordOutput(format!("写入失败: {}", e)))?;

        Ok(())
    }

    /// 渲染文档
    fn render_document(&self, doc: &mut Document) -> Result<()> {
        // 1. 先处理表格行迭代（从后往前处理）
        self.process_table_row_iterations(doc)?;

        // 2. 收集所有需要替换的占位符（非行迭代）
        let replacements = self.collect_single_replacements(doc)?;

        // 3. 使用 replace_all 批量替换
        let hash_map: HashMap<&str, &str> = replacements.iter()
            .map(|(k, v)| (k.as_str(), v.as_str()))
            .collect();
        doc.replace_all(&hash_map);

        Ok(())
    }

    /// 处理表格行迭代
    fn process_table_row_iterations(&self, doc: &mut Document) -> Result<()> {
        // 扫描表格，收集行迭代信息
        let tables_info = self.scan_tables_for_iteration(doc);

        // 从后往前处理表格（避免索引偏移）
        for (table_idx, iterate_rows) in tables_info.into_iter().rev() {
            self.process_table_iteration(doc, table_idx, iterate_rows)?;
        }

        Ok(())
    }

    /// 扫描表格，收集行迭代信息
    fn scan_tables_for_iteration(&self, doc: &Document) -> Vec<(usize, Vec<IterateRowInfo>)> {
        let mut result = Vec::new();

        let tables = doc.tables();
        for (t_idx, table) in tables.iter().enumerate() {
            let iterate_rows = self.scan_table_rows(table);
            if !iterate_rows.is_empty() {
                result.push((t_idx, iterate_rows));
            }
        }

        result
    }

    /// 扫描表格行，识别行迭代标记
    fn scan_table_rows(&self, table: &rdocx::TableRef<'_>) -> Vec<IterateRowInfo> {
        let mut result = Vec::new();

        for r_idx in 0..table.row_count() {
            if let Some(row) = table.row(r_idx) {
                for c_idx in 0..row.cell_count() {
                    if let Some(cell) = row.cell(c_idx) {
                        let text = cell.text();
                        // 检查行迭代标记（排除公式）
                        if text.contains("{{@") && !text.contains("{{@=") {
                            let parsed = Parser::parse_cell(&text);
                            if let Some(tag) = parsed.iterate_tag {
                                result.push(IterateRowInfo {
                                    row_idx: r_idx,
                                    tag,
                                });
                                break;
                            }
                        }
                    }
                }
            }
        }

        result
    }

    /// 处理单个表格的行迭代
    fn process_table_iteration(&self, doc: &mut Document, table_idx: usize, iterate_rows: Vec<IterateRowInfo>) -> Result<()> {
        // 从后往前处理行迭代（避免行索引偏移）
        for info in iterate_rows.into_iter().rev() {
            self.expand_table_rows(doc, table_idx, &info)?;
        }

        Ok(())
    }

    /// 扩展表格行
    fn expand_table_rows(&self, doc: &mut Document, table_idx: usize, info: &IterateRowInfo) -> Result<()> {
        // 获取数据源
        let source = self.context.get_source(&info.tag)
            .ok_or_else(|| ReportError::SourceNotFound { tag: info.tag.clone() })?;

        let rows = source.get_rows()
            .ok_or_else(|| ReportError::SourceNotIteratable { tag: info.tag.clone() })?;

        let template_row_idx = info.row_idx;
        let col_count = doc.get_table_column_count(table_idx);

        if rows.is_empty() {
            // 没有数据，清空模板行
            unsafe {
                for c in 0..col_count {
                    doc.set_table_cell_text(table_idx, template_row_idx, c, "");
                }
            }
            return Ok(());
        }

        // 先保存模板行的原始文本（在插入新行之前）
        let template_texts: Vec<String> = (0..col_count)
            .map(|c| doc.get_table_cell_text(table_idx, template_row_idx, c))
            .collect();

        // 插入新行（使用 unsafe 方法）
        // 从后往前插入，避免索引偏移
        unsafe {
            for _ in (1..rows.len()).rev() {
                doc.insert_table_row(table_idx, template_row_idx + 1, template_row_idx);
            }

            // 填充所有行的数据（使用保存的模板文本）
            for (data_idx, data_row) in rows.iter().enumerate() {
                let actual_row_idx = template_row_idx + data_idx;

                for c in 0..col_count {
                    let template_text = &template_texts[c];

                    if template_text.contains("{{") {
                        let rendered = self.render_row_text(template_text, data_row, data_idx)?;
                        doc.set_table_cell_text(table_idx, actual_row_idx, c, &rendered);
                    } else if data_idx > 0 {
                        // 非第一行，复制模板行的静态文本
                        doc.set_table_cell_text(table_idx, actual_row_idx, c, template_text);
                    }
                    // 第一行的静态文本保持不变（已在模板中）
                }
            }
        }

        Ok(())
    }

    /// 收集单值替换（非行迭代）
    fn collect_single_replacements(&self, doc: &Document) -> Result<Vec<(String, String)>> {
        let mut replacements = Vec::new();

        // 扫描段落
        for para in doc.paragraphs() {
            let text = para.text();
            self.extract_non_iterate_placeholders(&text, &mut replacements)?;
        }

        // 扫描表格
        let tables = doc.tables();
        for (t_idx, table) in tables.iter().enumerate() {
            for r_idx in 0..table.row_count() {
                if let Some(row) = table.row(r_idx) {
                    for c_idx in 0..row.cell_count() {
                        if let Some(cell) = row.cell(c_idx) {
                            let text = cell.text();
                            self.extract_non_iterate_placeholders(&text, &mut replacements)?;
                        }
                    }
                }
            }
        }

        Ok(replacements)
    }

    /// 提取非行迭代的占位符
    fn extract_non_iterate_placeholders(&self, text: &str, replacements: &mut Vec<(String, String)>) -> Result<()> {
        if !text.contains("{{") {
            return Ok(());
        }

        let parsed = Parser::parse_cell(text);

        // 跳过行迭代
        if parsed.iterate_tag.is_some() {
            return Ok(());
        }

        for node in &parsed.nodes {
            if let AstNode::Marker { modifier, expr, format } = node {
                // 跳过公式和行迭代
                if matches!(modifier, Modifier::Formula | Modifier::RowFormula | Modifier::RowIterate) {
                    continue;
                }

                let placeholder = self.build_placeholder(modifier, expr, format);
                let value = self.eval_and_format(expr, format)?;
                replacements.push((placeholder, value));
            }
        }

        Ok(())
    }

    /// 渲染行迭代文本
    fn render_row_text(&self, text: &str, data: &DataRow, idx: usize) -> Result<String> {
        let parsed = Parser::parse_cell(text);
        let mut result = String::new();

        for node in &parsed.nodes {
            match node {
                AstNode::Text(s) => result.push_str(s),
                AstNode::Marker { modifier, expr, format } => {
                    if matches!(modifier, Modifier::Formula | Modifier::RowFormula) {
                        continue;
                    }

                    match modifier {
                        Modifier::RowIterate => {
                            result.push_str(&self.eval_row(expr, data, idx, format)?);
                        },
                        Modifier::VariableOp => {
                            result.push_str(&self.eval_row(expr, data, idx, format)?);
                        },
                        Modifier::None => {
                            if expr.is_aggregate() {
                                result.push_str(&self.eval_and_format(expr, format)?);
                            } else {
                                result.push_str(&self.try_eval_from_row(expr, data, idx, format)?);
                            }
                        },
                        _ => {
                            result.push_str(&self.eval_and_format(expr, format)?);
                        }
                    }
                }
            }
        }

        Ok(result)
    }

    fn try_eval_from_row(&self, expr: &Expression, data: &DataRow, idx: usize, fmt: &Option<FormatSpec>) -> Result<String> {
        match expr {
            Expression::FieldRef(r) => {
                if let Some(v) = data.fields.get(&r.field) {
                    self.format_val(v.clone(), fmt)
                } else {
                    self.eval_and_format(expr, fmt)
                }
            },
            _ => self.eval_and_format(expr, fmt)
        }
    }

    fn build_placeholder(&self, modifier: &Modifier, expr: &Expression, format: &Option<FormatSpec>) -> String {
        let mut s = String::from("{{");

        match modifier {
            Modifier::None => {}
            Modifier::RowIterate => s.push('@'),
            Modifier::Formula => s.push('='),
            Modifier::RowFormula => s.push_str("@="),
            Modifier::VariableOp => s.push('#'),
            _ => {}
        }

        s.push_str(&self.expr_to_string(expr));

        if let Some(fmt) = format {
            s.push(':');
            s.push_str(&self.format_to_string(&fmt.format_type));
        }

        s.push_str("}}");
        s
    }

    fn expr_to_string(&self, expr: &Expression) -> String {
        match expr {
            Expression::FieldRef(r) => format!("{}.{}", r.tag, r.field),
            Expression::ValueExpr { base, op, operand } => {
                format!("{}.{}{}{}", base.tag, base.field, self.op_to_string(*op), operand)
            },
            Expression::BinaryOp { left, op, right } => {
                format!("({}{}{})", self.expr_to_string(left), self.op_to_string(*op), self.expr_to_string(right))
            },
            Expression::Aggregate { func, target } => {
                let func_str = match func {
                    AggFunc::Sum => "sum",
                    AggFunc::Count => "count",
                    AggFunc::Avg => "avg",
                    AggFunc::Max => "max",
                    AggFunc::Min => "min",
                };
                match target {
                    AggTarget::Tag(t) => format!("{}({})", func_str, t),
                    AggTarget::Field(r) => format!("{}({}.{})", func_str, r.tag, r.field),
                }
            },
            Expression::Number(n) => format!("{}", n),
            _ => String::new(),
        }
    }

    fn op_to_string(&self, op: Operator) -> &'static str {
        match op {
            Operator::Add => "+",
            Operator::Sub => "-",
            Operator::Mul => "*",
            Operator::Div => "/",
        }
    }

    fn format_to_string(&self, ft: &FormatType) -> String {
        match ft {
            FormatType::Int => "int".into(),
            FormatType::Float(d) => if *d == 2 { "float".into() } else { format!("float{}", d) },
            FormatType::Percent(d) => if *d == 2 { "pct".into() } else { format!("pct{}", d) },
            FormatType::Currency(s) => if s.is_empty() { "currency".into() } else { format!("currency:{}", s) },
            FormatType::Pad(w, c) => format!("pad:{}{}", w, c),
            FormatType::Date(d) => if d.is_empty() { "date".into() } else { format!("date{}", d) },
            FormatType::Text => "text".into(),
            FormatType::Custom(c) => c.clone(),
        }
    }

    fn eval_and_format(&self, expr: &Expression, fmt: &Option<FormatSpec>) -> Result<String> {
        match expr {
            Expression::FieldRef(r) => {
                let src = self.context.get_source(&r.tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: r.tag.clone() })?;
                let v = src.get_value(&r.field)
                    .ok_or_else(|| ReportError::FieldNotFound { tag: r.tag.clone(), field: r.field.clone() })?;
                self.format_val(v, fmt)
            },
            Expression::ValueExpr { base, op, operand } => {
                let src = self.context.get_source(&base.tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: base.tag.clone() })?;
                let v = src.get_value(&base.field)
                    .ok_or_else(|| ReportError::FieldNotFound { tag: base.tag.clone(), field: base.field.clone() })?;
                self.format_num(self.apply_op(self.to_f64(&v), *op, *operand), fmt)
            },
            Expression::BinaryOp { left, op, right } => {
                let l = self.eval_expr(left)?;
                let r = self.eval_expr(right)?;
                self.format_num(self.apply_op(self.to_f64(&l), *op, self.to_f64(&r)), fmt)
            },
            Expression::Number(n) => self.format_num(*n, fmt),
            Expression::Aggregate { func, target } => {
                let (tag, field) = match target {
                    AggTarget::Tag(t) => (t.clone(), None),
                    AggTarget::Field(r) => (r.tag.clone(), Some(r.field.clone())),
                };
                let src = self.context.get_source(&tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: tag.clone() })?;
                let data_rows = src.get_rows()
                    .ok_or_else(|| ReportError::SourceNotIteratable { tag: tag.clone() })?;
                self.format_val(self.aggregate(*func, data_rows, &field)?, fmt)
            },
            _ => Ok(String::new()),
        }
    }

    fn eval_expr(&self, expr: &Expression) -> Result<Value> {
        match expr {
            Expression::FieldRef(r) => {
                let src = self.context.get_source(&r.tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: r.tag.clone() })?;
                Ok(src.get_value(&r.field).unwrap_or(Value::Null))
            },
            Expression::Number(n) => Ok(Value::Number(serde_json::Number::from_f64(*n).unwrap_or_else(|| serde_json::Number::from(0)))),
            _ => Ok(Value::Null),
        }
    }

    fn eval_row(&self, expr: &Expression, data: &DataRow, idx: usize, fmt: &Option<FormatSpec>) -> Result<String> {
        match expr {
            Expression::FieldRef(r) => {
                let v = data.fields.get(&r.field).cloned().unwrap_or(Value::Null);
                self.format_val(v, fmt)
            },
            Expression::ValueExpr { base, op, operand } => {
                let v = data.fields.get(&base.field).unwrap_or(&Value::Null);
                self.format_num(self.apply_op(self.to_f64(v), *op, *operand), fmt)
            },
            Expression::BinaryOp { left, op, right } => {
                let l = self.eval_binop_side(left, data, idx)?;
                let r = self.eval_binop_side(right, data, idx)?;
                self.format_num(self.apply_op(self.to_f64(&l), *op, self.to_f64(&r)), fmt)
            },
            Expression::Number(n) => self.format_num(*n, fmt),
            _ => Ok(String::new()),
        }
    }

    fn eval_binop_side(&self, expr: &Expression, data: &DataRow, idx: usize) -> Result<Value> {
        match expr {
            Expression::FieldRef(r) => {
                if data.fields.contains_key(&r.field) {
                    Ok(data.fields.get(&r.field).cloned().unwrap_or(Value::Null))
                } else {
                    let src = self.context.get_source(&r.tag)
                        .ok_or_else(|| ReportError::SourceNotFound { tag: r.tag.clone() })?;
                    match src.get_rows() {
                        Some(rows) if idx < rows.len() => Ok(rows[idx].fields.get(&r.field).cloned().unwrap_or(Value::Null)),
                        _ => Ok(src.get_value(&r.field).unwrap_or(Value::Null)),
                    }
                }
            },
            Expression::Number(n) => Ok(Value::Number(serde_json::Number::from_f64(*n).unwrap_or_else(|| serde_json::Number::from(0)))),
            _ => Ok(Value::Null),
        }
    }

    fn aggregate(&self, func: AggFunc, rows: &[DataRow], field: &Option<String>) -> Result<Value> {
        match func {
            AggFunc::Count => Ok(Value::Number(serde_json::Number::from(rows.len()))),
            AggFunc::Sum | AggFunc::Avg | AggFunc::Max | AggFunc::Min => {
                let f = field.as_ref().ok_or_else(|| ReportError::TokenParse("聚合缺少字段".into()))?;
                let vals: Vec<f64> = rows.iter().filter_map(|r| r.fields.get(f).and_then(|v| v.as_f64())).collect();
                if vals.is_empty() { return Ok(Value::Null); }
                let res = match func {
                    AggFunc::Sum => vals.iter().sum(),
                    AggFunc::Avg => vals.iter().sum::<f64>() / vals.len() as f64,
                    AggFunc::Max => vals.iter().cloned().fold(f64::NEG_INFINITY, f64::max),
                    AggFunc::Min => vals.iter().cloned().fold(f64::INFINITY, f64::min),
                    AggFunc::Count => unreachable!(),
                };
                Ok(Value::Number(serde_json::Number::from_f64(res).unwrap_or_else(|| serde_json::Number::from(0))))
            }
        }
    }

    fn format_val(&self, v: Value, fmt: &Option<FormatSpec>) -> Result<String> {
        match fmt {
            Some(f) => self.apply_fmt(&v, &f.format_type),
            None => Ok(self.val_to_str(&v)),
        }
    }

    fn format_num(&self, n: f64, fmt: &Option<FormatSpec>) -> Result<String> {
        self.format_val(Value::Number(serde_json::Number::from_f64(n).unwrap_or_else(|| serde_json::Number::from(0))), fmt)
    }

    fn apply_fmt(&self, v: &Value, ft: &FormatType) -> Result<String> {
        let n = self.to_f64(v);
        Ok(match ft {
            FormatType::Int => format!("{}", n.round() as i64),
            FormatType::Float(d) => format!("{:.1$}", n, *d as usize),
            FormatType::Percent(d) => format!("{:.1$}%", n * 100.0, *d as usize),
            FormatType::Currency(s) => if s.is_empty() { format!("{:.2}", n) } else { format!("{}{:.2}", s, n) },
            FormatType::Pad(w, c) => if n != n.round() { format!("{} (pad仅支持整数)", n) } else {
                let s = format!("{:0>1$}", n as i64, *w);
                if *c == '0' { s } else { s.replace('0', &c.to_string()) }
            },
            FormatType::Date(_) | FormatType::Text | FormatType::Custom(_) => self.val_to_str(v),
        })
    }

    fn to_f64(&self, v: &Value) -> f64 {
        match v { Value::Number(n) => n.as_f64().unwrap_or(0.0), Value::String(s) => s.parse().ok().unwrap_or(0.0), Value::Bool(b) => if *b { 1.0 } else { 0.0 }, _ => 0.0 }
    }

    fn val_to_str(&self, v: &Value) -> String {
        match v {
            Value::Null => String::new(),
            Value::Bool(b) => b.to_string(),
            Value::Number(n) => { let x = n.as_f64().unwrap_or(0.0); if x == x.round() { format!("{}", x as i64) } else { format!("{:.2}", x) } },
            Value::String(s) => s.clone(),
            _ => "[复杂类型]".into(),
        }
    }

    fn apply_op(&self, l: f64, op: Operator, r: f64) -> f64 {
        match op { Operator::Add => l + r, Operator::Sub => l - r, Operator::Mul => l * r, Operator::Div => if r == 0.0 { 0.0 } else { l / r } }
    }
}

struct IterateRowInfo {
    row_idx: usize,
    tag: String,
}