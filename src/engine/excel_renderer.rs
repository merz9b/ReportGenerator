//! Excel 渲染器
//!
//! 使用AST解析结果渲染Excel模板

use crate::data::{DataContext, DataRow};
use crate::error::{ReportError, Result};
use crate::template::{
    AggFunc, AggTarget, AstNode, Expression, FieldRef, FormatSpec,
    FormatType, Modifier, Parser, replace_row_placeholders,
    CellValue, FormattedCell
};
use umya_spreadsheet::{Spreadsheet, Worksheet, Style, reader::xlsx};
use serde_json::Value;
use std::path::Path;
use std::collections::HashMap;

/// Excel 渲染器
pub struct ExcelRenderer<'a> {
    context: &'a DataContext,
}

impl<'a> ExcelRenderer<'a> {
    pub fn new(context: &'a DataContext) -> Self {
        Self { context }
    }

    /// 从模板加载 workbook
    pub fn load_template(template_path: &Path) -> Result<Spreadsheet> {
        if !template_path.exists() {
            return Err(ReportError::TemplateNotFound(template_path.to_path_buf()));
        }

        xlsx::read(template_path)
            .map_err(|e| ReportError::ExcelTemplateRead(format!("{}: {}", template_path.display(), e)))
    }

    /// 渲染指定 sheet
    pub fn render_sheet(&self, sheet: &mut Worksheet) -> Result<()> {
        // 扫描 sheet 中的标记
        let scan_result = self.scan_sheet(sheet)?;

        // 记录行偏移（行迭代插入新行后，后续行的位置需要调整）
        let mut row_offsets: HashMap<usize, usize> = HashMap::new();

        // 处理行迭代（从下往上）
        for row_info in scan_result.iterate_rows.iter().rev() {
            let original_row = row_info.row;
            let inserted_rows = self.process_row_iteration(sheet, row_info)?;

            if inserted_rows > 0 {
                for (row, offset) in row_offsets.iter_mut() {
                    if *row > original_row {
                        *offset += inserted_rows;
                    }
                }
                row_offsets.insert(original_row + 1, inserted_rows);
            }

            tracing::info!("行迭代完成: tag={}, 原模板行={}, 插入行数={}",
                row_info.tag, original_row, inserted_rows);
        }

        // 处理单值变量
        for cell_info in &scan_result.variable_cells {
            let adjusted_row = self.adjust_row_position(cell_info.row, &row_offsets);
            self.process_marker(sheet, cell_info.col, adjusted_row, &cell_info.node)?;
        }

        // 处理聚合函数
        for cell_info in &scan_result.aggregate_cells {
            let adjusted_row = self.adjust_row_position(cell_info.row, &row_offsets);
            self.process_aggregate(sheet, cell_info.col, adjusted_row, &cell_info.node)?;
        }

        // 处理公式
        for cell_info in &scan_result.formula_cells {
            let adjusted_row = self.adjust_row_position(cell_info.row, &row_offsets);
            self.process_formula(sheet, cell_info.col, adjusted_row, &cell_info.node)?;
        }

        Ok(())
    }

    /// 根据偏移调整行位置
    fn adjust_row_position(&self, original_row: usize, offsets: &HashMap<usize, usize>) -> usize {
        let mut adjusted = original_row;
        for (row, offset) in offsets.iter() {
            if *row <= original_row {
                adjusted += *offset;
            }
        }
        adjusted
    }

    /// 扫描 sheet，识别所有标记
    fn scan_sheet(&self, sheet: &Worksheet) -> Result<ScanResult> {
        let mut iterate_rows: Vec<IterateRowInfo> = Vec::new();
        let mut variable_cells: Vec<MarkerCellInfo> = Vec::new();
        let mut aggregate_cells: Vec<MarkerCellInfo> = Vec::new();
        let mut formula_cells: Vec<MarkerCellInfo> = Vec::new();

        let max_row = sheet.get_highest_row() as usize;
        let max_col = sheet.get_highest_column() as usize;

        for row in 1..=max_row {
            let mut row_has_iterate = false;
            let mut row_columns: Vec<IterateColumnInfo> = Vec::new();

            for col in 1..=max_col {
                let value_str = sheet.get_value((col as u32, row as u32)).to_string();
                if value_str.is_empty() {
                    continue;
                }

                let parsed = Parser::parse_cell(&value_str);

                for node in &parsed.nodes {
                    match node {
                        AstNode::Marker { modifier, expr, format } => {
                            match modifier {
                                Modifier::RowIterate => {
                                    row_has_iterate = true;
                                    row_columns.push(IterateColumnInfo {
                                        col: col - 1,
                                        expr: expr.clone(),
                                        format: format.clone(),
                                    });
                                },
                                Modifier::RowFormula => {
                                    row_has_iterate = true;
                                    row_columns.push(IterateColumnInfo {
                                        col: col - 1,
                                        expr: expr.clone(),
                                        format: format.clone(),
                                    });
                                },
                                Modifier::Formula => {
                                    formula_cells.push(MarkerCellInfo {
                                        row: row - 1,
                                        col: col - 1,
                                        node: node.clone(),
                                    });
                                },
                                Modifier::None | Modifier::VariableOp => {
                                    if expr.is_aggregate() {
                                        aggregate_cells.push(MarkerCellInfo {
                                            row: row - 1,
                                            col: col - 1,
                                            node: node.clone(),
                                        });
                                    } else {
                                        variable_cells.push(MarkerCellInfo {
                                            row: row - 1,
                                            col: col - 1,
                                            node: node.clone(),
                                        });
                                    }
                                },
                                Modifier::Conditional => {
                                    // 条件表达式暂不处理
                                },
                            }
                        },
                        AstNode::Text(_) => {}
                    }
                }
            }

            if row_has_iterate && !row_columns.is_empty() {
                // 获取主 tag
                let primary_tag = row_columns.iter()
                    .filter_map(|c| c.expr.get_tags().first().cloned())
                    .next()
                    .unwrap_or_default();

                iterate_rows.push(IterateRowInfo {
                    row: row - 1,
                    tag: primary_tag,
                    columns: row_columns,
                });
            }
        }

        Ok(ScanResult {
            iterate_rows,
            variable_cells,
            aggregate_cells,
            formula_cells,
        })
    }

    /// 处理行迭代 - 返回插入的行数
    fn process_row_iteration(&self, sheet: &mut Worksheet, row_info: &IterateRowInfo) -> Result<usize> {
        // 获取数据源
        let rows = if !row_info.tag.is_empty() {
            let source = self.context.get_source(&row_info.tag)
                .ok_or_else(|| ReportError::SourceNotFound { tag: row_info.tag.clone() })?;

            if source.get_rows().is_none() {
                return Err(ReportError::SourceNotIteratable { tag: row_info.tag.clone() });
            }
            source.get_rows().unwrap()
        } else {
            // 纯公式行迭代，没有数据源 - 无法处理
            return Ok(0);
        };

        let row_count = rows.len();

        if row_count == 0 {
            // 清空标记行
            for col_info in &row_info.columns {
                sheet.get_cell_mut((col_info.col as u32 + 1, row_info.row as u32 + 1))
                    .set_value("");
            }
            return Ok(0);
        }

        // 收集模板行样式
        let template_row = row_info.row as u32 + 1;
        let max_col = sheet.get_highest_column();
        let template_styles: Vec<(u32, Style)> = (1..=max_col)
            .filter_map(|col| {
                let value = sheet.get_value((col, template_row)).to_string();
                if !value.is_empty() {
                    let style = sheet.get_style((col, template_row)).clone();
                    Some((col, style))
                } else {
                    None
                }
            })
            .collect();

        // 第一行：替换模板行
        {
            let data_row = &rows[0];
            for col_info in &row_info.columns {
                let col = col_info.col as u32 + 1;
                self.render_row_iterate_cell(sheet, col, template_row, &col_info.expr, &col_info.format, data_row, 0)?;
            }
        }

        // 插入新行
        for (data_idx, data_row) in rows.iter().skip(1).enumerate() {
            let insert_row = template_row + 1 + data_idx as u32;

            sheet.insert_new_row(&insert_row, &1);

            // 复制样式
            for (col, style) in &template_styles {
                sheet.set_style((*col, insert_row), style.clone());
            }

            // 设置数据
            for col_info in &row_info.columns {
                let col = col_info.col as u32 + 1;
                self.render_row_iterate_cell(sheet, col, insert_row, &col_info.expr, &col_info.format, data_row, data_idx + 1)?;
            }
        }

        Ok(row_count - 1)
    }

    /// 渲染行迭代单元格
    fn render_row_iterate_cell(
        &self,
        sheet: &mut Worksheet,
        col: u32,
        row: u32,
        expr: &Expression,
        format: &Option<FormatSpec>,
        data_row: &DataRow,
        row_index: usize,
    ) -> Result<()> {
        match expr {
            Expression::FieldRef(refr) => {
                let value = data_row.fields.get(&refr.field).unwrap_or(&Value::Null);
                let formatted = self.format_value(value, format)?;
                self.write_formatted_cell(sheet, col, row, &formatted)?;
            },
            Expression::ValueExpr { base, op, operand } => {
                let base_value = data_row.fields.get(&base.field).unwrap_or(&Value::Null);
                let base_num = self.value_to_f64(base_value);
                let result = self.apply_operator(base_num, *op, *operand);
                let formatted = self.format_number(result, format)?;
                self.write_formatted_cell(sheet, col, row, &formatted)?;
            },
            Expression::BinaryOp { left, op, right } => {
                let left_value = self.evaluate_binop_side(left, data_row, row_index)?;
                let right_value = self.evaluate_binop_side(right, data_row, row_index)?;
                let left_num = self.value_to_f64(&left_value);
                let right_num = self.value_to_f64(&right_value);
                let result = self.apply_operator(left_num, *op, right_num);
                let formatted = self.format_number(result, format)?;
                self.write_formatted_cell(sheet, col, row, &formatted)?;
            },
            Expression::ExcelFormula { formula } => {
                let actual_formula = replace_row_placeholders(formula, row);
                let cell = sheet.get_cell_mut((col, row));
                cell.set_formula(&actual_formula);
                if let Some(fmt) = format {
                    self.apply_format_spec(cell, &fmt.format_type)?;
                }
            },
            _ => {
                sheet.get_cell_mut((col, row)).set_value("");
            },
        }
        Ok(())
    }

    /// 计算二元运算的一边（可能来自当前行或其他数据源）
    fn evaluate_binop_side(&self, expr: &Expression, data_row: &DataRow, row_index: usize) -> Result<Value> {
        match expr {
            Expression::FieldRef(refr) => {
                // 优先检查当前数据行是否有该字段（同tag情况）
                if data_row.fields.contains_key(&refr.field) {
                    Ok(data_row.fields.get(&refr.field).cloned().unwrap_or(Value::Null))
                } else {
                    // 来自其他 tag，需要获取对应行的数据
                    let source = self.context.get_source(&refr.tag)
                        .ok_or_else(|| ReportError::SourceNotFound { tag: refr.tag.clone() })?;
                    if let Some(rows) = source.get_rows() {
                        if row_index < rows.len() {
                            Ok(rows[row_index].fields.get(&refr.field).cloned().unwrap_or(Value::Null))
                        } else {
                            Ok(Value::Null)
                        }
                    } else {
                        Ok(source.get_value(&refr.field).unwrap_or(Value::Null))
                    }
                }
            },
            Expression::Number(n) => Ok(Value::Number(serde_json::Number::from_f64(*n)
                .unwrap_or_else(|| serde_json::Number::from(0)))),
            _ => Ok(Value::Null),
        }
    }

    /// 处理标记节点（单值变量）
    fn process_marker(&self, sheet: &mut Worksheet, col: usize, row: usize, node: &AstNode) -> Result<()> {
        if let AstNode::Marker { modifier, expr, format } = node {
            match modifier {
                Modifier::None | Modifier::VariableOp => {
                    let formatted = self.evaluate_and_format(expr, format)?;
                    self.write_formatted_cell(sheet, col as u32 + 1, row as u32 + 1, &formatted)?;
                },
                _ => {}
            }
        }
        Ok(())
    }

    /// 处理聚合函数
    fn process_aggregate(&self, sheet: &mut Worksheet, col: usize, row: usize, node: &AstNode) -> Result<()> {
        if let AstNode::Marker { expr, format, .. } = node {
            if let Expression::Aggregate { func, target } = expr {
                let (tag, field) = match target {
                    AggTarget::Tag(t) => (t.clone(), None),
                    AggTarget::Field(refr) => (refr.tag.clone(), Some(refr.field.clone())),
                };

                let source = self.context.get_source(&tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: tag.clone() })?;

                if source.get_rows().is_none() {
                    return Err(ReportError::SourceNotIteratable { tag: tag.clone() });
                }

                let rows = source.get_rows().unwrap();
                let result = self.compute_aggregate(*func, rows, &field)?;

                let formatted = self.format_value(&result, format)?;
                self.write_formatted_cell(sheet, col as u32 + 1, row as u32 + 1, &formatted)?;
            }
        }
        Ok(())
    }

    /// 处理公式
    fn process_formula(&self, sheet: &mut Worksheet, col: usize, row: usize, node: &AstNode) -> Result<()> {
        if let AstNode::Marker { expr, format, .. } = node {
            if let Expression::ExcelFormula { formula } = expr {
                let actual_row = row as u32 + 1;
                let actual_formula = replace_row_placeholders(formula, actual_row);

                let cell = sheet.get_cell_mut((col as u32 + 1, actual_row));
                cell.set_formula(&actual_formula);

                if let Some(fmt) = format {
                    self.apply_format_spec(cell, &fmt.format_type)?;
                }
            }
        }
        Ok(())
    }

    /// 计算聚合值
    fn compute_aggregate(&self, func: AggFunc, rows: &[DataRow], field: &Option<String>) -> Result<Value> {
        match func {
            AggFunc::Count => {
                Ok(Value::Number(serde_json::Number::from(rows.len())))
            }
            AggFunc::Sum | AggFunc::Avg | AggFunc::Max | AggFunc::Min => {
                let field_name = field.as_ref()
                    .ok_or_else(|| ReportError::TokenParse("聚合函数缺少字段名".to_string()))?;

                let values: Vec<f64> = rows.iter()
                    .filter_map(|row| {
                        row.fields.get(field_name)
                            .and_then(|v| v.as_f64())
                    })
                    .collect();

                if values.is_empty() {
                    return Ok(Value::Null);
                }

                let result = match func {
                    AggFunc::Sum => values.iter().sum(),
                    AggFunc::Avg => values.iter().sum::<f64>() / values.len() as f64,
                    AggFunc::Max => values.iter().cloned().fold(f64::NEG_INFINITY, |a, b| a.max(b)),
                    AggFunc::Min => values.iter().cloned().fold(f64::INFINITY, |a, b| a.min(b)),
                    AggFunc::Count => unreachable!(),
                };

                Ok(Value::Number(serde_json::Number::from_f64(result)
                    .unwrap_or_else(|| serde_json::Number::from(0))))
            }
        }
    }

    /// 计算表达式并格式化
    fn evaluate_and_format(&self, expr: &Expression, format: &Option<FormatSpec>) -> Result<FormattedCell> {
        match expr {
            Expression::FieldRef(refr) => {
                let source = self.context.get_source(&refr.tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: refr.tag.clone() })?;
                let value = source.get_value(&refr.field)
                    .ok_or_else(|| ReportError::FieldNotFound {
                        tag: refr.tag.clone(),
                        field: refr.field.clone(),
                    })?;
                self.format_value(&value, format)
            },
            Expression::ValueExpr { base, op, operand } => {
                let source = self.context.get_source(&base.tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: base.tag.clone() })?;
                let value = source.get_value(&base.field)
                    .ok_or_else(|| ReportError::FieldNotFound {
                        tag: base.tag.clone(),
                        field: base.field.clone(),
                    })?;
                let num = self.value_to_f64(&value);
                let result = self.apply_operator(num, *op, *operand);
                self.format_number(result, format)
            },
            Expression::BinaryOp { left, op, right } => {
                let left_value = self.evaluate_expr_value(left)?;
                let right_value = self.evaluate_expr_value(right)?;
                let left_num = self.value_to_f64(&left_value);
                let right_num = self.value_to_f64(&right_value);
                let result = self.apply_operator(left_num, *op, right_num);
                self.format_number(result, format)
            },
            Expression::Number(n) => {
                self.format_number(*n, format)
            },
            _ => Ok(FormattedCell::text(String::new())),
        }
    }

    /// 计算表达式值（用于单值）
    fn evaluate_expr_value(&self, expr: &Expression) -> Result<Value> {
        match expr {
            Expression::FieldRef(refr) => {
                let source = self.context.get_source(&refr.tag)
                    .ok_or_else(|| ReportError::SourceNotFound { tag: refr.tag.clone() })?;
                Ok(source.get_value(&refr.field).unwrap_or(Value::Null))
            },
            Expression::Number(n) => {
                Ok(Value::Number(serde_json::Number::from_f64(*n)
                    .unwrap_or_else(|| serde_json::Number::from(0))))
            },
            _ => Ok(Value::Null),
        }
    }

    /// 格式化值
    fn format_value(&self, value: &Value, format: &Option<FormatSpec>) -> Result<FormattedCell> {
        if let Some(fmt) = format {
            self.apply_format_type(value, &fmt.format_type)
        } else {
            match value {
                Value::Number(n) => Ok(FormattedCell::number(n.as_f64().unwrap_or(0.0), None)),
                Value::String(s) => Ok(FormattedCell::text(s.clone())),
                Value::Null => Ok(FormattedCell::text(String::new())),
                _ => Ok(FormattedCell::text("[复杂类型]".to_string())),
            }
        }
    }

    /// 格式化数值
    fn format_number(&self, num: f64, format: &Option<FormatSpec>) -> Result<FormattedCell> {
        if let Some(fmt) = format {
            self.apply_format_type(&Value::Number(serde_json::Number::from_f64(num)
                .unwrap_or_else(|| serde_json::Number::from(0))), &fmt.format_type)
        } else {
            Ok(FormattedCell::number(num, None))
        }
    }

    /// 应用格式类型
    fn apply_format_type(&self, value: &Value, format_type: &FormatType) -> Result<FormattedCell> {
        match format_type {
            FormatType::Int => {
                let num = self.value_to_f64(value).round();
                Ok(FormattedCell::number(num, Some("0".to_string())))
            },
            FormatType::Float(digits) => {
                let num = self.value_to_f64(value);
                let zeros = "0".repeat(*digits as usize);
                Ok(FormattedCell::number(num, Some(format!("0.{}", zeros))))
            },
            FormatType::Percent(digits) => {
                let num = self.value_to_f64(value);
                let zeros = "0".repeat(*digits as usize);
                Ok(FormattedCell::number(num, Some(format!("0.{}%", zeros))))
            },
            FormatType::Date(code) => {
                let serial = self.date_to_excel_serial(value);
                Ok(FormattedCell::number(serial, Some(code.clone())))
            },
            FormatType::Currency(symbol) => {
                let num = self.value_to_f64(value);
                let format_code = if symbol.is_empty() {
                    "#,##0.00".to_string()
                } else {
                    format!("{}#,##0.00", symbol)
                };
                Ok(FormattedCell::number(num, Some(format_code)))
            },
            FormatType::Pad(width, pad_char) => {
                let num = self.value_to_f64(value);
                if num != num.round() {
                    Ok(FormattedCell::text(format!("{} (pad仅支持整数)", num)))
                } else {
                    let int_val = num as i64;
                    let formatted = if *pad_char == '0' {
                        format!("{:0>width$}", int_val, width = *width)
                    } else {
                        let with_zeros = format!("{:0>width$}", int_val, width = *width);
                        with_zeros.replace('0', &pad_char.to_string())
                    };
                    Ok(FormattedCell::text(formatted))
                }
            },
            FormatType::Text => {
                Ok(FormattedCell::text(self.value_to_string(value)))
            },
            FormatType::Custom(code) => {
                let num = self.value_to_f64(value);
                Ok(FormattedCell::number(num, Some(code.clone())))
            },
        }
    }

    /// 应用格式规范到单元格
    fn apply_format_spec(&self, cell: &mut umya_spreadsheet::Cell, format_type: &FormatType) -> Result<()> {
        let format_code = match format_type {
            FormatType::Int => "0".to_string(),
            FormatType::Float(digits) => {
                let zeros = "0".repeat(*digits as usize);
                format!("0.{}", zeros)
            },
            FormatType::Percent(digits) => {
                let zeros = "0".repeat(*digits as usize);
                format!("0.{}%", zeros)
            },
            FormatType::Date(code) => code.clone(),
            FormatType::Currency(symbol) => {
                if symbol.is_empty() {
                    "#,##0.00".to_string()
                } else {
                    format!("{}#,##0.00", symbol)
                }
            },
            FormatType::Text => "@".to_string(),
            FormatType::Custom(code) => code.clone(),
            _ => "".to_string(),
        };

        if !format_code.is_empty() {
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format_code);
        }
        Ok(())
    }

    /// 写入格式化单元格
    fn write_formatted_cell(&self, sheet: &mut Worksheet, col: u32, row: u32, formatted: &FormattedCell) -> Result<()> {
        let cell = sheet.get_cell_mut((col, row));

        match &formatted.value {
            CellValue::Number(n) => {
                cell.set_value_number(*n);
            }
            CellValue::Text(s) => {
                cell.set_value(s);
            }
            CellValue::Formula(f) => {
                cell.set_formula(f);
            }
        }

        if let Some(format_code) = &formatted.format_code {
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format_code);
        }

        Ok(())
    }

    /// 应用运算符
    fn apply_operator(&self, left: f64, op: crate::template::lexer::Operator, right: f64) -> f64 {
        use crate::template::lexer::Operator;
        match op {
            Operator::Add => left + right,
            Operator::Sub => left - right,
            Operator::Mul => left * right,
            Operator::Div => {
                if right == 0.0 { 0.0 } else { left / right }
            },
        }
    }

    /// 将 Value 转换为 f64
    fn value_to_f64(&self, value: &Value) -> f64 {
        match value {
            Value::Number(n) => n.as_f64().unwrap_or(0.0),
            Value::String(s) => s.parse::<f64>().ok().unwrap_or(0.0),
            Value::Bool(b) => if *b { 1.0 } else { 0.0 },
            Value::Null => 0.0,
            _ => 0.0,
        }
    }

    /// 将 Value 转换为字符串
    fn value_to_string(&self, value: &Value) -> String {
        match value {
            Value::Null => String::new(),
            Value::Bool(b) => b.to_string(),
            Value::Number(n) => n.to_string(),
            Value::String(s) => s.clone(),
            _ => "[复杂类型]".to_string(),
        }
    }

    /// 将日期转换为 Excel 日期序列数
    fn date_to_excel_serial(&self, value: &Value) -> f64 {
        use chrono::{NaiveDate, NaiveDateTime};

        match value {
            Value::String(s) => {
                let date = if let Ok(d) = NaiveDate::parse_from_str(s, "%Y/%m/%d") {
                    d
                } else if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
                    d
                } else if let Ok(d) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
                    d.date()
                } else {
                    return 0.0;
                };

                let excel_epoch = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
                let days = date.signed_duration_since(excel_epoch).num_days();
                days as f64
            },
            Value::Number(n) => n.as_f64().unwrap_or(0.0),
            _ => 0.0,
        }
    }
}

/// 扫描结果
struct ScanResult {
    iterate_rows: Vec<IterateRowInfo>,
    variable_cells: Vec<MarkerCellInfo>,
    aggregate_cells: Vec<MarkerCellInfo>,
    formula_cells: Vec<MarkerCellInfo>,
}

/// 行迭代信息
struct IterateRowInfo {
    row: usize,
    tag: String,
    columns: Vec<IterateColumnInfo>,
}

/// 行迭代列信息
struct IterateColumnInfo {
    col: usize,
    expr: Expression,
    format: Option<FormatSpec>,
}

/// 标记单元格信息
struct MarkerCellInfo {
    row: usize,
    col: usize,
    node: AstNode,
}