//! 抽象语法树 (AST) 定义
//!
//! 定义模板标记的AST节点结构，支持嵌套表达式和运算符优先级。

use crate::template::lexer::Operator;
use std::fmt;

/// 聚合函数类型
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum AggFunc {
    Sum,
    Count,
    Avg,
    Max,
    Min,
}

impl AggFunc {
    pub fn from_str(s: &str) -> Option<Self> {
        match s.to_lowercase().as_str() {
            "sum" => Some(AggFunc::Sum),
            "count" => Some(AggFunc::Count),
            "avg" => Some(AggFunc::Avg),
            "max" => Some(AggFunc::Max),
            "min" => Some(AggFunc::Min),
            _ => None,
        }
    }
}

/// 标记修饰符
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Modifier {
    /// 无修饰符 - 单值引用
    None,
    /// @ - 行迭代
    RowIterate,
    /// # - 变量运算
    VariableOp,
    /// = - Excel公式
    Formula,
    /// @= - 行迭代公式
    RowFormula,
    /// ? - 条件判断 (预留)
    Conditional,
}

/// 字段引用 tag.field
#[derive(Debug, Clone, PartialEq)]
pub struct FieldRef {
    pub tag: String,
    pub field: String,
}

impl FieldRef {
    pub fn new(tag: String, field: String) -> Self {
        Self { tag, field }
    }
}

impl fmt::Display for FieldRef {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}.{}", self.tag, self.field)
    }
}

/// 聚合函数目标
#[derive(Debug, Clone, PartialEq)]
pub enum AggTarget {
    /// 仅tag: count(tag)
    Tag(String),
    /// tag.field: sum(tag.field)
    Field(FieldRef),
}

/// 格式类型
#[derive(Debug, Clone, PartialEq)]
pub enum FormatType {
    /// 整数
    Int,
    /// 浮点数（指定小数位数）
    Float(u8),
    /// 百分比（指定小数位数）
    Percent(u8),
    /// 日期格式
    Date(String),
    /// 金额格式
    Currency(String),
    /// 数字填充（宽度，填充字符）
    Pad(usize, char),
    /// 文本
    Text,
    /// 自定义格式
    Custom(String),
}

impl FormatType {
    pub fn from_str(s: &str) -> Option<Self> {
        let lower = s.to_lowercase();

        if lower == "int" {
            return Some(FormatType::Int);
        }

        if lower == "float" || lower.starts_with("float:") {
            let digits = if lower == "float" {
                2
            } else {
                lower[6..].parse::<u8>().ok()?
            };
            return Some(FormatType::Float(digits));
        }

        if lower == "percent" || lower.starts_with("percent:") {
            let digits = if lower == "percent" {
                2
            } else {
                lower[8..].parse::<u8>().ok()?
            };
            return Some(FormatType::Percent(digits));
        }

        // pct 作为 percent 的简写
        if lower == "pct" || lower.starts_with("pct:") {
            let digits = if lower == "pct" {
                2
            } else {
                lower[4..].parse::<u8>().ok()?
            };
            return Some(FormatType::Percent(digits));
        }

        // pad 格式
        if lower.starts_with("pad:") {
            let rest = &lower[4..];
            let parts: Vec<&str> = rest.split(':').collect();
            let width: usize = parts[0].parse().ok()?;
            let pad_char = if parts.len() > 1 {
                parts[1].chars().next().unwrap_or('0')
            } else {
                '0'
            };
            return Some(FormatType::Pad(width, pad_char));
        }

        // date 格式
        if lower == "date" || lower.starts_with("date:") {
            let format_code = if lower == "date" {
                "yyyy-mm-dd".to_string()
            } else if lower == "date:ymd" {
                "yyyy/mm/dd".to_string()
            } else if lower == "date:dmy" {
                "dd-mm-yyyy".to_string()
            } else if lower == "date:dms" {
                "dd/mm/yyyy".to_string()
            } else if lower.starts_with("date(") {
                let end = lower.find(')').unwrap_or(lower.len());
                lower[5..end].to_string()
            } else {
                "yyyy-mm-dd".to_string()
            };
            return Some(FormatType::Date(format_code));
        }

        // currency 格式
        if lower == "currency" || lower.starts_with("currency:") {
            let symbol = if lower == "currency" {
                ""
            } else {
                &lower[9..]
            };
            return Some(FormatType::Currency(symbol.to_string()));
        }

        if lower == "text" {
            return Some(FormatType::Text);
        }

        // 默认作为自定义格式
        Some(FormatType::Custom(s.to_string()))
    }

    /// 获取 Excel 格式代码
    pub fn to_excel_format_code(&self) -> Option<String> {
        match self {
            FormatType::Int => Some("0".to_string()),
            FormatType::Float(digits) => {
                let zeros = "0".repeat(*digits as usize);
                Some(format!("0.{}", zeros))
            },
            FormatType::Percent(digits) => {
                let zeros = "0".repeat(*digits as usize);
                Some(format!("0.{}%", zeros))
            },
            FormatType::Date(code) => Some(code.clone()),
            FormatType::Currency(symbol) => {
                if symbol.is_empty() {
                    Some("#,##0.00".to_string())
                } else if symbol == "$" {
                    Some("$#,##0.00".to_string())
                } else if symbol == "¥" {
                    Some("¥#,##0.00".to_string())
                } else {
                    Some(format!("{}#,##0.00", symbol))
                }
            },
            FormatType::Pad(_, _) => Some("@".to_string()),
            FormatType::Custom(code) => Some(code.clone()),
            FormatType::Text => Some("@".to_string()),
        }
    }
}

/// 格式规范
#[derive(Debug, Clone, PartialEq)]
pub struct FormatSpec {
    pub format_type: FormatType,
    pub raw: String,
}

impl FormatSpec {
    pub fn from_str(s: &str) -> Self {
        let format_type = FormatType::from_str(s).unwrap_or(FormatType::Custom(s.to_string()));
        Self {
            format_type,
            raw: s.to_string(),
        }
    }

    pub fn format_code(&self) -> Option<String> {
        self.format_type.to_excel_format_code()
    }
}

/// 表达式类型
#[derive(Debug, Clone, PartialEq)]
pub enum Expression {
    /// 变量引用: tag.field
    FieldRef(FieldRef),

    /// 值运算: tag.field/N 或 tag.field*N
    /// 由 Parser 将运算合并到 ValueExpr
    ValueExpr {
        base: FieldRef,
        op: Operator,
        operand: f64,
    },

    /// 二元运算: left op right
    /// 支持变量间运算和嵌套运算
    BinaryOp {
        left: Box<Expression>,
        op: Operator,
        right: Box<Expression>,
    },

    /// 数字字面量
    Number(f64),

    /// 聚合函数: sum(tag.field), count(tag)
    Aggregate {
        func: AggFunc,
        target: AggTarget,
    },

    /// Excel公式: =A{r}*B{r}
    ExcelFormula {
        formula: String,
    },

    /// 条件表达式: ?cond:expr (预留)
    Conditional {
        cond: Box<Expression>,
        then_expr: Box<Expression>,
        else_expr: Option<Box<Expression>>,
    },
}

impl Expression {
    /// 判断是否是聚合函数
    pub fn is_aggregate(&self) -> bool {
        matches!(self, Expression::Aggregate { .. })
    }

    /// 判断是否涉及行迭代
    pub fn needs_row_iteration(&self) -> bool {
        // 递归检查是否有 FieldRef
        match self {
            Expression::FieldRef(_) => true,
            Expression::ValueExpr { .. } => true,
            Expression::BinaryOp { left, right, .. } => {
                left.needs_row_iteration() || right.needs_row_iteration()
            },
            Expression::Number(_) => false,
            Expression::Aggregate { .. } => false,
            Expression::ExcelFormula { .. } => false,
            Expression::Conditional { .. } => false,
        }
    }

    /// 获取表达式中的所有 tag
    pub fn get_tags(&self) -> Vec<String> {
        let mut tags = Vec::new();
        self.collect_tags(&mut tags);
        tags
    }

    fn collect_tags(&self, tags: &mut Vec<String>) {
        match self {
            Expression::FieldRef(refr) => {
                if !tags.contains(&refr.tag) {
                    tags.push(refr.tag.clone());
                }
            },
            Expression::ValueExpr { base, .. } => {
                if !tags.contains(&base.tag) {
                    tags.push(base.tag.clone());
                }
            },
            Expression::BinaryOp { left, right, .. } => {
                left.collect_tags(tags);
                right.collect_tags(tags);
            },
            Expression::Aggregate { target, .. } => {
                match target {
                    AggTarget::Tag(tag) => {
                        if !tags.contains(tag) {
                            tags.push(tag.clone());
                        }
                    },
                    AggTarget::Field(refr) => {
                        if !tags.contains(&refr.tag) {
                            tags.push(refr.tag.clone());
                        }
                    },
                }
            },
            _ => {}
        }
    }
}

impl fmt::Display for Expression {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Expression::FieldRef(refr) => write!(f, "{}", refr),
            Expression::ValueExpr { base, op, operand } => {
                write!(f, "{}{}{}", base, op.to_char(), operand)
            },
            Expression::BinaryOp { left, op, right } => {
                write!(f, "({}{}{})", left, op.to_char(), right)
            },
            Expression::Number(n) => write!(f, "{}", n),
            Expression::Aggregate { func, target } => {
                let func_name = match func {
                    AggFunc::Sum => "sum",
                    AggFunc::Count => "count",
                    AggFunc::Avg => "avg",
                    AggFunc::Max => "max",
                    AggFunc::Min => "min",
                };
                match target {
                    AggTarget::Tag(tag) => write!(f, "{}({})", func_name, tag),
                    AggTarget::Field(refr) => write!(f, "{}({})", func_name, refr),
                }
            },
            Expression::ExcelFormula { formula } => write!(f, "={}", formula),
            Expression::Conditional { cond, then_expr, .. } => {
                write!(f, "?{}:{}", cond, then_expr)
            },
        }
    }
}

/// AST节点
#[derive(Debug, Clone, PartialEq)]
pub enum AstNode {
    /// 标记节点 {{ ... }}
    Marker {
        modifier: Modifier,
        expr: Expression,
        format: Option<FormatSpec>,
    },

    /// 普通文本
    Text(String),
}

impl AstNode {
    /// 判断是否是标记节点
    pub fn is_marker(&self) -> bool {
        matches!(self, AstNode::Marker { .. })
    }

    /// 判断是否是行迭代标记
    pub fn is_row_iterate(&self) -> bool {
        match self {
            AstNode::Marker { modifier, .. } => {
                matches!(modifier, Modifier::RowIterate | Modifier::RowFormula)
            },
            _ => false,
        }
    }

    /// 判断是否包含聚合函数
    pub fn has_aggregate(&self) -> bool {
        match self {
            AstNode::Marker { expr, .. } => expr.is_aggregate(),
            _ => false,
        }
    }

    /// 获取标记的表达式（如果是标记节点）
    pub fn get_expression(&self) -> Option<&Expression> {
        match self {
            AstNode::Marker { expr, .. } => Some(expr),
            _ => None,
        }
    }

    /// 获取标记的修饰符
    pub fn get_modifier(&self) -> Option<Modifier> {
        match self {
            AstNode::Marker { modifier, .. } => Some(*modifier),
            _ => None,
        }
    }

    /// 获取标记的格式
    pub fn get_format(&self) -> Option<&FormatSpec> {
        match self {
            AstNode::Marker { format, .. } => format.as_ref(),
            _ => None,
        }
    }
}

impl fmt::Display for AstNode {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            AstNode::Marker { modifier, expr, format } => {
                let mod_str = match modifier {
                    Modifier::None => "",
                    Modifier::RowIterate => "@",
                    Modifier::VariableOp => "#",
                    Modifier::Formula => "=",
                    Modifier::RowFormula => "@=",
                    Modifier::Conditional => "?",
                };
                let fmt_str = format
                    .as_ref()
                    .map(|s| format!(":{}", s.raw))
                    .unwrap_or_default();
                write!(f, "{{{{{}{}{}}}}}", mod_str, expr, fmt_str)
            },
            AstNode::Text(s) => write!(f, "{}", s),
        }
    }
}

/// 解析后的单元格结果
#[derive(Debug)]
pub struct ParsedCell {
    /// AST节点列表
    pub nodes: Vec<AstNode>,
    /// 是否包含行迭代标记
    pub is_row_iterate: bool,
    /// 行迭代的 tag（如果有）
    pub iterate_tag: Option<String>,
    /// 是否包含聚合函数
    pub has_aggregate: bool,
}

impl ParsedCell {
    pub fn from_nodes(nodes: Vec<AstNode>) -> Self {
        let is_row_iterate = nodes.iter().any(|n| n.is_row_iterate());
        let has_aggregate = nodes.iter().any(|n| n.has_aggregate());

        // 获取行迭代的 tag
        let iterate_tag = nodes.iter()
            .filter_map(|n| {
                if n.is_row_iterate() {
                    n.get_expression()?.get_tags().first().cloned()
                } else {
                    None
                }
            })
            .next();

        Self {
            nodes,
            is_row_iterate,
            iterate_tag,
            has_aggregate,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_field_ref_display() {
        let refr = FieldRef::new("tag".to_string(), "field".to_string());
        assert_eq!(format!("{}", refr), "tag.field");
    }

    #[test]
    fn test_format_type_from_str() {
        assert!(matches!(FormatType::from_str("int"), Some(FormatType::Int)));
        assert!(matches!(FormatType::from_str("float"), Some(FormatType::Float(2))));
        assert!(matches!(FormatType::from_str("float:4"), Some(FormatType::Float(4))));
        assert!(matches!(FormatType::from_str("pct"), Some(FormatType::Percent(2))));
        assert!(matches!(FormatType::from_str("pct:1"), Some(FormatType::Percent(1))));
    }

    #[test]
    fn test_expression_display() {
        let expr = Expression::FieldRef(FieldRef::new("tag".to_string(), "field".to_string()));
        assert_eq!(format!("{}", expr), "tag.field");
    }
}