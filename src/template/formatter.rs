//! 格式化器
//!
//! 根据AST表达式和数据上下文计算并格式化单元格值。

use crate::template::ast::*;
use crate::template::lexer::Operator;
use crate::template::value::{CellValue, FormattedCell};
use chrono::{NaiveDate, NaiveDateTime};
use serde_json::Value;

/// 格式化AST表达式
pub fn format_expression(expr: &Expression, data: &impl DataContext, format: Option<&FormatSpec>) -> FormattedCell {
    // 1. 计算表达式值
    let value = evaluate_expression(expr, data);

    // 2. 应用格式化
    if let Some(fmt) = format {
        apply_format(&value, &fmt.format_type)
    } else {
        match &value {
            Value::Number(n) => FormattedCell::number(n.as_f64().unwrap_or(0.0), None),
            Value::String(s) => FormattedCell::text(s.clone()),
            Value::Null => FormattedCell::text(String::new()),
            _ => FormattedCell::text("[复杂类型]".to_string()),
        }
    }
}

/// 计算表达式值
pub fn evaluate_expression(expr: &Expression, data: &impl DataContext) -> Value {
    match expr {
        Expression::FieldRef(refr) => {
            data.get_field_value(&refr.tag, &refr.field)
        },

        Expression::ValueExpr { base, op, operand } => {
            let base_value = data.get_field_value(&base.tag, &base.field);
            let base_num = value_to_f64(&base_value);
            let result = apply_operator(base_num, *op, *operand);
            number_to_value(result)
        },

        Expression::BinaryOp { left, op, right } => {
            let left_value = evaluate_expression(left, data);
            let right_value = evaluate_expression(right, data);
            let left_num = value_to_f64(&left_value);
            let right_num = value_to_f64(&right_value);
            let result = apply_operator(left_num, *op, right_num);
            number_to_value(result)
        },

        Expression::Number(n) => {
            number_to_value(*n)
        },

        Expression::Aggregate { func, target } => {
            // 聚合函数需要特殊处理，返回占位值
            // 实际计算在 renderer 中完成
            Value::Null
        },

        Expression::ExcelFormula { formula } => {
            // 公式不在这里计算，返回公式字符串
            Value::String(formula.clone())
        },

        Expression::Conditional { .. } => {
            // 条件表达式暂未实现
            Value::Null
        },
    }
}

/// 计算聚合函数
pub fn compute_aggregate(func: AggFunc, tag: &str, field: Option<&str>, data: &impl DataContext, format: Option<&FormatSpec>) -> FormattedCell {
    // 获取数据行
    let rows = data.get_rows(tag);

    let values: Vec<f64> = if let Some(field_name) = field {
        rows.iter()
            .filter_map(|row| row.get(field_name))
            .map(|v| value_to_f64(v))
            .collect()
    } else {
        // count 只需要行数
        vec![]
    };

    let result = match func {
        AggFunc::Sum => values.iter().sum(),
        AggFunc::Count => rows.len() as f64,
        AggFunc::Avg => {
            if values.is_empty() {
                0.0
            } else {
                values.iter().sum::<f64>() / values.len() as f64
            }
        },
        AggFunc::Max => values.iter().copied().fold(0.0_f64, |a: f64, b: f64| a.max(b)),
        AggFunc::Min => values.iter().copied().fold(0.0_f64, |a: f64, b: f64| a.min(b)),
    };

    if let Some(fmt) = format {
        apply_format(&number_to_value(result), &fmt.format_type)
    } else {
        FormattedCell::number(result, None)
    }
}

/// 应用格式化
fn apply_format(value: &Value, format_type: &FormatType) -> FormattedCell {
    match format_type {
        FormatType::Int => {
            let num = value_to_f64(value).round();
            FormattedCell::number(num, Some("0".to_string()))
        },

        FormatType::Float(digits) => {
            let num = value_to_f64(value);
            let zeros = "0".repeat(*digits as usize);
            FormattedCell::number(num, Some(format!("0.{}", zeros)))
        },

        FormatType::Percent(digits) => {
            let num = value_to_f64(value);
            let zeros = "0".repeat(*digits as usize);
            FormattedCell::number(num, Some(format!("0.{}%", zeros)))
        },

        FormatType::Date(code) => {
            let serial = date_to_excel_serial(value);
            FormattedCell::number(serial, Some(code.clone()))
        },

        FormatType::Currency(symbol) => {
            let num = value_to_f64(value);
            let format_code = if symbol.is_empty() {
                "#,##0.00".to_string()
            } else {
                format!("{}#,##0.00", symbol)
            };
            FormattedCell::number(num, Some(format_code))
        },

        FormatType::Pad(width, pad_char) => {
            let num = value_to_f64(value);
            if num != num.round() {
                FormattedCell::text(format!("{} (pad仅支持整数)", num))
            } else {
                let int_val = num as i64;
                let formatted = if *pad_char == '0' {
                    format!("{:0>width$}", int_val, width = *width)
                } else {
                    let with_zeros = format!("{:0>width$}", int_val, width = *width);
                    with_zeros.replace('0', &pad_char.to_string())
                };
                FormattedCell::text(formatted)
            }
        },

        FormatType::Text => {
            FormattedCell::text(value_to_string(value))
        },

        FormatType::Custom(code) => {
            let num = value_to_f64(value);
            FormattedCell::number(num, Some(code.clone()))
        },
    }
}

/// 应用运算符
fn apply_operator(left: f64, op: Operator, right: f64) -> f64 {
    match op {
        Operator::Add => left + right,
        Operator::Sub => left - right,
        Operator::Mul => left * right,
        Operator::Div => {
            if right == 0.0 {
                0.0
            } else {
                left / right
            }
        },
    }
}

/// 将 Value 转换为 f64
fn value_to_f64(value: &Value) -> f64 {
    match value {
        Value::Number(n) => n.as_f64().unwrap_or(0.0),
        Value::String(s) => s.parse::<f64>().ok().unwrap_or(0.0),
        Value::Bool(b) => if *b { 1.0 } else { 0.0 },
        Value::Null => 0.0,
        _ => 0.0,
    }
}

/// 将 f64 转换为 Value
fn number_to_value(n: f64) -> Value {
    Value::Number(serde_json::Number::from_f64(n).unwrap_or_else(|| serde_json::Number::from(0)))
}

/// 将 Value 转换为字符串
fn value_to_string(value: &Value) -> String {
    match value {
        Value::Null => String::new(),
        Value::Bool(b) => b.to_string(),
        Value::Number(n) => n.to_string(),
        Value::String(s) => s.clone(),
        Value::Array(_) | Value::Object(_) => "[复杂类型]".to_string(),
    }
}

/// 将日期转换为 Excel 日期序列数
fn date_to_excel_serial(value: &Value) -> f64 {
    match value {
        Value::String(s) => {
            let date = if let Ok(d) = NaiveDate::parse_from_str(s, "%Y/%m/%d") {
                d
            } else if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
                d
            } else if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
                d
            } else if let Ok(d) = NaiveDateTime::parse_from_str(s, "%Y/%m/%d %H:%M:%S") {
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

/// 数据上下文 trait
/// 用于从数据源获取值
pub trait DataContext {
    /// 获取字段值（单值或行数据）
    fn get_field_value(&self, tag: &str, field: &str) -> Value;

    /// 获取数据行（用于聚合函数和行迭代）
    fn get_rows(&self, tag: &str) -> Vec<Value>;
}

/// 简单的数据上下文实现（用于测试）
pub struct SimpleDataContext {
    values: std::collections::HashMap<String, Value>,
}

impl SimpleDataContext {
    pub fn new() -> Self {
        Self {
            values: std::collections::HashMap::new(),
        }
    }

    pub fn insert(&mut self, key: String, value: Value) {
        self.values.insert(key, value);
    }
}

impl DataContext for SimpleDataContext {
    fn get_field_value(&self, tag: &str, field: &str) -> Value {
        let key = format!("{}.{}", tag, field);
        self.values.get(&key).cloned().unwrap_or(Value::Null)
    }

    fn get_rows(&self, _tag: &str) -> Vec<Value> {
        vec![]
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_int() {
        let ctx = SimpleDataContext::new();
        let expr = Expression::Number(3.14);
        let fmt = FormatSpec::from_str("int");
        let cell = format_expression(&expr, &ctx, Some(&fmt));
        assert!(matches!(cell.value, CellValue::Number(3.0)));
    }

    #[test]
    fn test_format_float() {
        let ctx = SimpleDataContext::new();
        let expr = Expression::Number(3.14159);
        let fmt = FormatSpec::from_str("float:2");
        let cell = format_expression(&expr, &ctx, Some(&fmt));
        assert!(matches!(cell.value, CellValue::Number(_)));
        assert_eq!(cell.format_code, Some("0.00".to_string()));
    }

    #[test]
    fn test_format_percent() {
        let ctx = SimpleDataContext::new();
        let expr = Expression::Number(0.025);
        let fmt = FormatSpec::from_str("pct");
        let cell = format_expression(&expr, &ctx, Some(&fmt));
        assert_eq!(cell.format_code, Some("0.00%".to_string()));
    }
}