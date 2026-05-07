/// Token 类型
#[derive(Debug, Clone, PartialEq)]
pub enum Token {
    /// 普通文本
    Text(String),

    /// 单值引用: {{tag.field}} 或 {{tag.field/expr:format}}
    Variable {
        tag: String,
        field: String,
        /// 值运算表达式，如 "/1000", "*100", "(+100)/2"
        expression: Option<String>,
        format: Option<String>,
    },

    /// 行迭代标记: {{@tag.field}} 或 {{@tag.field/expr:format}}
    RowIterate {
        tag: String,
        field: String,
        expression: Option<String>,
        format: Option<String>,
    },

    /// 单值变量间运算: {{tag1.field1*tag2.field2:format}}
    VariableBinOp {
        tag1: String,
        field1: String,
        operator: String,
        tag2: String,
        field2: String,
        format: Option<String>,
    },

    /// 行迭代变量间运算: {{@(tag1.field1*tag2.field2):format}}
    RowIterateBinOp {
        tag1: String,
        field1: String,
        operator: String,
        tag2: String,
        field2: String,
        format: Option<String>,
    },

    /// Excel 公式: {{=公式}}
    Formula {
        /// 公式内容，如 "A{r}*B{r}"
        formula: String,
        format: Option<String>,
    },

    /// 行迭代公式: {{@=公式}}
    RowIterateFormula {
        formula: String,
        format: Option<String>,
    },

    /// 聚合函数: {{func(tag.field)}} 或 {{func(tag)}} (count)
    Aggregate {
        func: AggFunc,
        tag: String,
        field: Option<String>,
        format: Option<String>,
    },
}

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

/// 单元格值类型（用于渲染）
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// 数值
    Number(f64),
    /// 文本
    Text(String),
    /// Excel 公式
    Formula(String),
}

/// 格式化后的单元格结果
#[derive(Debug, Clone)]
pub struct FormattedCell {
    /// 值
    pub value: CellValue,
    /// Excel 格式代码（如 "0.00", "0.00%", "yyyy-mm-dd"）
    pub format_code: Option<String>,
}

impl FormattedCell {
    /// 创建数值单元格
    pub fn number(value: f64, format_code: Option<String>) -> Self {
        Self {
            value: CellValue::Number(value),
            format_code,
        }
    }

    /// 创建文本单元格
    pub fn text(value: String) -> Self {
        Self {
            value: CellValue::Text(value),
            format_code: None,
        }
    }

    /// 创建公式单元格
    pub fn formula(formula: String, format_code: Option<String>) -> Self {
        Self {
            value: CellValue::Formula(formula),
            format_code,
        }
    }
}