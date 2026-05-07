//! 单元格值类型定义
//!
//! 定义单元格值的类型和格式化结果结构。

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

    /// 获取显示字符串（用于调试）
    pub fn to_display_string(&self) -> String {
        match &self.value {
            CellValue::Number(n) => format!("{}", n),
            CellValue::Text(s) => s.clone(),
            CellValue::Formula(f) => format!("={}", f),
        }
    }
}