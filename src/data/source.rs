use serde_json::Value;
use std::collections::HashMap;

/// 数据源类型
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SourceType {
    /// Excel 表格数据，支持行迭代
    Excel,
    /// JSON 键值对数据，仅单值替换
    Json,
}

/// 数据行（Excel 表格）
#[derive(Debug, Clone)]
pub struct DataRow {
    /// 行索引（0-based）
    pub index: usize,
    /// 字段名 -> 值
    pub fields: HashMap<String, Value>,
}

/// 数据源 trait
pub trait DataSource {
    /// 数据源类型
    fn source_type(&self) -> SourceType;

    /// 获取单值（JSON 数据或 Excel 首行）
    fn get_value(&self, field: &str) -> Option<Value>;

    /// 获取所有行数据（仅 Excel）
    fn get_rows(&self) -> Option<&[DataRow]>;

    /// 获取行数（仅 Excel）
    fn row_count(&self) -> usize;
}