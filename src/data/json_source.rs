use crate::error::{ReportError, Result};
use crate::data::source::{DataSource, SourceType};
use serde_json::Value;
use std::collections::HashMap;
use std::path::Path;

/// JSON 数据源（键值对）
pub struct JsonSource {
    /// 键值对数据
    data: HashMap<String, Value>,
}

impl JsonSource {
    /// 从文件加载 JSON 数据
    pub fn load(path: &Path) -> Result<Self> {
        if !path.exists() {
            return Err(ReportError::DataFileNotFound(path.to_path_buf()));
        }

        let content = std::fs::read_to_string(path)
            .map_err(|e| ReportError::JsonDataRead(format!("{}: {}", path.display(), e)))?;

        // 解析 JSON 对象
        let json: serde_json::Value = serde_json::from_str(&content)
            .map_err(|e| ReportError::JsonDataRead(format!("JSON 解析失败: {}", e)))?;

        // 只支持 JSON 对象（键值对）
        let obj = json.as_object()
            .ok_or_else(|| ReportError::JsonDataRead("JSON 数据必须是对象格式".to_string()))?;

        // 转换为 HashMap
        let data: HashMap<String, Value> = obj.iter()
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect();

        Ok(Self { data })
    }
}

impl DataSource for JsonSource {
    fn source_type(&self) -> SourceType {
        SourceType::Json
    }

    fn get_value(&self, field: &str) -> Option<Value> {
        self.data.get(field).cloned()
    }

    fn get_rows(&self) -> Option<&[crate::data::source::DataRow]> {
        None // JSON 不支持行迭代
    }

    fn row_count(&self) -> usize {
        0
    }
}