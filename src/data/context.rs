use crate::error::Result;
use crate::data::source::DataSource;
use crate::data::excel_source::ExcelSource;
use crate::data::json_source::JsonSource;
use std::collections::HashMap;
use std::path::Path;

/// 数据上下文，管理所有数据源
pub struct DataContext {
    /// tag -> 数据源
    sources: HashMap<String, Box<dyn DataSource>>,
}

impl DataContext {
    /// 创建空上下文
    pub fn new() -> Self {
        Self {
            sources: HashMap::new(),
        }
    }

    /// 加载 Excel 数据源
    pub fn load_excel(&mut self, tag: &str, path: &Path) -> Result<()> {
        let source = ExcelSource::load(path)?;
        self.sources.insert(tag.to_string(), Box::new(source));
        Ok(())
    }

    /// 加载 JSON 数据源
    pub fn load_json(&mut self, tag: &str, path: &Path) -> Result<()> {
        let source = JsonSource::load(path)?;
        self.sources.insert(tag.to_string(), Box::new(source));
        Ok(())
    }

    /// 获取数据源
    pub fn get_source(&self, tag: &str) -> Option<&dyn DataSource> {
        self.sources.get(tag).map(|s| s.as_ref())
    }
}

impl Default for DataContext {
    fn default() -> Self {
        Self::new()
    }
}