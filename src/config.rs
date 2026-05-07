use crate::error::{ReportError, Result};
use serde::Deserialize;
use std::path::{Path, PathBuf};

/// 数据源类型
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SourceType {
    Excel,
    Json,
}

/// 数据源映射
#[derive(Debug, Deserialize)]
pub struct DataMapping {
    /// 变量名，用于模板中引用
    pub tag: String,
    /// Excel 数据文件路径（与 json 互斥）
    pub file: Option<PathBuf>,
    /// JSON 数据文件路径（与 file 互斥）
    pub json: Option<PathBuf>,
}

impl DataMapping {
    /// 获取数据源类型
    pub fn source_type(&self) -> Option<SourceType> {
        if self.file.is_some() {
            Some(SourceType::Excel)
        } else if self.json.is_some() {
            Some(SourceType::Json)
        } else {
            None
        }
    }

    /// 获取文件路径
    pub fn path(&self) -> Option<&PathBuf> {
        self.file.as_ref().or(self.json.as_ref())
    }

    /// 验证数据映射有效性
    pub fn validate(&self) -> Result<()> {
        // 检查是否同时指定了 file 和 json（互斥）
        if self.file.is_some() && self.json.is_some() {
            return Err(ReportError::MutuallyExclusiveData { tag: self.tag.clone() });
        }

        // 检查是否都没有指定
        if self.file.is_none() && self.json.is_none() {
            return Err(ReportError::MissingDataPath { tag: self.tag.clone() });
        }

        Ok(())
    }
}

/// Tab 配置项
#[derive(Debug, Deserialize)]
pub struct TabConfig {
    /// 输出 sheet 名称，"*" 表示复制模板所有 sheet
    pub tab: String,
    /// 模板文件路径
    pub template: PathBuf,
    /// 数据源映射（tab="*" 时可选）
    pub data: Option<Vec<DataMapping>>,
}

/// 配置文件结构
#[derive(Debug, Deserialize)]
pub struct Config {
    /// Tab 配置列表
    tabs: Vec<TabConfig>,
}

impl Config {
    /// 从文件加载配置
    pub fn load(path: &Path) -> Result<Self> {
        if !path.exists() {
            return Err(ReportError::ConfigNotFound(path.to_path_buf()));
        }

        let content = std::fs::read_to_string(path)
            .map_err(|e| ReportError::ConfigParse(e.to_string()))?;

        // JSON 配置文件是一个数组，直接解析为 Vec<TabConfig>
        let tabs: Vec<TabConfig> = serde_json::from_str(&content)
            .map_err(|e| ReportError::ConfigParse(format!("{}: {}", path.display(), e)))?;

        // 验证每个数据映射
        for tab_config in &tabs {
            if let Some(data) = &tab_config.data {
                for mapping in data {
                    mapping.validate()?;
                }
            }
        }

        Ok(Config { tabs })
    }

    /// 获取所有 tab 配置
    pub fn tabs(&self) -> &[TabConfig] {
        &self.tabs
    }

    /// 获取 tab 数量
    pub fn len(&self) -> usize {
        self.tabs.len()
    }

    /// 是否为空
    pub fn is_empty(&self) -> bool {
        self.tabs.is_empty()
    }
}

impl IntoIterator for Config {
    type Item = TabConfig;
    type IntoIter = std::vec::IntoIter<TabConfig>;

    fn into_iter(self) -> Self::IntoIter {
        self.tabs.into_iter()
    }
}

impl<'a> IntoIterator for &'a Config {
    type Item = &'a TabConfig;
    type IntoIter = std::slice::Iter<'a, TabConfig>;

    fn into_iter(self) -> Self::IntoIter {
        self.tabs.iter()
    }
}