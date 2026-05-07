use std::path::PathBuf;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum ReportError {
    #[error("配置文件解析失败: {0}")]
    ConfigParse(String),

    #[error("配置文件不存在: {0}")]
    ConfigNotFound(PathBuf),

    #[error("模板文件不存在: {0}")]
    TemplateNotFound(PathBuf),

    #[error("数据文件不存在: {0}")]
    DataFileNotFound(PathBuf),

    #[error("标记解析错误: {0}")]
    TokenParse(String),

    #[error("数据源 '{tag}' 不存在")]
    SourceNotFound { tag: String },

    #[error("数据源 '{tag}' 不支持行迭代，需要 Excel 表格数据")]
    SourceNotIteratable { tag: String },

    #[error("字段不存在: {tag}.{field}")]
    FieldNotFound { tag: String, field: String },

    #[error("不支持的输出类型: {0}，仅支持 xlsx/docx")]
    UnsupportedType(String),

    #[error("Excel 模板读取失败: {0}")]
    ExcelTemplateRead(String),

    #[error("Excel 渲染失败: {0}")]
    ExcelRender(String),

    #[error("Excel 输出失败: {0}")]
    ExcelOutput(String),

    #[error("Excel 数据文件读取失败: {0}")]
    ExcelDataRead(String),

    #[error("JSON 数据文件读取失败: {0}")]
    JsonDataRead(String),

    #[error("数据映射缺少文件路径: tag '{tag}'")]
    MissingDataPath { tag: String },

    #[error("数据映射 '{tag}' 同时指定了 file 和 json，两者互斥")]
    MutuallyExclusiveData { tag: String },

    #[error("IO 错误: {0}")]
    IoError(String),

    #[error("XML 解析错误: {0}")]
    XmlParseError(String),

    #[error("Word 模板读取失败: {0}")]
    WordTemplateRead(String),

    #[error("Word 渲染失败: {0}")]
    WordRender(String),

    #[error("Word 输出失败: {0}")]
    WordOutput(String),
}

pub type Result<T> = std::result::Result<T, ReportError>;