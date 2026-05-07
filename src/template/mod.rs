//! 模板解析模块
//!
//! 提供模板标记的词法分析、语法分析和格式化功能。
//!
//! ## 架构
//!
//! 采用 AST (抽象语法树) 架构：
//!
//! ```
//! 模板字符串 → Lexer(词法分析) → LexToken流 → Parser(语法分析) → AST节点
//! ```
//!
//! ## 模块结构
//!
//! - `lexer`: 词法分析器，将字符串转换为词法单元
//! - `ast`: AST节点定义，包含表达式、修饰符、格式等
//! - `parser`: 语法分析器，构建AST节点（支持运算符优先级）
//! - `formatter`: 格式化器，计算表达式值并格式化
//! - `value`: 单元格值类型定义

pub mod ast;
pub mod lexer;
pub mod parser;
pub mod formatter;
pub mod value;

// 公共导出
pub use ast::{AggFunc, AggTarget, AstNode, Expression, FieldRef, FormatSpec, FormatType, Modifier, ParsedCell};
pub use formatter::{compute_aggregate, format_expression, DataContext, SimpleDataContext};
pub use lexer::{LexToken, Lexer, Operator};
pub use parser::{Parser, replace_row_placeholders};
pub use value::{CellValue, FormattedCell};