use crate::error::{ReportError, Result};
use crate::data::source::{DataRow, DataSource, SourceType};
use umya_spreadsheet::reader::xlsx;
use serde_json::Value;
use std::collections::HashMap;
use std::path::Path;

/// Excel 数据源
pub struct ExcelSource {
    /// 行数据
    rows: Vec<DataRow>,
}

impl ExcelSource {
    /// 从文件加载 Excel 数据
    pub fn load(path: &Path) -> Result<Self> {
        if !path.exists() {
            return Err(ReportError::DataFileNotFound(path.to_path_buf()));
        }

        let book = xlsx::read(path)
            .map_err(|e| ReportError::ExcelDataRead(format!("{}: {}", path.display(), e)))?;

        // 读取第一个 sheet
        let sheet = book.get_sheet(&0)
            .ok_or_else(|| ReportError::ExcelDataRead("Excel 文件无 sheet".to_string()))?;

        // 解析数据
        let rows = Self::parse_sheet(sheet);

        Ok(Self { rows })
    }

    /// 解析 Sheet
    fn parse_sheet(sheet: &umya_spreadsheet::Worksheet) -> Vec<DataRow> {
        let mut rows = Vec::new();

        // 获取 sheet 的行数和列数
        let max_row = sheet.get_highest_row() as usize;
        let max_col = sheet.get_highest_column() as usize;

        if max_row <= 1 {
            return rows;
        }

        // 第一行为列名（表头）
        let headers: Vec<String> = (1..=max_col)
            .map(|col| {
                sheet.get_value((col as u32, 1u32))
                    .to_string()
            })
            .collect();

        // 解析数据行（从第 2 行开始）
        for row_idx in 2..=max_row {
            let mut fields = HashMap::new();

            for (col_idx, header) in headers.iter().enumerate() {
                if header.is_empty() {
                    continue;
                }

                let col = (col_idx + 1) as u32;
                let value_str = sheet.get_value((col, row_idx as u32)).to_string();

                let value = if value_str.is_empty() {
                    Value::Null
                } else if let Ok(num) = value_str.parse::<f64>() {
                    Value::Number(serde_json::Number::from_f64(num)
                        .unwrap_or_else(|| serde_json::Number::from(0)))
                } else {
                    Value::String(value_str)
                };

                fields.insert(header.clone(), value);
            }

            rows.push(DataRow {
                index: row_idx - 2,
                fields,
            });
        }

        rows
    }
}

impl DataSource for ExcelSource {
    fn source_type(&self) -> SourceType {
        SourceType::Excel
    }

    fn get_value(&self, field: &str) -> Option<Value> {
        // 返回首行数据
        self.rows.first()?.fields.get(field).cloned()
    }

    fn get_rows(&self) -> Option<&[DataRow]> {
        Some(&self.rows)
    }

    fn row_count(&self) -> usize {
        self.rows.len()
    }
}