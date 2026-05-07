//! rdocx Document 扩展 - 提可变表格访问
//!
//! 使用 unsafe 指针转换操作表格内部结构

use rdocx::Document;
use rdocx_oxml::table::{CellContent, CT_Row, CT_Tbl, CT_Tc};
use rdocx_oxml::text::CT_P;

/// 扩展 Document 以支持表格行插入
pub trait DocumentExt {
    /// 在指定表格的指定位置插入行，复制模板行的样式
    unsafe fn insert_table_row(&mut self, table_idx: usize, row_idx: usize, template_row_idx: usize);

    /// 设置表格单元格文本（保留段落样式）
    unsafe fn set_table_cell_text(&mut self, table_idx: usize, row_idx: usize, cell_idx: usize, text: &str);

    /// 获取表格行数
    fn get_table_row_count(&self, table_idx: usize) -> usize;

    /// 获取表格单元格文本
    fn get_table_cell_text(&self, table_idx: usize, row_idx: usize, cell_idx: usize) -> String;

    /// 获取表格列数
    fn get_table_column_count(&self, table_idx: usize) -> usize;
}

impl DocumentExt for Document {
    unsafe fn insert_table_row(&mut self, table_idx: usize, row_idx: usize, template_row_idx: usize) {
        let tables = self.tables();

        if table_idx >= tables.len() {
            return;
        }

        let table_ref = &tables[table_idx];
        let table_ref_ptr = table_ref as *const rdocx::TableRef<'_> as *const *const CT_Tbl;
        let ct_tbl_ptr = *table_ref_ptr as *mut CT_Tbl;
        let ct_tbl = &mut *ct_tbl_ptr;

        if template_row_idx >= ct_tbl.rows.len() {
            return;
        }

        // 复制模板行
        let template_row = &ct_tbl.rows[template_row_idx];
        let mut new_row = CT_Row::new();

        // 复制行属性（高度等）
        new_row.properties = template_row.properties.clone();

        // 复制单元格结构，保留段落样式
        for template_cell in &template_row.cells {
            // 克隆整个单元格（包括段落样式）
            let mut new_cell = template_cell.clone();

            // 清空段落内容但保留段落样式
            for content in &mut new_cell.content {
                if let CellContent::Paragraph(p) = content {
                    // 保留段落属性（对齐方式等），清空文本内容
                    p.runs.clear();
                    p.hyperlinks.clear();
                }
            }

            new_row.cells.push(new_cell);
        }

        // 插入新行
        if row_idx <= ct_tbl.rows.len() {
            ct_tbl.rows.insert(row_idx, new_row);
        }
    }

    unsafe fn set_table_cell_text(&mut self, table_idx: usize, row_idx: usize, cell_idx: usize, text: &str) {
        let tables = self.tables();

        if table_idx >= tables.len() {
            return;
        }

        let table_ref = &tables[table_idx];
        let table_ref_ptr = table_ref as *const rdocx::TableRef<'_> as *const *const CT_Tbl;
        let ct_tbl_ptr = *table_ref_ptr as *mut CT_Tbl;
        let ct_tbl = &mut *ct_tbl_ptr;

        if row_idx >= ct_tbl.rows.len() || cell_idx >= ct_tbl.rows[row_idx].cells.len() {
            return;
        }

        let cell = &mut ct_tbl.rows[row_idx].cells[cell_idx];

        // 找到第一个段落并设置文本（保留段落属性如对齐方式）
        let first_para = cell.content.iter_mut().find_map(|c| {
            if let CellContent::Paragraph(p) = c {
                Some(p)
            } else {
                None
            }
        });

        if let Some(para) = first_para {
            para.runs.clear();
            if !text.is_empty() {
                para.add_run(text);
            }
        } else {
            // 如果没有段落，创建一个新的
            let mut p = CT_P::new();
            if !text.is_empty() {
                p.add_run(text);
            }
            cell.content.insert(0, CellContent::Paragraph(p));
        }
    }

    fn get_table_row_count(&self, table_idx: usize) -> usize {
        let tables = self.tables();
        if table_idx >= tables.len() { return 0; }
        tables[table_idx].row_count()
    }

    fn get_table_cell_text(&self, table_idx: usize, row_idx: usize, cell_idx: usize) -> String {
        let tables = self.tables();
        if table_idx >= tables.len() { return String::new(); }
        if let Some(row) = tables[table_idx].row(row_idx) {
            if let Some(cell) = row.cell(cell_idx) {
                cell.text()
            } else { String::new() }
        } else { String::new() }
    }

    fn get_table_column_count(&self, table_idx: usize) -> usize {
        let tables = self.tables();
        if table_idx >= tables.len() { return 0; }
        tables[table_idx].column_count()
    }
}