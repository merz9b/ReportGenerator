//! rdocx Table 扩展 - 提供 insert_row 功能
//!
//! 由于 rdocx 的 Table.inner 是 pub(crate)，我们使用 unsafe 指针转换来访问底层 CT_Tbl

use rdocx::Table;
use rdocx_oxml::table::{CT_Row, CT_Tc};

/// 为 Table 实现扩展方法
pub trait TableExt {
    /// 在指定位置插入一行，复制模板行的结构
    fn insert_row(&mut self, index: usize, template_row_idx: usize) -> usize;

    /// 获取行数
    fn row_count_ext(&self) -> usize;
}

impl<'a> TableExt for Table<'a> {
    fn insert_row(&mut self, index: usize, template_row_idx: usize) -> usize {
        // Table 的布局: { inner: &'a mut CT_Tbl }
        // 我们通过指针转换来访问它
        let table_ptr = self as *mut Table<'a> as *mut *mut rdocx_oxml::table::CT_Tbl;

        unsafe {
            let ct_tbl = *table_ptr;
            let ct_tbl_ref = &mut *ct_tbl;

            // 确保模板行存在
            if template_row_idx >= ct_tbl_ref.rows.len() {
                return ct_tbl_ref.rows.len();
            }

            // 获取模板行并克隆其结构
            let template_row = &ct_tbl_ref.rows[template_row_idx];

            // 创建新行：复制模板行的单元格数量和属性
            let mut new_row = CT_Row::new();
            new_row.properties = template_row.properties.clone();

            // 复制单元格结构（保留属性如宽度、样式）
            for template_cell in &template_row.cells {
                let mut new_cell = CT_Tc::new();
                new_cell.properties = template_cell.properties.clone();
                new_row.cells.push(new_cell);
            }

            // 在指定位置插入新行
            if index <= ct_tbl_ref.rows.len() {
                ct_tbl_ref.rows.insert(index, new_row);
            } else {
                ct_tbl_ref.rows.push(new_row);
            }

            ct_tbl_ref.rows.len() - 1
        }
    }

    fn row_count_ext(&self) -> usize {
        unsafe {
            let table_ptr = self as *const Table<'a> as *const *const rdocx_oxml::table::CT_Tbl;
            let ct_tbl = *table_ptr;
            (*ct_tbl).rows.len()
        }
    }
}