//! 创建 Word 模板示例文件

use rdocx::Document;

fn main() {
    // 创建新文档
    let mut doc = Document::new();

    // 添加段落（单值替换）
    doc.add_paragraph("报告日期: {{info.date}}");
    doc.add_paragraph("制作人: {{info.name}}");

    // 添加聚合函数段落
    doc.add_paragraph("总授信额度: {{sum(limits.limits):currency:¥}}");
    doc.add_paragraph("客户数量: {{count(limits)}}");

    // 创建表格（4列2行）
    let mut table = doc.add_table(2, 4);

    // 设置表头
    if let Some(mut cell) = table.cell(0, 0) {
        cell.set_text("客户");
    }
    if let Some(mut cell) = table.cell(0, 1) {
        cell.set_text("授信额度");
    }
    if let Some(mut cell) = table.cell(0, 2) {
        cell.set_text("使用量");
    }
    if let Some(mut cell) = table.cell(0, 3) {
        cell.set_text("到期时间");
    }

    // 设置模板行（行迭代标记）
    if let Some(mut cell) = table.cell(1, 0) {
        cell.set_text("{{@limits.client}}");
    }
    if let Some(mut cell) = table.cell(1, 1) {
        cell.set_text("{{@limits.limits:int}}");
    }
    if let Some(mut cell) = table.cell(1, 2) {
        cell.set_text("{{@limits.use:int}}");
    }
    if let Some(mut cell) = table.cell(1, 3) {
        cell.set_text("{{@limits.ttm}}");
    }

    // 保存
    doc.save("example/templates/template.docx").unwrap();
    println!("创建 Word 模板成功: example/templates/template.docx");
}