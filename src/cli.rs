use clap::Parser;
use std::path::PathBuf;

/// 报告生成工具 - 根据配置文件和模板生成报告
///
/// 该工具读取 JSON 配置文件，加载 Excel/JSON 数据源，
/// 使用模板文件渲染生成最终的报告文件。
///
/// 支持的输出类型：
///   - xlsx: Excel 报告
///   - docx: Word 报告（尚未实现）
///
/// 模板语法详见 UserHandbook.md
#[derive(Parser, Debug)]
#[command(name = "report_gen")]
#[command(version = "0.1.0")]
#[command(about = "报告生成工具")]
pub struct Cli {
    /// 配置文件路径 (JSON 格式)
    ///
    /// 配置文件定义了每个 Sheet 的渲染规则，包括：
    /// - tab: Sheet 名称（"*" 表示复制模板所有 Sheet）
    /// - template: 模板文件路径
    /// - data: 数据源映射数组
    #[arg(short = 'c', long = "config", value_name = "FILE")]
    pub config: PathBuf,

    /// 输出文件类型
    ///
    /// 目前支持：
    ///   xlsx - Excel 报告
    ///   docx - Word 报告（尚未实现）
    #[arg(short = 't', long = "type", value_name = "TYPE", default_value = "xlsx")]
    pub r#type: String,

    /// 输出文件路径
    ///
    /// 生成的报告文件保存位置
    #[arg(short = 'o', long = "out", value_name = "FILE")]
    pub out: PathBuf,
}

/// 打印使用帮助（包含示例）
pub fn print_help() {
    println!("report_gen 0.1.0 - 报告生成工具");
    println!();
    println!("用法:");
    println!("  report_gen -c <CONFIG> -t <TYPE> -o <OUTPUT>");
    println!("  report_gen --config <CONFIG> --type <TYPE> --out <OUTPUT>");
    println!();
    println!("参数:");
    println!("  -c, --config <FILE>    配置文件路径 (JSON 格式，必需)");
    println!("  -t, --type <TYPE>      输出类型: xlsx/docx (默认: xlsx)");
    println!("  -o, --out <FILE>       输出文件路径 (必需)");
    println!("  -h, --help             显示帮助信息");
    println!("  -V, --version          显示版本信息");
    println!();
    println!("示例:");
    println!("  report_gen -c config/config.json -t xlsx -o output/result.xlsx");
    println!("  report_gen --config example/config.json --out output/report.xlsx");
    println!();
    println!("更多帮助请查看 UserHandbook.md");
}