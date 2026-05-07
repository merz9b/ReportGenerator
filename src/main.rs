mod cli;
mod config;
mod data;
mod engine;
mod error;
mod template;

use clap::Parser;
use cli::Cli;
use config::{Config, TabConfig};
use data::DataContext;
use engine::{ExcelRenderer, WordRenderer};
use error::Result;
use umya_spreadsheet::{new_file, Spreadsheet, writer::xlsx as xlsx_writer};
use std::path::Path;
use tracing_subscriber::{layer::SubscriberExt, util::SubscriberInitExt};

fn main() {
    // 解析 CLI 参数
    let cli = match Cli::try_parse() {
        Ok(cli) => cli,
        Err(e) => {
            // 如果是帮助请求
            if e.kind() == clap::error::ErrorKind::DisplayHelp {
                cli::print_help();
                std::process::exit(0);
            }
            // 如果是版本请求
            if e.kind() == clap::error::ErrorKind::DisplayVersion {
                println!("report_gen 0.1.0");
                std::process::exit(0);
            }
            // 其他错误，显示友好提示
            eprintln!("错误: {}", e);
            eprintln!();
            eprintln!("用法: report_gen -c <CONFIG> -t <TYPE> -o <OUTPUT>");
            eprintln!("运行 report_gen --help 查看更多帮助");
            std::process::exit(1);
        }
    };

    // 初始化日志（只在解析成功后）
    tracing_subscriber::registry()
        .with(tracing_subscriber::fmt::layer())
        .with(tracing_subscriber::EnvFilter::from_default_env()
            .add_directive(tracing::Level::INFO.into()))
        .init();

    tracing::info!("开始处理: config={}, type={}, out={}",
        cli.config.display(), cli.r#type, cli.out.display());

    // 执行主逻辑
    if let Err(e) = run(cli) {
        tracing::error!("处理失败: {}", e);
        eprintln!("错误: {}", e);
        std::process::exit(1);
    }
}

/// 主逻辑
fn run(cli: Cli) -> Result<()> {

    // 加载配置文件
    let config = Config::load(&cli.config)?;
    tracing::info!("配置文件加载成功，共 {} 个 tab 配置", config.len());

    // 根据输出类型处理
    match cli.r#type.as_str() {
        "xlsx" => process_excel(&config, &cli.out)?,
        "docx" => process_docx(&config, &cli.out)?,
        other => return Err(error::ReportError::UnsupportedType(other.to_string())),
    }

    tracing::info!("报告生成完成: {}", cli.out.display());
    Ok(())
}

/// 处理 Excel 输出
fn process_excel(config: &Config, output: &Path) -> Result<()> {
    tracing::info!("开始处理 Excel 输出");

    // 创建空的 workbook
    let mut target_book = new_file();

    // 按 config 顺序处理每个 tab
    for (idx, tab_config) in config.tabs().iter().enumerate() {
        if tab_config.tab == "*" {
            // tab="*": 复制模板的所有 sheets（不渲染）
            process_star_tab(&mut target_book, tab_config, idx)?;
        } else {
            // 普通 tab: 渲染后添加
            process_render_tab(&mut target_book, tab_config, idx)?;
        }
    }

    // 删除可能的默认空 sheet（Sheet1）
    let sheet_count = target_book.get_sheet_count();
    if sheet_count > config.tabs().len() {
        // 检查第一个 sheet 是否是默认空 sheet
        if let Some(sheet) = target_book.get_sheet(&0) {
            let sheet_name = sheet.get_name();
            if sheet_name == "Sheet1" && sheet.get_highest_row() <= 1 {
                let _ = target_book.remove_sheet(0);
            }
        }
    }

    // 设置默认激活第一个 sheet
    target_book.set_active_sheet(0);
    if let Some(sheet) = target_book.get_sheet_mut(&0) {
        sheet.set_active_cell("A1");
    }

    // 保存
    xlsx_writer::write(&target_book, output)
        .map_err(|e| error::ReportError::ExcelOutput(format!("保存失败: {}", e)))?;

    tracing::info!("保存文件: {}", output.display());
    Ok(())
}

/// 处理 tab="*" - 复制模板的所有 sheets
fn process_star_tab(target_book: &mut Spreadsheet, tab_config: &TabConfig, idx: usize) -> Result<()> {
    tracing::info!("处理 tab='*' (配置 {}): 模板 {}", idx, tab_config.template.display());

    // 从模板加载 workbook
    let template_book = ExcelRenderer::load_template(&tab_config.template)?;

    // 复制所有 sheets
    let sheet_count = template_book.get_sheet_count();
    for i in 0..sheet_count {
        let template_sheet = template_book.get_sheet(&i)
            .ok_or_else(|| error::ReportError::ExcelTemplateRead("模板 sheet 获取失败".to_string()))?;

        let cloned_sheet = template_sheet.clone();
        let _ = target_book.add_sheet(cloned_sheet);

        tracing::info!("复制 sheet: {}", template_sheet.get_name());
    }

    tracing::info!("tab='*' 处理完成，共复制 {} 个 sheets", sheet_count);
    Ok(())
}

/// 处理渲染 tab
fn process_render_tab(target_book: &mut Spreadsheet, tab_config: &TabConfig, idx: usize) -> Result<()> {
    let name = tab_config.tab.as_str();
    tracing::info!("处理 tab='{}' (配置 {}): 模板 {}", name, idx, tab_config.template.display());

    // 加载数据源
    let mut context = DataContext::new();

    if let Some(data) = &tab_config.data {
        for mapping in data {
            if mapping.validate().is_err() {
                tracing::warn!("数据映射验证失败: {}", mapping.tag);
                continue;
            }

            let path = mapping.path().unwrap();
            let source_type = mapping.source_type().unwrap();

            match source_type {
                config::SourceType::Excel => {
                    tracing::info!("加载 Excel 数据源: tag={}, file={}", mapping.tag, path.display());
                    context.load_excel(&mapping.tag, path)?;
                }
                config::SourceType::Json => {
                    tracing::info!("加载 JSON 数据源: tag={}, file={}", mapping.tag, path.display());
                    context.load_json(&mapping.tag, path)?;
                }
            }
        }
    }

    // 从模板加载 workbook
    let template_book = ExcelRenderer::load_template(&tab_config.template)?;

    // 克隆模板的第一个 sheet 到目标 workbook
    let template_sheet = template_book.get_sheet(&0)
        .ok_or_else(|| error::ReportError::ExcelTemplateRead("模板无 sheet".to_string()))?;

    let mut cloned_sheet = template_sheet.clone();
    cloned_sheet.set_name(name);

    // 添加到目标 workbook
    let _ = target_book.add_sheet(cloned_sheet);

    // 获取新添加的 sheet 进行渲染
    let sheet_index = target_book.get_sheet_count() - 1;
    let sheet = target_book.get_sheet_mut(&sheet_index)
        .ok_or_else(|| error::ReportError::ExcelOutput("无法获取 sheet".to_string()))?;

    // 渲染
    tracing::info!("渲染 sheet: {}", name);
    let renderer = ExcelRenderer::new(&context);
    renderer.render_sheet(sheet)?;

    tracing::info!("Sheet '{}' 处理完成", name);
    Ok(())
}

/// 处理 Word 输出
fn process_docx(config: &Config, output: &Path) -> Result<()> {
    tracing::info!("开始处理 Word 输出");

    // Word 目前只支持单个模板文件
    if config.tabs().len() > 1 {
        tracing::warn!("Word 输出仅支持单个模板，将只处理第一个配置");
    }

    let tab_config = config.tabs().first()
        .ok_or_else(|| error::ReportError::ConfigParse("配置文件无有效 tab".to_string()))?;

    // tab="*" 不支持 Word
    if tab_config.tab == "*" {
        return Err(error::ReportError::WordRender("Word 不支持 tab='*' 配置".to_string()));
    }

    tracing::info!("处理模板: {}", tab_config.template.display());

    // 加载数据源
    let mut context = DataContext::new();

    if let Some(data) = &tab_config.data {
        for mapping in data {
            if mapping.validate().is_err() {
                tracing::warn!("数据映射验证失败: {}", mapping.tag);
                continue;
            }

            let path = mapping.path().unwrap();
            let source_type = mapping.source_type().unwrap();

            match source_type {
                config::SourceType::Excel => {
                    tracing::info!("加载 Excel 数据源: tag={}, file={}", mapping.tag, path.display());
                    context.load_excel(&mapping.tag, path)?;
                }
                config::SourceType::Json => {
                    tracing::info!("加载 JSON 数据源: tag={}, file={}", mapping.tag, path.display());
                    context.load_json(&mapping.tag, path)?;
                }
            }
        }
    }

    // 渲染
    tracing::info!("渲染 Word 文档");
    let renderer = WordRenderer::new(&context);
    renderer.render_file(&tab_config.template, output)?;

    tracing::info!("保存文件: {}", output.display());
    Ok(())
}