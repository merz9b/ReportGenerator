# Report Generator - 示例与使用指南

报告生成工具 - 根据 JSON 配置和模板文件自动渲染生成 Excel/Word 报告。

---

## 快速开始

### 1. 构建项目

```bash
# 克隆项目
git clone https://github.com/xxx/report_generator.git
cd report_generator

# 构建发布版本
cargo build --release
```

### 2. 运行 Excel 渲染

```bash
# Windows
.\target\release\report_generator.exe -c example/config/config.json -t xlsx -o example/output/result.xlsx

# Linux/Mac
./target/release/report_generator -c example/config/config.json -t xlsx -o example/output/result.xlsx
```

### 3. 运行 Word 渲染

```bash
# Windows
.\target\release\report_generator.exe -c example/config/config_docx.json -t docx -o example/output/output.docx

# Linux/Mac
./target/release/report_generator -c example/config/config_docx.json -t docx -o example/output/output.docx
```

---

## 依赖库

| 库 | 版本 | 用途 |
|---|------|------|
| clap | 4.x | CLI 参数解析 |
| serde / serde_json | 1.x | JSON 序列化 |
| umya-spreadsheet | 2.3.3 | Excel 读写 + 样式保留 |
| rdocx | 0.1 | Word 读写 + 文本替换 |
| rdocx-oxml | 0.1 | Word XML 结构操作 |
| thiserror | 1.x | 错误处理 |
| tracing | 0.1.x | 日志 |
| regex | 1.x | 公式行号替换 |
| chrono | 0.4 | 日期处理 |

---

## CLI 参数

| 参数 | 简写 | 说明 | 示例 |
|---|------|------|------|
| `--config` | `-c` | JSON 配置文件（必需） | `config/config.json` |
| `--type` | `-t` | 输出类型：xlsx/docx（默认 xlsx） | `docx` |
| `--out` | `-o` | 输出文件路径（必需） | `output/result.xlsx` |
| `--help` | `-h` | 显示帮助 | |
| `--version` | `-V` | 显示版本 | |

---

## 目录结构

```
example/
├── config/
│   ├── config.json          # Excel 配置
│   └── config_docx.json     # Word 配置
├── templates/
│   ├── template.xlsx        # Excel 模板
│   ├── template_1.xlsx      # Excel 复制模板
│   └── template.docx        # Word 模板
├── data/
│   ├── data_1.xlsx          # Excel 数据源
│   ├── data_2.xlsx          # Excel 数据源
│   └── data_1.json          # JSON 数据源
└── output/
    ├── result.xlsx          # Excel 输出（运行后生成）
    └── output.docx          # Word 输出（运行后生成）
```

---

## 配置文件格式

```json
[
    {
        "tab": "sheet_name",
        "template": "templates/template.xlsx",
        "data": [
            {"tag": "limits", "file": "data/data_1.xlsx"},
            {"tag": "meta", "json": "data/data_1.json"}
        ]
    },
    {
        "tab": "*",
        "template": "templates/template_1.xlsx"
    }
]
```

| 字段 | 说明 |
|---|------|
| `tab` | Sheet 名称；`"*"` 表示复制模板所有 Sheet |
| `template` | 模板文件路径 |
| `data` | 数据源数组，每个元素有 `tag` + `file`/`json` |

---

## 模板语法

### 基础标记

| 语法 | 说明 | 示例 |
|---|------|------|
| `{{tag.field}}` | 单值替换 | `{{meta.date}}` |
| `{{@tag.field}}` | 行迭代 | `{{@limits.client}}` |
| `{{sum(tag.field)}}` | 聚合求和 | `{{sum(limits.limits)}}` |
| `{{count(tag)}}` | 聚合计数 | `{{count(limits)}}` |

### 值运算

| 语法 | 说明 | 示例 |
|---|------|------|
| `{{tag.field/1000}}` | 除法 | `{{limits.limits/1000}}` |
| `{{tag.field*100}}` | 乘法 | `{{limits.limits*100}}` |
| `{{#(tag1.f1*tag2.f2)}}` | 变量间运算 | `{{#(limits.use/limits.limits):pct}}` |

### 格式化

| 格式码 | 说明 | 示例 |
|---|------|------|
| `:int` | 整数 | `{{v:int}}` |
| `:float:2` | 2位小数 | `{{v:float:2}}` |
| `:pct:2` | 百分比 | `{{v:pct:2}}` → 2.50% |
| `:currency:¥` | 人民币 | `{{v:currency:¥}}` |
| `:date` | 日期 | `{{d:date}}` |

### Excel 公式

| 语法 | 说明 | 示例 |
|---|------|------|
| `{{=A{r}*B{r}}}` | 单格公式 | `{{=C{r}/B{r}:pct}}` |
| `{{@=A{r}*B{r}}}` | 行迭代公式 | 每行插入公式 |
| `{r}` | 当前行号 | A5 行时 `{r}=5` |
| `{r+1}` | 当前行+1 | A5 行时 `{r+1}=6` |

---

## Excel 模板示例

**template.xlsx 布局：**

| 行 | A列 | B列 | C列 | D列 | 说明 |
|---|-----|-----|-----|-----|------|
| 1 | `{{meta.date:date}}` | `{{meta.name}}` | | | 单值 |
| 3 | 客户 | 限额 | 占用 | 到期 | 表头 |
| 4 | `{{@limits.client}}` | `{{@limits.limits:int}}` | `{{@limits.use}}` | `{{@limits.ttm}}` | 行迭代 |
| 8 | 合计 | `{{sum(limits.limits):int}}` | | | 聚合 |

---

## Word 模板示例

**template.docx 内容：**

```
报告日期: {{info.date}}
制作人: {{info.name}}

总授信额度: {{sum(limits.limits):float:2}}万元
客户数量: {{count(limits)}}家

[表格]
| 客户 | 授信额度 | 使用量 | 到期时间 |
| {{@limits.client}} | {{@limits.limits:int}} | {{@limits.use:int}} | {{@limits.ttm}} |
```

**Word 与 Excel 差异：**
- Word 不支持 Excel 公式 `{{=}}`
- 聚合函数输出为文本而非数值
- 行迭代自动扩展表格行并保留样式

---

## 数据文件格式

### Excel 数据源 (data_1.xlsx)

| client | limits | use | ttm |
|--------|--------|-----|-----|
| x1 | 2966 | 0 | 2023/12/14 |
| x2 | 2833 | 0 | 2023/12/20 |
| x3 | 2833 | 0 | 2023/12/20 |

### JSON 数据源 (data_1.json)

```json
{
    "date": "2026-04-23",
    "name": "maker",
    "version": "1.0.0"
}
```

---

## 预期输出

### Excel (result.xlsx)

- **test_1 sheet** - 渲染后的数据表格
  - 标题行：日期和名称
  - 授信占用表：3行数据 + 聚合
  - 指数区间表：4行数据
- **summary sheet** - 来自 template_1.xlsx
- **notes sheet** - 来自 template_1.xlsx

### Word (output.docx)

- 段落中的占位符替换为数据值
- 表格自动扩展为 3 行数据
- 保留模板样式（对齐方式等）

---


---

## 许可证

MIT License