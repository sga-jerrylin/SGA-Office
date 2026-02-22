# SGA-Office 重构报告

> **报告日期**: 2026-02-22  
> **执行者**: Augment Agent  
> **仓库**: https://github.com/sga-jerrylin/SGA-Office  
> **最终提交**: `b686fb4` — 44 files changed, +4915 lines

---

## 一、背景与问题诊断

项目由前一版本（Gemini 编写）移交。经过全面审计，发现以下严重问题：

| 问题类型 | 具体描述 |
|---|---|
| **VIS 模块完全是假实现** | `chart_renderer.py` 内所有图表函数均返回硬编码假数据，从未真正生成图片 |
| **PDF-01 无法使用** | 原实现用 PIL 截图 Word 窗口，在无界面服务器上完全不可用 |
| **PDF-04 伪 OCR** | `extract_text_from_scanned_pdf()` 只调用了普通文字提取，并非真正 OCR |
| **Celery/Redis 过度设计** | 项目引入了完整的异步任务队列，但所有任务都是毫秒级操作，完全没必要 |
| **DRY 违反** | COS 上传/下载逻辑在多处重复，URL 校验逻辑散布在各 route 文件中 |
| **URL 校验 Bug** | 腾讯云 COS 签名 URL 含 `?sign=xxx` 参数，原校验逻辑直接检查整个 URL 末尾，导致 `.docx?sign=...` 被误判为非法 |
| **Schema 字段命名不一致** | `AddWatermarkResult` 用 `pdf_url`，其他模块用 `file_url`，AI Agent 调用时易出错 |
| **app/ 目录从未提交** | 原开发者只提交了配置文件，整个源代码从未进入 git 版本管理 |

---

## 二、重构内容（P0 ~ P5）

### P0 — 基础设施清理

**删除的文件**：
- `app/api/endpoints/task_routes.py` — Celery 任务路由
- `app/core/celery_app.py` — Celery 应用实例
- `app/worker/background_tasks.py` — 后台任务定义
- `app/worker/__init__.py`
- `app/services/chart_renderer.py` — 假实现的图表渲染器

**修改**：
- `docker-compose.yml` — 删除 Redis、Celery Worker 服务
- `app/main.py` — 移除 Celery 相关导入和任务路由注册

### P1 — Bug 修复与代码质量

- **修复 URL 校验 Bug**：`FillTemplateRequest.validate_template_url()` 改用 `urlparse(v).path` 提取路径部分再检查后缀，使带签名参数的 COS URL 能正常通过验证
- **统一 COS 操作**：提取 `get_cos_service()` 工厂函数，所有 route 通过统一入口获取 COS 客户端
- **统一响应字段**：`AddWatermarkResult`、`MergeSplitResult` 的 `pdf_url` 改为 `file_url`，与其他模块保持一致

### P2 — VIS 模块重建（全新真实实现）

新建 `app/services/vis_renderer.py`（410 行），包含 5 个真实实现：

| 函数 | 实现方式 | 说明 |
|---|---|---|
| `render_mermaid_to_image()` | 调用 [mermaid.ink](https://mermaid.ink) 公共渲染 API | AI 模型训练数据里有大量 Mermaid，天然支持 |
| `render_chart_from_data()` | matplotlib 本地渲染 | 支持 bar/line/pie/scatter/radar/heatmap/funnel/gauge |
| `generate_qrcode()` | qrcode 库 | 支持 L/M/Q/H 四档纠错级别 |
| `generate_barcode()` | python-barcode 库 | 支持 Code128/Code39/EAN13/EAN8/ISBN/UPC |
| `generate_wordcloud()` | wordcloud + jieba | 中文自动分词，支持多种色彩方案 |

### P3 — PDF-01 重写

删除原 PIL 截图方式，改用 **LibreOffice headless** 进行高保真 Word→PDF 转换：

```
convert_docx_to_pdf():
  1. 从 COS 下载 .docx 到临时目录
  2. 调用 soffice --headless --convert-to pdf
  3. 读取生成的 .pdf 上传 COS
  4. 清理临时文件
```

### P4 — 依赖与基础设施更新

**`requirements.txt`** 重写（移除 Celery/Redis，新增 VIS 依赖）：
```
新增: matplotlib, qrcode[pil], python-barcode, wordcloud, jieba
移除: celery, redis, flower
```

**`dockerfile`** 新增 LibreOffice 安装步骤：
```dockerfile
RUN apt-get install -y libreoffice --no-install-recommends
```

### P5 — 测试套件 + Git 初始化提交

- 编写 54 个 pytest 测试用例（详见第三节）
- 首次将完整 `app/` 源代码提交到 git 版本管理
- 更新 `.gitignore` 增加 `.pytest_cache/`、`instance/`

---

## 三、测试报告

### 测试环境

```
Python: 3.13.3
pytest: 9.0.2
平台: Windows (win32)
运行命令: python -m pytest tests/ -v --tb=short
```

### 最终结果

```
=== 54 passed, 8 warnings in 11.78s ===
```

**54/54 全部通过，0 失败，0 错误** ✅

### 测试文件分布

| 测试文件 | 测试数量 | 覆盖内容 |
|---|---|---|
| `tests/test_schemas.py` | 20 | Pydantic Schema 校验规则 |
| `tests/test_services.py` | 14 | Service 层纯业务逻辑 |
| `tests/test_routes.py` | 10 | API 端点集成测试（DOC/EXC/VIS/System） |
| `tests/test_vis_pdf_routes.py` | 10 | VIS 和 PDF 端点集成测试 |
| **合计** | **54** | |

### test_schemas.py — Schema 校验层（20 个）

| 测试用例 | 验证内容 | 结果 |
|---|---|---|
| `test_valid_request` | 正常 Markdown 请求通过 | ✅ |
| `test_blank_content_rejected` | 纯空白内容被 422 拒绝 | ✅ |
| `test_empty_string_rejected` | 空字符串被拒绝 | ✅ |
| `test_filename_sanitized` | 文件名非法字符自动清除 | ✅ |
| `test_valid_docx_url` | 合法 .docx URL 通过 | ✅ |
| `test_non_docx_url_rejected` | .pdf URL 被拒绝 | ✅ |
| `test_signed_cos_url_accepted` | 带签名参数的 COS URL 通过（P1 修复验证）| ✅ |
| `test_empty_variables_rejected` | 空 variables 字典被拒绝 | ✅ |
| `test_valid_excel_request` | 合法 Excel 请求通过 | ✅ |
| `test_non_string_header_rejected` | 非字符串表头被拒绝 | ✅ |
| `test_single_row_rejected` | 只有表头无数据行被拒绝 | ✅ |
| `test_valid_mermaid` | 合法 Mermaid 代码通过 | ✅ |
| `test_too_short_code_rejected` | 过短 Mermaid 代码被拒绝 | ✅ |
| `test_width_range` | 图片宽度超出范围被拒绝 | ✅ |
| `test_wordcloud_valid_request` | 合法词云请求通过 | ✅ |
| `test_too_short_text_rejected` | 文本太短被拒绝 | ✅ |
| `test_watermark_valid_config` | 合法水印配置通过 | ✅ |
| `test_invalid_color_rejected` | 非法 HEX 颜色被拒绝 | ✅ |
| `test_valid_page_range` | 合法页码区间通过 | ✅ |
| `test_zero_page_rejected` | 页码为 0 被拒绝（ge=1）| ✅ |

### test_services.py — Service 层单元测试（14 个）

| 测试类 | 测试用例 | 验证内容 | 结果 |
|---|---|---|---|
| `TestRenderMarkdownToDocx` | `test_basic_heading_and_paragraph` | 标题和正文正确渲染 | ✅ |
| | `test_markdown_with_table` | Markdown 表格转为 Word 表格 | ✅ |
| | `test_markdown_with_bold_and_italic` | 加粗/斜体格式保留 | ✅ |
| | `test_escaped_newlines_converted` | `\n` 转义符正确处理 | ✅ |
| | `test_empty_result_is_valid_docx` | 输出是合法 ZIP（docx 格式）| ✅ |
| `TestFillDocxTemplate` | `test_simple_placeholder_replacement` | `{{name}}` 占位符正确替换 | ✅ |
| | `test_unmatched_placeholders_remain` | 未匹配的占位符保持原样 | ✅ |
| `TestCreateExcelFromArray` | `test_basic_creation` | 输出是合法 ZIP（xlsx 格式）| ✅ |
| | `test_custom_sheet_name` | 自定义 Sheet 名称生效 | ✅ |
| `TestGenerateComplexExcel` | `test_multi_sheet_creation` | 多 Sheet 报表正确创建 | ✅ |
| | `test_with_formula` | Excel 公式字符串被保留 | ✅ |
| `TestHexToRgb` | `test_black` | `#000000` → `(0,0,0)` | ✅ |
| | `test_white` | `#FFFFFF` → `(1,1,1)` | ✅ |
| | `test_no_hash` | 不带 `#` 的 hex 也能解析 | ✅ |

### test_routes.py — 端点集成测试（10 个）

| 测试用例 | HTTP 方法 & 路径 | 验证内容 | 结果 |
|---|---|---|---|
| `test_root` | GET `/` | 服务名称包含 "SGA-Office" | ✅ |
| `test_health` | GET `/health` | status == "healthy" | ✅ |
| `test_doc01_render_markdown` | POST `/api/v1/docx/render_markdown` | 返回 file_url + filename | ✅ |
| `test_doc01_empty_content_rejected` | POST `/api/v1/docx/render_markdown` | 空白内容返回 422 | ✅ |
| `test_doc02_fill_template` | POST `/api/v1/docx/fill_template` | 模板变量替换完成 | ✅ |
| `test_exc01_create_from_array` | POST `/api/v1/excel/create_from_array` | 返回 file_url | ✅ |
| `test_exc01_invalid_headers` | POST `/api/v1/excel/create_from_array` | 非字符串表头返回 422 | ✅ |
| `test_exc03_generate_complex` | POST `/api/v1/excel/generate_complex` | 复杂报表生成成功 | ✅ |
| `test_vis01_render_mermaid` | POST `/api/v1/vis/render_mermaid` | 返回 .png 文件 URL | ✅ |
| `test_vis02_render_chart` | POST `/api/v1/vis/render_chart` | code == 200 | ✅ |

### test_vis_pdf_routes.py — VIS/PDF 端点补充测试（10 个）

| 测试用例 | HTTP 方法 & 路径 | 验证内容 | 结果 |
|---|---|---|---|
| `test_vis03a_generate_qrcode` | POST `/api/v1/vis/generate_qrcode` | QR 码生成成功 | ✅ |
| `test_vis03b_generate_barcode` | POST `/api/v1/vis/generate_barcode` | 条形码生成成功 | ✅ |
| `test_vis04_generate_wordcloud` | POST `/api/v1/vis/generate_wordcloud` | 词云生成成功 | ✅ |
| `test_vis01_invalid_code_too_short` | POST `/api/v1/vis/render_mermaid` | 短代码返回 422 | ✅ |
| `test_vis02_invalid_chart_type` | POST `/api/v1/vis/render_chart` | 非法图表类型返回 422 | ✅ |
| `test_pdf01_convert_from_docx` | POST `/api/v1/pdf/convert_from_docx` | 返回 .pdf 文件 URL | ✅ |
| `test_pdf02_add_watermark` | POST `/api/v1/pdf/add_watermark` | code == 200 | ✅ |
| `test_pdf03_merge` | POST `/api/v1/pdf/merge_split` | page_count 正确返回 | ✅ |
| `test_pdf03_split` | POST `/api/v1/pdf/merge_split` | 页码截取成功 | ✅ |
| `test_pdf02_invalid_color_rejected` | POST `/api/v1/pdf/add_watermark` | 非法颜色返回 422 | ✅ |

### 测试中发现并修复的问题

| 问题 | 原因 | 解决方案 |
|---|---|---|
| `test_services.py` 占位符测试失败 | 测试用 `{{ name }}` 有空格，实现用 `{{name}}` 无空格 | 统一测试格式为无空格 |
| `test_vis_pdf_routes.py` PDF mock 不生效 | mock 路径指向 `app.services.pdf_manipulator.*`，但 route 已 import 到局部 | 改为 mock `app.api.endpoints.pdf_routes.*` |
| pytest 采集报错 `AttributeError: 'Package' object has no attribute 'obj'` | `tests/__init__.py` 存在导致 pytest 将 tests 视为 Package | 删除 `tests/__init__.py`，同时卸载不需要的 `pytest-asyncio` |

---

## 四、当前 API 接口文档

**Base URL**: `http://<host>:<port>`  
**API 前缀**: `/api/v1`  
**统一响应格式**:
```json
{
  "code": 200,
  "message": "操作成功",
  "data": { ... }
}
```

---

### 系统接口

#### `GET /`
服务根路由，返回服务名称和版本。

**响应示例**:
```json
{
  "code": 200,
  "message": "SGA-Office API 运行中",
  "data": { "service": "SGA-Office: Agent-First 办公微服务平台", "version": "v1", "docs_url": "/docs" }
}
```

#### `GET /health`
健康检查。

**响应示例**:
```json
{
  "code": 200,
  "message": "healthy",
  "data": { "status": "healthy", "timestamp": "2026-02-22T10:00:00" }
}
```

---

### Word 模块 (DOC)

#### `POST /api/v1/docx/render_markdown` — [DOC-01]
**功能**: 将 Markdown 文本渲染为正式 Word 文档 (.docx)。

**请求体**:
```json
{
  "markdown_content": "# 年终总结\n\n本年度完成了...",
  "filename": "年终总结报告"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `markdown_content` | string | ✅ | Markdown 原文，支持标题/列表/表格/图片/代码块 |
| `filename` | string | ❌ | 输出文件名（不含扩展名），非法字符自动清除 |

**响应**:
```json
{
  "code": 200,
  "message": "Word 文档生成成功",
  "data": {
    "file_url": "https://cos.xxx.com/documents/年终总结报告_20260222_a1b2.docx",
    "filename": "年终总结报告_20260222_a1b2.docx"
  }
}
```

---

#### `POST /api/v1/docx/fill_template` — [DOC-02]
**功能**: 下载 Word 模板，将 `{{变量名}}` 占位符替换为实际值后上传。

**请求体**:
```json
{
  "template_url": "https://cos.xxx.com/templates/录用通知书.docx",
  "variables": {
    "姓名": "张三",
    "部门": "技术部",
    "入职日期": "2026-03-01",
    "薪资": "20000"
  },
  "filename": "录用通知书_张三"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `template_url` | string (URL) | ✅ | 模板 .docx 文件的可下载链接，支持带签名参数的 COS URL |
| `variables` | object | ✅ | Key-Value 变量字典，Key 对应模板中 `{{Key}}` 占位符 |
| `filename` | string | ❌ | 输出文件名 |

**占位符格式**: 模板中使用 `{{变量名}}`（无空格）

**响应**:
```json
{
  "code": 200,
  "message": "模板填充成功",
  "data": {
    "file_url": "https://cos.xxx.com/documents/录用通知书_张三_20260222_a1b2.docx",
    "filename": "录用通知书_张三_20260222_a1b2.docx"
  }
}
```

---

### Excel 模块 (EXC)

#### `POST /api/v1/excel/create_from_array` — [EXC-01]
**功能**: 将二维数组 JSON 创建为 Excel 文件。

**请求体**:
```json
{
  "title": "2026年Q1销售数据",
  "data": [
    ["区域", "产品", "销售额", "增长率"],
    ["华东", "产品A", 320000, "12%"],
    ["华南", "产品B", 280000, "8%"]
  ],
  "sheet_name": "销售数据",
  "filename": "Q1销售汇总"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `title` | string | ✅ | 表格标题（写入第一行） |
| `data` | array\[\[...\]\] | ✅ | 二维数组，第一行为表头（必须全为字符串），后续为数据行 |
| `sheet_name` | string | ❌ | Sheet 名称，默认 "Sheet1" |
| `filename` | string | ❌ | 输出文件名 |

**校验规则**:
- `data` 至少需要 2 行（表头 + 1 行数据）
- 第一行（表头）所有元素必须为字符串，否则返回 422

---

#### `POST /api/v1/excel/append_rows` — [EXC-02]
**功能**: 向已有 Excel 文件追加新数据行（不覆盖原内容）。

**请求体**:
```json
{
  "source_excel_url": "https://cos.xxx.com/excel_documents/Q1数据.xlsx",
  "rows": [
    ["西部", "产品C", 150000, "5%"],
    ["华北", "产品A", 210000, "15%"]
  ],
  "sheet_name": "销售数据"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `source_excel_url` | string (URL) | ✅ | 源 Excel 文件的可下载链接 |
| `rows` | array\[\[...\]\] | ✅ | 要追加的数据行列表 |
| `sheet_name` | string | ❌ | 目标 Sheet 名称，默认追加到第一个 Sheet |

---

#### `POST /api/v1/excel/generate_complex` — [EXC-03]
**功能**: 生成包含多 Sheet、可预埋 Excel 公式的复杂报表。

**请求体**:
```json
{
  "title": "2026年财务报告",
  "sheets": [
    {
      "sheet_name": "收入",
      "headers": ["月份", "收入", "环比增长"],
      "data": [
        ["1月", 100000, "=B2/B1-1"],
        ["2月", 120000, "=B3/B2-1"]
      ]
    },
    {
      "sheet_name": "支出",
      "headers": ["月份", "支出"],
      "data": [["1月", 80000], ["2月", 85000]]
    }
  ],
  "filename": "财务报告_2026"
}
```

| sheets 子字段 | 类型 | 说明 |
|---|---|---|
| `sheet_name` | string | Sheet 名称 |
| `headers` | string[] | 表头列表 |
| `data` | array | 数据行，可包含 Excel 公式字符串（如 `"=SUM(B2:B10)"`) |

---

#### `POST /api/v1/excel/extract_range` — [EXC-04]
**功能**: 从大型 Excel 精准提取指定区域数据，返回结构化 JSON。

**请求体**:
```json
{
  "source_excel_url": "https://cos.xxx.com/excel_documents/大表格.xlsx",
  "sheet_name": "数据汇总",
  "cell_range": "A1:E20",
  "keyword": "合计"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `source_excel_url` | string (URL) | ✅ | 源 Excel 文件链接 |
| `sheet_name` | string | ❌ | 目标 Sheet，默认第一个 |
| `cell_range` | string | ❌ | 单元格区间如 `"A1:E20"` |
| `keyword` | string | ❌ | 关键词搜索，返回包含该关键词的行 |

---

### VIS 模块（素材工厂）

> VIS 模块定位：产出可嵌入 Word 文档的图片素材。典型工作流：AI 生成 Mermaid/数据 → 调用 VIS 接口 → 得到图片 URL → 嵌入 Markdown → 调用 DOC-01 生成 Word

#### `POST /api/v1/vis/render_mermaid` — [VIS-01]
**功能**: 将 Mermaid DSL 代码渲染为图片（流程图/时序图/甘特图/思维导图等）。

**请求体**:
```json
{
  "code": "graph TD;\n    A[用户下单] --> B{库存充足?};\n    B -->|是| C[确认发货];\n    B -->|否| D[发送缺货通知];",
  "output_format": "png",
  "theme": "default",
  "width": 1200,
  "height": 800
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `code` | string | ✅ | Mermaid 语法代码，最短 5 字符 |
| `output_format` | `"png"` \| `"svg"` | ❌ | 输出格式，默认 png |
| `theme` | string | ❌ | 主题：`default` \| `dark` \| `forest` \| `neutral` |
| `width` | integer (200~4000) | ❌ | 图片宽度（像素），默认 1200 |
| `height` | integer (200~4000) | ❌ | 图片高度（像素），默认 800 |

**支持的 Mermaid 图类型**: flowchart, sequence, gantt, classDiagram, stateDiagram, er, pie, mindmap, timeline

**实现**: 调用 [mermaid.ink](https://mermaid.ink) 公共渲染服务

---

#### `POST /api/v1/vis/render_chart` — [VIS-02]
**功能**: 根据结构化数据生成统计图表图片（matplotlib 本地渲染）。

**请求体**:
```json
{
  "chart_type": "bar",
  "title": "2026年Q1各区域销售业绩",
  "categories": ["华东", "华南", "华北", "西部"],
  "series": [
    {"name": "产品A", "values": [320000, 280000, 150000, 90000]},
    {"name": "产品B", "values": [200000, 350000, 120000, 60000]}
  ],
  "output_format": "png",
  "width": 900,
  "height": 600
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `chart_type` | enum | ✅ | `bar`/`line`/`pie`/`scatter`/`radar`/`heatmap`/`funnel`/`gauge` |
| `title` | string | ❌ | 图表标题 |
| `categories` | string[] | ✅ | X 轴标签或饼图扇区名称 |
| `series` | array | ✅ | 数据系列，每项含 `name` 和 `values` |
| `output_format` | `"png"` \| `"svg"` | ❌ | 默认 png |
| `width` / `height` | integer | ❌ | 图片尺寸（像素） |

**AI Agent 选型建议**: 趋势用 `line`，占比用 `pie`，对比用 `bar`，分布用 `scatter`

---

#### `POST /api/v1/vis/generate_qrcode` — [VIS-03a]
**功能**: 将文本或 URL 生成为 QR 二维码图片。

**请求体**:
```json
{
  "content": "https://www.example.com/product/123",
  "size": 10,
  "error_correction": "M"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `content` | string (1~4000字符) | ✅ | 要编码的文本或 URL |
| `size` | integer (1~40) | ❌ | 方块尺寸，默认 10 |
| `error_correction` | `L`/`M`/`Q`/`H` | ❌ | 纠错级别，嵌入 logo 建议 H，默认 M |

---

#### `POST /api/v1/vis/generate_barcode` — [VIS-03b]
**功能**: 生成条形码图片。

**请求体**:
```json
{
  "content": "1234567890128",
  "barcode_type": "ean13"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `content` | string | ✅ | 要编码的文本（EAN13 需 12-13 位数字） |
| `barcode_type` | enum | ❌ | `code128`（默认）/`code39`/`ean13`/`ean8`/`isbn13`/`isbn10`/`upca` |

---

#### `POST /api/v1/vis/generate_wordcloud` — [VIS-04]
**功能**: 根据文本内容生成词云图片，支持中文自动分词。

**请求体**:
```json
{
  "text": "人工智能 机器学习 深度学习 自然语言处理 大模型 数据分析 云计算",
  "width": 800,
  "height": 600,
  "max_words": 200,
  "background_color": "white",
  "colormap": "viridis",
  "use_jieba": true
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `text` | string (≥10字符) | ✅ | 输入文本，中文自动 jieba 分词 |
| `width` / `height` | integer | ❌ | 图片尺寸，默认 800×600 |
| `max_words` | integer (10~1000) | ❌ | 最大词数，默认 200 |
| `background_color` | string | ❌ | 背景色（CSS 名或 hex），默认 white |
| `colormap` | string | ❌ | matplotlib 色彩方案，默认 viridis |
| `use_jieba` | boolean | ❌ | 中文分词开关，默认 true |

---

### PDF 模块

#### `POST /api/v1/pdf/convert_from_docx` — [PDF-01]
**功能**: 使用 LibreOffice headless 将 .docx 高保真转换为 PDF。

**典型工作流**: DOC-01 生成 Word → 用户审阅确认 → PDF-01 转换 → PDF-02 盖章归档

**请求体**:
```json
{
  "source_docx_url": "https://cos.xxx.com/documents/周报_20260222.docx",
  "filename": "周报_20260222"
}
```

| 字段 | 类型 | 必填 | 说明 |
|---|---|---|---|
| `source_docx_url` | string (URL) | ✅ | Word 文档的可下载链接 |
| `filename` | string | ❌ | 输出文件名（不含扩展名） |

**依赖环境**: 需要在部署环境安装 LibreOffice（`apt-get install -y libreoffice`）

---

#### `POST /api/v1/pdf/add_watermark` — [PDF-02]
**功能**: 为 PDF 添加文字水印和/或图章盖印（不破坏正文）。

**请求体**:
```json
{
  "source_pdf_url": "https://cos.xxx.com/pdf/合同.pdf",
  "watermark": {
    "text": "内部资料 严禁外传",
    "font_size": 40,
    "opacity": 0.12,
    "angle": -45,
    "color": "#808080"
  },
  "stamp": {
    "stamp_image_url": "https://cos.xxx.com/stamps/合同专用章.png",
    "x": 430,
    "y": 750,
    "width": 120,
    "target_pages": [1, 3]
  }
}
```

| watermark 字段 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `text` | string (1~50) | 必填 | 水印文字 |
| `font_size` | integer (10~100) | 40 | 文字大小 |
| `opacity` | float (0.01~0.5) | 0.15 | 透明度，合同建议 0.1~0.2 |
| `angle` | float (-90~90) | -45 | 倾斜角度 |
| `color` | string (#RRGGBB) | #808080 | 水印颜色 |

| stamp 字段 | 类型 | 说明 |
|---|---|---|
| `stamp_image_url` | URL | 印章图片（建议透明背景 PNG）|
| `x`, `y` | float | 印章左上角坐标（pt），A4 页面 595×842pt |
| `width` | float | 印章宽度（pt），通常 120pt |
| `target_pages` | int[] | 指定盖章页码（1-indexed），null 表示仅最后一页 |

---

#### `POST /api/v1/pdf/merge_split` — [PDF-03]
**功能**: 合并多个 PDF 文件，或从单个 PDF 截取页码区间。

**合并示例**:
```json
{
  "source_pdf_urls": [
    "https://cos.xxx.com/pdf/第一周报.pdf",
    "https://cos.xxx.com/pdf/第二周报.pdf"
  ],
  "output_filename": "双周报合集"
}
```

**拆分（截取）示例**:
```json
{
  "source_pdf_urls": ["https://cos.xxx.com/pdf/完整报告.pdf"],
  "page_ranges": [{"start": 1, "end": 10}],
  "output_filename": "报告摘要_前10页"
}
```

| 字段 | 类型 | 说明 |
|---|---|---|
| `source_pdf_urls` | URL[] | 源 PDF 列表，多个文件按顺序合并 |
| `page_ranges` | array | 页码截取区间，仅单文件时生效；为空则合并所有页 |
| `output_filename` | string | 输出文件名 |

**响应**额外含 `page_count` 字段（最终 PDF 总页数）。

---

### 兼容接口（Legacy）

> 这三个接口保留用于向后兼容，新开发请使用上方标准 API

| 路径 | 说明 |
|---|---|
| `POST /generate-doc` | 旧版文档生成，接受 `{title, content}` |
| `POST /generate_excel` | 旧版 Excel 生成，从 Markdown 表格解析数据 |
| `POST /generate_pdf` | 旧版 PDF 生成，接受 `{docx_url}` |

---

## 五、工程架构概览

```
sga-office/
├── app/
│   ├── main.py                    # FastAPI 应用入口，路由注册，CORS 配置
│   ├── core/
│   │   ├── config.py              # 环境变量配置（pydantic-settings）
│   │   └── cos_client.py          # 腾讯云 COS 客户端初始化
│   ├── api/endpoints/
│   │   ├── doc_routes.py          # DOC-01, DOC-02
│   │   ├── excel_routes.py        # EXC-01 ~ EXC-04
│   │   ├── vis_routes.py          # VIS-01 ~ VIS-04
│   │   ├── pdf_routes.py          # PDF-01 ~ PDF-03
│   │   └── legacy_routes.py       # 兼容接口
│   ├── schemas/
│   │   ├── base.py                # 统一 ApiResponse[T] 泛型结构
│   │   ├── payload_docx.py        # DOC 请求/响应 Schema
│   │   ├── payload_excel.py       # EXC 请求/响应 Schema
│   │   ├── payload_vis.py         # VIS 请求/响应 Schema
│   │   └── payload_pdf.py         # PDF 请求/响应 Schema
│   └── services/
│       ├── cos_storage.py         # COS 上传/下载封装
│       ├── doc_builder.py         # Markdown→Word 渲染，模板填充
│       ├── excel_handler.py       # Excel 创建/追加/复杂报表/读取
│       ├── vis_renderer.py        # VIS 五个真实实现
│       └── pdf_manipulator.py     # LibreOffice 转换，水印/盖章，合并/拆分
├── tests/
│   ├── conftest.py                # 共享 fixtures：TestClient，Mock COS
│   ├── test_schemas.py            # Schema 校验层测试（20个）
│   ├── test_services.py           # Service 层单元测试（14个）
│   ├── test_routes.py             # 端点集成测试（10个）
│   └── test_vis_pdf_routes.py     # VIS/PDF 补充测试（10个）
├── docs/                          # 产品/技术/API 文档
├── pytest.ini                     # pytest 配置
├── requirements.txt               # Python 依赖
├── dockerfile                     # 含 LibreOffice 安装
└── docker-compose.yml             # 精简版（移除了 Redis/Celery Worker）
```

### 关键依赖

| 库 | 版本要求 | 用途 |
|---|---|---|
| fastapi | ≥0.109.0 | API 框架 |
| pydantic | ≥2.5.0 | 数据校验与序列化 |
| python-docx | ≥1.1.0 | Word 文档生成 |
| mistune | ==3.0.2 | Markdown AST 解析 |
| openpyxl | ≥3.1.0 | Excel 操作 |
| pymupdf | ≥1.24.0 | PDF 处理（PyMuPDF/fitz） |
| matplotlib | ≥3.8.0 | 统计图表渲染（VIS-02） |
| qrcode[pil] | ≥7.4.0 | QR 码生成（VIS-03a） |
| python-barcode | ≥0.15.0 | 条形码生成（VIS-03b） |
| wordcloud | ≥1.9.0 | 词云生成（VIS-04） |
| jieba | ≥0.42.0 | 中文分词（VIS-04） |
| cos-python-sdk-v5 | latest | 腾讯云 COS 存储 |
| LibreOffice (系统) | 任意 | Word→PDF 转换（PDF-01） |

---

## 六、部署注意事项

1. **LibreOffice**: PDF-01 依赖系统安装的 LibreOffice。Docker 镜像已包含安装步骤，本地开发环境需手动安装
2. **环境变量**: 需配置 `.env` 文件（参考 `.env.example`）：
   ```
   COS_SECRET_ID=xxx
   COS_SECRET_KEY=xxx
   COS_REGION=ap-guangzhou
   COS_BUCKET_NAME=xxx
   ```
3. **VIS-01 网络**: `render_mermaid_to_image()` 调用 `mermaid.ink` 公共 API，需要服务器能访问外网
4. **中文字体**: 词云生成（VIS-04）在 Linux 环境下需确保系统安装了中文字体，否则中文显示为方块

---

## 七、快速启动

```bash
# 安装依赖
pip install -r requirements.txt

# 配置环境变量
cp .env.example .env
# 编辑 .env 填入 COS 凭证

# 启动服务
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload

# 运行测试
python -m pytest tests/ -v

# 访问 API 文档
open http://localhost:8000/docs
```

---

*报告由 Augment Agent 自动生成，反映 commit `b686fb4` 的代码状态。*

