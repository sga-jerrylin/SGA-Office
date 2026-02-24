# SGA-Office 富文档渲染引擎设计

> **日期**: 2026-02-24
> **状态**: Approved
> **范围**: Docx 增强渲染 + Excel 样式引擎 + Agent 友好报错

---

## 1 背景与动机

当前 SGA-Office 的 Docx 渲染仅支持基础 Markdown 转换（标题、段落、表格、列表、图片），缺少封面页、页眉页脚、目录、高亮框、彩色表格等高级排版能力。Excel 模块仅有基础样式（灰色表头、细边框），缺少主题色、分组着色、冻结窗格、甘特图等企业级效果。

Claude 网页端通过直接写 python-docx 代码可以生成高质量文档，但速度慢、不可局部修改。本设计的目标是将这些排版能力内建到 SGA-Office，Agent 只需输出 Markdown + 轻量元数据，服务端毫秒级渲染出同等质量的文档。

### 设计原则

- **Agent-First**：Agent 最自然的输出是 Markdown，不强迫输出复杂 JSON
- **向后兼容**：所有现有调用零影响，新功能通过可选字段渐进启用
- **报错即文档**：422 错误带完整修正示例，Agent 可自动纠正重试

---

## 2 Docx 增强：Frontmatter 驱动的富排版

### 2.1 输入格式

Agent 在 Markdown 开头加 YAML frontmatter（可选）：

```markdown
---
cover:
  title: "拓竹销售智能体"
  subtitle: "前端入参与上下文快照对接建议书"
  meta:
    - "交付方：炎云团队"
    - "版本：v2.0 | 日期：2026-02"
    - "密级：保密"

header: "拓竹销售智能体 — 前端入参与上下文快照对接建议书"
footer: page_number
toc: true
theme: business_blue
---

# 1 概述

正文内容...
```

### 2.2 Frontmatter 字段规范

| 字段 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `cover` | object | 否 | null | 有则生成封面页 |
| `cover.title` | string | cover 存在时必填 | - | 封面主标题 |
| `cover.subtitle` | string | 否 | null | 封面副标题 |
| `cover.meta` | list[string] | 否 | [] | 元信息行（居中展示） |
| `header` | string | 否 | null | 页眉文字 |
| `footer` | string | 否 | null | `page_number` / `custom_text` / `both` |
| `footer_text` | string | 否 | null | footer 为 custom_text 或 both 时使用 |
| `toc` | bool | 否 | false | 自动生成目录页 |
| `theme` | string | 否 | `business_blue` | 主题色系名称 |

### 2.3 正文增强语法

| Markdown 语法 | 渲染效果 |
|---------------|----------|
| `---`（三个以上横线，独占一行，非 frontmatter 分隔符） | 分页符 |
| `> [!INFO] 标题` | 蓝色高亮框 |
| `> [!NOTE] 标题` | 绿色提示框 |
| `> [!WARNING] 标题` | 橙色警告框 |
| 标准 Markdown 表格 | 自动应用主题色表头 + 交替行背景 |

### 2.4 主题色系

| 主题名 | 标题色 | 表头色 | 高亮框色 | 适用场景 |
|--------|--------|--------|----------|----------|
| `business_blue` | #2E75B6 | #2E75B6 | #D6E4F0 | 商务报告（默认） |
| `government_red` | #C00000 | #C00000 | #FDE9D9 | 政务/公文 |
| `tech_dark` | #404040 | #505050 | #F2F2F2 | 技术文档 |
| `academic_green` | #548235 | #548235 | #E2EFDA | 学术/研究 |
| `minimal` | #333333 | #F2F2F2 | #F7F7F7 | 简约风格 |

---

## 3 Excel 增强：可选 style 引擎

### 3.1 style 字段（EXC-01 / EXC-03 可选）

```json
{
  "style": {
    "theme": "business_blue",
    "header_style": "colored",
    "freeze_panes": "A3",
    "auto_filter": true,
    "alternating_rows": true,
    "column_widths": {"A": 6, "B": 12, "C": 30},
    "row_groups": {
      "group_column": "B",
      "colors": {
        "环境部署": "#2E75B6",
        "知识库问答": "#7030A0"
      }
    },
    "gantt": {
      "date_columns": ["D", "E"],
      "timeline_start": "2026-03-01",
      "timeline_end": "2026-06-30",
      "granularity": "week",
      "bar_color_column": "B"
    }
  }
}
```

### 3.2 style 字段规范

| 字段 | 类型 | 必填 | 默认值 | 说明 |
|------|------|------|--------|------|
| `theme` | string | 否 | `business_blue` | 主题色系 |
| `header_style` | string | 否 | `colored` | `colored` / `minimal` / `bold_only` |
| `freeze_panes` | string | 否 | null | 冻结窗格位置如 `"A3"` |
| `auto_filter` | bool | 否 | false | 表头自动筛选 |
| `alternating_rows` | bool | 否 | true | 交替行背景色 |
| `column_widths` | dict | 否 | null | 手动列宽，不指定走自动 |
| `row_groups` | object | 否 | null | 按列值分组着色 |
| `row_groups.group_column` | string | row_groups 时必填 | - | 分组依据列字母 |
| `row_groups.colors` | dict | 否 | 自动分配 | 分组值→颜色映射 |
| `gantt` | object | 否 | null | 甘特图时间线（可选） |
| `gantt.date_columns` | [string, string] | gantt 时必填 | - | 开始/结束日期列 |
| `gantt.timeline_start` | string | gantt 时必填 | - | 时间轴起始日期 |
| `gantt.timeline_end` | string | gantt 时必填 | - | 时间轴结束日期 |
| `gantt.granularity` | string | 否 | `week` | `day` / `week` / `month` |
| `gantt.bar_color_column` | string | 否 | null | 条形颜色跟随哪列的分组色 |

### 3.3 渐进式启用

```
不传 style         → 现有默认样式（零影响）
传 style 不传 gantt → 增强样式（彩色表头、分组、冻结...）
传 style + gantt   → 增强样式 + 右侧甘特图时间线
```

---

## 4 Agent 友好报错机制

### 4.1 统一错误响应结构

```json
{
  "code": 422,
  "message": "frontmatter 中 cover.title 是必填字段",
  "agent_hint": {
    "error_type": "missing_field",
    "field": "cover.title",
    "fix_suggestion": "在 YAML frontmatter 的 cover 块中添加 title 字段",
    "correct_example": "---\ncover:\n  title: \"你的文档标题\"\n---\n# 正文",
    "docs_url": "/api/v1/docs/frontmatter-spec"
  }
}
```

### 4.2 错误类型枚举

| error_type | 场景 | agent_hint 内容 |
|------------|------|-----------------|
| `invalid_frontmatter` | YAML 语法错误 | 指出错误行 + 正确 YAML 示例 |
| `missing_field` | 必填字段缺失 | 缺哪个字段 + 补全示例 |
| `invalid_value` | 字段值不合法 | 合法值枚举列表 |
| `type_error` | 类型错误 | 期望类型 + 正确示例 |
| `empty_content` | 内容为空 | 最小可用示例 |
| `excel_style_error` | style 配置错误 | 完整 style 配置示例 |

### 4.3 实现方式

- Pydantic `field_validator` 抛出带结构化信息的 ValueError
- FastAPI 全局 exception handler 捕获后包装为 `{code, message, agent_hint}`
- 每个错误类型预置 `correct_example` 模板

---

## 5 文件变更清单

```
修改:
  app/services/doc_builder.py       # frontmatter 解析 + 封面/页眉页脚/TOC/高亮框/分页符/彩色表格
  app/services/excel_handler.py     # style 引擎 + 分组着色 + 甘特图
  app/schemas/payload_docx.py       # frontmatter 容错校验
  app/schemas/payload_excel.py      # style 字段 Schema
  app/schemas/base.py               # 统一错误响应结构
  app/api/endpoints/doc_routes.py   # 错误处理增强
  app/api/endpoints/excel_routes.py # 错误处理增强
  app/main.py                       # 全局 422 exception handler

新增:
  app/core/themes.py                # 主题色系定义（Docx + Excel 共用）
  app/core/error_hints.py           # agent_hint 生成器
  tests/test_frontmatter.py         # frontmatter 解析 + 富渲染测试
  tests/test_excel_style.py         # Excel 样式引擎测试
  tests/test_error_hints.py         # 报错机制测试
```

## 6 新增依赖

```
pyyaml    # YAML frontmatter 解析
```

## 7 向后兼容保证

| 现有调用方式 | 影响 |
|-------------|------|
| DOC-01 传纯 Markdown（无 frontmatter） | 零影响 |
| EXC-01 传 data 不传 style | 零影响 |
| EXC-03 传 sheets 不传 style | 零影响 |
| Legacy 接口 | 零影响 |
| 现有 54 个测试 | 必须全部通过 |
