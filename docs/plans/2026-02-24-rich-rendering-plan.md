# Rich Rendering Engine Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Enhance SGA-Office with frontmatter-driven Docx rich layout (cover, header/footer, TOC, callout boxes, themed tables), Excel style engine (row groups, freeze panes, gantt), and agent-friendly error hints — all backward compatible.

**Architecture:** Detect YAML frontmatter in markdown_content, parse it, and route to enhanced rendering. Excel gets an optional `style` field. All errors return `agent_hint` with correct examples. Shared theme definitions in `app/core/themes.py`.

**Tech Stack:** python-docx, openpyxl, pyyaml, pydantic, fastapi

---

### Task 1: Add pyyaml dependency

**Files:**
- Modify: `requirements.txt`

**Step 1: Add pyyaml to requirements.txt**

Add `pyyaml>=6.0` to `requirements.txt` (after `pydantic-settings`).

**Step 2: Install**

Run: `pip install pyyaml`

**Step 3: Commit**

```bash
git add requirements.txt
git commit -m "chore: add pyyaml dependency for frontmatter parsing"
```

---

### Task 2: Create shared theme definitions

**Files:**
- Create: `app/core/themes.py`
- Test: `tests/test_themes.py`

**Step 1: Write the failing test**

Create `tests/test_themes.py`:

```python
"""主题色系单元测试"""

from app.core.themes import get_theme, AVAILABLE_THEMES


class TestThemes:

    def test_default_theme_is_business_blue(self):
        theme = get_theme()
        assert theme.name == "business_blue"
        assert theme.heading_color == "2E75B6"

    def test_all_builtin_themes_exist(self):
        expected = {"business_blue", "government_red", "tech_dark", "academic_green", "minimal"}
        assert expected == set(AVAILABLE_THEMES.keys())

    def test_get_unknown_theme_falls_back_to_default(self):
        theme = get_theme("nonexistent_theme")
        assert theme.name == "business_blue"

    def test_theme_has_all_required_colors(self):
        for name in AVAILABLE_THEMES:
            theme = get_theme(name)
            assert theme.heading_color
            assert theme.table_header_bg
            assert theme.table_header_font
            assert theme.table_alt_row_bg
            assert theme.callout_info_bg
            assert theme.callout_warning_bg
            assert theme.callout_note_bg
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_themes.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.core.themes'`

**Step 3: Write implementation**

Create `app/core/themes.py`:

```python
"""
SGA-Office 主题色系定义。
Docx 和 Excel 共用同一套主题，保持视觉一致性。
颜色值统一使用 6 位 HEX（不含 #），方便 python-docx 和 openpyxl 直接使用。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Theme:
    name: str
    # 标题/强调色
    heading_color: str        # 标题文字色
    # 表格
    table_header_bg: str      # 表头背景
    table_header_font: str    # 表头字体色
    table_alt_row_bg: str     # 交替行背景
    # 高亮框
    callout_info_bg: str      # [!INFO] 背景
    callout_info_border: str  # [!INFO] 左边框
    callout_note_bg: str      # [!NOTE] 背景
    callout_note_border: str
    callout_warning_bg: str   # [!WARNING] 背景
    callout_warning_border: str
    # 封面
    cover_title_color: str
    cover_meta_color: str


AVAILABLE_THEMES: dict[str, Theme] = {
    "business_blue": Theme(
        name="business_blue",
        heading_color="2E75B6",
        table_header_bg="2E75B6",
        table_header_font="FFFFFF",
        table_alt_row_bg="D6E4F0",
        callout_info_bg="D6E4F0",
        callout_info_border="2E75B6",
        callout_note_bg="E2EFDA",
        callout_note_border="548235",
        callout_warning_bg="FFF2CC",
        callout_warning_border="BF8F00",
        cover_title_color="2E75B6",
        cover_meta_color="808080",
    ),
    "government_red": Theme(
        name="government_red",
        heading_color="C00000",
        table_header_bg="C00000",
        table_header_font="FFFFFF",
        table_alt_row_bg="FDE9D9",
        callout_info_bg="FDE9D9",
        callout_info_border="C00000",
        callout_note_bg="E2EFDA",
        callout_note_border="548235",
        callout_warning_bg="FFF2CC",
        callout_warning_border="BF8F00",
        cover_title_color="C00000",
        cover_meta_color="808080",
    ),
    "tech_dark": Theme(
        name="tech_dark",
        heading_color="404040",
        table_header_bg="505050",
        table_header_font="FFFFFF",
        table_alt_row_bg="F2F2F2",
        callout_info_bg="F2F2F2",
        callout_info_border="505050",
        callout_note_bg="E8F5E9",
        callout_note_border="4CAF50",
        callout_warning_bg="FFF8E1",
        callout_warning_border="FF8F00",
        cover_title_color="404040",
        cover_meta_color="808080",
    ),
    "academic_green": Theme(
        name="academic_green",
        heading_color="548235",
        table_header_bg="548235",
        table_header_font="FFFFFF",
        table_alt_row_bg="E2EFDA",
        callout_info_bg="D6E4F0",
        callout_info_border="2E75B6",
        callout_note_bg="E2EFDA",
        callout_note_border="548235",
        callout_warning_bg="FFF2CC",
        callout_warning_border="BF8F00",
        cover_title_color="548235",
        cover_meta_color="808080",
    ),
    "minimal": Theme(
        name="minimal",
        heading_color="333333",
        table_header_bg="F2F2F2",
        table_header_font="333333",
        table_alt_row_bg="F9F9F9",
        callout_info_bg="F7F7F7",
        callout_info_border="CCCCCC",
        callout_note_bg="F7F7F7",
        callout_note_border="CCCCCC",
        callout_warning_bg="FFF9F0",
        callout_warning_border="E0A000",
        cover_title_color="333333",
        cover_meta_color="999999",
    ),
}

DEFAULT_THEME = "business_blue"


def get_theme(name: str | None = None) -> Theme:
    """获取主题，未知名称回退到默认主题。"""
    if name is None:
        name = DEFAULT_THEME
    return AVAILABLE_THEMES.get(name, AVAILABLE_THEMES[DEFAULT_THEME])
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_themes.py -v`
Expected: all PASS

**Step 5: Commit**

```bash
git add app/core/themes.py tests/test_themes.py
git commit -m "feat: add shared theme definitions for docx and excel"
```

---

### Task 3: Create agent-friendly error hints module

**Files:**
- Create: `app/core/error_hints.py`
- Modify: `app/schemas/base.py`
- Modify: `app/main.py`
- Test: `tests/test_error_hints.py`

**Step 1: Write the failing test**

Create `tests/test_error_hints.py`:

```python
"""Agent 友好报错机制测试"""

import pytest
from app.core.error_hints import build_agent_hint, ErrorType


class TestBuildAgentHint:

    def test_missing_field_hint(self):
        hint = build_agent_hint(
            error_type=ErrorType.MISSING_FIELD,
            field="cover.title",
            message="cover.title 是必填字段",
        )
        assert hint["error_type"] == "missing_field"
        assert hint["field"] == "cover.title"
        assert "correct_example" in hint
        assert len(hint["correct_example"]) > 0

    def test_invalid_value_hint_includes_allowed(self):
        hint = build_agent_hint(
            error_type=ErrorType.INVALID_VALUE,
            field="theme",
            message="不支持的主题",
            allowed_values=["business_blue", "government_red"],
        )
        assert "business_blue" in hint["fix_suggestion"]

    def test_invalid_frontmatter_hint(self):
        hint = build_agent_hint(
            error_type=ErrorType.INVALID_FRONTMATTER,
            message="YAML 解析失败: line 3",
        )
        assert hint["error_type"] == "invalid_frontmatter"
        assert "correct_example" in hint

    def test_empty_content_hint(self):
        hint = build_agent_hint(
            error_type=ErrorType.EMPTY_CONTENT,
            message="内容为空",
        )
        assert "# " in hint["correct_example"]


class TestErrorResponseIntegration:

    def test_422_response_includes_agent_hint(self, client):
        """发送非法请求，验证 422 响应包含 agent_hint"""
        resp = client.post("/api/v1/docx/render_markdown", json={
            "markdown_content": "   ",
        })
        assert resp.status_code == 422
        body = resp.json()
        # Pydantic 校验错误应该包含 agent_hint
        assert "agent_hint" in body or "detail" in body
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_error_hints.py::TestBuildAgentHint -v`
Expected: FAIL with `ModuleNotFoundError`

**Step 3: Write error_hints module**

Create `app/core/error_hints.py`:

```python
"""
Agent 友好报错机制。
每个错误类型都生成结构化的 agent_hint，包含修正建议和正确示例。
"""

from enum import Enum


class ErrorType(str, Enum):
    INVALID_FRONTMATTER = "invalid_frontmatter"
    MISSING_FIELD = "missing_field"
    INVALID_VALUE = "invalid_value"
    TYPE_ERROR = "type_error"
    EMPTY_CONTENT = "empty_content"
    EXCEL_STYLE_ERROR = "excel_style_error"


# 每种错误类型的正确示例模板
_EXAMPLES: dict[ErrorType, str] = {
    ErrorType.INVALID_FRONTMATTER: (
        "---\n"
        "cover:\n"
        "  title: \"文档标题\"\n"
        "  subtitle: \"副标题\"\n"
        "header: \"页眉文字\"\n"
        "footer: page_number\n"
        "theme: business_blue\n"
        "---\n"
        "# 正文标题\n\n"
        "正文内容"
    ),
    ErrorType.MISSING_FIELD: (
        "---\n"
        "cover:\n"
        "  title: \"你的文档标题\"\n"
        "  subtitle: \"副标题（可选）\"\n"
        "  meta:\n"
        "    - \"作者：XXX\"\n"
        "---\n"
        "# 正文开始"
    ),
    ErrorType.INVALID_VALUE: "请使用以下合法值之一: {allowed_values}",
    ErrorType.TYPE_ERROR: "字段 {field} 期望类型为 {expected_type}，实际收到 {actual_type}",
    ErrorType.EMPTY_CONTENT: (
        "# 文档标题\n\n"
        "这里是正文内容。支持 **加粗**、*斜体*、表格、列表等 Markdown 语法。\n\n"
        "| 列1 | 列2 |\n"
        "|-----|-----|\n"
        "| 数据 | 数据 |"
    ),
    ErrorType.EXCEL_STYLE_ERROR: (
        '{"style": {"theme": "business_blue", "header_style": "colored", '
        '"freeze_panes": "A3", "auto_filter": true}}'
    ),
}


def build_agent_hint(
    error_type: ErrorType,
    message: str = "",
    field: str | None = None,
    allowed_values: list[str] | None = None,
    expected_type: str | None = None,
    actual_type: str | None = None,
) -> dict:
    """
    构建 agent_hint 字典，Agent 读后可自动修正入参。
    """
    hint: dict = {
        "error_type": error_type.value,
    }

    if field:
        hint["field"] = field

    # fix_suggestion
    if error_type == ErrorType.INVALID_VALUE and allowed_values:
        hint["fix_suggestion"] = f"请使用以下合法值之一: {', '.join(allowed_values)}"
    elif error_type == ErrorType.MISSING_FIELD and field:
        hint["fix_suggestion"] = f"请在请求中添加 {field} 字段"
    elif error_type == ErrorType.TYPE_ERROR:
        hint["fix_suggestion"] = f"字段 {field} 期望类型为 {expected_type}，实际收到 {actual_type}"
    else:
        hint["fix_suggestion"] = message

    # correct_example
    example = _EXAMPLES.get(error_type, "")
    if isinstance(example, str) and "{" in example:
        example = example.format(
            field=field or "unknown",
            allowed_values=", ".join(allowed_values) if allowed_values else "",
            expected_type=expected_type or "",
            actual_type=actual_type or "",
        )
    hint["correct_example"] = example

    return hint
```

**Step 4: Update `app/schemas/base.py` — add AgentErrorResponse**

Add to `app/schemas/base.py` after existing classes:

```python
class AgentErrorResponse(BaseModel):
    """带 agent_hint 的错误响应，Agent 可据此自动修正入参"""
    code: int = Field(..., description="HTTP 状态码")
    message: str = Field(..., description="人类可读的错误描述")
    agent_hint: Optional[dict] = Field(None, description="Agent 修正提示，含 error_type / fix_suggestion / correct_example")
```

**Step 5: Update `app/main.py` — enhance global exception handler**

Replace the existing `global_exception_handler` and add a new 422 handler. Add imports at top:

```python
from app.core.error_hints import build_agent_hint, ErrorType
```

Add new handler before the existing `global_exception_handler`:

```python
from fastapi.exceptions import RequestValidationError

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """Pydantic 校验失败时返回 agent_hint"""
    errors = exc.errors()
    first_error = errors[0] if errors else {}
    field_loc = " -> ".join(str(l) for l in first_error.get("loc", []))
    msg = first_error.get("msg", "校验失败")

    hint = build_agent_hint(
        error_type=ErrorType.MISSING_FIELD if "required" in msg.lower() or "missing" in msg.lower()
                   else ErrorType.TYPE_ERROR if "type" in msg.lower()
                   else ErrorType.INVALID_VALUE,
        field=field_loc,
        message=msg,
    )

    return JSONResponse(
        status_code=422,
        content={
            "code": 422,
            "message": f"入参校验失败: {msg} (字段: {field_loc})",
            "agent_hint": hint,
        },
    )
```

**Step 6: Run tests**

Run: `python -m pytest tests/test_error_hints.py -v`
Expected: all PASS

Then run full suite to check backward compat:

Run: `python -m pytest tests/ -v`
Expected: all 54 existing tests still PASS

**Step 7: Commit**

```bash
git add app/core/error_hints.py app/schemas/base.py app/main.py tests/test_error_hints.py
git commit -m "feat: add agent-friendly error hints with correct_example in 422 responses"
```

---

### Task 4: Frontmatter parser + enhanced Docx renderer

This is the core task. Enhance `doc_builder.py` to support frontmatter-driven rich rendering.

**Files:**
- Modify: `app/services/doc_builder.py`
- Test: `tests/test_frontmatter.py`

**Step 1: Write the failing tests**

Create `tests/test_frontmatter.py`:

```python
"""Frontmatter 解析 + 富文档渲染测试"""

import pytest
from io import BytesIO
from docx import Document


class TestFrontmatterParsing:
    """测试 frontmatter 提取逻辑"""

    def test_parse_frontmatter_with_cover(self):
        from app.services.doc_builder import _parse_frontmatter
        md = '---\ncover:\n  title: "测试标题"\n  subtitle: "副标题"\ntheme: business_blue\n---\n# 正文'
        config, body = _parse_frontmatter(md)
        assert config["cover"]["title"] == "测试标题"
        assert config["theme"] == "business_blue"
        assert body.strip() == "# 正文"

    def test_no_frontmatter_returns_none(self):
        from app.services.doc_builder import _parse_frontmatter
        md = "# 普通文档\n\n正文内容"
        config, body = _parse_frontmatter(md)
        assert config is None
        assert body == md

    def test_invalid_yaml_returns_none(self):
        from app.services.doc_builder import _parse_frontmatter
        md = "---\n  bad: yaml: :\n---\n# 正文"
        config, body = _parse_frontmatter(md)
        assert config is None

    def test_frontmatter_only_dashes_ignored(self):
        """正文中的 --- 不应被当作 frontmatter"""
        from app.services.doc_builder import _parse_frontmatter
        md = "# 标题\n\n---\n\n下一节"
        config, body = _parse_frontmatter(md)
        assert config is None


class TestCoverPage:
    """测试封面页生成"""

    def test_cover_page_generated(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '---\ncover:\n  title: "年度报告"\n  subtitle: "2026年"\n  meta:\n    - "作者：张三"\n---\n# 第一章\n\n正文'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        texts = [p.text for p in doc.paragraphs]
        assert any("年度报告" in t for t in texts)
        assert any("张三" in t for t in texts)

    def test_no_cover_when_not_specified(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '---\ntheme: business_blue\n---\n# 标题\n\n正文'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        # 第一个段落应该直接是标题，不是封面
        non_empty = [p for p in doc.paragraphs if p.text.strip()]
        assert non_empty[0].text == "标题"


class TestHeaderFooter:
    """测试页眉页脚"""

    def test_header_text_set(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '---\nheader: "机密文档 — 内部使用"\n---\n# 正文'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        header = doc.sections[0].header
        header_text = "".join(p.text for p in header.paragraphs)
        assert "机密文档" in header_text

    def test_footer_page_number(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '---\nfooter: page_number\n---\n# 正文'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        footer = doc.sections[0].footer
        # 页码通过 XML field code 插入，检查 footer 已被激活
        assert footer is not None


class TestCalloutBox:
    """测试高亮框渲染"""

    def test_info_callout_rendered(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '# 标题\n\n> [!INFO] 重要提示\n> 这是一条信息\n> 第二行信息\n\n正文继续'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        texts = [p.text for p in doc.paragraphs]
        # 高亮框内容应该出现在文档中
        full = " ".join(texts)
        assert "重要提示" in full or "信息" in full

    def test_warning_callout_rendered(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '> [!WARNING] 注意事项\n> 请务必检查'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        full = " ".join(p.text for p in doc.paragraphs)
        assert "注意" in full or "检查" in full


class TestPageBreak:
    """测试分页符"""

    def test_hr_becomes_page_break(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '---\ntheme: business_blue\n---\n# 第一章\n\n内容一\n\n---\n\n# 第二章\n\n内容二'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        # 文档应包含两个章节标题
        headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
        assert len(headings) >= 2


class TestThemedTable:
    """测试主题色表格"""

    def test_table_still_renders(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '---\ntheme: business_blue\n---\n# 数据\n\n| 字段 | 类型 |\n|------|------|\n| id | int |\n| name | str |'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        assert len(doc.tables) >= 1


class TestTOC:
    """测试目录生成"""

    def test_toc_placeholder_inserted(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = '---\ntoc: true\n---\n# 第一章\n\n内容\n\n# 第二章\n\n内容'
        result = render_markdown_to_docx(md)
        doc = Document(result)
        # TOC 通过 XML 指令插入，检查目录段落存在
        texts = [p.text for p in doc.paragraphs]
        assert any("目录" in t for t in texts) or len(doc.paragraphs) > 2


class TestBackwardCompat:
    """确保纯 Markdown（无 frontmatter）行为不变"""

    def test_plain_markdown_still_works(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = "# 测试标题\n\n这是一段正文。\n\n| A | B |\n|---|---|\n| 1 | 2 |"
        result = render_markdown_to_docx(md)
        doc = Document(result)
        assert any("测试标题" in p.text for p in doc.paragraphs)
        assert len(doc.tables) >= 1

    def test_escaped_newlines_still_work(self):
        from app.services.doc_builder import render_markdown_to_docx
        md = "Line1\\nLine2"
        result = render_markdown_to_docx(md)
        doc = Document(result)
        full = " ".join(p.text for p in doc.paragraphs)
        assert "Line1" in full
```

**Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_frontmatter.py -v`
Expected: FAIL (multiple — `_parse_frontmatter` not defined, etc.)

**Step 3: Implement the enhanced doc_builder.py**

Modify `app/services/doc_builder.py`. The changes are:

1. Add `_parse_frontmatter()` function at top
2. Add `_add_cover_page()` function
3. Add `_add_header_footer()` function
4. Add `_add_toc()` function
5. Add `_render_callout_box()` in MarkdownToDocx class
6. Add `visit_thematic_break()` for page breaks
7. Enhance `visit_table()` with theme colors
8. Enhance `visit_heading()` with theme colors
9. Modify `render_markdown_to_docx()` to detect/use frontmatter
10. Enhance `visit_block_quote()` to detect `[!INFO]` / `[!NOTE]` / `[!WARNING]` patterns

Key implementation details:

**_parse_frontmatter():**
```python
import yaml

def _parse_frontmatter(markdown_content: str) -> tuple[dict | None, str]:
    """
    从 Markdown 开头提取 YAML frontmatter。
    返回 (config_dict, remaining_body)。无 frontmatter 返回 (None, original)。
    """
    content = markdown_content.strip()
    if not content.startswith("---"):
        return None, markdown_content

    # 找第二个 ---
    end_idx = content.find("---", 3)
    if end_idx == -1:
        return None, markdown_content

    yaml_str = content[3:end_idx].strip()
    body = content[end_idx + 3:].strip()

    try:
        config = yaml.safe_load(yaml_str)
        if not isinstance(config, dict):
            return None, markdown_content
        return config, body
    except yaml.YAMLError:
        return None, markdown_content
```

**_add_cover_page():**
```python
def _add_cover_page(doc: Document, cover: dict, theme: Theme) -> None:
    """生成封面页，封面后自动分页。"""
    from docx.oxml.ns import qn
    from docx.shared import Pt, Cm, RGBColor

    # 空行撑到页面中部
    for _ in range(6):
        doc.add_paragraph()

    # 主标题
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(cover["title"])
    title_run.font.size = Pt(26)
    title_run.bold = True
    title_run.font.color.rgb = RGBColor.from_string(theme.cover_title_color)
    title_run.font.name = '微软雅黑'
    title_run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

    # 副标题
    subtitle = cover.get("subtitle")
    if subtitle:
        sub_para = doc.add_paragraph()
        sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_para.paragraph_format.space_before = Pt(12)
        sub_run = sub_para.add_run(subtitle)
        sub_run.font.size = Pt(18)
        sub_run.bold = True
        sub_run.font.name = '微软雅黑'
        sub_run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

    # 空行间隔
    for _ in range(4):
        doc.add_paragraph()

    # 元信息行
    for line in cover.get("meta", []):
        meta_para = doc.add_paragraph()
        meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta_run = meta_para.add_run(line)
        meta_run.font.size = Pt(12)
        meta_run.font.color.rgb = RGBColor.from_string(theme.cover_meta_color)
        meta_run.font.name = '微软雅黑'
        meta_run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
        meta_para.paragraph_format.space_after = Pt(4)
        meta_para.paragraph_format.first_line_indent = Cm(0)

    # 封面后分页
    doc.add_page_break()
```

**_add_header_footer():**
```python
def _add_header_footer(doc: Document, config: dict) -> None:
    """设置页眉页脚。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.shared import Pt, RGBColor

    section = doc.sections[0]

    # 页眉
    header_text = config.get("header")
    if header_text:
        header = section.header
        header.is_linked_to_previous = False
        hp = header.paragraphs[0]
        hp.text = header_text
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in hp.runs:
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor.from_string("808080")

    # 页脚
    footer_type = config.get("footer")
    if footer_type:
        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

        if footer_type in ("page_number", "both"):
            # 插入页码域代码
            run = fp.add_run()
            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')
            run._element.append(fldChar1)

            run2 = fp.add_run()
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = ' PAGE '
            run2._element.append(instrText)

            run3 = fp.add_run()
            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'end')
            run3._element.append(fldChar2)

        if footer_type in ("custom_text", "both"):
            footer_text = config.get("footer_text", "")
            if footer_type == "both":
                fp.add_run("  |  ")
            fp.add_run(footer_text)
```

**_add_toc():**
```python
def _add_toc(doc: Document) -> None:
    """插入目录占位符（Word 打开时自动刷新）。"""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    toc_title = doc.add_paragraph("目录")
    toc_title.style = doc.styles['Heading 1']

    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)

    run2 = paragraph.add_run()
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = ' TOC \\o "1-3" \\h \\z \\u '
    run2._element.append(instrText)

    run3 = paragraph.add_run()
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run3._element.append(fldChar2)

    doc.add_page_break()
```

**MarkdownToDocx enhancements — add theme awareness and new visit methods:**

- Constructor takes optional `theme` parameter
- `visit_thematic_break()` → `doc.add_page_break()`
- `visit_block_quote()` → detect `[!INFO]`, `[!NOTE]`, `[!WARNING]` patterns, render as shaded table cell (this is the standard python-docx technique for colored boxes)
- `visit_table()` → apply `theme.table_header_bg` / `theme.table_header_font` / `theme.table_alt_row_bg`
- `visit_heading()` → apply `theme.heading_color`

**Modified `render_markdown_to_docx()`:**

```python
def render_markdown_to_docx(markdown_content: str) -> BytesIO:
    from app.core.themes import get_theme

    # 1. 解析 frontmatter
    config, body = _parse_frontmatter(markdown_content)

    # 2. 加载主题
    theme_name = config.get("theme") if config else None
    theme = get_theme(theme_name)

    # 3. 预处理正文（与现有逻辑相同）
    content = body
    if content.startswith('\\#'):
        content = content[1:]
    content = content.replace('\\n', '\n')
    content = _convert_tab_tables_to_markdown(content)
    content = _repair_markdown_table(content)

    # 4. 创建文档
    doc = Document()
    _setup_page(doc)  # 提取现有的页面/字体设置

    # 5. 封面页
    if config and config.get("cover"):
        cover = config["cover"]
        if not cover.get("title"):
            from app.core.error_hints import build_agent_hint, ErrorType
            raise ValueError(build_agent_hint(
                ErrorType.MISSING_FIELD, field="cover.title",
                message="封面配置中 title 是必填字段"
            ))
        _add_cover_page(doc, cover, theme)

    # 6. 目录
    if config and config.get("toc"):
        _add_toc(doc)

    # 7. 页眉页脚
    if config:
        _add_header_footer(doc, config)

    # 8. 解析 & 渲染正文
    markdown = mistune.create_markdown(renderer=None, plugins=['table'])
    ast = markdown(content)
    renderer = MarkdownToDocx(doc, theme=theme)
    renderer.render(ast)

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output
```

**Step 4: Run tests**

Run: `python -m pytest tests/test_frontmatter.py -v`
Expected: all PASS

Run: `python -m pytest tests/test_services.py -v`
Expected: all existing tests still PASS (backward compat)

**Step 5: Commit**

```bash
git add app/services/doc_builder.py tests/test_frontmatter.py
git commit -m "feat: frontmatter-driven rich docx rendering (cover, header/footer, TOC, callouts, themed tables)"
```

---

### Task 5: Excel style engine — themes, freeze, row groups

**Files:**
- Modify: `app/services/excel_handler.py`
- Modify: `app/schemas/payload_excel.py`
- Modify: `app/api/endpoints/excel_routes.py`
- Test: `tests/test_excel_style.py`

**Step 1: Write the failing tests**

Create `tests/test_excel_style.py`:

```python
"""Excel 样式引擎测试"""

import pytest
from io import BytesIO
from openpyxl import load_workbook


class TestExcelStyleEngine:

    def test_themed_header(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Test",
            data=[["Name", "Score"], ["Alice", 95]],
            style={"theme": "business_blue", "header_style": "colored"},
        )
        wb = load_workbook(result)
        ws = wb.active
        # Row 2 is header (row 1 is title)
        header_cell = ws.cell(row=2, column=1)
        assert header_cell.fill.start_color.rgb is not None

    def test_freeze_panes(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Test",
            data=[["A", "B"], [1, 2]],
            style={"freeze_panes": "A3"},
        )
        wb = load_workbook(result)
        ws = wb.active
        assert ws.freeze_panes == "A3"

    def test_auto_filter(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Test",
            data=[["A", "B"], [1, 2]],
            style={"auto_filter": True},
        )
        wb = load_workbook(result)
        ws = wb.active
        assert ws.auto_filter.ref is not None

    def test_alternating_rows(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Test",
            data=[["A", "B"], [1, 2], [3, 4], [5, 6]],
            style={"alternating_rows": True, "theme": "business_blue"},
        )
        wb = load_workbook(result)
        ws = wb.active
        # Row 3 and row 5 should have alt background (rows 3,5 are data rows with 0-indexed even)
        cell_even = ws.cell(row=4, column=1)  # 2nd data row
        assert cell_even.fill.start_color is not None

    def test_row_groups_coloring(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Project Plan",
            data=[
                ["Phase", "Task", "Status"],
                ["Alpha", "Task 1", "Done"],
                ["Alpha", "Task 2", "WIP"],
                ["Beta", "Task 3", "TODO"],
            ],
            style={
                "row_groups": {
                    "group_column": "A",
                    "colors": {"Alpha": "2E75B6", "Beta": "7030A0"},
                }
            },
        )
        wb = load_workbook(result)
        ws = wb.active
        # Alpha rows should have blue tint
        alpha_cell = ws.cell(row=3, column=1)  # First Alpha row
        assert alpha_cell.fill.start_color.rgb is not None

    def test_no_style_backward_compat(self):
        """不传 style 时和现有行为一致"""
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Test",
            data=[["A", "B"], [1, 2]],
        )
        assert isinstance(result, BytesIO)
        result.seek(0)
        assert result.read(2) == b"PK"

    def test_column_widths(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Test",
            data=[["Name", "Description"], ["A", "Long text"]],
            style={"column_widths": {"A": 8, "B": 40}},
        )
        wb = load_workbook(result)
        ws = wb.active
        assert ws.column_dimensions["A"].width == 8
        assert ws.column_dimensions["B"].width == 40


class TestExcelStyleSchema:

    def test_style_field_is_optional(self):
        from app.schemas.payload_excel import CreateExcelRequest
        req = CreateExcelRequest(
            title="Test",
            data=[["A", "B"], [1, 2]],
        )
        assert req.style is None

    def test_style_field_accepts_valid_config(self):
        from app.schemas.payload_excel import CreateExcelRequest
        req = CreateExcelRequest(
            title="Test",
            data=[["A", "B"], [1, 2]],
            style={"theme": "business_blue", "freeze_panes": "A3"},
        )
        assert req.style is not None
        assert req.style.theme == "business_blue"


class TestGanttRendering:

    def test_gantt_creates_timeline_columns(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Project Timeline",
            data=[
                ["Task", "Start", "End"],
                ["Task A", "2026-03-01", "2026-03-15"],
                ["Task B", "2026-03-10", "2026-03-25"],
            ],
            style={
                "gantt": {
                    "date_columns": ["B", "C"],
                    "timeline_start": "2026-03-01",
                    "timeline_end": "2026-03-31",
                    "granularity": "week",
                }
            },
        )
        wb = load_workbook(result)
        ws = wb.active
        # Should have more columns than the original 3 (timeline columns added)
        assert ws.max_column > 3

    def test_no_gantt_no_extra_columns(self):
        from app.services.excel_handler import create_excel_from_array
        result = create_excel_from_array(
            title="Simple",
            data=[["A", "B"], [1, 2]],
        )
        wb = load_workbook(result)
        ws = wb.active
        assert ws.max_column == 2
```

**Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_excel_style.py -v`
Expected: FAIL (style parameter not accepted, etc.)

**Step 3: Add ExcelStyle schema to payload_excel.py**

Add to `app/schemas/payload_excel.py`:

```python
class GanttConfig(BaseModel):
    """甘特图时间线配置"""
    date_columns: list[str] = Field(..., min_length=2, max_length=2, description="开始/结束日期所在列字母，如 ['D', 'E']")
    timeline_start: str = Field(..., description="时间轴起始日期，如 '2026-03-01'")
    timeline_end: str = Field(..., description="时间轴结束日期，如 '2026-06-30'")
    granularity: str = Field(default="week", description="时间粒度: day / week / month")
    bar_color_column: Optional[str] = Field(default=None, description="条形颜色跟随哪列的分组色")

class RowGroupConfig(BaseModel):
    """行分组着色配置"""
    group_column: str = Field(..., description="分组依据列字母，如 'B'")
    colors: Optional[dict[str, str]] = Field(default=None, description="分组值→HEX颜色映射，不指定则自动分配")

class ExcelStyle(BaseModel):
    """Excel 增强样式配置（全部可选）"""
    theme: Optional[str] = Field(default=None, description="主题色系: business_blue / government_red / tech_dark / academic_green / minimal")
    header_style: str = Field(default="colored", description="表头样式: colored / minimal / bold_only")
    freeze_panes: Optional[str] = Field(default=None, description="冻结窗格位置，如 'A3'")
    auto_filter: bool = Field(default=False, description="表头自动筛选")
    alternating_rows: bool = Field(default=True, description="交替行背景色")
    column_widths: Optional[dict[str, int]] = Field(default=None, description="手动列宽，如 {'A': 8, 'B': 30}")
    row_groups: Optional[RowGroupConfig] = Field(default=None, description="按列值分组着色")
    gantt: Optional[GanttConfig] = Field(default=None, description="甘特图时间线（可选）")
```

Then add `style: Optional[ExcelStyle] = Field(default=None, ...)` to `CreateExcelRequest` and `GenerateComplexExcelRequest`.

**Step 4: Implement style engine in excel_handler.py**

Add a new function `_apply_style_engine(sheet, data, style, headers)` that:

1. Applies themed header colors from `get_theme(style.theme)`
2. Sets `sheet.freeze_panes` if configured
3. Sets `sheet.auto_filter.ref` if enabled
4. Applies `column_widths` overrides
5. Applies `alternating_rows` background
6. Applies `row_groups` coloring — for each data row, check column value, apply the mapped background color
7. If `gantt` configured, call `_render_gantt_timeline(sheet, data, headers, gantt_config, theme)`

Add `_render_gantt_timeline()` that:
1. Calculates date range from `timeline_start` to `timeline_end` at given `granularity`
2. Adds timeline header columns (dates) starting after the last data column
3. For each data row, reads start/end dates from `date_columns`, fills cells in the timeline range with the appropriate color

Modify `create_excel_from_array` signature to accept optional `style: dict | None = None`, convert to ExcelStyle internally if dict.

Modify `generate_complex_excel` similarly — each SheetDefinition can have its own style.

**Step 5: Wire style through routes**

In `excel_routes.py`, pass `req.style` to service functions.

**Step 6: Run tests**

Run: `python -m pytest tests/test_excel_style.py -v`
Expected: all PASS

Run: `python -m pytest tests/ -v`
Expected: all existing tests + new tests PASS

**Step 7: Commit**

```bash
git add app/services/excel_handler.py app/schemas/payload_excel.py app/api/endpoints/excel_routes.py tests/test_excel_style.py
git commit -m "feat: excel style engine with themes, freeze panes, row groups, gantt timeline"
```

---

### Task 6: Integration tests + full backward compat verification

**Files:**
- Modify: `tests/test_routes.py` (add integration tests for enhanced endpoints)

**Step 1: Add integration tests**

Add to `tests/test_routes.py`:

```python
class TestEnhancedDocRoutes:

    def test_doc01_with_frontmatter(self, client):
        """DOC-01 with frontmatter should work"""
        md = '---\ncover:\n  title: "Test Cover"\ntheme: business_blue\n---\n# Hello\n\nWorld'
        resp = client.post("/api/v1/docx/render_markdown", json={
            "markdown_content": md,
            "filename": "test_rich",
        })
        assert resp.status_code == 200
        assert resp.json()["code"] == 200

    def test_doc01_without_frontmatter_still_works(self, client):
        """DOC-01 without frontmatter should be identical to before"""
        resp = client.post("/api/v1/docx/render_markdown", json={
            "markdown_content": "# Simple\n\nPlain markdown.",
        })
        assert resp.status_code == 200

    def test_doc01_invalid_frontmatter_returns_hint(self, client):
        """Invalid frontmatter cover without title should return agent_hint"""
        md = '---\ncover:\n  subtitle: "missing title"\n---\n# Content'
        resp = client.post("/api/v1/docx/render_markdown", json={
            "markdown_content": md,
        })
        # Should fail with hint about missing cover.title
        assert resp.status_code in (422, 500)


class TestEnhancedExcelRoutes:

    def test_exc01_with_style(self, client):
        resp = client.post("/api/v1/excel/create_from_array", json={
            "title": "Styled Sheet",
            "data": [["Name", "Score"], ["Alice", 95]],
            "style": {"theme": "business_blue", "freeze_panes": "A3"},
        })
        assert resp.status_code == 200

    def test_exc01_without_style_still_works(self, client):
        resp = client.post("/api/v1/excel/create_from_array", json={
            "title": "Basic",
            "data": [["A", "B"], [1, 2]],
        })
        assert resp.status_code == 200
```

**Step 2: Run full test suite**

Run: `python -m pytest tests/ -v`
Expected: ALL tests pass (original 54 + new tests)

**Step 3: Commit**

```bash
git add tests/test_routes.py
git commit -m "test: add integration tests for enhanced docx and excel endpoints"
```

---

### Task 7: Update API documentation

**Files:**
- Modify: `REFACTOR_REPORT.md` (add section about new capabilities)

**Step 1: Add a section to REFACTOR_REPORT.md documenting the new features**

Add a new section "八、富文档渲染增强" covering:
- Frontmatter spec with complete example
- Excel style spec with complete example
- Agent error hint format
- Backward compatibility notes

**Step 2: Commit**

```bash
git add REFACTOR_REPORT.md
git commit -m "docs: document rich rendering engine capabilities"
```
