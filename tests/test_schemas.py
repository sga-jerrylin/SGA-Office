"""
Pydantic Schema 校验测试。
验证各模块的请求体校验规则是否正确拦截非法输入。
"""

import pytest
from pydantic import ValidationError


# =====================================================
#  DOC schemas
# =====================================================

class TestRenderMarkdownRequest:

    def test_valid_request(self):
        from app.schemas.payload_docx import RenderMarkdownRequest
        req = RenderMarkdownRequest(markdown_content="# Hello\nWorld")
        assert req.markdown_content == "# Hello\nWorld"

    def test_blank_content_rejected(self):
        from app.schemas.payload_docx import RenderMarkdownRequest
        with pytest.raises(ValidationError, match="纯空白"):
            RenderMarkdownRequest(markdown_content="   \n  ")

    def test_empty_string_rejected(self):
        from app.schemas.payload_docx import RenderMarkdownRequest
        with pytest.raises(ValidationError):
            RenderMarkdownRequest(markdown_content="")

    def test_filename_sanitized(self):
        from app.schemas.payload_docx import RenderMarkdownRequest
        req = RenderMarkdownRequest(
            markdown_content="# Test",
            filename='test file/name:bad*"chars',
        )
        # 非法字符被移除
        assert "/" not in (req.filename or "")
        assert ":" not in (req.filename or "")
        assert "*" not in (req.filename or "")


class TestFillTemplateRequest:

    def test_valid_docx_url(self):
        from app.schemas.payload_docx import FillTemplateRequest
        req = FillTemplateRequest(
            template_url="https://cos.example.com/templates/report.docx",
            variables={"name": "test"},
        )
        assert "report.docx" in req.template_url

    def test_non_docx_url_rejected(self):
        from app.schemas.payload_docx import FillTemplateRequest
        with pytest.raises(ValidationError, match="docx"):
            FillTemplateRequest(
                template_url="https://example.com/file.pdf",
                variables={"name": "test"},
            )

    def test_signed_cos_url_accepted(self):
        """带签名参数的 COS URL 应该通过验证（P1 修复的 bug）"""
        from app.schemas.payload_docx import FillTemplateRequest
        req = FillTemplateRequest(
            template_url="https://cos.example.com/templates/report.docx?sign=abc123&expire=9999",
            variables={"name": "test"},
        )
        assert req.template_url.startswith("https://")

    def test_empty_variables_rejected(self):
        from app.schemas.payload_docx import FillTemplateRequest
        with pytest.raises(ValidationError):
            FillTemplateRequest(
                template_url="https://cos.example.com/t.docx",
                variables={},
            )


# =====================================================
#  Excel schemas
# =====================================================

class TestCreateExcelRequest:

    def test_valid_request(self):
        from app.schemas.payload_excel import CreateExcelRequest
        req = CreateExcelRequest(
            title="Test",
            data=[["Name", "Age"], ["Alice", 30]],
        )
        assert len(req.data) == 2

    def test_non_string_header_rejected(self):
        from app.schemas.payload_excel import CreateExcelRequest
        with pytest.raises(ValidationError, match="字符串"):
            CreateExcelRequest(
                title="Test",
                data=[[123, "Age"], ["Alice", 30]],
            )

    def test_single_row_rejected(self):
        """至少需要表头 + 一行数据"""
        from app.schemas.payload_excel import CreateExcelRequest
        with pytest.raises(ValidationError):
            CreateExcelRequest(
                title="Test",
                data=[["Header"]],
            )


# =====================================================
#  VIS schemas
# =====================================================

class TestRenderMermaidRequest:

    def test_valid_mermaid(self):
        from app.schemas.payload_vis import RenderMermaidRequest
        req = RenderMermaidRequest(code="graph TD;\n    A-->B;")
        assert "graph" in req.code

    def test_too_short_code_rejected(self):
        from app.schemas.payload_vis import RenderMermaidRequest
        with pytest.raises(ValidationError):
            RenderMermaidRequest(code="ab")  # min_length=5

    def test_width_range(self):
        from app.schemas.payload_vis import RenderMermaidRequest
        with pytest.raises(ValidationError):
            RenderMermaidRequest(code="graph TD;\n    A-->B;", width=100)  # min 200


class TestGenerateWordCloudRequest:

    def test_valid_request(self):
        from app.schemas.payload_vis import GenerateWordCloudRequest
        req = GenerateWordCloudRequest(text="人工智能 机器学习 深度学习 大模型 数据分析")
        assert req.use_jieba is True  # 默认值

    def test_too_short_text_rejected(self):
        from app.schemas.payload_vis import GenerateWordCloudRequest
        with pytest.raises(ValidationError):
            GenerateWordCloudRequest(text="短")  # min_length=10


# =====================================================
#  PDF schemas
# =====================================================

class TestWatermarkConfig:

    def test_valid_config(self):
        from app.schemas.payload_pdf import WatermarkConfig
        wm = WatermarkConfig(text="机密文件")
        assert wm.opacity == 0.15  # 默认值
        assert wm.angle == -45.0

    def test_invalid_color_rejected(self):
        from app.schemas.payload_pdf import WatermarkConfig
        with pytest.raises(ValidationError):
            WatermarkConfig(text="test", color="not-a-hex")


class TestPageRange:

    def test_valid_range(self):
        from app.schemas.payload_pdf import PageRange
        pr = PageRange(start=1, end=10)
        assert pr.start == 1

    def test_zero_page_rejected(self):
        from app.schemas.payload_pdf import PageRange
        with pytest.raises(ValidationError):
            PageRange(start=0, end=5)  # ge=1

