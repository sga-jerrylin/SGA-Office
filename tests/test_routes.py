"""
API 端点集成测试。
使用 FastAPI TestClient + mock 服务层，验证所有端点的请求/响应格式。
"""

import pytest
from unittest.mock import patch, MagicMock
from io import BytesIO


# =====================================================
#  系统端点
# =====================================================

class TestSystemEndpoints:

    def test_root(self, client):
        resp = client.get("/")
        assert resp.status_code == 200
        data = resp.json()
        assert data["code"] == 200
        assert "SGA-Office" in data["data"]["service"]

    def test_health(self, client):
        resp = client.get("/health")
        assert resp.status_code == 200
        data = resp.json()
        assert data["data"]["status"] == "healthy"


# =====================================================
#  DOC 端点
# =====================================================

class TestDocRoutes:

    def test_doc01_render_markdown(self, client):
        """DOC-01: Markdown → Word"""
        resp = client.post("/api/v1/docx/render_markdown", json={
            "markdown_content": "# Hello World\n\nThis is a test document.",
            "filename": "test_doc",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["code"] == 200
        assert "file_url" in data["data"]
        assert "filename" in data["data"]

    def test_doc01_empty_content_rejected(self, client):
        """DOC-01: 空白内容应被 422 拒绝"""
        resp = client.post("/api/v1/docx/render_markdown", json={
            "markdown_content": "   ",
        })
        assert resp.status_code == 422

    @patch("app.api.endpoints.doc_routes.get_cos_service")
    def test_doc02_fill_template(self, mock_cos, client):
        """DOC-02: 模板填充（需要 mock 下载模板字节）"""
        from docx import Document
        # 创建一个真实的 docx 模板
        doc = Document()
        doc.add_paragraph("Hello {{name}}, your role is {{role}}.")
        buf = BytesIO()
        doc.save(buf)
        template_bytes = buf.getvalue()

        mock_svc = MagicMock()
        mock_svc.download_to_bytes.return_value = template_bytes
        mock_svc.upload_bytes.return_value = "https://cos.test/filled.docx"
        mock_svc.generate_cos_key.return_value = "documents/filled_20250222_abc.docx"
        mock_cos.return_value = mock_svc

        resp = client.post("/api/v1/docx/fill_template", json={
            "template_url": "https://cos.example.com/template.docx",
            "variables": {"name": "Alice", "role": "Engineer"},
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["code"] == 200


# =====================================================
#  Excel 端点
# =====================================================

class TestExcelRoutes:

    def test_exc01_create_from_array(self, client):
        """EXC-01: 二维数组 → Excel"""
        resp = client.post("/api/v1/excel/create_from_array", json={
            "title": "Test Sheet",
            "data": [["Name", "Score"], ["Alice", 95], ["Bob", 87]],
            "filename": "test_excel",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["code"] == 200
        assert "file_url" in data["data"]

    def test_exc01_invalid_headers(self, client):
        """EXC-01: 非字符串表头应被拒绝"""
        resp = client.post("/api/v1/excel/create_from_array", json={
            "title": "Bad",
            "data": [[123, 456], ["a", "b"]],
        })
        assert resp.status_code == 422

    @patch("app.api.endpoints.excel_routes.generate_complex_excel")
    def test_exc03_generate_complex(self, mock_gen, client):
        """EXC-03: 复杂 Excel"""
        mock_gen.return_value = BytesIO(b"PK\x03\x04fake_xlsx")
        resp = client.post("/api/v1/excel/generate_complex", json={
            "title": "Financial Report",
            "sheets": [{
                "sheet_name": "Revenue",
                "headers": ["Month", "Amount"],
                "data": [["Jan", 10000]],
            }],
        })
        assert resp.status_code == 200


# =====================================================
#  VIS 端点
# =====================================================

class TestVisRoutes:

    @patch("app.api.endpoints.vis_routes.render_mermaid_to_image")
    def test_vis01_render_mermaid(self, mock_render, client):
        """VIS-01: Mermaid → 图片"""
        mock_render.return_value = {
            "file_url": "https://cos.test/mermaid.png",
            "filename": "mermaid_20250222_abc.png",
        }
        resp = client.post("/api/v1/vis/render_mermaid", json={
            "code": "graph TD;\n    A-->B;\n    B-->C;",
        })
        assert resp.status_code == 200
        assert resp.json()["data"]["file_url"].endswith(".png")

    @patch("app.api.endpoints.vis_routes.render_chart_from_data")
    def test_vis02_render_chart(self, mock_render, client):
        """VIS-02: 数据 → 统计图表"""
        mock_render.return_value = {
            "file_url": "https://cos.test/chart.png",
            "filename": "chart_20250222_abc.png",
        }
        resp = client.post("/api/v1/vis/render_chart", json={
            "chart_type": "bar",
            "categories": ["Q1", "Q2", "Q3"],
            "series": [{"name": "Revenue", "values": [100, 200, 150]}],
            "title": "Quarterly Revenue",
        })
        assert resp.status_code == 200
        assert resp.json()["code"] == 200

