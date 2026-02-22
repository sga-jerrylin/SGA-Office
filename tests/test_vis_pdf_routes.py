"""
VIS 和 PDF 端点的补充集成测试。
"""

import pytest
from unittest.mock import patch, MagicMock
from io import BytesIO


# =====================================================
#  VIS 端点（续）
# =====================================================

class TestVisRoutesExtra:

    @patch("app.services.vis_renderer.generate_qrcode")
    def test_vis03a_generate_qrcode(self, mock_qr, client):
        """VIS-03a: QR Code 生成"""
        mock_qr.return_value = {
            "file_url": "https://cos.test/qr.png",
            "filename": "qr_20250222_abc.png",
        }
        resp = client.post("/api/v1/vis/generate_qrcode", json={
            "content": "https://www.example.com",
            "size": 10,
            "error_correction": "M",
        })
        assert resp.status_code == 200
        assert resp.json()["code"] == 200

    @patch("app.services.vis_renderer.generate_barcode")
    def test_vis03b_generate_barcode(self, mock_bc, client):
        """VIS-03b: 条形码生成"""
        mock_bc.return_value = {
            "file_url": "https://cos.test/barcode.png",
            "filename": "barcode_20250222_abc.png",
        }
        resp = client.post("/api/v1/vis/generate_barcode", json={
            "content": "1234567890",
            "barcode_type": "code128",
        })
        assert resp.status_code == 200

    @patch("app.services.vis_renderer.generate_wordcloud")
    def test_vis04_generate_wordcloud(self, mock_wc, client):
        """VIS-04: 词云生成"""
        mock_wc.return_value = {
            "file_url": "https://cos.test/wordcloud.png",
            "filename": "wordcloud_20250222_abc.png",
        }
        resp = client.post("/api/v1/vis/generate_wordcloud", json={
            "text": "人工智能 机器学习 深度学习 大模型 数据分析 自然语言处理 计算机视觉",
            "width": 800,
            "height": 600,
        })
        assert resp.status_code == 200

    def test_vis01_invalid_code_too_short(self, client):
        """VIS-01: 代码太短应被 422 拒绝"""
        resp = client.post("/api/v1/vis/render_mermaid", json={
            "code": "ab",
        })
        assert resp.status_code == 422

    def test_vis02_invalid_chart_type(self, client):
        """VIS-02: 不支持的图表类型应被 422 拒绝"""
        resp = client.post("/api/v1/vis/render_chart", json={
            "chart_type": "unknown_type",
            "categories": ["A"],
            "series": [{"name": "X", "values": [1]}],
        })
        assert resp.status_code == 422


# =====================================================
#  PDF 端点
# =====================================================

class TestPdfRoutes:

    @patch("app.api.endpoints.pdf_routes.convert_docx_to_pdf")
    def test_pdf01_convert_from_docx(self, mock_conv, client):
        """PDF-01: Word → PDF"""
        mock_conv.return_value = {
            "file_url": "https://cos.test/output.pdf",
            "filename": "converted_20250222_abc.pdf",
        }
        resp = client.post("/api/v1/pdf/convert_from_docx", json={
            "source_docx_url": "https://cos.example.com/documents/report.docx",
            "filename": "report_final",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["code"] == 200
        assert data["data"]["file_url"].endswith(".pdf")

    @patch("app.api.endpoints.pdf_routes.add_watermark_and_sign")
    def test_pdf02_add_watermark(self, mock_wm, client):
        """PDF-02: 水印/盖章"""
        mock_wm.return_value = {
            "file_url": "https://cos.test/watermarked.pdf",
            "filename": "watermarked_20250222_abc.pdf",
        }
        resp = client.post("/api/v1/pdf/add_watermark", json={
            "source_pdf_url": "https://cos.example.com/pdf/report.pdf",
            "watermark": {
                "text": "机密文件",
                "opacity": 0.15,
                "angle": -45,
                "color": "#808080",
            },
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["code"] == 200

    @patch("app.api.endpoints.pdf_routes.merge_and_split_pdf")
    def test_pdf03_merge(self, mock_merge, client):
        """PDF-03: 合并多个 PDF"""
        mock_merge.return_value = {
            "file_url": "https://cos.test/merged.pdf",
            "filename": "merged_20250222_abc.pdf",
            "page_count": 5,
        }
        resp = client.post("/api/v1/pdf/merge_split", json={
            "source_pdf_urls": [
                "https://cos.example.com/pdf/part1.pdf",
                "https://cos.example.com/pdf/part2.pdf",
            ],
            "output_filename": "combined",
        })
        assert resp.status_code == 200
        data = resp.json()
        assert data["data"]["page_count"] == 5

    @patch("app.api.endpoints.pdf_routes.merge_and_split_pdf")
    def test_pdf03_split(self, mock_split, client):
        """PDF-03: 拆分（页码截取）"""
        mock_split.return_value = {
            "file_url": "https://cos.test/split.pdf",
            "filename": "split_20250222_abc.pdf",
            "page_count": 3,
        }
        resp = client.post("/api/v1/pdf/merge_split", json={
            "source_pdf_urls": ["https://cos.example.com/pdf/full.pdf"],
            "page_ranges": [{"start": 1, "end": 3}],
        })
        assert resp.status_code == 200

    def test_pdf02_invalid_color_rejected(self, client):
        """PDF-02: 非法颜色格式应被 422 拒绝"""
        resp = client.post("/api/v1/pdf/add_watermark", json={
            "source_pdf_url": "https://cos.example.com/pdf/report.pdf",
            "watermark": {
                "text": "test",
                "color": "not-hex",
            },
        })
        assert resp.status_code == 422

