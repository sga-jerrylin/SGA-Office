"""
PDF 操作服务层。
- PDF-01: Word→PDF (LibreOffice headless)
- PDF-02: 水印/盖章
- PDF-03: 合并/拆分
"""

import os
import logging
import subprocess
import tempfile
from io import BytesIO
from typing import Optional, Any

import fitz

from app.services.cos_storage import get_cos_service

logger = logging.getLogger(__name__)


# =====================================================
#  PDF-01: Word → PDF (LibreOffice headless)
# =====================================================

def convert_docx_to_pdf(
    source_docx_url: str,
    filename: Optional[str] = None,
) -> dict[str, str]:
    """
    将 Word 文档转换为 PDF。
    使用 LibreOffice headless 模式，保证格式和中文字体的高保真转换。

    Args:
        source_docx_url: Word 文档的可下载 URL
        filename: 输出文件名（不含扩展名）

    Returns:
        dict with file_url and filename
    """
    cos = get_cos_service()
    docx_bytes = cos.download_to_bytes(source_docx_url)

    with tempfile.TemporaryDirectory() as tmpdir:
        # 写入临时 docx 文件
        docx_path = os.path.join(tmpdir, "input.docx")
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)

        # 调用 LibreOffice headless 转换
        cmd = [
            "soffice",
            "--headless",
            "--norestore",
            "--convert-to", "pdf",
            "--outdir", tmpdir,
            docx_path,
        ]
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
            )
        except FileNotFoundError:
            raise RuntimeError(
                "LibreOffice 未安装或不在 PATH 中。"
                "请确保 Docker 镜像已安装 libreoffice。"
            )
        except subprocess.TimeoutExpired:
            raise RuntimeError("LibreOffice 转换超时（120 秒）。文档可能过大或过于复杂。")

        if result.returncode != 0:
            logger.error("soffice stderr: %s", result.stderr)
            raise RuntimeError(f"LibreOffice 转换失败: {result.stderr[:500]}")

        # 读取输出 PDF
        pdf_path = os.path.join(tmpdir, "input.pdf")
        if not os.path.exists(pdf_path):
            raise RuntimeError(
                f"LibreOffice 转换后未找到 PDF 文件。stdout: {result.stdout[:300]}"
            )

        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

    # 上传到 COS
    output_name = filename or "converted"
    cos_key = cos.generate_cos_key("pdf_documents", output_name, "pdf")
    file_url = cos.upload_bytes(pdf_bytes, cos_key)

    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
    }


def _hex_to_rgb(hex_color: str) -> tuple[float, float, float]:
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16) / 255
    g = int(hex_color[2:4], 16) / 255
    b = int(hex_color[4:6], 16) / 255
    return r, g, b


def add_watermark_and_sign(
    source_pdf_url: str,
    watermark: Optional[dict[str, Any]] = None,
    stamp: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    cos = get_cos_service()
    pdf_bytes = cos.download_to_bytes(source_pdf_url)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    for idx in range(len(doc)):
        page = doc[idx]
        if watermark:
            rect = page.rect
            text = str(watermark.get("text", ""))
            if text:
                color = _hex_to_rgb(watermark.get("color", "#808080"))
                fontsize = float(watermark.get("font_size", 40))
                rotate = float(watermark.get("angle", -45))
                page.insert_textbox(
                    rect,
                    text,
                    fontsize=fontsize,
                    color=color,
                    rotate=rotate,
                    align=1,
                )
    if stamp:
        stamp_bytes = cos.download_to_bytes(str(stamp.get("stamp_image_url")))
        pix = fitz.Pixmap(stream=stamp_bytes)
        x = float(stamp.get("x", 430))
        y = float(stamp.get("y", 750))
        width = float(stamp.get("width", 120))
        pages = stamp.get("target_pages")
        if pages:
            target_pages = [p - 1 for p in pages if isinstance(p, int) and p >= 1]
        else:
            target_pages = [len(doc) - 1] if len(doc) else []
        for p_idx in target_pages:
            if p_idx < 0 or p_idx >= len(doc):
                continue
            page = doc[p_idx]
            rect = fitz.Rect(x, y, x + width, y + width)
            page.insert_image(rect, pixmap=pix)
    output = BytesIO()
    doc.save(output)
    doc.close()
    cos_key = cos.generate_cos_key("pdf_documents", "watermarked", "pdf")
    file_url = cos.upload_bytes(output.getvalue(), cos_key)
    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
    }


def merge_and_split_pdf(
    source_pdf_urls: list[str],
    page_ranges: Optional[list[dict[str, int]]] = None,
    output_filename: Optional[str] = None,
) -> dict[str, Any]:
    cos = get_cos_service()
    out_doc = fitz.open()
    if len(source_pdf_urls) == 1 and page_ranges:
        pdf_bytes = cos.download_to_bytes(source_pdf_urls[0])
        src_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for pr in page_ranges:
            start = max(int(pr.get("start", 1)) - 1, 0)
            end = max(int(pr.get("end", 1)) - 1, 0)
            out_doc.insert_pdf(src_doc, from_page=start, to_page=end)
        src_doc.close()
    else:
        for url in source_pdf_urls:
            pdf_bytes = cos.download_to_bytes(url)
            src_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            out_doc.insert_pdf(src_doc)
            src_doc.close()
    output = BytesIO()
    out_doc.save(output)
    page_count = out_doc.page_count
    out_doc.close()
    filename = output_filename or "merged_pdf"
    cos_key = cos.generate_cos_key("pdf_documents", filename, "pdf")
    file_url = cos.upload_bytes(output.getvalue(), cos_key)
    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
        "page_count": page_count,
    }

