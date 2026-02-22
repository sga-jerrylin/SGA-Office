"""
PDF-X 系列接口的请求/响应 Schema 定义。
面向 Compliance & PDF Agent 的 MCP 工具契约。

覆盖:
  - [PDF-01] Word → PDF (LibreOffice headless) — P3 阶段添加
  - [PDF-02] 水印/盖章
  - [PDF-03] 合并/拆分
"""

from typing import Optional
from pydantic import BaseModel, Field, HttpUrl


# ========== PDF-01: Word → PDF (LibreOffice headless) ==========

class ConvertDocxToPdfRequest(BaseModel):
    """
    [PDF-01] convert_docx_to_pdf
    将 Word 文档转换为 PDF。使用 LibreOffice headless 高保真转换。
    典型流程: 先用 DOC-01/02 生成 Word → 用户审阅 → 调用此接口转为 PDF。
    """
    source_docx_url: HttpUrl = Field(
        ...,
        description="Word 文档 (.docx) 的云端可下载链接。"
    )
    filename: Optional[str] = Field(
        default=None,
        max_length=100,
        description="输出 PDF 文件名（不含扩展名）。为空则自动生成。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "source_docx_url": "https://cos.example.com/documents/周报_20250221.docx",
                    "filename": "周报_20250221"
                }
            ]
        }
    }


class ConvertDocxToPdfResult(BaseModel):
    file_url: str = Field(..., description="转换后 PDF 文件的云端 URL")
    filename: str = Field(..., description="实际存储的文件名")


# ========== PDF-02: 文件组密层防伪加印 (Watermark & Stamp) ==========

class WatermarkConfig(BaseModel):
    """文字水印配置"""
    text: str = Field(
        ...,
        min_length=1,
        max_length=50,
        description="水印文字内容。例如: '内部资料 仅供对账' 或 请求人姓名。"
    )
    font_size: int = Field(
        default=40,
        ge=10,
        le=100,
        description="水印文字大小 (pt)"
    )
    opacity: float = Field(
        default=0.15,
        ge=0.01,
        le=0.5,
        description="水印透明度。0.01 为几乎不可见，0.5 为半透明。"
                    "Agent 建议: 合同/内审文件建议 0.1~0.2，正式外发文件建议 0.08。"
    )
    angle: float = Field(
        default=-45.0,
        ge=-90.0,
        le=90.0,
        description="水印文字倾斜角度。-45 为标准的对角线布局。"
    )
    color: str = Field(
        default="#808080",
        pattern=r"^#[0-9A-Fa-f]{6}$",
        description="水印颜色 (HEX 格式)。默认灰色。"
    )


class StampConfig(BaseModel):
    """图章/公章配置"""
    stamp_image_url: HttpUrl = Field(
        ...,
        description="印章图片的云端 URL (推荐透明背景 PNG)。"
    )
    x: float = Field(
        ...,
        ge=0,
        le=595,
        description="印章左上角 X 坐标 (pt)。标准 A4 页面宽度为 595pt。"
                    "Agent 提示: 右下角盖章建议 X=400~450。"
    )
    y: float = Field(
        ...,
        ge=0,
        le=842,
        description="印章左上角 Y 坐标 (pt)。标准 A4 页面高度为 842pt。"
                    "Agent 提示: 底部盖章建议 Y=700~780。"
    )
    width: float = Field(
        default=120,
        ge=20,
        le=300,
        description="印章渲染宽度 (pt)。通常公章直径约 120pt。"
    )
    target_pages: Optional[list[int]] = Field(
        default=None,
        description="指定需要盖章的页码列表 (1-indexed)。为空表示仅盖在最后一页。"
    )


class AddWatermarkRequest(BaseModel):
    """
    [PDF-02] add_watermark_and_sign
    向 PDF 文件添加文字水印和/或图章盖印。不破坏文档正文。
    """
    source_pdf_url: HttpUrl = Field(
        ...,
        description="源 PDF 文件的云端可下载链接。"
    )
    watermark: Optional[WatermarkConfig] = Field(
        default=None,
        description="文字水印配置。为空则不添加水印。"
    )
    stamp: Optional[StampConfig] = Field(
        default=None,
        description="图章/公章盖印配置。为空则不盖章。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "source_pdf_url": "https://cos.example.com/pdf/审计报告.pdf",
                    "watermark": {
                        "text": "内部审计 机密文件",
                        "opacity": 0.12,
                        "angle": -45,
                        "color": "#808080"
                    },
                    "stamp": {
                        "stamp_image_url": "https://cos.example.com/stamps/合同专用章.png",
                        "x": 430,
                        "y": 750,
                        "width": 120
                    }
                }
            ]
        }
    }


class AddWatermarkResult(BaseModel):
    file_url: str = Field(..., description="处理后的 PDF 文件云端下载链接")
    filename: str = Field(..., description="实际存储的文件名")


# ========== PDF-03: 物理拓扑层拼拆重组 (Merge & Split) ==========

class PageRange(BaseModel):
    """页码区间定义"""
    start: int = Field(..., ge=1, description="起始页码 (1-indexed)")
    end: int = Field(..., ge=1, description="结束页码 (1-indexed, 含)")


class MergeSplitRequest(BaseModel):
    """
    [PDF-03] merge_and_split_pdf
    合并多个 PDF 文件为一个，或从单一 PDF 中截取指定页码区间。
    """
    source_pdf_urls: list[HttpUrl] = Field(
        ...,
        min_length=1,
        description="源 PDF 文件 URL 列表。多个文件将按顺序合并。"
    )
    page_ranges: Optional[list[PageRange]] = Field(
        default=None,
        description="页码截取区间。仅在单文件时生效。为空则合并所有页面。"
                    "例如: 仅取第1~10页摘要。"
    )
    output_filename: Optional[str] = Field(
        default=None,
        max_length=100,
        description="输出文件名（不含扩展名）"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "source_pdf_urls": [
                        "https://cos.example.com/pdf/周报_第一周.pdf",
                        "https://cos.example.com/pdf/周报_第二周.pdf"
                    ],
                    "output_filename": "双周报合集"
                }
            ]
        }
    }


class MergeSplitResult(BaseModel):
    """PDF-03 响应数据"""
    file_url: str = Field(..., description="合并/截取后的 PDF 文件云端下载链接")
    filename: str = Field(..., description="实际存储的文件名")
    page_count: int = Field(..., description="最终 PDF 总页数")

