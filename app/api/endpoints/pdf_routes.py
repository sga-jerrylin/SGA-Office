"""
PDF-X 系列 API 端点（同步模式）。
- PDF-01: Word→PDF（LibreOffice headless）
- PDF-02: 水印/盖章
- PDF-03: 合并/拆分
"""

import logging
from fastapi import APIRouter, HTTPException

from app.schemas.base import ApiResponse
from app.schemas.payload_pdf import (
    ConvertDocxToPdfRequest, ConvertDocxToPdfResult,
    AddWatermarkRequest, AddWatermarkResult,
    MergeSplitRequest, MergeSplitResult,
)
from app.services.pdf_manipulator import (
    convert_docx_to_pdf,
    add_watermark_and_sign,
    merge_and_split_pdf,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/pdf", tags=["PDF - Compliance & PDF Agent"])


# =====================================================
#  PDF-01: convert_docx_to_pdf (Word→PDF)
# =====================================================

@router.post(
    "/convert_from_docx",
    response_model=ApiResponse[ConvertDocxToPdfResult],
    summary="[PDF-01] Word 文档转 PDF",
    description="使用 LibreOffice headless 将 .docx 文件高保真转换为 PDF。",
)
async def pdf01_convert_from_docx(req: ConvertDocxToPdfRequest):
    """Word (.docx) → PDF 高保真转换。"""
    try:
        result = convert_docx_to_pdf(
            source_docx_url=str(req.source_docx_url),
            filename=req.filename,
        )
        return ApiResponse(
            code=200,
            message="Word 转 PDF 成功",
            data=ConvertDocxToPdfResult(**result),
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except RuntimeError as e:
        logger.exception("PDF-01 convert_docx_to_pdf 失败")
        raise HTTPException(status_code=500, detail=str(e))
    except Exception as e:
        logger.exception("PDF-01 convert_docx_to_pdf 失败")
        raise HTTPException(status_code=500, detail=f"Word 转 PDF 失败: {str(e)}")


# =====================================================
#  PDF-02: add_watermark_and_sign (水印/盖章)
# =====================================================

@router.post(
    "/add_watermark",
    response_model=ApiResponse[AddWatermarkResult],
    summary="[PDF-02] 文件组密层防伪加印",
    description="为 PDF 添加文字水印和/或图章盖印，不破坏文档正文。",
)
async def pdf02_add_watermark(req: AddWatermarkRequest):
    """同步处理 PDF 水印/盖章。"""
    try:
        result = add_watermark_and_sign(
            source_pdf_url=str(req.source_pdf_url),
            watermark=req.watermark.model_dump() if req.watermark else None,
            stamp=req.stamp.model_dump() if req.stamp else None,
        )
        return ApiResponse(
            code=200,
            message="水印/盖章处理成功",
            data=AddWatermarkResult(**result),
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("PDF-02 add_watermark 失败")
        raise HTTPException(status_code=500, detail=f"PDF 水印/盖章失败: {str(e)}")


# =====================================================
#  PDF-03: merge_and_split_pdf (合并/拆分)
# =====================================================

@router.post(
    "/merge_split",
    response_model=ApiResponse[MergeSplitResult],
    summary="[PDF-03] 物理拓扑层拼拆重组",
    description="合并多个 PDF 文件为一个，或从单一 PDF 中截取指定页码区间。",
)
async def pdf03_merge_split(req: MergeSplitRequest):
    """同步处理 PDF 合并/拆分。"""
    try:
        result = merge_and_split_pdf(
            source_pdf_urls=[str(u) for u in req.source_pdf_urls],
            page_ranges=[pr.model_dump() for pr in req.page_ranges] if req.page_ranges else None,
            output_filename=req.output_filename,
        )
        return ApiResponse(
            code=200,
            message="PDF 合并/截取成功",
            data=MergeSplitResult(**result),
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("PDF-03 merge_split 失败")
        raise HTTPException(status_code=500, detail=f"PDF 合并/截取失败: {str(e)}")
