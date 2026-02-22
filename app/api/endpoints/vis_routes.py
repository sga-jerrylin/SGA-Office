"""
VIS-X 系列 API 端点（同步模式）。
素材工厂 — 生成可嵌入文档的可视化素材。
- VIS-01: Mermaid → 流程图/时序图/甘特图等
- VIS-02: 数据 → 统计图表 (bar/line/pie 等)
- VIS-03: QR Code / Barcode 生成
- VIS-04: 词云生成
"""

import logging
from fastapi import APIRouter, HTTPException

from app.schemas.base import ApiResponse
from app.schemas.payload_vis import (
    RenderMermaidRequest,
    RenderChartRequest,
    GenerateQRCodeRequest,
    GenerateBarcodeRequest,
    GenerateWordCloudRequest,
)
from app.services.vis_renderer import (
    render_mermaid_to_image,
    render_chart_from_data,
    generate_qrcode,
    generate_barcode,
    generate_wordcloud,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/vis", tags=["Visualization - 素材工厂"])


# =====================================================
#  VIS-01: Mermaid → Image
# =====================================================

@router.post(
    "/render_mermaid",
    summary="[VIS-01] Mermaid 代码渲染为图片",
    description="将 Mermaid DSL 代码渲染为 PNG/SVG 图片。支持 flowchart, sequence, gantt, mindmap 等。",
)
async def vis01_render_mermaid(req: RenderMermaidRequest):
    """Mermaid 代码 → 可视化图片。"""
    try:
        result = render_mermaid_to_image(
            code=req.code,
            output_format=req.output_format.value,
            theme=req.theme or "default",
            width=req.width or 1200,
            height=req.height or 800,
        )
        return ApiResponse(code=200, message="Mermaid 图片渲染成功", data=result)
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("VIS-01 render_mermaid 失败")
        raise HTTPException(status_code=500, detail=f"Mermaid 渲染失败: {str(e)}")


# =====================================================
#  VIS-02: 数据 → 统计图表
# =====================================================

@router.post(
    "/render_chart",
    summary="[VIS-02] 结构化数据生成统计图表",
    description="根据数据和图表类型生成统计图表图片 (bar/line/pie/scatter/radar 等)。",
)
async def vis02_render_chart(req: RenderChartRequest):
    """结构化数据 → 统计图表图片。"""
    try:
        result = render_chart_from_data(
            chart_type=req.chart_type.value,
            categories=req.categories,
            series=req.series,
            title=req.title,
            output_format=req.output_format.value,
            width=req.width or 900,
            height=req.height or 600,
            custom_options=req.custom_options,
        )
        return ApiResponse(code=200, message="图表生成成功", data=result)
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("VIS-02 render_chart 失败")
        raise HTTPException(status_code=500, detail=f"图表生成失败: {str(e)}")


# =====================================================
#  VIS-03a: QR Code 生成
# =====================================================

@router.post(
    "/generate_qrcode",
    summary="[VIS-03a] 生成 QR 二维码",
    description="将文本或 URL 编码为 QR 二维码图片。",
)
async def vis03a_generate_qrcode(req: GenerateQRCodeRequest):
    """文本/URL → QR Code 图片。"""
    try:
        result = generate_qrcode(
            content=req.content,
            size=req.size,
            error_correction=req.error_correction.value,
        )
        return ApiResponse(code=200, message="QR 码生成成功", data=result)
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("VIS-03a generate_qrcode 失败")
        raise HTTPException(status_code=500, detail=f"QR 码生成失败: {str(e)}")


# =====================================================
#  VIS-03b: Barcode 生成
# =====================================================

@router.post(
    "/generate_barcode",
    summary="[VIS-03b] 生成条形码",
    description="生成 Code128/Code39/EAN13 等格式的条形码图片。",
)
async def vis03b_generate_barcode(req: GenerateBarcodeRequest):
    """编码文本 → 条形码图片。"""
    try:
        result = generate_barcode(
            content=req.content,
            barcode_type=req.barcode_type.value,
        )
        return ApiResponse(code=200, message="条形码生成成功", data=result)
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("VIS-03b generate_barcode 失败")
        raise HTTPException(status_code=500, detail=f"条形码生成失败: {str(e)}")


# =====================================================
#  VIS-04: 词云生成
# =====================================================

@router.post(
    "/generate_wordcloud",
    summary="[VIS-04] 文本生成词云图片",
    description="根据输入文本生成词云图片。支持中文自动分词。",
)
async def vis04_generate_wordcloud(req: GenerateWordCloudRequest):
    """文本 → 词云图片。"""
    try:
        result = generate_wordcloud(
            text=req.text,
            width=req.width,
            height=req.height,
            max_words=req.max_words,
            background_color=req.background_color,
            colormap=req.colormap,
            use_jieba=req.use_jieba,
        )
        return ApiResponse(code=200, message="词云生成成功", data=result)
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("VIS-04 generate_wordcloud 失败")
        raise HTTPException(status_code=500, detail=f"词云生成失败: {str(e)}")
