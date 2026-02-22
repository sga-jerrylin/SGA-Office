"""
DOC-X 系列 API 端点。
Content & Word Agent 的 MCP 工具绑定入口。
"""

import logging
from fastapi import APIRouter, HTTPException

from app.schemas.base import ApiResponse
from app.schemas.payload_docx import (
    RenderMarkdownRequest,
    FillTemplateRequest,
)
from app.services.doc_builder import render_markdown_to_docx, fill_docx_template
from app.services.cos_storage import get_cos_service

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/docx", tags=["Word - Content & Word Agent"])


# =====================================================
#  DOC-01: render_markdown_to_docx
# =====================================================

class DocResult(ApiResponse):
    """DOC 系列通用文件结果"""
    pass


@router.post(
    "/render_markdown",
    summary="[DOC-01] Markdown 结构化文本极速排版",
    description="将 Markdown 文本渲染为带标准标题层级和排版的正式 .docx 文件。",
)
async def doc01_render_markdown(req: RenderMarkdownRequest):
    """Markdown → Docx 全自动排版渲染。"""
    try:
        # 1. 调用 service 渲染
        docx_bytes = render_markdown_to_docx(req.markdown_content)

        # 2. 上传到 COS
        cos = get_cos_service()
        filename = req.filename or "未命名文档"
        cos_key = cos.generate_cos_key("documents", filename, "docx")
        file_url = cos.upload_bytes(docx_bytes.getvalue(), cos_key)
        actual_filename = cos_key.rsplit("/", 1)[-1]

        return ApiResponse(
            code=200,
            message="Word 文档生成成功",
            data={
                "file_url": file_url,
                "filename": actual_filename,
            },
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("DOC-01 render_markdown_to_docx 失败")
        raise HTTPException(status_code=500, detail=f"文档生成失败: {str(e)}")


# =====================================================
#  DOC-02: fill_docx_template (占位符模板注水)
# =====================================================

@router.post(
    "/fill_template",
    summary="[DOC-02] 占位符模板动态注水",
    description="基于 Word 模板文件中的 {{ 占位符 }}，进行变量数据的精确替换。",
)
async def doc02_fill_template(req: FillTemplateRequest):
    """下载模板 → 替换变量 → 上传结果。"""
    try:
        cos = get_cos_service()

        # 1. 下载模板文件
        template_bytes = cos.download_to_bytes(str(req.template_url))

        # 2. 调用 service 替换占位符
        output = fill_docx_template(template_bytes, req.variables)

        # 3. 上传结果
        filename = req.filename or "模板填充文档"
        cos_key = cos.generate_cos_key("documents", filename, "docx")
        file_url = cos.upload_bytes(output.getvalue(), cos_key)
        actual_filename = cos_key.rsplit("/", 1)[-1]

        return ApiResponse(
            code=200,
            message="模板填充成功",
            data={
                "file_url": file_url,
                "filename": actual_filename,
            },
        )
    except Exception as e:
        logger.exception("DOC-02 fill_docx_template 失败")
        raise HTTPException(status_code=500, detail=f"模板填充失败: {str(e)}")

