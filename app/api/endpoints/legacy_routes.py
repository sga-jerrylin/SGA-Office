import logging
from typing import Any

from fastapi import APIRouter, HTTPException, Request

import mistune

from app.services.doc_builder import render_markdown_to_docx
from app.services.excel_handler import create_excel_from_array
from app.services.pdf_manipulator import convert_docx_to_pdf
from app.services.cos_storage import get_cos_service

logger = logging.getLogger(__name__)

router = APIRouter(tags=["Legacy - 传统生成接口"])


def _extract_text(node: dict[str, Any]) -> str:
    if "raw" in node:
        return node["raw"]
    if "children" in node:
        return "".join(_extract_text(child) for child in node["children"])
    return ""


def _parse_markdown_table(content: str) -> tuple[list[str], list[list[str]]]:
    markdown = mistune.create_markdown(renderer=None, plugins=["table"])
    ast = markdown(content)
    table_node = None
    for node in ast:
        if node.get("type") == "table":
            table_node = node
            break
    if not table_node:
        return [], []
    headers: list[str] = []
    rows: list[list[str]] = []
    children = table_node.get("children", [])
    thead = next((c for c in children if c.get("type") == "table_head"), None)
    tbody = next((c for c in children if c.get("type") == "table_body"), None)
    if thead:
        for row in thead.get("children", []):
            for cell in row.get("children", []):
                headers.append(_extract_text(cell))
    if tbody:
        for row in tbody.get("children", []):
            current_row = []
            for cell in row.get("children", []):
                current_row.append(_extract_text(cell))
            rows.append(current_row)
    return headers, rows


@router.post("/generate-doc")
async def legacy_generate_doc(request: Request):
    data = await request.json()
    if not data:
        raise HTTPException(status_code=400, detail="没有提供数据")
    if "doc" in data and isinstance(data["doc"], dict):
        data = data["doc"]
    filename_input = str(data.get("filename") or data.get("title") or "默认文档").strip()
    content = str(data.get("content") or data.get("Content") or "")
    if not content or content.strip() == "" or content == "None":
        raise HTTPException(status_code=400, detail="内容不能为空")
    try:
        docx_bytes = render_markdown_to_docx(content)
        cos = get_cos_service()
        cos_key = cos.generate_cos_key("documents", filename_input, "docx")
        file_url = cos.upload_bytes(docx_bytes.getvalue(), cos_key)
        return {"message": "生成成功", "file_url": file_url}
    except Exception as e:
        logger.exception("legacy generate-doc 失败")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/generate_excel")
async def legacy_generate_excel(request: Request):
    data = await request.json()
    if not data:
        raise HTTPException(status_code=400, detail="No JSON data received")
    filename_input = str(data.get("filename") or "未命名表格").strip()
    content = str(data.get("content") or "")
    if not content:
        raise HTTPException(status_code=400, detail="内容不能为空")
    headers, rows = _parse_markdown_table(content.replace("\\n", "\n"))
    if not headers and not rows:
        raise HTTPException(status_code=400, detail="未找到表格内容")
    excel_bytes = create_excel_from_array(
        title=filename_input,
        data=[headers] + rows,
        sheet_name="Sheet1",
    )
    cos = get_cos_service()
    cos_key = cos.generate_cos_key("excel_documents", filename_input, "xlsx")
    file_url = cos.upload_bytes(excel_bytes.getvalue(), cos_key)
    return {"message": "Excel文件生成成功", "file_url": file_url, "filename": cos_key.rsplit("/", 1)[-1]}


@router.post("/generate_pdf")
async def legacy_generate_pdf(request: Request):
    """兼容接口: 接收 Word URL，转为 PDF。"""
    data = await request.json()
    if not data:
        raise HTTPException(status_code=400, detail="没有提供数据")
    docx_url = data.get("docx_url") or data.get("file_url") or ""
    if not docx_url:
        raise HTTPException(status_code=400, detail="缺少 docx_url 或 file_url 参数")
    filename = str(data.get("filename") or "转换文档").strip()
    try:
        result = convert_docx_to_pdf(
            source_docx_url=docx_url,
            filename=filename,
        )
        return {"message": "PDF 生成成功", "file_url": result["file_url"]}
    except Exception as e:
        logger.exception("legacy generate_pdf 失败")
        raise HTTPException(status_code=500, detail=str(e))
