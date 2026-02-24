"""
EXC-X 系列 API 端点。
Data & Excel Agent 的 MCP 工具绑定入口。
"""

import logging
from datetime import datetime

from fastapi import APIRouter, HTTPException

from app.schemas.base import ApiResponse
from app.schemas.payload_excel import (
    CreateExcelRequest, CreateExcelResult,
    AppendRowsRequest, AppendRowsResult,
    GenerateComplexExcelRequest,
    ExtractExcelRangeRequest, ExtractExcelRangeResult,
)
from app.services.excel_handler import (
    create_excel_from_array,
    append_rows_to_excel,
    generate_complex_excel,
    extract_excel_range,
)
from app.services.cos_storage import get_cos_service

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/excel", tags=["Excel - Data & Excel Agent"])


# =====================================================
#  EXC-01: create_excel_from_array
# =====================================================

@router.post(
    "/create_from_array",
    response_model=ApiResponse[CreateExcelResult],
    summary="[EXC-01] 简单二维归档入库",
    description="将二维数组型 JSON 极速创建为一个基础的 Excel 表格文件。",
)
async def exc01_create_excel(req: CreateExcelRequest):
    """接收二维数组数据，生成 Excel 并上传至 COS，返回下载链接。"""
    try:
        # 1. 调用 service 生成 Excel
        excel_bytes = create_excel_from_array(
            title=req.title,
            data=req.data,
            sheet_name=req.sheet_name or "Sheet1",
            style=req.style.model_dump() if req.style else None,
        )

        # 2. 上传到 COS
        cos = get_cos_service()
        filename = req.filename or "未命名表格"
        cos_key = cos.generate_cos_key("excel_documents", filename, "xlsx")
        file_url = cos.upload_bytes(excel_bytes.getvalue(), cos_key)

        actual_filename = cos_key.rsplit("/", 1)[-1]

        return ApiResponse(
            code=200,
            message="Excel 文件生成成功",
            data=CreateExcelResult(
                file_url=file_url,
                filename=actual_filename,
            ),
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("EXC-01 create_excel_from_array 失败")
        raise HTTPException(status_code=500, detail=f"生成 Excel 失败: {str(e)}")


# =====================================================
#  EXC-02: append_rows_to_excel
# =====================================================

@router.post(
    "/append_rows",
    response_model=ApiResponse[AppendRowsResult],
    summary="[EXC-02] 存量表格增量追加",
    description="向已存在的云端 Excel 文件的指定 Sheet 追加新行（不覆盖原内容）。",
)
async def exc02_append_rows(req: AppendRowsRequest):
    """下载已有 Excel，追加数据行后重新上传。"""
    try:
        # 1. 调用 service 追加行
        updated_bytes = append_rows_to_excel(
            source_excel_url=str(req.source_excel_url),
            rows=req.rows,
            sheet_name=req.sheet_name,
        )

        # 2. 上传更新后的文件
        cos = get_cos_service()
        cos_key = cos.generate_cos_key("excel_documents", "appended", "xlsx")
        file_url = cos.upload_bytes(updated_bytes.getvalue(), cos_key)

        return ApiResponse(
            code=200,
            message=f"成功追加 {len(req.rows)} 行数据",
            data=AppendRowsResult(
                success=True,
                rows_appended=len(req.rows),
                file_url=file_url,
            ),
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("EXC-02 append_rows_to_excel 失败")
        raise HTTPException(status_code=500, detail=f"追加行失败: {str(e)}")


# =====================================================
#  EXC-03: generate_complex_excel
# =====================================================

@router.post(
    "/generate_complex",
    response_model=ApiResponse[CreateExcelResult],
    summary="[EXC-03] 多维报表与公式生成",
    description="生成包含多 Sheet、合并单元格、预埋计算公式的行业级报表。",
)
async def exc03_generate_complex(req: GenerateComplexExcelRequest):
    """根据多 Sheet 结构化定义生成复杂 Excel 报表。"""
    try:
        # 1. 转换为 service 需要的 dict 格式
        sheets_data = [s.model_dump() for s in req.sheets]

        # 2. 调用 service 生成
        excel_bytes = generate_complex_excel(
            title=req.title,
            sheets_def=sheets_data,
            style=req.style.model_dump() if req.style else None,
        )

        # 3. 上传
        cos = get_cos_service()
        filename = req.filename or req.title
        cos_key = cos.generate_cos_key("excel_documents", filename, "xlsx")
        file_url = cos.upload_bytes(excel_bytes.getvalue(), cos_key)
        actual_filename = cos_key.rsplit("/", 1)[-1]

        return ApiResponse(
            code=200,
            message="复杂 Excel 报表生成成功",
            data=CreateExcelResult(
                file_url=file_url,
                filename=actual_filename,
            ),
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("EXC-03 generate_complex_excel 失败")
        raise HTTPException(status_code=500, detail=f"复杂报表生成失败: {str(e)}")


# =====================================================
#  EXC-04: extract_excel_named_range
# =====================================================

@router.post(
    "/extract_range",
    response_model=ApiResponse[ExtractExcelRangeResult],
    summary="[EXC-04] 命名区域精准解析",
    description="从大型 Excel 文件中精准提取局部数据区域，返回结构化 JSON。",
)
async def exc04_extract_range(req: ExtractExcelRangeRequest):
    """从远程 Excel 精准读取指定区域数据。"""
    try:
        result = extract_excel_range(
            source_excel_url=str(req.source_excel_url),
            sheet_name=req.sheet_name,
            cell_range=req.cell_range,
            keyword=req.keyword,
        )

        return ApiResponse(
            code=200,
            message="数据提取成功",
            data=ExtractExcelRangeResult(**result),
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))
    except Exception as e:
        logger.exception("EXC-04 extract_excel_range 失败")
        raise HTTPException(status_code=500, detail=f"数据提取失败: {str(e)}")
