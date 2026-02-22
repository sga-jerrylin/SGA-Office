"""
EXC-X 系列接口的请求/响应 Schema 定义。
面向 Data & Excel Agent 的 MCP 工具契约。
"""

from typing import Optional, Any, Union
from pydantic import BaseModel, Field, HttpUrl, field_validator


# ========== EXC-01: 简单二维归档入库 (2D Array → Excel) ==========

class CreateExcelRequest(BaseModel):
    """
    [EXC-01] create_excel_from_array
    将二维数组型 JSON 极速创建为一个基础的 Excel 表格文件。
    """
    title: str = Field(
        ...,
        min_length=1,
        max_length=200,
        description="表格标题，将显示在 Excel 的第一行合并单元格中。"
    )
    data: list[list[Union[str, int, float, None]]] = Field(
        ...,
        min_length=2,
        description="二维数组。第一层是行，第二层是列。"
                    "data[0] 必须为表头行（字符串数组），后续行为数据内容。"
                    "Agent 注意：严禁传入三维以上嵌套数组！每行的列数应与表头一致。"
    )
    sheet_name: Optional[str] = Field(
        default="Sheet1",
        max_length=31,
        description="Sheet 页签名称。Excel 限制最长 31 个字符。"
    )
    filename: Optional[str] = Field(
        default="未命名表格",
        max_length=100,
        description="生成的文件名（不含扩展名）。"
    )

    @field_validator("data")
    @classmethod
    def validate_data_structure(cls, v):
        """确保表头行存在且都是字符串"""
        if not v or not v[0]:
            raise ValueError("data 数组不能为空，且第一行（表头）不能为空行。")
        # 验证表头行全部为字符串
        for i, header in enumerate(v[0]):
            if not isinstance(header, str):
                raise ValueError(
                    f"data[0][{i}] 必须是字符串（表头），但收到了 {type(header).__name__}。"
                    f"请确保二维数组的第一行全部为字符串类型的列标题。"
                )
        return v

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "title": "2025年Q1新闻采集汇总",
                    "data": [
                        ["序号", "新闻标题", "来源", "发布日期"],
                        [1, "AI 技术突破", "新华社", "2025-01-15"],
                        [2, "经济数据好转", "财经日报", "2025-01-20"]
                    ],
                    "filename": "Q1新闻汇总"
                }
            ]
        }
    }


class CreateExcelResult(BaseModel):
    """EXC-01 响应数据"""
    file_url: str = Field(..., description="生成的 Excel 文件云端下载链接")
    filename: str = Field(..., description="实际存储的文件名")


# ========== EXC-02: 存量表格增量追加 (Append Rows) ==========

class AppendRowsRequest(BaseModel):
    """
    [EXC-02] append_rows_to_excel
    向已存在的云端 Excel 文件的指定 Sheet 追加新行数据。
    不会覆盖原有内容。
    """
    source_excel_url: HttpUrl = Field(
        ...,
        description="已存在的 .xlsx 文件的云端可下载链接。"
    )
    sheet_name: str = Field(
        default="Sheet1",
        max_length=31,
        description="目标 Sheet 页签名称。"
    )
    rows: list[list[Union[str, int, float, None]]] = Field(
        ...,
        min_length=1,
        description="要追加的行数据。每个子数组代表一行。"
                    "Agent 注意：请确保每行的列数与目标 Sheet 的表头列数一致。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "source_excel_url": "https://cos.example.com/excel/漏洞年度总表.xlsx",
                    "sheet_name": "Sheet1",
                    "rows": [
                        ["192.168.1.10", "高", "SQL注入", "2025-01-22"],
                        ["192.168.1.15", "中", "XSS跨站", "2025-01-22"]
                    ]
                }
            ]
        }
    }


class AppendRowsResult(BaseModel):
    """EXC-02 响应数据"""
    success: bool = Field(..., description="操作是否成功")
    rows_appended: int = Field(..., description="实际追加的行数")
    file_url: str = Field(..., description="更新后的文件云端下载链接")


# ========== EXC-03: 多维报表与公式生成 (Complex Excel) ==========

class SheetDefinition(BaseModel):
    """单个 Sheet 的定义"""
    sheet_name: str = Field(
        ...,
        max_length=31,
        description="Sheet 页签名称"
    )
    headers: list[str] = Field(
        ...,
        min_length=1,
        description="表头列名数组"
    )
    data: list[list[Union[str, int, float, None]]] = Field(
        ...,
        description="数据行（不含表头）。支持直接写入 Excel 公式字符串，"
                    "如 '=SUM(B2:B10)' 或 '=AVERAGE(C2:C10)'。"
    )
    merge_cells: Optional[list[str]] = Field(
        default=None,
        description="需要合并的单元格区间列表。格式: ['A1:D1', 'E5:E8']"
    )


class GenerateComplexExcelRequest(BaseModel):
    """
    [EXC-03] generate_complex_excel
    生成包含多 Sheet、合并单元格、预埋计算公式的行业级报表。
    """
    title: str = Field(
        ...,
        min_length=1,
        max_length=200,
        description="报表总标题"
    )
    sheets: list[SheetDefinition] = Field(
        ...,
        min_length=1,
        description="Sheet 定义列表。每个元素描述一个独立的 Sheet 页。"
    )
    filename: Optional[str] = Field(
        default=None,
        max_length=100,
        description="生成文件名（不含扩展名）"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "title": "2025年度财务月度结算单",
                    "sheets": [
                        {
                            "sheet_name": "收入明细",
                            "headers": ["月份", "产品线A", "产品线B", "合计"],
                            "data": [
                                ["1月", 50000, 30000, "=B2+C2"],
                                ["2月", 55000, 32000, "=B3+C3"],
                                ["合计", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)"]
                            ]
                        }
                    ],
                    "filename": "2025年度结算单"
                }
            ]
        }
    }


# ========== EXC-04: 命名区域精准解析 (Selective Read) ==========

class ExtractExcelRangeRequest(BaseModel):
    """
    [EXC-04] extract_excel_named_range
    从大型 Excel 文件中精准提取局部数据区域。
    """
    source_excel_url: HttpUrl = Field(
        ...,
        description="云端 .xlsx 文件的可下载链接。"
    )
    sheet_name: str = Field(
        default="Sheet1",
        max_length=31,
        description="目标 Sheet 页签名称。"
    )
    cell_range: Optional[str] = Field(
        default=None,
        pattern=r"^[A-Z]{1,3}\d+:[A-Z]{1,3}\d+$",
        description="要提取的单元格区间。例如 'A1:D10'。"
                    "格式必须为 '列字母+行号:列字母+行号'。"
    )
    keyword: Optional[str] = Field(
        default=None,
        max_length=200,
        description="内容检索关键词。系统将定位包含此关键词的区域并返回上下文。"
                    "与 cell_range 二选一使用。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "source_excel_url": "https://cos.example.com/excel/财报总表.xlsx",
                    "sheet_name": "Sheet2",
                    "cell_range": "A1:E20"
                }
            ]
        }
    }


class ExtractExcelRangeResult(BaseModel):
    """EXC-04 响应数据"""
    sheet_name: str = Field(..., description="实际读取的 Sheet 名称")
    headers: list[str] = Field(..., description="提取区域的表头")
    data: list[list[Any]] = Field(..., description="提取的扁平化数据行")
    total_rows: int = Field(..., description="提取的总行数")
