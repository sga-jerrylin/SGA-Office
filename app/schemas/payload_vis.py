"""
VIS-X 系列接口的请求/响应 Schema 定义。
素材工厂 — 生成可嵌入文档的可视化素材。

覆盖:
  - [VIS-01] Mermaid → 流程图/时序图/甘特图等
  - [VIS-02] 数据 → 统计图表 (bar/line/pie 等)
  - [VIS-03] QR Code / Barcode 生成 — P2 阶段添加
  - [VIS-04] 词云生成 — P2 阶段添加
"""

from typing import Optional, Any
from enum import Enum
from pydantic import BaseModel, Field


# ========== 图表类型枚举 ==========

class ChartType(str, Enum):
    """支持的统计图表类型"""
    BAR = "bar"
    LINE = "line"
    PIE = "pie"
    SCATTER = "scatter"
    RADAR = "radar"
    HEATMAP = "heatmap"
    FUNNEL = "funnel"
    GAUGE = "gauge"


class ImageFormat(str, Enum):
    """输出图片格式"""
    PNG = "png"
    SVG = "svg"


# ========== VIS-01: Mermaid → Image ==========

class RenderMermaidRequest(BaseModel):
    """
    [VIS-01] render_mermaid_to_image
    将 Mermaid 语法的纯文本描述渲染为可视化图形图片。
    支持: flowchart, sequence, gantt, class, state, er, pie, mindmap, timeline 等。
    """
    code: str = Field(
        ...,
        min_length=5,
        max_length=50000,
        description="Mermaid 语法的图形描述代码。\n"
                    "Agent 注意：请在发送前严格验证语法标签闭合、括号匹配，\n"
                    "以及避免中文半角符号替代英文符号。哪怕差一个字符后端也无法渲染。"
    )
    output_format: ImageFormat = Field(
        default=ImageFormat.PNG,
        description="输出图片格式。PNG 用于嵌入文档，SVG 用于高清缩放场景。"
    )
    theme: Optional[str] = Field(
        default="default",
        description="Mermaid 主题配色: default | dark | forest | neutral"
    )
    width: Optional[int] = Field(
        default=1200,
        ge=200,
        le=4000,
        description="输出图片宽度 (像素)"
    )
    height: Optional[int] = Field(
        default=800,
        ge=200,
        le=4000,
        description="输出图片高度 (像素)"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "code": "graph TD;\n    A[下单] --> B{是否有库存};\n    B -->|是| C[发货];\n    B -->|否| D[缺货通知];",
                    "output_format": "png",
                    "theme": "default"
                }
            ]
        }
    }


class RenderMermaidResult(BaseModel):
    """VIS-01 响应数据"""
    file_url: str = Field(..., description="渲染输出图片的云端 URL")
    filename: str = Field(..., description="实际存储的文件名")


# ========== VIS-02: 数据 → 统计图表 (matplotlib) ==========

class RenderChartRequest(BaseModel):
    """
    [VIS-02] render_chart_from_data
    根据结构化数据源和图表类型，使用 matplotlib 生成统计图表图片。
    """
    chart_type: ChartType = Field(
        ...,
        description="图表类型。Agent 应根据数据特征选择最佳可视化方式："
                    "趋势用 line，占比用 pie，对比用 bar，分布用 scatter。"
    )
    title: str = Field(
        default="",
        max_length=200,
        description="图表标题，将渲染在图表上方。"
    )
    categories: list[str] = Field(
        ...,
        min_length=1,
        description="X 轴分类标签 (或饼图的各扇区名称)。"
    )
    series: list[dict[str, Any]] = Field(
        ...,
        min_length=1,
        description="数据系列。每个元素包含 'name' (系列名称) 和 'values' (数值数组)。"
                    "例如: [{\"name\": \"产品A\", \"values\": [120, 200, 150]}]"
    )
    output_format: ImageFormat = Field(
        default=ImageFormat.PNG,
        description="输出图片格式"
    )
    width: Optional[int] = Field(default=900, ge=300, le=2000, description="图表宽度 (像素)")
    height: Optional[int] = Field(default=600, ge=200, le=1500, description="图表高度 (像素)")
    custom_options: Optional[dict[str, Any]] = Field(
        default=None,
        description="高级选项（预留扩展）。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "chart_type": "bar",
                    "title": "2025年Q1各区域销售业绩",
                    "categories": ["华东", "华南", "华北", "西部"],
                    "series": [
                        {"name": "产品A", "values": [320, 280, 150, 90]},
                        {"name": "产品B", "values": [200, 350, 120, 60]}
                    ],
                    "output_format": "png",
                    "width": 900,
                    "height": 600
                }
            ]
        }
    }


class RenderChartResult(BaseModel):
    file_url: str = Field(..., description="渲染输出图片的云端 URL")
    filename: str = Field(..., description="实际存储的文件名")


# ========== VIS-03: QR Code / Barcode ==========

class ErrorCorrectionLevel(str, Enum):
    """QR Code 纠错级别"""
    L = "L"  # ~7% 纠错
    M = "M"  # ~15% 纠错
    Q = "Q"  # ~25% 纠错
    H = "H"  # ~30% 纠错


class BarcodeType(str, Enum):
    """条形码类型"""
    CODE128 = "code128"
    CODE39 = "code39"
    EAN13 = "ean13"
    EAN8 = "ean8"
    ISBN13 = "isbn13"
    ISBN10 = "isbn10"
    UPC = "upca"


class GenerateQRCodeRequest(BaseModel):
    """
    [VIS-03a] generate_qrcode
    生成 QR 二维码图片。可用于名片、链接分享、设备标识等。
    """
    content: str = Field(
        ...,
        min_length=1,
        max_length=4000,
        description="要编码的文本或 URL。"
    )
    size: int = Field(
        default=10,
        ge=1,
        le=40,
        description="QR 码方块尺寸 (box_size)，越大图片越大。"
    )
    error_correction: ErrorCorrectionLevel = Field(
        default=ErrorCorrectionLevel.M,
        description="纠错级别。嵌入 logo 时建议用 H。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "content": "https://www.example.com",
                    "size": 10,
                    "error_correction": "M"
                }
            ]
        }
    }


class GenerateQRCodeResult(BaseModel):
    file_url: str = Field(..., description="QR Code 图片的云端 URL")
    filename: str = Field(..., description="实际存储的文件名")


class GenerateBarcodeRequest(BaseModel):
    """
    [VIS-03b] generate_barcode
    生成条形码图片。常用于商品编码、快递单号等。
    """
    content: str = Field(
        ...,
        min_length=1,
        max_length=200,
        description="要编码的文本。不同条形码类型对内容有格式要求（如 EAN13 需要 12-13 位数字）。"
    )
    barcode_type: BarcodeType = Field(
        default=BarcodeType.CODE128,
        description="条形码类型。code128 最通用，ean13 用于商品。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "content": "1234567890128",
                    "barcode_type": "ean13"
                }
            ]
        }
    }


class GenerateBarcodeResult(BaseModel):
    file_url: str = Field(..., description="条形码图片的云端 URL")
    filename: str = Field(..., description="实际存储的文件名")


# ========== VIS-04: 词云生成 ==========

class GenerateWordCloudRequest(BaseModel):
    """
    [VIS-04] generate_wordcloud
    根据文本内容生成词云图片。支持中英文自动分词。
    """
    text: str = Field(
        ...,
        min_length=10,
        max_length=100000,
        description="输入文本。中文会自动使用 jieba 分词。"
    )
    width: int = Field(default=800, ge=200, le=2000, description="词云图片宽度 (像素)")
    height: int = Field(default=600, ge=200, le=1500, description="词云图片高度 (像素)")
    max_words: int = Field(default=200, ge=10, le=1000, description="最大显示词数")
    background_color: str = Field(
        default="white",
        description="背景颜色 (CSS 颜色名或 hex 值)"
    )
    colormap: str = Field(
        default="viridis",
        description="matplotlib 色彩方案: viridis/plasma/inferno/magma/cividis/Set2/tab10 等"
    )
    use_jieba: bool = Field(
        default=True,
        description="是否使用 jieba 对中文文本进行分词。纯英文文本可设为 false。"
    )

    model_config = {
        "json_schema_extra": {
            "examples": [
                {
                    "text": "人工智能 机器学习 深度学习 自然语言处理 计算机视觉 大模型 数据分析 云计算 边缘计算 物联网 区块链 量子计算",
                    "width": 800,
                    "height": 600,
                    "max_words": 200,
                    "background_color": "white",
                    "colormap": "viridis"
                }
            ]
        }
    }


class GenerateWordCloudResult(BaseModel):
    file_url: str = Field(..., description="词云图片的云端 URL")
    filename: str = Field(..., description="实际存储的文件名")
