"""
VIS-X 系列可视化素材工厂。
生成可嵌入文档的可视化素材（流程图、图表、二维码、词云等）。

- VIS-01: Mermaid → Image (使用 mermaid.ink 在线渲染)
- VIS-02: 数据 → 统计图表 (使用 matplotlib)
- VIS-03: QR Code / Barcode (使用 qrcode + python-barcode)
- VIS-04: 词云 (使用 wordcloud + jieba)
"""

import base64
import logging
from io import BytesIO
from typing import Any, Optional

import requests

from app.services.cos_storage import get_cos_service

logger = logging.getLogger(__name__)


# =====================================================
#  VIS-01: Mermaid → Image
# =====================================================

def render_mermaid_to_image(
    code: str,
    output_format: str = "png",
    theme: str = "default",
    width: int = 1200,
    height: int = 800,
) -> dict[str, str]:
    """
    将 Mermaid DSL 代码渲染为 PNG/SVG 图片。
    使用 mermaid.ink 公共渲染服务。

    Args:
        code: Mermaid 语法代码
        output_format: 输出格式 "png" 或 "svg"
        theme: Mermaid 主题 (default/dark/forest/neutral)
        width: 图片宽度
        height: 图片高度

    Returns:
        dict with file_url and filename
    """
    # 构建 mermaid.ink 渲染参数
    # mermaid.ink 接受 base64 编码的 Mermaid 代码
    encoded = base64.urlsafe_b64encode(code.encode("utf-8")).decode("utf-8")

    if output_format == "svg":
        render_url = f"https://mermaid.ink/svg/{encoded}?theme={theme}"
        ext = "svg"
    else:
        render_url = f"https://mermaid.ink/img/{encoded}?theme={theme}&width={width}&height={height}"
        ext = "png"

    # 请求渲染
    response = requests.get(
        render_url,
        timeout=30,
        headers={"User-Agent": "SGA-Office/1.0"},
    )
    if response.status_code != 200:
        raise ValueError(
            f"Mermaid 渲染失败 (HTTP {response.status_code})。"
            "请检查 Mermaid 语法是否正确。"
        )

    image_bytes = response.content
    if len(image_bytes) < 100:
        raise ValueError("Mermaid 渲染结果异常（内容过小），请检查语法是否正确。")

    # 上传到 COS
    cos = get_cos_service()
    cos_key = cos.generate_cos_key("vis_diagrams", "mermaid", ext)
    file_url = cos.upload_bytes(image_bytes, cos_key)

    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
    }


# =====================================================
#  VIS-02: 数据 → 统计图表 (matplotlib)
# =====================================================

def render_chart_from_data(
    chart_type: str,
    categories: list[str],
    series: list[dict[str, Any]],
    title: str = "",
    output_format: str = "png",
    width: int = 900,
    height: int = 600,
    custom_options: Optional[dict[str, Any]] = None,
) -> dict[str, str]:
    """
    根据结构化数据生成统计图表图片。

    Args:
        chart_type: 图表类型 (bar/line/pie/scatter/radar/heatmap/funnel/gauge)
        categories: X 轴分类标签
        series: 数据系列列表，每个元素 {"name": str, "values": list}
        title: 图表标题
        output_format: "png" 或 "svg"
        width: 宽度（像素）
        height: 高度（像素）
        custom_options: 高级选项（预留）

    Returns:
        dict with file_url and filename
    """
    import matplotlib
    matplotlib.use("Agg")  # 非交互后端
    import matplotlib.pyplot as plt
    import numpy as np

    # 中文字体支持
    plt.rcParams["font.sans-serif"] = ["SimHei", "Noto Sans CJK SC", "WenQuanYi Micro Hei", "Arial Unicode MS"]
    plt.rcParams["axes.unicode_minus"] = False

    fig_w = width / 100
    fig_h = height / 100
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=100)

    _render_chart_type(ax, chart_type, categories, series, title)

    plt.tight_layout()

    # 保存到 BytesIO
    buf = BytesIO()
    ext = "svg" if output_format == "svg" else "png"
    fig.savefig(buf, format=ext, bbox_inches="tight", dpi=150)
    plt.close(fig)
    buf.seek(0)

    # 上传
    cos = get_cos_service()
    cos_key = cos.generate_cos_key("vis_charts", f"chart_{chart_type}", ext)
    file_url = cos.upload_bytes(buf.getvalue(), cos_key)

    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
    }


def _render_chart_type(ax, chart_type: str, categories: list[str],
                       series: list[dict[str, Any]], title: str) -> None:
    """根据图表类型在 axes 上绘制对应的图表。"""
    import numpy as np

    if title:
        ax.set_title(title, fontsize=14, fontweight="bold", pad=12)

    if chart_type == "bar":
        x = np.arange(len(categories))
        bar_width = 0.8 / max(len(series), 1)
        for i, s in enumerate(series):
            offset = (i - len(series) / 2 + 0.5) * bar_width
            ax.bar(x + offset, s["values"], bar_width, label=s.get("name", f"系列{i+1}"))
        ax.set_xticks(x)
        ax.set_xticklabels(categories, rotation=30, ha="right")
        if len(series) > 1:
            ax.legend()

    elif chart_type == "line":
        for s in series:
            ax.plot(categories, s["values"], marker="o", label=s.get("name", ""))
        ax.set_xticks(range(len(categories)))
        ax.set_xticklabels(categories, rotation=30, ha="right")
        if len(series) > 1:
            ax.legend()

    elif chart_type == "pie":
        # 饼图只用第一个系列
        values = series[0]["values"] if series else []
        ax.pie(values, labels=categories, autopct="%1.1f%%", startangle=90)
        ax.axis("equal")

    elif chart_type == "scatter":
        for s in series:
            ax.scatter(categories, s["values"], label=s.get("name", ""), alpha=0.7)
        if len(series) > 1:
            ax.legend()

    elif chart_type == "radar":
        _render_radar(ax, categories, series, title)

    elif chart_type in ("heatmap", "funnel", "gauge"):
        # 回退到柱状图
        x = np.arange(len(categories))
        for s in series:
            ax.bar(x, s["values"], label=s.get("name", ""))
        ax.set_xticks(x)
        ax.set_xticklabels(categories, rotation=30, ha="right")
        if len(series) > 1:
            ax.legend()
        ax.set_xlabel(f"({chart_type} 暂用柱状图展示)")

    else:
        raise ValueError(f"不支持的图表类型: {chart_type}")


def _render_radar(ax, categories: list[str], series: list[dict[str, Any]], title: str) -> None:
    """绘制雷达图（需要 polar axes）。"""
    import numpy as np

    fig = ax.get_figure()
    ax.remove()
    ax = fig.add_subplot(111, polar=True)

    n = len(categories)
    angles = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    angles += angles[:1]

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories)

    for s in series:
        values = s["values"] + s["values"][:1]
        ax.plot(angles, values, "o-", label=s.get("name", ""))
        ax.fill(angles, values, alpha=0.15)

    if title:
        ax.set_title(title, fontsize=14, fontweight="bold", pad=20)
    if len(series) > 1:
        ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.0))



# =====================================================
#  VIS-03: QR Code / Barcode
# =====================================================

def generate_qrcode(
    content: str,
    size: int = 10,
    error_correction: str = "M",
) -> dict[str, str]:
    """
    生成 QR Code 图片。

    Args:
        content: 要编码的文本/URL
        size: QR 码尺寸 (box_size)
        error_correction: 纠错级别 L/M/Q/H

    Returns:
        dict with file_url and filename
    """
    import qrcode

    ec_map = {
        "L": qrcode.constants.ERROR_CORRECT_L,
        "M": qrcode.constants.ERROR_CORRECT_M,
        "Q": qrcode.constants.ERROR_CORRECT_Q,
        "H": qrcode.constants.ERROR_CORRECT_H,
    }

    qr = qrcode.QRCode(
        version=None,
        error_correction=ec_map.get(error_correction.upper(), qrcode.constants.ERROR_CORRECT_M),
        box_size=size,
        border=4,
    )
    qr.add_data(content)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)

    cos = get_cos_service()
    cos_key = cos.generate_cos_key("vis_qrcode", "qrcode", "png")
    file_url = cos.upload_bytes(buf.getvalue(), cos_key)

    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
    }


def generate_barcode(
    content: str,
    barcode_type: str = "code128",
) -> dict[str, str]:
    """
    生成条形码图片。

    Args:
        content: 要编码的文本
        barcode_type: 条形码类型 (code128/code39/ean13/ean8/isbn13/isbn10/upc)

    Returns:
        dict with file_url and filename
    """
    import barcode
    from barcode.writer import ImageWriter

    try:
        bc_class = barcode.get_barcode_class(barcode_type)
    except barcode.errors.BarcodeNotFoundError:
        raise ValueError(f"不支持的条形码类型: {barcode_type}")

    bc = bc_class(content, writer=ImageWriter())
    buf = BytesIO()
    bc.write(buf)
    buf.seek(0)

    cos = get_cos_service()
    cos_key = cos.generate_cos_key("vis_barcode", f"barcode_{barcode_type}", "png")
    file_url = cos.upload_bytes(buf.getvalue(), cos_key)

    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
    }


# =====================================================
#  VIS-04: 词云生成
# =====================================================

def generate_wordcloud(
    text: str,
    width: int = 800,
    height: int = 600,
    max_words: int = 200,
    background_color: str = "white",
    colormap: str = "viridis",
    use_jieba: bool = True,
) -> dict[str, str]:
    """
    生成词云图片。

    Args:
        text: 输入文本（中英文均可）
        width: 图片宽度
        height: 图片高度
        max_words: 最大词数
        background_color: 背景颜色
        colormap: matplotlib 色彩方案
        use_jieba: 是否使用 jieba 分词（中文文本建议开启）

    Returns:
        dict with file_url and filename
    """
    from wordcloud import WordCloud

    # 中文分词
    if use_jieba:
        import jieba
        words = " ".join(jieba.cut(text))
    else:
        words = text

    # 尝试查找中文字体
    font_path = _find_cjk_font()

    wc = WordCloud(
        width=width,
        height=height,
        max_words=max_words,
        background_color=background_color,
        colormap=colormap,
        font_path=font_path,
        margin=10,
    )
    wc.generate(words)

    buf = BytesIO()
    wc.to_image().save(buf, format="PNG")
    buf.seek(0)

    cos = get_cos_service()
    cos_key = cos.generate_cos_key("vis_wordcloud", "wordcloud", "png")
    file_url = cos.upload_bytes(buf.getvalue(), cos_key)

    return {
        "file_url": file_url,
        "filename": cos_key.rsplit("/", 1)[-1],
    }


def _find_cjk_font() -> Optional[str]:
    """查找系统可用的 CJK 字体路径。"""
    import os

    candidates = [
        # Linux
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        # macOS
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        # Windows
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None