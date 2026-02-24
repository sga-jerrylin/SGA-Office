"""
SGA-Office 主题色系定义。
Docx 和 Excel 共用同一套主题，保持视觉一致性。
颜色值统一使用 6 位 HEX（不含 #），方便 python-docx 和 openpyxl 直接使用。
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Theme:
    name: str
    # 标题/强调色
    heading_color: str        # 标题文字色
    # 表格
    table_header_bg: str      # 表头背景
    table_header_font: str    # 表头字体色
    table_alt_row_bg: str     # 交替行背景
    # 高亮框
    callout_info_bg: str      # [!INFO] 背景
    callout_info_border: str  # [!INFO] 左边框
    callout_note_bg: str      # [!NOTE] 背景
    callout_note_border: str
    callout_warning_bg: str   # [!WARNING] 背景
    callout_warning_border: str
    # 封面
    cover_title_color: str
    cover_meta_color: str


AVAILABLE_THEMES: dict[str, Theme] = {
    "business_blue": Theme(
        name="business_blue",
        heading_color="2E75B6",
        table_header_bg="2E75B6",
        table_header_font="FFFFFF",
        table_alt_row_bg="D6E4F0",
        callout_info_bg="D6E4F0",
        callout_info_border="2E75B6",
        callout_note_bg="E2EFDA",
        callout_note_border="548235",
        callout_warning_bg="FFF2CC",
        callout_warning_border="BF8F00",
        cover_title_color="2E75B6",
        cover_meta_color="808080",
    ),
    "government_red": Theme(
        name="government_red",
        heading_color="C00000",
        table_header_bg="C00000",
        table_header_font="FFFFFF",
        table_alt_row_bg="FDE9D9",
        callout_info_bg="FDE9D9",
        callout_info_border="C00000",
        callout_note_bg="E2EFDA",
        callout_note_border="548235",
        callout_warning_bg="FFF2CC",
        callout_warning_border="BF8F00",
        cover_title_color="C00000",
        cover_meta_color="808080",
    ),
    "tech_dark": Theme(
        name="tech_dark",
        heading_color="404040",
        table_header_bg="505050",
        table_header_font="FFFFFF",
        table_alt_row_bg="F2F2F2",
        callout_info_bg="F2F2F2",
        callout_info_border="505050",
        callout_note_bg="E8F5E9",
        callout_note_border="4CAF50",
        callout_warning_bg="FFF8E1",
        callout_warning_border="FF8F00",
        cover_title_color="404040",
        cover_meta_color="808080",
    ),
    "academic_green": Theme(
        name="academic_green",
        heading_color="548235",
        table_header_bg="548235",
        table_header_font="FFFFFF",
        table_alt_row_bg="E2EFDA",
        callout_info_bg="D6E4F0",
        callout_info_border="2E75B6",
        callout_note_bg="E2EFDA",
        callout_note_border="548235",
        callout_warning_bg="FFF2CC",
        callout_warning_border="BF8F00",
        cover_title_color="548235",
        cover_meta_color="808080",
    ),
    "minimal": Theme(
        name="minimal",
        heading_color="333333",
        table_header_bg="F2F2F2",
        table_header_font="333333",
        table_alt_row_bg="F9F9F9",
        callout_info_bg="F7F7F7",
        callout_info_border="CCCCCC",
        callout_note_bg="F7F7F7",
        callout_note_border="CCCCCC",
        callout_warning_bg="FFF9F0",
        callout_warning_border="E0A000",
        cover_title_color="333333",
        cover_meta_color="999999",
    ),
}

DEFAULT_THEME = "business_blue"


def get_theme(name: str | None = None) -> Theme:
    """获取主题，未知名称回退到默认主题。"""
    if name is None:
        name = DEFAULT_THEME
    return AVAILABLE_THEMES.get(name, AVAILABLE_THEMES[DEFAULT_THEME])
