"""主题色系单元测试"""

from app.core.themes import get_theme, AVAILABLE_THEMES


class TestThemes:

    def test_default_theme_is_business_blue(self):
        theme = get_theme()
        assert theme.name == "business_blue"
        assert theme.heading_color == "2E75B6"

    def test_all_builtin_themes_exist(self):
        expected = {"business_blue", "government_red", "tech_dark", "academic_green", "minimal"}
        assert expected == set(AVAILABLE_THEMES.keys())

    def test_get_unknown_theme_falls_back_to_default(self):
        theme = get_theme("nonexistent_theme")
        assert theme.name == "business_blue"

    def test_theme_has_all_required_colors(self):
        for name in AVAILABLE_THEMES:
            theme = get_theme(name)
            assert theme.heading_color
            assert theme.table_header_bg
            assert theme.table_header_font
            assert theme.table_alt_row_bg
            assert theme.callout_info_bg
            assert theme.callout_warning_bg
            assert theme.callout_note_bg
