"""
Agent 友好报错机制。
每个错误类型都生成结构化的 agent_hint，包含修正建议和正确示例。
"""

from enum import Enum


class ErrorType(str, Enum):
    INVALID_FRONTMATTER = "invalid_frontmatter"
    MISSING_FIELD = "missing_field"
    INVALID_VALUE = "invalid_value"
    TYPE_ERROR = "type_error"
    EMPTY_CONTENT = "empty_content"
    EXCEL_STYLE_ERROR = "excel_style_error"


# 每种错误类型的正确示例模板
_EXAMPLES: dict[ErrorType, str] = {
    ErrorType.INVALID_FRONTMATTER: (
        "---\n"
        "cover:\n"
        "  title: \"文档标题\"\n"
        "  subtitle: \"副标题\"\n"
        "header: \"页眉文字\"\n"
        "footer: page_number\n"
        "theme: business_blue\n"
        "---\n"
        "# 正文标题\n\n"
        "正文内容"
    ),
    ErrorType.MISSING_FIELD: (
        "---\n"
        "cover:\n"
        "  title: \"你的文档标题\"\n"
        "  subtitle: \"副标题（可选）\"\n"
        "  meta:\n"
        "    - \"作者：XXX\"\n"
        "---\n"
        "# 正文开始"
    ),
    ErrorType.INVALID_VALUE: "请使用以下合法值之一: {allowed_values}",
    ErrorType.TYPE_ERROR: "字段 {field} 期望类型为 {expected_type}，实际收到 {actual_type}",
    ErrorType.EMPTY_CONTENT: (
        "# 文档标题\n\n"
        "这里是正文内容。支持 **加粗**、*斜体*、表格、列表等 Markdown 语法。\n\n"
        "| 列1 | 列2 |\n"
        "|-----|-----|\n"
        "| 数据 | 数据 |"
    ),
    ErrorType.EXCEL_STYLE_ERROR: (
        '{"style": {"theme": "business_blue", "header_style": "colored", '
        '"freeze_panes": "A3", "auto_filter": true}}'
    ),
}


def build_agent_hint(
    error_type: ErrorType,
    message: str = "",
    field: str | None = None,
    allowed_values: list[str] | None = None,
    expected_type: str | None = None,
    actual_type: str | None = None,
) -> dict:
    """
    构建 agent_hint 字典，Agent 读后可自动修正入参。
    """
    hint: dict = {
        "error_type": error_type.value,
    }

    if field:
        hint["field"] = field

    # fix_suggestion
    if error_type == ErrorType.INVALID_VALUE and allowed_values:
        hint["fix_suggestion"] = f"请使用以下合法值之一: {', '.join(allowed_values)}"
    elif error_type == ErrorType.MISSING_FIELD and field:
        hint["fix_suggestion"] = f"请在请求中添加 {field} 字段"
    elif error_type == ErrorType.TYPE_ERROR:
        hint["fix_suggestion"] = f"字段 {field} 期望类型为 {expected_type}，实际收到 {actual_type}"
    else:
        hint["fix_suggestion"] = message

    # correct_example
    example = _EXAMPLES.get(error_type, "")
    if isinstance(example, str) and "{" in example:
        example = example.format(
            field=field or "unknown",
            allowed_values=", ".join(allowed_values) if allowed_values else "",
            expected_type=expected_type or "",
            actual_type=actual_type or "",
        )
    hint["correct_example"] = example

    return hint
