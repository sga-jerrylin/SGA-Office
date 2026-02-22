"""
DOC-X 系列接口的 Pydantic 请求体定义 (面向 MCP Agent 的超严格校验层)。
覆盖:
  - [DOC-01] render_markdown_to_docx
  - [DOC-02] fill_docx_template

设计要点:
  1. 每个字段的 description 必须精确到大模型一看就懂要怎么填。
  2. 利用 Pydantic 的 regex/max_length 等机制，拦截大模型幻觉产出的脏数据。
"""

from typing import Optional, Any
from urllib.parse import urlparse
from pydantic import BaseModel, Field, field_validator
import re


class RenderMarkdownRequest(BaseModel):
    """
    [DOC-01] Markdown 结构化文本极速排版请求体。
    Agent 将一整段 Markdown 文本（可含表格、图片链接）发给系统，
    系统自动渲染为带有标准标题层级和自动编号的 .docx 文件。
    """
    markdown_content: str = Field(
        ...,
        min_length=1,
        max_length=500000,
        description=(
            "需要排版的 Markdown 原文。支持标准 Markdown 语法，包括：\n"
            "- 标题 (# ~ ######)\n"
            "- 加粗/斜体\n"
            "- 有序/无序列表\n"
            "- 代码块\n"
            "- 表格 (Markdown 表格语法)\n"
            "- 图片 (![alt](url) — 图片 URL 必须为公网可达的 HTTP/HTTPS 链接)\n"
            "注意: 不要传递空字符串或纯空白内容。"
        ),
    )
    filename: Optional[str] = Field(
        default=None,
        max_length=100,
        description=(
            "期望的输出文件名 (不含扩展名)。\n"
            "例如: '年终总结报告' 或 'Q4营收分析'。\n"
            "如果为空，系统会自动根据 Markdown 首行标题提取。"
        ),
    )

    @field_validator("markdown_content")
    @classmethod
    def content_must_not_be_blank(cls, v: str) -> str:
        if not v.strip():
            raise ValueError("markdown_content 不能为纯空白内容")
        return v

    @field_validator("filename")
    @classmethod
    def sanitize_filename(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        # 移除文件名中的非法字符
        cleaned = re.sub(r'[\\/:*?"<>|\s]', '', v.strip())
        if not cleaned:
            return None
        return cleaned[:60]  # 限制长度


class FillTemplateRequest(BaseModel):
    """
    [DOC-02] 占位符模板动态注水请求体。
    基于模板文件中的 {{ 变量名 }} 占位符进行精确替换。
    """
    template_url: str = Field(
        ...,
        description=(
            "云端 Word 模板文件的可下载 URL。\n"
            "模板中使用 Jinja2 语法: {{ 变量名 }}。\n"
            "例如: 'https://cos.xxx.com/templates/录用通知书.docx'\n"
            "URL 必须以 .docx 结尾且为公网可达。"
        ),
    )
    variables: dict[str, Any] = Field(
        ...,
        min_length=1,
        description=(
            "Key-Value 变量字典。Key 对应模板中的占位符名称。\n"
            "示例: {\"姓名\": \"张三\", \"部门\": \"技术部\", \"入职日期\": \"2025-03-01\"}\n"
            "支持的 Value 类型:\n"
            "  - 字符串: \"张三\"\n"
            "  - 数字: 40000\n"
            "  - 列表 (用于模板中的循环渲染): [{\"条目\": \"费用1\", \"金额\": 500}]\n"
            "注意: Key 不要传入含 {{ }} 花括号的字段名，只需传入纯名称即可。"
        ),
    )
    filename: Optional[str] = Field(
        default=None,
        max_length=100,
        description=(
            "期望的输出文件名 (不含扩展名)。\n"
            "如果为空，系统会在模板文件名后拼接 '_filled' 后缀。"
        ),
    )

    @field_validator("template_url")
    @classmethod
    def validate_template_url(cls, v: str) -> str:
        v = v.strip()
        if not v.lower().startswith(("http://", "https://")):
            raise ValueError("template_url 必须是 http:// 或 https:// 开头的合法 URL")
        # 用 urlparse 提取 path 部分，忽略 query params（如 COS 签名参数）
        parsed_path = urlparse(v).path.lower()
        if not parsed_path.endswith(".docx"):
            raise ValueError(
                "template_url 必须指向 .docx 文件。"
                "请确认链接路径为 .docx 扩展名（URL query 参数不影响判断）。"
            )
        return v

    @field_validator("filename")
    @classmethod
    def sanitize_filename(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        cleaned = re.sub(r'[\\/:*?"<>|\s]', '', v.strip())
        if not cleaned:
            return None
        return cleaned[:60]



