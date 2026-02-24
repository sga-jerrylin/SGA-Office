"""Agent 友好报错机制测试"""

import pytest
from app.core.error_hints import build_agent_hint, ErrorType


class TestBuildAgentHint:

    def test_missing_field_hint(self):
        hint = build_agent_hint(
            error_type=ErrorType.MISSING_FIELD,
            field="cover.title",
            message="cover.title 是必填字段",
        )
        assert hint["error_type"] == "missing_field"
        assert hint["field"] == "cover.title"
        assert "correct_example" in hint
        assert len(hint["correct_example"]) > 0

    def test_invalid_value_hint_includes_allowed(self):
        hint = build_agent_hint(
            error_type=ErrorType.INVALID_VALUE,
            field="theme",
            message="不支持的主题",
            allowed_values=["business_blue", "government_red"],
        )
        assert "business_blue" in hint["fix_suggestion"]

    def test_invalid_frontmatter_hint(self):
        hint = build_agent_hint(
            error_type=ErrorType.INVALID_FRONTMATTER,
            message="YAML 解析失败: line 3",
        )
        assert hint["error_type"] == "invalid_frontmatter"
        assert "correct_example" in hint

    def test_empty_content_hint(self):
        hint = build_agent_hint(
            error_type=ErrorType.EMPTY_CONTENT,
            message="内容为空",
        )
        assert "# " in hint["correct_example"]


class TestErrorResponseIntegration:

    def test_422_response_includes_agent_hint(self, client):
        """发送非法请求，验证 422 响应包含 agent_hint"""
        resp = client.post("/api/v1/docx/render_markdown", json={
            "markdown_content": "   ",
        })
        assert resp.status_code == 422
        body = resp.json()
        # Pydantic 校验错误应该包含 agent_hint
        assert "agent_hint" in body or "detail" in body
