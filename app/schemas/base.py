"""
SGA-Office 统一 API 响应结构定义。
所有端点均为同步即时响应 (code=200)。
"""

from typing import Any, Optional, Generic, TypeVar
from pydantic import BaseModel, Field

T = TypeVar("T")


class SgaResponse(BaseModel):
    """通用统一响应体"""
    code: int = Field(200, description="HTTP 状态码")
    message: str = Field("success", description="操作结果描述")
    data: Optional[Any] = Field(None, description="业务数据体")


class ApiResponse(BaseModel, Generic[T]):
    """泛型统一响应体，data 字段类型由 T 决定"""
    code: int = Field(200, description="HTTP 状态码")
    message: str = Field("success", description="操作结果描述")
    data: Optional[T] = Field(None, description="业务数据体")


class AgentErrorResponse(BaseModel):
    """带 agent_hint 的错误响应，Agent 可据此自动修正入参"""
    code: int = Field(..., description="HTTP 状态码")
    message: str = Field(..., description="人类可读的错误描述")
    agent_hint: Optional[dict] = Field(None, description="Agent 修正提示，含 error_type / fix_suggestion / correct_example")
