"""
SGA-Office: Agent-First 办公微服务平台
FastAPI ASGI 入口文件 — 注册所有 Router，配置中间件与生命周期。
"""

import logging
from contextlib import asynccontextmanager
from datetime import datetime

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.core.error_hints import build_agent_hint, ErrorType

from app.core.config import get_settings
from app.api.endpoints import excel_routes, doc_routes, vis_routes, pdf_routes, legacy_routes

# ---------- 日志配置 ----------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("sga-office")


# ---------- 生命周期 ----------
@asynccontextmanager
async def lifespan(app: FastAPI):
    settings = get_settings()
    logger.info(f"🚀 {settings.app_name} 启动中...")
    logger.info(f"   COS Region : {settings.cos_region}")
    logger.info(f"   COS Bucket : {settings.cos_bucket_name}")
    logger.info(f"   API Version: {settings.api_version}")
    yield
    logger.info("🛑 SGA-Office 正在关闭...")


# ---------- FastAPI 实例 ----------
settings = get_settings()

app = FastAPI(
    title="SGA-Office: Agent-First 办公微服务平台",
    description=(
        "将办公软件能力封装为面向 AI Agent 自动化调用的原子化 API 工具集。\n\n"
        "## 模块\n"
        "- **Word** (DOC-01~02): Markdown 排版、模板注水\n"
        "- **Excel** (EXC-01~04): 表格创建、追加、复杂报表、区域提取\n"
        "- **VIS** (VIS-01~04): Mermaid 流程图、统计图表、QR/条形码、词云\n"
        "- **PDF** (PDF-01~03): Word→PDF、水印/盖章、合并/拆分\n\n"
        "## 响应格式\n"
        "所有接口统一同步返回 `{code, message, data}` 结构体。"
    ),
    version=settings.api_version,
    lifespan=lifespan,
)

# ---------- CORS 中间件 ----------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------- 全局异常处理 ----------
@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """Pydantic 校验失败时返回 agent_hint"""
    errors = exc.errors()
    first_error = errors[0] if errors else {}
    field_loc = " -> ".join(str(l) for l in first_error.get("loc", []))
    msg = first_error.get("msg", "校验失败")

    hint = build_agent_hint(
        error_type=ErrorType.MISSING_FIELD if "required" in msg.lower() or "missing" in msg.lower()
                   else ErrorType.TYPE_ERROR if "type" in msg.lower()
                   else ErrorType.INVALID_VALUE,
        field=field_loc,
        message=msg,
    )

    return JSONResponse(
        status_code=422,
        content={
            "code": 422,
            "message": f"入参校验失败: {msg} (字段: {field_loc})",
            "agent_hint": hint,
        },
    )


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.exception(f"未捕获异常: {request.method} {request.url}")
    return JSONResponse(
        status_code=500,
        content={
            "code": 500,
            "message": f"服务器内部错误: {str(exc)}",
            "data": None,
        },
    )


# ---------- 注册路由 ----------
API_PREFIX = f"/api/{settings.api_version}"

app.include_router(doc_routes.router, prefix=API_PREFIX)
app.include_router(excel_routes.router, prefix=API_PREFIX)
app.include_router(vis_routes.router, prefix=API_PREFIX)
app.include_router(pdf_routes.router, prefix=API_PREFIX)
app.include_router(legacy_routes.router)


# ---------- 根路由和健康检查 ----------
@app.get("/", tags=["System"])
async def root():
    return {
        "code": 200,
        "message": "SGA-Office API 运行中",
        "data": {
            "service": settings.app_name,
            "version": settings.api_version,
            "docs_url": "/docs",
        },
    }


@app.get("/health", tags=["System"])
async def health_check():
    return {
        "code": 200,
        "message": "healthy",
        "data": {
            "status": "healthy",
            "timestamp": datetime.now().isoformat(),
            "service": settings.app_name,
        },
    }
