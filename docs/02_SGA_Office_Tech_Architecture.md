# SGA-Office: Agent-First 办公微服务 - 技术架构设计 (Tech Spec)

---

## 1. 架构目标与约束

基于 PRD 的 `Agent-First` 原则，本系统的架构设计必须满足以下三个核心非功能性需求 (NFR)：

1.  **极度稳定与鲁棒性**：大模型生成的数据经常存在“幻觉”（格式缺失、类型错误）。API 必须能在最外壳层挡住所有脏数据，绝不能让核心渲染进程因为一个标点符号崩溃。
2.  **异步防阻塞**：图形渲染和文档转码属于 CPU 密集型任务，极易导致网关超时。必须采用异步任务队列机制，将 I/O 与计算分离。
3.  **MCP (模型上下文协议) 完美兼容**：必须能极低成本地直接转换为大模型的工具集 (Tools)。这意味着接口出入参必须具备严格的强类型和丰富的 Schema 语义描述。

---

## 2. 核心技术栈选型

*   **API 接入层与网关**: Python **FastAPI**
    *   *选型理由*: Pydantic 提供的强类型校验和自动生成的 OpenAPI Schema 是最佳的 MCP 适配器。利用其原生的 `async/await` 处理高发兵的轻量级请求（如上传/下载链接）。
*   **任务编排与执行引擎**: Python **Celery** + **Redis** (Broker & Backend)
    *   *选型理由*: 解耦长耗时的渲染任务（如 `VIS-02` 的 Headless 画图，`PDF-01` 的格式转换）。网关负责接单收发 `task_id`，Worker 集群负责在后台重负载死磕。
*   **持久化与文件中转**: 腾讯云 **COS (Cloud Object Storage)**
    *   *选型理由*: 彻底剥离系统对宿主机本地文件系统的依赖。这使得 Web 节点和 Worker 节点可以无状态地自由横向扩展 (Stateless horizontally scalable)。
*   **底层文档处理套件**:
    *   **Word**: `python-docx` + `docxtpl` (处理模板注水 `DOC-02`)
    *   **Excel**: `openpyxl`
    *   **图形渲染**: `pyppeteer` 或 `playwright` (无头浏览器，用于生成 `VIS-02` 复杂 ECharts 或 Mermaid)
    *   **PDF**: `WeasyPrint` (HTML直转PDF)、`PyMuPDF` (处理拆合/防伪/文本解析)

---

## 3. 系统架构拓扑 (Architecture Topology)

系统分为三个独立的物理/逻辑群集：

1.  **API Gateway (FastAPI)**：暴露 RESTful 接口，负责鉴权、入参 Schema 极严校验、拦截非法请求。校验通过后，将标准 JSON Payload 序列化并压入 Redis 队列。
2.  **Task Queue (Redis)**：作为消息中间件缓存待办队列，并存储任务最终的执行状态 (Pending, Success, Failed) 及其结果 URL。
3.  **Worker Cluster (Celery)**：装备了所有的底层依赖包及虚拟环境（包含字体、浏览器驱动等胖资源）。它们以 Pull 的方式从队列监听任务，下载源文件到内存 -> 密集计算渲染组装 -> 将产出物上传直传云端 (COS) -> 向 Redis 报告完结并附上长效 URL。

---

## 4. MCP Agent 交互协议设计模式 (The Contract)

为了让 Agent 不出错地使用这些接口，我们必须制定全系统统一的交互模式。

### 模式 A：即时响应型 (短时 I/O, 如读取轻量 Excel 或查询 Task 状态)
*   网关接客，立即返回数据体。必须返回 JSON。
*   **标准响应结构**：
    ```json
    {
      "code": 200,
      "message": "success",
      "data": { ...具体的业务结构... }
    }
    ```

### 模式 B：长轮询型 (202 Accepted, Rendering Tasks)
*   **Step 1. 提交任务**:
    Agent Post payload 给如 `/api/v1/docx/render_template`。
    **网关立刻返回 HTTP 202**:
    ```json
    {
      "code": 202,
      "message": "Task submitted to background worker.",
      "data": {
        "task_id": "9b1deb4d-3b7d-4bad-9bdd-2b0d7b3dcb6d",
        "status_url": "/api/v1/tasks/9b1deb4d-3b7d-4bad-9bdd-2b0d7b3dcb6d"
      }
    }
    ```
*   **Step 2. Agent 轮询**:
    Agent 携带 task_id 根据业务逻辑 `sleep` 几秒后访问 `status_url`。
    如果完成，返回 HTTP 200：
    ```json
    {
      "code": 200,
      "data": {
        "status": "SUCCESS",
        "result_url": "https://sga-cos.xxx.myqcloud.com/out/final_doc.docx"
      }
    }
    ```
    如果错误，返回明确的错误原因 (让大模型自我检查逻辑是否出错)。

---

## 5. 目录结构规划 (Modular Breakdown)

对现有的单体 `main.py` 执行手术，拆解为以下微服务包结构：

```text
sga-office/
├── docs/                           # 包含已有的 PRD 和此架构设计
├── app/
│   ├── main.py                     # FastAPI ASGI 入口文件，注册所有 Router
│   ├── core/                       # 核心基建
│   │   ├── config.py               # Pydantic BaseSettings 环境变量管理 (包含COS凭证)
│   │   └── celery_app.py           # Celery 实例编排连接池
│   ├── api/                        # 网关路由层 (接单、参数拦截)
│   │   ├── dependencies.py         # 共享依赖注入 (如安全效验)
│   │   ├── endpoints/
│   │   │   ├── doc_routes.py       # DOC-X 系列 API 端点
│   │   │   ├── excel_routes.py     # EXC-X 系列 API 端点
│   │   │   ├── vis_routes.py       # VIS-X 系列 API 端点
│   │   │   ├── pdf_routes.py       # PDF-X 系列 API 端点
│   │   │   └── task_routes.py      # 通用统一任务轮询与状态机端点
│   ├── schemas/                    # 重中之重: 面向 MCP 的接口强校验层
│   │   ├── base.py                 # 基础统一返回字典的结构定义
│   │   ├── payload_docx.py         # 例如定义的 TemplateFillRequest
│   │   └── payload_excel.py
│   ├── services/                   # 面向对象思想解耦的核心业务层算法包
│   │   ├── doc_builder.py          # 原代码 Markdown->Docx 提取放此
│   │   ├── excel_handler.py
│   │   ├── chart_renderer.py
│   │   └── pdf_manipulator.py
│   └── worker/                     # Celery Tasks 定义 (胶水层)
│       └── background_tasks.py     # 把 Schema 参数拆包传入 services 跑重活
├── .env.example
├── docker-compose.yml              # 定义 Redis, API Web, Celery Worker 容器群
├── Dockerfile.web                  # 瘦镜像: 只装 FastAPI 等基础包
└── Dockerfile.worker               # 胖镜像: 需安装 libpango, pyppeteer 及无头驱动
```

---

## 6. 数据安全与容错 (Pydantic 的极致运用)

为了防止大模型幻觉，开发时需要在 `app/schemas/` 给出极其严格、带有正则表达式校验的模型定义。

**开发范例 (Excel 表格增量 payload 校验)**:
```python
from pydantic import BaseModel, Field, HttpUrl

class AppendRowRequest(BaseModel):
    source_excel_url: HttpUrl = Field(..., description="必须是一个合法的以 xlsx 结尾的可下载外链。")
    sheet_name: str = Field(..., max_length=31, description="目标 Sheet 页签的名字。")
    row_data: list[str | int | float] = Field(
        ...,
        min_items=1,
        description="一维数组。传入你想要追加的这一行各列的具体值。Agent 注意不要传成嵌套数组。"
    )
```
*说明：通过这段代码，FastAPI 会自动拒绝所有试图传两维数组或提供本地非法路径的 Agent 调用，并返回 `422 Unprocessable Entity` 和精准的出错列提示。大模型根据错误信息会自动启动重试策略 (Reflection) 修正参数。这就是整个系统抗震防线的本质。*
