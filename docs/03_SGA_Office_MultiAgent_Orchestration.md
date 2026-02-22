# SGA-Office: 多智能体协同作战规范 (Multi-Agent Orchestration)

---

## 1. 单 Agent vs 多 Agent 抉择

我们极力**反对使用单个“上帝 Agent”** 来调用 SGA-Office 的所有接口。

### 为什么不推荐单 Agent？
1. **Tool 爆炸与幻觉**：如果给一个 Agent 注册 Word、Excel、画图、PDF 等几十个 MCP Tool，在面临复杂用户指令（如：“帮我抓取数据分析后画个饼图，最后出一份排版精美的 PDF 报告并盖章”）时，单 Agent 极易不知道该先调哪个接口，或者弄错接口所需的 JSON 结构，导致任务无限期重试并失败。
2. **Context (上下文) 污染**：画图工具需要的时序逻辑，和排版工具需要的 Markdown 结构是两套思维模式。单模型在同一个 Session 里处理容易发生知识混合。

### 结论：采用多 Agent 专家路由架构 (Multi-Agent Routing Architecture)
系统外围应该由一个轻量级的 **Router Agent (任务拆解专家)** 进行统筹，将任务分发给具体的 **Worker Agents (专项能力专家)**。SGA-Office 的 MCP Tools 应该被精确地注册给对应的 Worker Agent。

---

## 2. Agent 角色定义与 Tools 绑定清单

以下是我们定义的 5 种标准 Worker Agent 角色，以及它们与 SGA-Office 微服务接口的具体绑定关系和职责边界。

### 🤖 角色一：Manager Agent (总线/路由智能体)
*   **定位**：系统的“项目经理”与网关。它**不拥有**任何 SGA-Office 的操作 Tools。
*   **核心职责**：
    1.  接收人类自然语言指令（如：“分析上个月的财报，画个利润图，做成红头文件发给我”）。
    2.  将其拆解为子任务 DAG (有向无环图)，例如：`数据提取 -> 图表生成 -> 文本撰写 -> PDF排版`。
    3.  调度后方的专家 Agent 依次拿手里的工具干活，并在各个 Agent 间流转信息（例如把数据 Agent 生成的 COS 云端数据链接，丢给视图 Agent 去画图）。
*   **绑定 SGA-Office Tools**：**无**（只负责调度编排）。

---

### 🤖 角色二：Data & Excel Agent (数据与表格分析员)
*   **定位**：负责所有跟数字、结构化二维表相关的脏活累活。它对 JSON Schema 和逻辑计算极其敏锐。
*   **绑定 SGA-Office Tools**：
    *   `[EXC-01] create_excel_from_array`
    *   `[EXC-02] append_rows_to_excel`
    *   `[EXC-03] generate_complex_excel`
    *   `[EXC-04] extract_excel_named_range`
*   **具体任务流 (Task Execution)**：
    1.  被 Manager 唤醒，接收到原始非结构化数据或第三方 API 爬取结果。
    2.  利用自身强大的计算能力，清洗数据，并映射构造出符合 `generate_complex_excel` 接口所要求的嵌套 JSON。
    3.  调用接口生成云端 Excel。
    4.  向 Manager 返回：“*表格生成完毕，URL 为 `https://cos...`，其中包含关键数据摘要为：营收同比增长 20%*”。

---

### 🤖 角色三：Visualization Agent (灵魂画手/图表专家)
*   **定位**：空间感和逻辑拓扑专家。它专精于写 Mermaid 代码或 ECharts 配置参数。
*   **绑定 SGA-Office Tools**：
    *   `[VIS-01] render_mermaid_to_image`
    *   `[VIS-02] render_echarts_from_data`
    *   `[VIS-03] export_drawio_xml_to_png`
*   **具体任务流 (Task Execution)**：
    1.  收到 Manager 传来的逻辑文本或 Data Agent 处理后的纯数字摘要。
    2.  思考最适合的表达方式（时序逻辑选 Mermaid，数据对比选 ECharts 柱状图）。
    3.  精准吐出渲染配置，调用对应 `VIS` 接口。
    4.  如果是长耗时任务，该 Agent 将进入休眠/轮询模式，直到获取生成的 PNG 的下载链接。
    5.  将图片 URL 返回给 Manager 备用。

---

### 🤖 角色四：Content & Word Agent (文案与排版编辑)
*   **定位**：精通八股文、公文和总结陈词的大语言模型，负责将散落的素材（数字、图片 URL）拼装成优美的 Markdown 并在合适的位置“注水”。
*   **绑定 SGA-Office Tools**：
    *   `[DOC-01] render_markdown_to_docx`
    *   `[DOC-02] fill_docx_template`
    *   `[DOC-03] insert_image_to_docx`
*   **具体任务流 (Task Execution)**：
    1.  从 Manager 处接收用户的核心意图，以及 Data / Vis Agent 产生的数据结论和图表 URL。
    2.  撰写带层级架构的 Markdown 长文，并在正文中插入 `![利润图](https://cos...)`。
    3.  或者根据指令，精准构造出 `{"姓名": "张三", "部门": "销售"}` 的小字典。
    4.  调用 `DOC-01` 或 `DOC-02`，生成初版 Word 定稿，将结果 URL 提交。

---

### 🤖 角色五：Compliance & PDF Agent (合规与归档专员)
*   **定位**：执行系统最后一道防线的严肃型智能体，拥有防篡改与高权限盖章的能力。
*   **绑定 SGA-Office Tools**：
    *   `[PDF-01] convert_to_pdf`
    *   `[PDF-02] add_watermark_and_sign`
    *   `[PDF-03] merge_and_split_pdf`
    *   `[PDF-04] extract_text_from_scanned_pdf`
*   **具体任务流 (Task Execution)**：
    1.  （归档流）收到 Content Agent 的 Word URL。调用 `PDF-01` 转换为不可编辑的 PDF格式。
    2.  （防伪流）读取用户上下文的权限组（如“法务部”），调用 `PDF-02` 在文件四角盖上部门骑缝章，并在全篇打上斜向防伪水印。
    3.  （逆向阅读流）当人类向系统发来一张老旧扫描件时，由该 Agent 首当其冲调用 `PDF-04`，利用服务的底层 OCR 转为机器可读 Markdown，再喂给群里的其他 Agent。

---

## 3. 多 Agent 协作范例 (A Multi-Agent Workflow Example)

**人类指令**：“帮我出一份包含今天系统架构拓扑图的、带有红头文件样式的正式巡检报告，加盖系统部印章，并且需要是不给别人篡改的格式。”

**流转时序 (Orchestration Sequence)**：

1.  人类 -> `Manager Agent`: 拆解出 3 个连续任务。
2.  `Manager Agent` -> `Visualization Agent`: “用 Mermaid 画一个当前后端的微服务调用连拓扑图。”
3.  `Visualization Agent` -> SGA-Office `[VIS-01]`: 传入代码，拿到 `架构图.png` URL。回复 Manager。
4.  `Manager Agent` -> `Content Agent`: “这是图片链接，结合刚才跑测试脚本抓的日志，去那个《红头巡检模板.docx》里填空。”
5.  `Content Agent` -> SGA-Office `[DOC-02]`: 传入变量字典与图片，收到 `初稿_巡检报告.docx` URL。回复 Manager。
6.  `Manager Agent` -> `Compliance Agent`: “拿着这个 docx，去转成 PDF，并且给我盖个印章。”
7.  `Compliance Agent` -> SGA-Office `[PDF-01]` -> `[PDF-02]`: 经历两步管道调用，拿到最终 `带章防伪.pdf` URL。
8.  `Manager Agent` -> 人类：“老板您好，报告已出并锁定盖章。下载地址：https://cos... ”
