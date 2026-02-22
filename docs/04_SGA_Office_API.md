# SGA-Office API 文档

---

## 1. 基础信息
- 基础地址: `http://localhost:5101/api/v1`
- 在线文档: `http://localhost:5101/docs`
- 统一响应结构: `{code, message, data}`
- 异步任务查询: `GET /tasks/{task_id}`

---

## 2. Word / Docx 接口

### DOC-01 渲染 Markdown 为 Docx
- `POST /docx/render_markdown`
- 入参: `markdown_content`, `filename?`
- 返回: `file_url`, `filename`

### DOC-02 模板注水
- `POST /docx/fill_template`
- 入参: `template_url`, `variables`, `filename?`
- 返回: `file_url`, `filename`

### DOC-03 插入图片
- `POST /docx/insert_image`
- 入参: `source_docx_url`, `target_paragraph_keyword`, `image_url`, `image_width_cm?`, `filename?`
- 返回: `file_url`, `filename`

---

## 3. Excel 接口

### EXC-01 二维数组生成 Excel
- `POST /excel/create_from_array`
- 入参: `title`, `data`, `sheet_name?`, `filename?`
- 返回: `file_url`, `filename`

### EXC-02 追加行
- `POST /excel/append_rows`
- 入参: `source_excel_url`, `sheet_name`, `rows`
- 返回: `success`, `rows_appended`, `file_url`

### EXC-03 复杂报表
- `POST /excel/generate_complex`
- 入参: `title`, `sheets`, `filename?`
- 返回: `file_url`, `filename`

### EXC-04 区域提取
- `POST /excel/extract_range`
- 入参: `source_excel_url`, `sheet_name`, `cell_range?`, `keyword?`
- 返回: `sheet_name`, `headers`, `data`, `total_rows`

---

## 4. Visualization 接口

### VIS-01 渲染流程图
- `POST /vis/render_diagram`
- 入参: `code`, `syntax?`, `output_format?`, `theme?`, `width?`, `height?`
- 返回: `task_id`, `status_url`

### VIS-02 渲染统计图
- `POST /vis/render_chart`
- 入参: `chart_type`, `title?`, `categories`, `series`, `output_format?`, `width?`, `height?`, `custom_options?`
- 返回: `task_id`, `status_url`

### VIS-03 导出 Draw.io
- `POST /vis/export_drawio`
- 入参: `nodes`, `edges?`, `layout?`, `export_png?`
- 返回: `task_id`, `status_url`

---

## 5. PDF 接口

### PDF-01 转换为 PDF
- `POST /pdf/convert`
- 入参: `source_url`, `source_type`, `output_filename?`
- 返回: `task_id`, `status_url`

### PDF-02 水印与盖章
- `POST /pdf/add_watermark`
- 入参: `source_pdf_url`, `watermark?`, `stamp?`
- 返回: `task_id`, `status_url`

### PDF-03 合并或截取
- `POST /pdf/merge_split`
- 入参: `source_pdf_urls`, `page_ranges?`, `output_filename?`
- 返回: `task_id`, `status_url`

### PDF-04 扫描解析
- `POST /pdf/extract_text`
- 入参: `source_pdf_url`, `pages?`, `output_format?`
- 返回: `task_id`, `status_url`

---

## 6. 传统接口 (兼容旧调用)

### 传统 Word 生成
- `POST /generate-doc`
- 入参: `title?`, `filename?`, `content` (支持 Markdown)
- 返回: `message`, `file_url`

### 传统 Excel 生成
- `POST /generate_excel`
- 入参: `filename?`, `content` (Markdown 表格)
- 返回: `message`, `file_url`, `filename`

### 传统 PDF 生成
- `POST /generate_pdf`
- 入参: `source_url`, `source_type`, `output_filename?`
- 返回: `message`, `file_url`, `filename`, `page_count`
