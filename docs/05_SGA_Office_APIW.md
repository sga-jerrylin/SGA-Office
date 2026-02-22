# SGA-Office APIW 文档 (API Workflow)

---

## 1. APIW 定义
APIW 用于说明异步任务型 API 的统一调用流程，适用于 VIS 与 PDF 等长耗时能力。

---

## 2. 标准流程
1. 提交任务，服务返回 `202` 与 `task_id`
2. 轮询 `GET /api/v1/tasks/{task_id}`
3. `status=SUCCESS` 时读取 `result_url` 或 `result_data`

---

## 3. 请求与响应规范

### 3.1 提交任务响应
```json
{
  "code": 202,
  "message": "Task submitted to background worker.",
  "data": {
    "task_id": "uuid",
    "status_url": "/api/v1/tasks/uuid"
  }
}
```

### 3.2 轮询响应
```json
{
  "code": 200,
  "message": "任务执行完毕。",
  "data": {
    "status": "SUCCESS",
    "progress": 100,
    "result_url": "https://cos.xxx/out/file.png",
    "result_data": {
      "drawio_url": "https://cos.xxx/out/file.drawio",
      "png_url": "https://cos.xxx/out/file.png"
    }
  }
}
```

---

## 4. 典型场景

### 4.1 VIS-01 渲染流程图
- 提交: `POST /api/v1/vis/render_diagram`
- 轮询: `GET /api/v1/tasks/{task_id}`
- 成功: `result_url` 为图片链接

### 4.2 VIS-03 导出 Draw.io
- 提交: `POST /api/v1/vis/export_drawio`
- 轮询: `GET /api/v1/tasks/{task_id}`
- 成功: `result_data.drawio_url` 为工程文件，`result_data.png_url` 为预览图

### 4.3 PDF-02 盖章与水印
- 提交: `POST /api/v1/pdf/add_watermark`
- 轮询: `GET /api/v1/tasks/{task_id}`
- 成功: `result_url` 为新 PDF 下载地址
