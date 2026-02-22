# ========== SGA-Office API 镜像 ==========
FROM python:3.10-slim

WORKDIR /app

# 安装系统依赖：LibreOffice (PDF-01) + CJK 字体 (VIS 中文支持)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-noto-cjk \
    fonts-wqy-microhei \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 复制依赖清单并安装 Python 包
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制源代码
COPY . .

# 暴露 API 端口
EXPOSE 8000

# 使用 uvicorn 启动 FastAPI (生产模式)
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--workers", "4"]