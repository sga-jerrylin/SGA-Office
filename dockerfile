# 使用 Python 3 官方镜像作为基础镜像
FROM python:3.10-slim

# 设置工作目录
WORKDIR /app

ADD ./ /app/

# 安装依赖
RUN pip install -r requirements.txt

# 暴露 5001 端口，Gunicorn 默认监听该端口
EXPOSE 5001

# 定义容器启动时执行的命令，使用 Gunicorn 启动 Flask 应用
CMD ["gunicorn", "--timeout", "120","-w", "4", "-b", "0.0.0.0:5001", "main:app"]