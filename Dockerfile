# ============================================================
# Inventory Valuation Tracking - Dockerfile
# 基础镜像：python:3.11-slim（兼容 CentOS/Linux 环境）
# ============================================================

FROM python:3.11-slim

# 安装 pymssql 依赖的系统库（freetds-dev 提供 TDS 协议支持）
RUN apt-get update && apt-get install -y --no-install-recommends \
        freetds-dev \
        freetds-bin \
        gcc \
        g++ \
        libssl-dev \
        tzdata \
    && rm -rf /var/lib/apt/lists/*

# 时区设置（苏州/上海时区）
ENV TZ=Asia/Shanghai
RUN ln -snf /usr/share/zoneinfo/$TZ /etc/localtime && echo $TZ > /etc/timezone

# 工作目录
WORKDIR /app

# 安装 Python 依赖（先复制 requirements.txt，利用 Docker 层缓存）
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制应用代码
COPY inventory_tracking.py .
COPY config.ini .

# 创建报表输出目录
RUN mkdir -p /app/reports /app/Previous Balance

# 默认命令：守护模式，每天早上 09:00 自动执行
CMD ["python", "inventory_tracking.py", "--all-sites", "--daemon", "--prev-dir", "/app/Previous Balance", "--output-dir", "/app/reports"]
