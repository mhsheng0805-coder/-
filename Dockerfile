# ── 紡織所業務收支系統 Dockerfile ──────────────────────
FROM python:3.10-slim

WORKDIR /app

# 安裝系統依賴
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-cjk curl \
    && rm -rf /var/lib/apt/lists/*

# 安裝 Python 套件
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt python-docx

# 複製應用程式
COPY webapp/ ./webapp/

# 建立資料目錄（本機 SQLite 用）
RUN mkdir -p /data

# 啟動腳本：自動初始化 DB 後再啟動 gunicorn
COPY start.sh .
RUN chmod +x start.sh

EXPOSE 5001

CMD ["./start.sh"]
