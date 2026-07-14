# ── 紡織所業務收支系統 Dockerfile ──────────────────────
FROM python:3.10-slim

# 設定工作目錄
WORKDIR /app

# 安裝系統依賴（中文字型 for python-docx）
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# 複製需求檔並安裝 Python 套件
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt python-docx

# 複製應用程式
COPY webapp/ ./webapp/

# 建立資料目錄（SQLite DB 存放於此 volume）
RUN mkdir -p /data

# 設定環境變數（可由 docker-compose 覆蓋）
ENV FLASK_ENV=production
ENV DATA_DIR=/data

# 對外 Port
EXPOSE 5001

# 啟動（gunicorn，4 worker）
CMD ["gunicorn", "--workers", "4", "--bind", "0.0.0.0:5001", \
     "--timeout", "120", "--chdir", "/app/webapp", "app:app"]
