#!/bin/sh
# 自動初始化資料庫（首次或 PostgreSQL 尚未建表時）
cd /app/webapp
python -c "from app import init_db; init_db()" 2>&1 | head -5

# 啟動 gunicorn
exec gunicorn --workers 2 --bind 0.0.0.0:5001 --timeout 120 app:app
