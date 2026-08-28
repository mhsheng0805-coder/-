"""
SQLite → PostgreSQL 資料遷移腳本
使用方式：
  1. 將 PythonAnywhere 的 revenue.db 下載到與此腳本同一資料夾
  2. 設定 DATABASE_URL 環境變數（Render PostgreSQL 連線字串）
  3. 執行: python migrate_to_pg.py

Render 的 DATABASE_URL 格式：
  postgresql://user:password@host/dbname
"""

import sqlite3, os, sys

DATABASE_URL = os.environ.get('DATABASE_URL', '')
if not DATABASE_URL:
    print("錯誤：請設定 DATABASE_URL 環境變數")
    print("範例：set DATABASE_URL=postgresql://ttri:xxx@xxx.render.com/revenue")
    sys.exit(1)

if DATABASE_URL.startswith('postgres://'):
    DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)

try:
    import psycopg2, psycopg2.extras
except ImportError:
    print("安裝中：pip install psycopg2-binary")
    os.system("pip install psycopg2-binary")
    import psycopg2, psycopg2.extras

# 找 SQLite 檔案
DB_PATH = os.path.join(os.path.dirname(__file__), 'revenue.db')
if not os.path.exists(DB_PATH):
    DB_PATH = os.path.join(os.path.dirname(__file__), 'webapp', 'revenue.db')
if not os.path.exists(DB_PATH):
    print(f"找不到 revenue.db，請將檔案放到此腳本的同一資料夾")
    sys.exit(1)

print(f"來源：{DB_PATH}")
print(f"目標：{DATABASE_URL[:40]}...")

sqlite = sqlite3.connect(DB_PATH)
sqlite.row_factory = sqlite3.Row
pg = psycopg2.connect(DATABASE_URL)
pg_cur = pg.cursor()

TABLES = ['settings', 'users', 'revenue', 'contracts', 'unclaimed',
          'locks', 'audit_log', 'annual_goals', 'carry_updates', 'dept_groups']

for table in TABLES:
    try:
        rows = sqlite.execute(f"SELECT * FROM [{table}]").fetchall()
    except Exception:
        print(f"  跳過 {table}（不存在）")
        continue

    if not rows:
        print(f"  {table}: 0 筆（略過）")
        continue

    cols = rows[0].keys()
    col_list = ', '.join(cols)
    placeholders = ', '.join(['%s'] * len(cols))

    # 清除舊資料（重複執行時安全）
    pg_cur.execute(f"DELETE FROM {table}")

    count = 0
    for row in rows:
        try:
            pg_cur.execute(
                f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})",
                tuple(row)
            )
            count += 1
        except Exception as e:
            print(f"  [警告] {table} 某筆資料略過: {e}")

    print(f"  {table}: 已匯入 {count} 筆")

# 重設 PostgreSQL sequences（避免 ID 衝突）
for table in ['revenue', 'contracts', 'unclaimed', 'audit_log', 'annual_goals', 'carry_updates', 'dept_groups']:
    try:
        pg_cur.execute(f"""
            SELECT setval(pg_get_serial_sequence('{table}', 'id'),
                   COALESCE((SELECT MAX(id) FROM {table}), 0) + 1, false)
        """)
    except Exception:
        pass

pg.commit()
pg.close()
sqlite.close()
print("\n✅ 遷移完成！")
