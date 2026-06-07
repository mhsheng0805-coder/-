"""
紡織所 來自民間業務收支管理系統 - Flask Web App
支援多年度（民國年）
執行: python app.py  |  網址: http://[本機IP]:5001
"""
from flask import (Flask, render_template, request, jsonify, redirect,
                   url_for, send_file, session, g)
import sqlite3, os, secrets, hashlib
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
from functools import wraps

app = Flask(__name__)
# secret_key 固定化以確保重啟後 session 仍有效
_SECRET_FILE = os.path.join(os.path.dirname(__file__), '.secret_key')
if os.path.exists(_SECRET_FILE):
    app.secret_key = open(_SECRET_FILE).read().strip()
else:
    app.secret_key = secrets.token_hex(32)
    open(_SECRET_FILE, 'w').write(app.secret_key)

app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=8)

DB = os.path.join(os.path.dirname(__file__), 'revenue.db')

# 當前民國年（預設）
CURRENT_ROC_YEAR = datetime.now().year - 1911

def hash_pw(pw):
    return hashlib.sha256(pw.encode('utf-8')).hexdigest()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user'):
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorated

# ─────────────────────────────────────────────────────
DEPARTMENTS = ['原料部', '產品部', '檢驗部', '製程部', '雲分部', '產服部']
MONTHS = list(range(1, 13))

INCOME_ITEMS = [
    '檢測服務收入費', '技術服務收入費', '認証服務收入費',
    '其他業務收入費', '附屬業務-委外加工收入',
    '附屬業務-其他非民間收入', '附屬業務-其他民間收入',
    '其他小計(自民間)', '小型企業收入', '來自民間業務收入合計',
]
EXPENSE_ITEMS = [
    '人事費用', '業務費用', '維護費', '旅運費', '材料費',
    '租借設備使用費', '差旅費',
    '附屬業務-委外加工支出', '附屬業務-其他非民間支出',
    '附屬業務-其他民間支出', '其他小計(自民間)支出', '支出合計',
]
UNCLAIMED_ITEMS = ['業務費(未核銷)', '旅運費(未核銷)', '材料費(未核銷)', '維護費(未核銷)']
CONTRACT_STATUSES = ['洽談中', '新增簽約', '已簽約執行中', '完成']

# ── 資料庫 ─────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB)
    cur = con.cursor()
    cur.executescript('''
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL UNIQUE,
        password_hash TEXT NOT NULL,
        display_name TEXT,
        role TEXT DEFAULT 'user',
        reset_token TEXT,
        reset_expires DATETIME,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS revenue (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        dept TEXT NOT NULL,
        month INTEGER NOT NULL,
        item TEXT NOT NULL,
        amount REAL DEFAULT 0,
        goal REAL DEFAULT 0,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(year, dept, month, item)
    );
    CREATE TABLE IF NOT EXISTS contracts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        dept TEXT NOT NULL,
        month INTEGER NOT NULL,
        client TEXT,
        amount REAL DEFAULT 0,
        sign_date TEXT,
        due_date TEXT,
        status TEXT DEFAULT '洽談中',
        actual_amount REAL DEFAULT 0,
        note TEXT,
        carry_next INTEGER DEFAULT 0,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    );
    CREATE TABLE IF NOT EXISTS unclaimed (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER NOT NULL,
        dept TEXT NOT NULL,
        month INTEGER NOT NULL,
        item TEXT NOT NULL,
        amount REAL DEFAULT 0,
        note TEXT,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(year, dept, month, item)
    );
    ''')
    con.commit()

    # 初始年度設定
    cur.execute("INSERT OR IGNORE INTO settings (key,value) VALUES ('current_year',?)",
                (str(CURRENT_ROC_YEAR),))

    # 預設管理員
    cur.execute("SELECT COUNT(*) FROM users")
    if cur.fetchone()[0] == 0:
        cur.execute(
            "INSERT INTO users (username, password_hash, display_name, role) VALUES (?,?,?,?)",
            ('admin', hash_pw('admin1234'), '系統管理員', 'admin')
        )
        print('已建立預設帳號: admin / admin1234  (請登入後立即修改密碼)')
    con.commit()
    con.close()

def get_db():
    con = sqlite3.connect(DB)
    con.row_factory = sqlite3.Row
    return con

def get_current_year():
    """取得目前選用的年度（從 session 或 DB settings）"""
    if session.get('year'):
        return int(session['year'])
    con = get_db()
    row = con.execute("SELECT value FROM settings WHERE key='current_year'").fetchone()
    con.close()
    return int(row['value']) if row else CURRENT_ROC_YEAR

def get_all_years():
    """取得資料庫中有資料的所有年度（+ 當前年度）"""
    con = get_db()
    years = set()
    for tbl in ('revenue', 'contracts', 'unclaimed'):
        rows = con.execute(f"SELECT DISTINCT year FROM {tbl}").fetchall()
        years.update(r['year'] for r in rows)
    cur_yr = get_current_year()
    years.add(cur_yr)
    # 加入前後各 1 年供切換
    years.add(cur_yr - 1)
    years.add(cur_yr + 1)
    con.close()
    return sorted(years, reverse=True)

# ── 帳號路由 ───────────────────────────────────────────
@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('user'):
        return redirect(url_for('index'))
    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        con = get_db()
        user = con.execute(
            "SELECT * FROM users WHERE username=? AND password_hash=?",
            (username, hash_pw(password))
        ).fetchone()
        con.close()
        if user:
            session.permanent = False
            session['user'] = user['username']
            session['display_name'] = user['display_name'] or user['username']
            session['role'] = user['role']
            # 預設使用 DB 設定年度
            con2 = get_db()
            row = con2.execute("SELECT value FROM settings WHERE key='current_year'").fetchone()
            con2.close()
            session['year'] = int(row['value']) if row else CURRENT_ROC_YEAR
            return redirect(request.args.get('next') or url_for('index'))
        error = '帳號或密碼錯誤，請再試一次。'
    return render_template('login.html', error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/api/switch_year', methods=['POST'])
@login_required
def switch_year():
    year = request.json.get('year')
    if year:
        session['year'] = int(year)
    return jsonify({'status': 'ok', 'year': session.get('year')})

@app.route('/forgot_password', methods=['GET', 'POST'])
def forgot_password():
    msg = None
    token_shown = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        con = get_db()
        user = con.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        if user:
            token = secrets.token_urlsafe(24)
            expires = (datetime.now() + timedelta(hours=1)).strftime('%Y-%m-%d %H:%M:%S')
            con.execute("UPDATE users SET reset_token=?, reset_expires=? WHERE username=?",
                        (token, expires, username))
            con.commit()
            token_shown = token
            msg = '重設連結已產生（有效 1 小時）。'
        else:
            msg = '找不到此帳號，請確認帳號名稱。'
        con.close()
    return render_template('forgot_password.html', msg=msg, token_shown=token_shown)

@app.route('/reset_password/<token>', methods=['GET', 'POST'])
def reset_password(token):
    con = get_db()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user = con.execute(
        "SELECT * FROM users WHERE reset_token=? AND reset_expires > ?", (token, now)
    ).fetchone()
    if not user:
        con.close()
        return render_template('reset_password.html', error='連結無效或已過期。')
    error = None
    if request.method == 'POST':
        pw1 = request.form.get('password', '')
        pw2 = request.form.get('password2', '')
        if len(pw1) < 6:
            error = '密碼至少需要 6 個字元。'
        elif pw1 != pw2:
            error = '兩次輸入的密碼不一致。'
        else:
            con.execute(
                "UPDATE users SET password_hash=?, reset_token=NULL, reset_expires=NULL WHERE id=?",
                (hash_pw(pw1), user['id'])
            )
            con.commit()
            con.close()
            return render_template('reset_password.html', success=True)
    con.close()
    return render_template('reset_password.html', token=token,
                           username=user['username'], error=error)

@app.route('/change_password', methods=['GET', 'POST'])
@login_required
def change_password():
    error = None
    success = False
    if request.method == 'POST':
        old_pw  = request.form.get('old_password', '')
        new_pw  = request.form.get('new_password', '')
        new_pw2 = request.form.get('new_password2', '')
        con = get_db()
        user = con.execute(
            "SELECT * FROM users WHERE username=? AND password_hash=?",
            (session['user'], hash_pw(old_pw))
        ).fetchone()
        if not user:
            error = '目前密碼不正確。'
        elif len(new_pw) < 6:
            error = '新密碼至少需要 6 個字元。'
        elif new_pw != new_pw2:
            error = '兩次輸入的新密碼不一致。'
        else:
            con.execute("UPDATE users SET password_hash=? WHERE username=?",
                        (hash_pw(new_pw), session['user']))
            con.commit()
            success = True
        con.close()
    year = get_current_year()
    return render_template('change_password.html', error=error, success=success, year=year)

@app.route('/admin/users')
@login_required
def admin_users():
    if session.get('role') != 'admin':
        return redirect(url_for('index'))
    con = get_db()
    users = [dict(r) for r in con.execute(
        "SELECT id, username, display_name, role, created_at FROM users ORDER BY id"
    ).fetchall()]
    con.close()
    year = get_current_year()
    return render_template('admin_users.html', users=users, year=year)

@app.route('/admin/add_user', methods=['POST'])
@login_required
def add_user():
    if session.get('role') != 'admin':
        return jsonify({'error': 'forbidden'}), 403
    d = request.json
    username = d.get('username', '').strip()
    password = d.get('password', '')
    if not username or len(password) < 6:
        return jsonify({'error': '帳號或密碼格式不正確'}), 400
    con = get_db()
    try:
        con.execute("INSERT INTO users (username, password_hash, display_name, role) VALUES (?,?,?,?)",
                    (username, hash_pw(password), d.get('display_name',''), d.get('role','user')))
        con.commit()
    except sqlite3.IntegrityError:
        con.close()
        return jsonify({'error': '帳號已存在'}), 400
    con.close()
    return jsonify({'status': 'ok'})

@app.route('/admin/delete_user/<int:uid>', methods=['DELETE'])
@login_required
def delete_user(uid):
    if session.get('role') != 'admin':
        return jsonify({'error': 'forbidden'}), 403
    con = get_db()
    user = con.execute("SELECT username FROM users WHERE id=?", (uid,)).fetchone()
    if user and user['username'] == 'admin':
        con.close()
        return jsonify({'error': '不可刪除 admin 帳號'}), 400
    con.execute("DELETE FROM users WHERE id=?", (uid,))
    con.commit()
    con.close()
    return jsonify({'status': 'ok'})

# ── 主功能路由 ─────────────────────────────────────────
@app.route('/')
@login_required
def index():
    year = get_current_year()
    all_years = get_all_years()
    return render_template('index.html', departments=DEPARTMENTS, months=MONTHS,
                           year=year, all_years=all_years)

@app.route('/dept/<dept>')
@login_required
def dept_view(dept):
    if dept not in DEPARTMENTS:
        return redirect(url_for('index'))
    month = request.args.get('month', 1, type=int)
    year = get_current_year()
    all_years = get_all_years()
    return render_template('dept.html', dept=dept, month=month,
                           months=MONTHS, year=year, all_years=all_years,
                           income_items=INCOME_ITEMS,
                           expense_items=EXPENSE_ITEMS,
                           unclaimed_items=UNCLAIMED_ITEMS)

@app.route('/api/data/<dept>/<int:month>')
@login_required
def get_data(dept, month):
    year = get_current_year()
    con = get_db()
    rows = con.execute(
        'SELECT item, amount, goal FROM revenue WHERE year=? AND dept=? AND month=?',
        (year, dept, month)
    ).fetchall()
    data = {r['item']: {'amount': r['amount'], 'goal': r['goal']} for r in rows}
    unclaimed = con.execute(
        'SELECT item, amount FROM unclaimed WHERE year=? AND dept=? AND month=?',
        (year, dept, month)
    ).fetchall()
    unclaim_data = {r['item']: r['amount'] for r in unclaimed}
    contracts = [dict(r) for r in con.execute(
        'SELECT * FROM contracts WHERE year=? AND dept=? AND month=?',
        (year, dept, month)
    ).fetchall()]
    carry_forward = [dict(r) for r in con.execute(
        'SELECT * FROM contracts WHERE year=? AND dept=? AND month=? AND carry_next=1',
        (year, dept, month - 1)
    ).fetchall()] if month > 1 else []
    con.close()
    return jsonify({'revenue': data, 'unclaimed': unclaim_data,
                    'contracts': contracts, 'carry_forward': carry_forward})

@app.route('/api/save_revenue', methods=['POST'])
@login_required
def save_revenue():
    d = request.json
    year = get_current_year()
    dept, month = d['dept'], d['month']
    con = get_db()
    for item, vals in d['items'].items():
        con.execute('''INSERT INTO revenue (year, dept, month, item, amount, goal)
            VALUES (?,?,?,?,?,?)
            ON CONFLICT(year, dept, month, item) DO UPDATE
            SET amount=excluded.amount, goal=excluded.goal, updated_at=CURRENT_TIMESTAMP''',
            (year, dept, month, item, vals.get('amount', 0), vals.get('goal', 0)))
    con.commit()
    con.close()
    return jsonify({'status': 'ok'})

@app.route('/api/save_unclaimed', methods=['POST'])
@login_required
def save_unclaimed():
    d = request.json
    year = get_current_year()
    dept, month = d['dept'], d['month']
    con = get_db()
    for item, amount in d['items'].items():
        con.execute('''INSERT INTO unclaimed (year, dept, month, item, amount)
            VALUES (?,?,?,?,?)
            ON CONFLICT(year, dept, month, item) DO UPDATE
            SET amount=excluded.amount, updated_at=CURRENT_TIMESTAMP''',
            (year, dept, month, item, amount))
    con.commit()
    con.close()
    return jsonify({'status': 'ok'})

@app.route('/api/save_contract', methods=['POST'])
@login_required
def save_contract():
    d = request.json
    year = get_current_year()
    con = get_db()
    if d.get('id'):
        con.execute('''UPDATE contracts SET client=?, amount=?, sign_date=?, due_date=?,
            status=?, actual_amount=?, note=?, carry_next=?, updated_at=CURRENT_TIMESTAMP
            WHERE id=?''',
            (d['client'], d['amount'], d.get('sign_date',''), d.get('due_date',''),
             d['status'], d.get('actual_amount',0), d.get('note',''), d.get('carry_next',0), d['id']))
    else:
        con.execute('''INSERT INTO contracts
            (year, dept, month, client, amount, sign_date, due_date,
             status, actual_amount, note, carry_next)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
            (year, d['dept'], d['month'], d['client'], d['amount'],
             d.get('sign_date',''), d.get('due_date',''), d['status'],
             d.get('actual_amount',0), d.get('note',''), d.get('carry_next',0)))
    con.commit()
    con.close()
    return jsonify({'status': 'ok'})

@app.route('/api/delete_contract/<int:cid>', methods=['DELETE'])
@login_required
def delete_contract(cid):
    con = get_db()
    con.execute('DELETE FROM contracts WHERE id=?', (cid,))
    con.commit()
    con.close()
    return jsonify({'status': 'ok'})

@app.route('/api/overview')
@login_required
def api_overview():
    year = get_current_year()
    con = get_db()
    result = []
    for dept in DEPARTMENTS:
        income   = con.execute("SELECT COALESCE(SUM(amount),0) FROM revenue WHERE year=? AND dept=? AND item='來自民間業務收入合計'", (year, dept)).fetchone()[0]
        expense  = con.execute("SELECT COALESCE(SUM(amount),0) FROM revenue WHERE year=? AND dept=? AND item='支出合計'", (year, dept)).fetchone()[0]
        unclaim  = con.execute("SELECT COALESCE(SUM(amount),0) FROM unclaimed WHERE year=? AND dept=?", (year, dept)).fetchone()[0]
        contracts= con.execute("SELECT COUNT(*) FROM contracts WHERE year=? AND dept=?", (year, dept)).fetchone()[0]
        result.append({'dept': dept, 'income': income, 'expense': expense,
                       'unclaim': unclaim, 'contracts': contracts})
    con.close()
    return jsonify({'depts': result})

@app.route('/api/summary_data')
@login_required
def api_summary_data():
    year = get_current_year()
    con = get_db()
    rows = con.execute(
        "SELECT dept, month, item, amount FROM revenue WHERE year=? AND item IN ('來自民間業務收入合計','支出合計')",
        (year,)
    ).fetchall()
    ucl = con.execute(
        "SELECT dept, month, SUM(amount) as total FROM unclaimed WHERE year=? GROUP BY dept, month",
        (year,)
    ).fetchall()
    con.close()
    data = {}
    for r in rows:
        d = data.setdefault(r['dept'], {}).setdefault(r['month'], {})
        if r['item'] == '來自民間業務收入合計':
            d['income'] = r['amount']
        else:
            d['expense'] = r['amount']
    for r in ucl:
        data.setdefault(r['dept'], {}).setdefault(r['month'], {})['unclaim'] = r['total']
    return jsonify(data)

@app.template_filter('enumerate')
def enumerate_filter(iterable, start=0):
    return list(enumerate(iterable, start))

@app.route('/summary')
@login_required
def summary():
    year = get_current_year()
    all_years = get_all_years()
    con = get_db()
    rows = con.execute('''
        SELECT dept, month, item, SUM(amount) as total
        FROM revenue WHERE year=? AND item IN ('來自民間業務收入合計', '支出合計')
        GROUP BY dept, month, item
    ''', (year,)).fetchall()
    con.close()
    data = {}
    for r in rows:
        data.setdefault(r['dept'], {}).setdefault(r['month'], {})[r['item']] = r['total']
    return render_template('summary.html', departments=DEPARTMENTS, months=MONTHS,
                           year=year, all_years=all_years, data=data)

@app.route('/contracts')
@login_required
def contracts_view():
    year  = get_current_year()
    all_years = get_all_years()
    dept  = request.args.get('dept', '')
    month = request.args.get('month', 0, type=int)
    con = get_db()
    q, params = 'SELECT * FROM contracts WHERE year=?', [year]
    if dept:  q += ' AND dept=?';  params.append(dept)
    if month: q += ' AND month=?'; params.append(month)
    q += ' ORDER BY month, dept, id'
    contracts = [dict(r) for r in con.execute(q, params).fetchall()]
    con.close()
    return render_template('contracts.html', contracts=contracts,
                           departments=DEPARTMENTS, months=MONTHS,
                           year=year, all_years=all_years,
                           selected_dept=dept, selected_month=month,
                           statuses=CONTRACT_STATUSES)

@app.route('/export_excel')
@login_required
def export_excel():
    year = get_current_year()
    con = get_db()
    rev_rows      = con.execute('SELECT * FROM revenue WHERE year=? ORDER BY dept, month, item', (year,)).fetchall()
    unclaim_rows  = con.execute('SELECT * FROM unclaimed WHERE year=? ORDER BY dept, month, item', (year,)).fetchall()
    contract_rows = con.execute('SELECT * FROM contracts WHERE year=? ORDER BY dept, month', (year,)).fetchall()
    con.close()

    wb = openpyxl.Workbook()
    thin   = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hfill  = PatternFill('solid', start_color='BDD7EE', fgColor='BDD7EE')
    hfont  = Font(name='微軟正黑體', size=10, bold=True)
    ctr    = Alignment(horizontal='center', vertical='center')

    def make_sheet(ws, headers, widths, rows_data):
        for col, (h, w) in enumerate(zip(headers, widths), 1):
            c = ws.cell(row=1, column=col, value=h)
            c.font = hfont; c.fill = hfill; c.alignment = ctr; c.border = border
            ws.column_dimensions[get_column_letter(col)].width = w
        for ri, row in enumerate(rows_data, 2):
            for ci, val in enumerate(row, 1):
                ws.cell(row=ri, column=ci, value=val).border = border

    ws1 = wb.active; ws1.title = f'{year}年收支資料'
    make_sheet(ws1, ['年度','部門','月份','項目','金額','年度目標','更新時間'],
               [8,10,8,25,14,14,20],
               [(r['year'],r['dept'],r['month'],r['item'],r['amount'],r['goal'],r['updated_at']) for r in rev_rows])

    ws2 = wb.create_sheet(f'{year}年未核銷費用')
    make_sheet(ws2, ['年度','部門','月份','項目','金額','備註','更新時間'],
               [8,10,8,20,14,20,20],
               [(r['year'],r['dept'],r['month'],r['item'],r['amount'],r['note'] or '',r['updated_at']) for r in unclaim_rows])

    ws3 = wb.create_sheet(f'{year}年合約追蹤')
    make_sheet(ws3, ['年度','部門','月份','客戶/計畫','合約金額','簽約日期','預計完成','狀態','實收金額','備註','延續下月'],
               [8,10,6,25,12,12,12,12,12,20,8],
               [(r['year'],r['dept'],r['month'],r['client'],r['amount'],r['sign_date'],r['due_date'],
                 r['status'],r['actual_amount'],r['note'] or '','是' if r['carry_next'] else '') for r in contract_rows])

    output = io.BytesIO()
    wb.save(output); output.seek(0)
    return send_file(output, as_attachment=True,
                     download_name=f'{year}年業務收支資料_{datetime.now().strftime("%Y%m%d")}.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    init_db()
    import socket
    try:
        local_ip = socket.gethostbyname(socket.gethostname())
    except:
        local_ip = '127.0.0.1'
    current_year = CURRENT_ROC_YEAR
    print('=' * 55)
    print(f'  紡織所 來自民間業務收支管理系統（多年度版）')
    print('=' * 55)
    print(f'  本機網址 : http://127.0.0.1:5001')
    print(f'  區域網路 : http://{local_ip}:5001')
    print(f'  當前年度 : 民國 {current_year} 年')
    print(f'  預設帳號 : admin / admin1234')
    print('  (請登入後立即修改密碼)')
    print('=' * 55)
    app.run(host='0.0.0.0', port=5001, debug=False)
