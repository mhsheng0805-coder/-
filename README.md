# 紡織所 來自民間業務收支管理系統

> 支援多年度（民國年）| 六部門 | 合約追蹤 | 未核銷費用 | Web 填寫

## 功能

- **六部門收支表**：原料部、產品部、檢驗部、製程部、雲分部、產服部
- **1–12 月逐月填寫**：收入（檢測/技術/認証/其他/附屬業務）+ 支出
- **已申請未核銷費用**：業務費、旅運費、材料費、維護費
- **合約追蹤**：每月新增簽約，洽談中／執行中自動延續至下月
- **六部門彙整**：總覽報表 + 趨勢圖
- **多年度切換**：右上角下拉切換民國年，各年資料獨立
- **帳號管理**：登入驗證、忘記密碼、修改密碼
- **匯出 Excel**：一鍵下載當年度完整資料
- **區域網路存取**：同 WiFi 下輸入伺服器 IP 即可填寫

## 快速啟動

### Windows（雙擊）
```
啟動系統.bat
```

### 命令列
```bash
cd webapp
pip install -r ../requirements.txt
python app.py
```

瀏覽器開啟 `http://127.0.0.1:5001`

> **預設帳號：** `admin` / `admin1234`  
> 首次登入後請立即至「🔑 修改密碼」更換。

## 目錄結構

```
業務收支系統/
├── webapp/
│   ├── app.py              # Flask 主程式
│   ├── templates/          # HTML 模板
│   │   ├── base.html
│   │   ├── login.html
│   │   ├── index.html
│   │   ├── dept.html
│   │   ├── summary.html
│   │   ├── contracts.html
│   │   ├── forgot_password.html
│   │   ├── reset_password.html
│   │   ├── change_password.html
│   │   └── admin_users.html
│   └── revenue.db          # SQLite 資料庫（不含於版控）
├── 啟動系統.bat
├── requirements.txt
├── .gitignore
└── README.md
```

## 多年度使用說明

1. 每年元旦後，系統自動以當年民國年為預設年度
2. 右上角下拉可隨時切換到歷史年度**查閱舊資料**
3. 各年度資料完全獨立，互不干擾

## GitHub 部署

```bash
git init
git add .
git commit -m "初始版本：紡織所業務收支管理系統"
git remote add origin https://github.com/你的帳號/業務收支系統.git
git push -u origin main
```

> ⚠️ `revenue.db`（含業務資料）已列入 `.gitignore`，不會上傳至 GitHub。

## 系統需求

- Python 3.10+
- Flask 3.x
- openpyxl 3.x
