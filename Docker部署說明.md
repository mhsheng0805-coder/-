# 紡織所業務收支系統 — Docker 部署說明

## 系統需求

| 項目 | 需求 |
|------|------|
| 作業系統 | Windows Server 2019+ / Ubuntu 20.04+ |
| Docker Engine | 24.x 以上 |
| Docker Compose | v2.x 以上（`docker compose` 指令）|
| RAM | 最低 512 MB |
| 硬碟 | 最低 2 GB |

---

## 一、初次部署步驟

### 1. 取得程式碼
```bash
# 方式一：從 GitHub clone
git clone https://github.com/mhsheng0805-coder/-.git revenue-system
cd revenue-system

# 方式二：直接複製資料夾至伺服器
# 將整個「業務收支系統」資料夾上傳到伺服器，進入該目錄
```

### 2. 修改 docker-compose.yml（重要）
開啟 `docker-compose.yml`，修改以下兩個設定：

```yaml
environment:
  - SECRET_KEY=請改為一段隨機長字串例如abc123xyz789   # ← 必改
  - DEPLOY_TOKEN=ttri2025deploy                        # 可維持原值
```

> SECRET_KEY 可用此指令產生：`python -c "import secrets; print(secrets.token_hex(32))"`

### 3. 建立並啟動容器
```bash
docker compose up -d --build
```

首次執行會下載 Python 映像檔（約 150 MB），需要幾分鐘。

### 4. 初始化資料庫（首次才需要）
```bash
docker compose exec webapp python -c "
import sys; sys.path.insert(0,'/app/webapp')
from app import init_db; init_db()
print('資料庫初始化完成')
"
```

### 5. 開啟系統
瀏覽器輸入：`http://伺服器IP:5001`

---

## 二、日常維運

### 啟動
```bash
docker compose up -d
```

### 停止
```bash
docker compose down
```

### 查看執行狀態
```bash
docker compose ps
docker compose logs -f webapp
```

### 更新程式碼後重新部署
```bash
git pull
docker compose up -d --build
```

---

## 三、備份資料庫

SQLite 資料庫存放於 Docker volume `db_data`，備份方式：

```bash
# 備份
docker compose exec webapp cp /data/revenue.db /data/revenue_backup_$(date +%Y%m%d).db

# 或複製到本機（Windows PowerShell）
docker cp ttri-revenue:/data/revenue.db ./revenue_backup.db
```

---

## 四、對外開放（選用）

若要用 80 port 對外，修改 `docker-compose.yml`：
```yaml
ports:
  - "80:5001"
```

若要加 HTTPS，建議在前方加裝 **Nginx** 反向代理（資管自行設定 SSL 憑證）。

---

## 五、Port 衝突處理

若 5001 port 已被佔用，修改 docker-compose.yml：
```yaml
ports:
  - "8080:5001"   # 改成其他未被佔用的 port
```

---

## 六、常見問題

| 症狀 | 解法 |
|------|------|
| 網頁打不開 | `docker compose ps` 確認容器是否 running |
| 資料遺失 | 確認 volume `db_data` 存在：`docker volume ls` |
| 登入失敗 | 確認 SECRET_KEY 在重新部署後未更換 |
| 容器一直重啟 | `docker compose logs webapp` 查看錯誤訊息 |
