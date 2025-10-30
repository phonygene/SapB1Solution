# SAP B1 SQL MCP Server

為 Claude Code 提供 SQL Server 資料庫操作能力的 MCP Server，具備完整的智能備份管理機制。

## ✨ 功能特色

- ✅ **SQL 查詢和寫入操作** - 支援參數化查詢，防止 SQL Injection
- ✅ **寫入前自動備份** - 每次寫入前自動建立備份，確保資料安全
- ✅ **智能備份策略** - 根據資料庫大小自動調整備份方案
- ✅ **流水備份** - 保留最近 N 個備份（預設 30 個）
- ✅ **每日備份** - 每天首次操作時建立，保留 N 天（預設 7 天）
- ✅ **自動清理舊備份** - 防止備份檔案占用過多空間
- ✅ **一鍵還原** - 快速從任意備份還原資料庫
- ✅ **安全檢查機制** - 黑名單關鍵字、影響行數上限
- ✅ **完整的日誌記錄** - 所有操作都有詳細記錄
- ✅ **操作協作模式** - 根據操作類型自動選擇確認方式

## 🤝 操作協作模式

為了確保資料安全，本 MCP Server 實施三種操作模式：

### 模式 1：直接執行
- **適用於**：SELECT 等純讀取操作
- **行為**：收到要求後直接執行，不需確認
- **範例**：查詢資料、列出資料表、檢視資料庫狀態

### 模式 2：顯示 SQL 並確認
- **適用於**：INSERT/UPDATE/DELETE 等寫入操作
- **行為**：執行前必須顯示完整 SQL，等待用戶確認
- **範例**：新增、修改、刪除資料

### 模式 3：顯示操作並確認
- **適用於**：BACKUP/RESTORE 等關鍵操作
- **行為**：執行前說明操作內容和影響，等待用戶確認
- **範例**：備份、還原資料庫

**📖 詳細規範請參閱**: [OPERATION_RULES.md](./OPERATION_RULES.md)

## 📋 系統需求

- Python 3.10 或更高版本
- SQL Server（支援 BACKUP/RESTORE 指令）
- ODBC Driver 17 for SQL Server
- uv（Python 套件管理工具）

## 🚀 快速開始

### 1. 安裝 ODBC Driver（如果尚未安裝）

```bash
# Ubuntu/Debian
curl https://packages.microsoft.com/keys/microsoft.asc | sudo apt-key add -
curl https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/prod.list | sudo tee /etc/apt/sources.list.d/mssql-release.list
sudo apt-get update
sudo ACCEPT_EULA=Y apt-get install -y msodbcsql17

# 或手動下載安裝
# https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server
```

### 2. 配置環境變數

```bash
# 複製範例檔案
cp .env.example .env

# 編輯 .env 檔案，填入你的資料庫連線資訊
nano .env
```

`.env` 檔案範例：

```env
DB_DRIVER=ODBC Driver 17 for SQL Server
DB_SERVER=localhost
DB_NAME=SBODemoTW
DB_USER=sa
DB_PASSWORD=your_password

BACKUP_ENABLED=true
BACKUP_DIR=./backups

LOG_LEVEL=INFO
LOG_DIR=./logs
```

### 3. 測試 MCP Server

```bash
# 使用 uv 執行測試
uv run python test_server.py
```

如果所有測試通過，你會看到：

```
🎉 所有測試通過！MCP Server 已準備好使用。
```

### 4. 配置到 Claude Code

找到 Claude Code 的配置檔案：
- **Windows**: `%APPDATA%\claude-code\settings.json`
- **Linux**: `~/.config/claude-code/settings.json`
- **Mac**: `~/Library/Application Support/claude-code/settings.json`

編輯 `settings.json`，加入：

```json
{
  "mcpServers": {
    "sapb1-sql": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "/absolute/path/to/mcp-sqlserver",
        "python",
        "src/server.py"
      ]
    }
  }
}
```

**重要：** 將 `/absolute/path/to/mcp-sqlserver` 替換為你的實際專案路徑。

### 5. 重啟 Claude Code

重新啟動 Claude Code，新的 MCP Server 就可以使用了！

## 🛠️ 可用工具

### 1. sql_query（模式 1：直接執行）
執行 SELECT 查詢，讀取資料。

**範例**：
```
用戶：幫我查詢 OITM 資料表的前 10 筆資料

Claude：[直接使用 sql_query 工具執行查詢]
        查詢成功，返回 10 筆記錄：
        [顯示查詢結果]
```

### 2. sql_write（模式 2：顯示 SQL 並確認）
執行 INSERT/UPDATE/DELETE，會自動備份。

**範例**：
```
用戶：在 TestTable 新增一筆資料，Name='測試', Value=100

Claude：我將執行以下 SQL 語句：

        INSERT INTO TestTable (Name, Value)
        VALUES ('測試', 100)

        這個操作將會：
        - 在 TestTable 新增 1 筆記錄
        - 寫入前會自動建立備份

        請確認是否執行？

用戶：確認

Claude：[執行 sql_write 工具]
        執行成功，已新增 1 筆記錄
        備份已建立：rolling_20250430_100523.bak
```

### 3. get_table_info
取得資料表結構資訊。

**範例**：
```
我：OITM 資料表有哪些欄位？
Claude：[使用 get_table_info 工具...]
```

### 4. list_tables
列出所有資料表名稱。

**範例**：
```
我：這個資料庫有哪些資料表？
Claude：[使用 list_tables 工具...]
```

### 5. list_backups
列出所有備份檔案。

**範例**：
```
我：列出所有可用的備份
Claude：[使用 list_backups 工具...]
```

### 6. restore_backup（模式 3：顯示操作並確認）
從備份還原資料庫。

**範例**：
```
用戶：還原到今天早上的備份

Claude：[使用 list_backups 工具查看]
        我找到以下備份檔案：
        - daily_20250430.bak (今天早上，50.2 MB)

        還原操作將會：
        1. 關閉所有現有連線
        2. 將資料庫還原到 2025-04-30 早上的狀態
        3. 當前的所有變更將會遺失
        4. 預計需要 1-2 分鐘

        ⚠️ 警告：此操作無法復原！

        請確認是否執行？

用戶：確認

Claude：[執行 restore_backup 工具]
        還原完成，資料庫已回復到 2025-04-30 早上的狀態
```

### 7. get_db_status（模式 1：直接執行）
查看資料庫狀態和備份策略。

**範例**：
```
我：資料庫目前的狀態如何？
Claude：[使用 get_db_status 工具...]
```

### 8. create_backup
手動建立備份。

**範例**：
```
我：手動建立一個備份
Claude：[使用 create_backup 工具...]
```

## 📊 備份策略

系統會根據資料庫大小自動選擇適合的備份策略：

| 資料庫大小 | 策略     | 流水備份數量 | 每日備份保留 | 說明                   |
|-----------|---------|-------------|-------------|----------------------|
| < 100 MB  | Small   | 30 個       | 7 天        | 適用於小型開發資料庫     |
| < 1 GB    | Medium  | 10 個       | 7 天        | 適用於中型測試資料庫     |
| < 10 GB   | Large   | 5 個        | 3 天        | 適用於大型生產資料庫     |
| > 10 GB   | Very Large | 停用     | 1 天        | 建議使用專業備份方案     |

你可以在 `config.json` 中自訂備份策略。

## 🔒 安全機制

### 1. SQL 注入防護
- 支援參數化查詢
- 自動檢查危險的 SQL 關鍵字

### 2. 操作限制
- 黑名單關鍵字（DROP DATABASE, TRUNCATE 等）
- 影響行數上限（預設 1000 筆）
- 超過限制自動回滾

### 3. 完整記錄
- 所有操作記錄到 `logs/mcp_server.log`
- 可追溯和審計

## 📚 uv 使用教學

### 常用命令

```bash
# 安裝或更新依賴
uv sync

# 執行 Python 腳本
uv run python script.py

# 進入虛擬環境
source .venv/bin/activate  # Linux/Mac
.venv\Scripts\activate     # Windows

# 新增套件
uv add package-name

# 查看已安裝套件
uv pip list

# 更新所有套件
uv lock --upgrade
```

### uv 的優勢

- ⚡ **極快速度** - 比 pip 快 10-100 倍
- 🎯 **一體化** - 管理 Python 版本、虛擬環境、套件
- 🔒 **依賴鎖定** - 自動生成 `uv.lock` 確保可重現性
- 📦 **標準格式** - 使用 `pyproject.toml`（PEP 621）

## 🔧 故障排除

### 連線失敗

**問題**：無法連線到 SQL Server

**解決**：
1. 檢查 `.env` 中的連線資訊是否正確
2. 確認 SQL Server 已啟動
3. 檢查防火牆設定
4. 測試：`uv run python -c "import pyodbc; print(pyodbc.drivers())"`

### 備份失敗

**問題**：BACKUP DATABASE 權限不足

**解決**：
1. 確保 SQL Server 帳號有 `BACKUP DATABASE` 權限
2. 執行：`GRANT BACKUP DATABASE TO [your_user];`
3. 或使用 `sa` 帳號

### ODBC Driver 找不到

**問題**：`[01000] [unixODBC][Driver Manager]Can't open lib...`

**解決**：
1. 安裝 ODBC Driver（見上方安裝指南）
2. 檢查驅動名稱：`odbcinst -q -d`
3. 更新 `.env` 中的 `DB_DRIVER`

### MCP Server 無法啟動

**問題**：Claude Code 無法連接到 MCP Server

**解決**：
1. 檢查 `logs/mcp_server.log` 查看錯誤訊息
2. 確認配置路徑是絕對路徑
3. 手動測試：`uv run python src/server.py`
4. 重啟 Claude Code

## 📁 專案結構

```
mcp-sqlserver/
├── src/
│   ├── __init__.py           # 模組初始化
│   ├── server.py             # MCP Server 主程式
│   ├── database.py           # 資料庫操作模組
│   └── backup_manager.py     # 備份管理模組
├── backups/
│   ├── rolling/              # 流水備份目錄
│   └── daily/                # 每日備份目錄
├── logs/                     # 日誌目錄
├── .venv/                    # 虛擬環境
├── pyproject.toml            # 專案配置
├── uv.lock                   # 依賴鎖定檔
├── config.json               # 備份策略配置
├── .env                      # 環境變數（需自行建立）
├── .env.example              # 環境變數範例
├── .gitignore                # Git 忽略檔案
├── test_server.py            # 測試腳本
└── README.md                 # 使用說明
```

## 🎓 進階使用

### 自訂備份策略

編輯 `config.json`：

```json
{
  "backup_strategies": {
    "small": {
      "max_size_mb": 100,
      "rolling_limit": 50,        // 改為保留 50 個
      "daily_retain_days": 14,    // 改為保留 14 天
      "enabled": true
    }
  }
}
```

### 停用自動備份

在 `.env` 中設定：

```env
BACKUP_ENABLED=false
```

### 調整安全限制

編輯 `config.json`：

```json
{
  "safe_mode": {
    "max_affected_rows": 5000,  // 提高影響行數上限
    "blacklist_keywords": [
      "DROP DATABASE",
      "TRUNCATE TABLE"
      // 新增或移除關鍵字
    ]
  }
}
```

## 🤝 貢獻

歡迎提交 Issue 或 Pull Request！

## 📄 授權

MIT License

## 🙏 致謝

- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) - Anthropic
- [uv](https://github.com/astral-sh/uv) - Astral
- [pyodbc](https://github.com/mkleehammer/pyodbc) - mkleehammer

---

**建立時間**: 2025-10-30
**版本**: 1.0.0
**作者**: Jason (with Claude Code)
