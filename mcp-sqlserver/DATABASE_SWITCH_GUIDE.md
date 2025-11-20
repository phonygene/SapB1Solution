# MCP Server 資料庫切換指南

## 概述

本系統允許您在不同的 SQL Server 資料庫之間快速切換 MCP Server 連接。

## 可用的資料庫配置

### 1. jtdb (原有生產環境)
- 伺服器: `172.17.16.1:12948`
- 資料庫: `jtdb`
- 用途: 原有的生產環境資料庫

### 2. JTTST (測試環境/新版本B1)
- 伺服器: `192.168.1.31:1433`
- 資料庫: `JTTST`
- 用途: 12月新版本 B1 測試環境，用於同步11月交易與主檔

## 切換方法

### 方法一: 使用切換腳本 (推薦)

#### Linux/WSL:
```bash
cd /home/jason/projects/vbnet/SapB1Solution/mcp-sqlserver
./switch-db.sh
```

或快速切換 (不需互動):
```bash
./switch-db.sh 1  # 切換到 jtdb
./switch-db.sh 2  # 切換到 JTTST
```

#### Windows:
```cmd
cd C:\path\to\mcp-sqlserver
switch-db.bat
```

### 方法二: 手動切換

直接複製對應的環境檔案：

```bash
# 切換到 jtdb
cp .env.jtdb .env

# 切換到 JTTST
cp .env.JTTST .env
```

## 重要事項

1. **切換後必須重啟 MCP Server**
   ```bash
   claude mcp restart
   ```

2. **備份目錄**
   - 每個資料庫都有獨立的備份目錄
   - jtdb: `C:\SQLBackups\jtdb`
   - JTTST: `C:\SQLBackups\JTTST`

3. **新增資料庫配置**
   - 建立新的 `.env.{資料庫名稱}` 檔案
   - 參考 `.env.example` 的格式
   - 切換腳本會自動偵測新的配置檔案

## 配置檔案格式

```env
# SQL Server 連線設定
DB_DRIVER=FreeTDS
DB_SERVER=伺服器IP
DB_PORT=連接埠
DB_NAME=資料庫名稱
DB_USER=使用者名稱
DB_PASSWORD=密碼

# 備份設定
BACKUP_ENABLED=true
BACKUP_DIR=備份目錄路徑

# 日誌設定
LOG_LEVEL=INFO
LOG_DIR=./logs
```

## 驗證連接

切換並重啟後，可以使用以下方式驗證：

```bash
# 查看資料庫狀態
claude # 然後使用 MCP 工具查詢資料庫狀態

# 或使用 SQL 查詢
SELECT @@SERVERNAME, DB_NAME()
```

## 故障排除

### 切換後無法連接
1. 確認已重啟 MCP Server
2. 檢查 `.env` 檔案內容是否正確
3. 驗證網路連接和防火牆設定
4. 檢查 SQL Server 是否運行且可訪問

### 找不到配置檔案
1. 確認 `.env.{資料庫名稱}` 檔案存在
2. 檢查檔案權限
3. 確認在正確的目錄下執行腳本

## 維護建議

1. 定期備份 `.env.*` 配置檔案
2. 不要將包含密碼的配置檔案提交到版本控制
3. 建立新環境時，先複製 `.env.example` 再修改
