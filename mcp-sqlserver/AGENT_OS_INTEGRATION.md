# Agent-OS 整合指南

本文件說明如何在 Agent-OS 環境中使用 SAP B1 SQL MCP Server。

## 📋 整合概覽

### 架構說明

```
┌─────────────────────────────────────┐
│         Claude (你)                 │
│    或 Agent-OS Agents               │
│  (database-engineer, api-engineer等)│
└──────────────┬──────────────────────┘
               │
               ├─ 使用 Claude Code 工具
               │
               ├─ 使用 MCP Server 工具 (新增)
               │   ├─ sql_query
               │   ├─ sql_write
               │   ├─ get_table_info
               │   ├─ list_tables
               │   ├─ list_backups
               │   ├─ restore_backup
               │   ├─ get_db_status
               │   └─ create_backup
               │
               ▼
        ┌──────────────┐
        │  jtdb 資料庫  │
        │ (本機測試環境) │
        └──────────────┘
```

### 配置檔案

1. **專案層級**：`.mcp.json`
   - 定義 MCP Server 配置
   - 所有 session 都能看到

2. **用戶層級**：`~/.claude/settings.json`
   - 批准使用 `sapb1-sql` server
   - 個人設定

3. **MCP Server**：`mcp-sqlserver/`
   - 實際的 server 程式碼
   - 獨立運行，按需啟動

## 🎯 使用場景

### 場景 1：你（主要 Claude）使用

當你需要：
- 查詢資料庫資訊
- 測試 SQL 語句
- 檢視資料表結構
- 管理資料庫備份

**操作方式**：
```
你：查詢 user_vendor_taxid 資料表的所有記錄

Claude：[使用 sql_query 工具]
        SELECT * FROM user_vendor_taxid

        查詢結果：[顯示結果]
```

### 場景 2：Database-Engineer Agent 使用

當 database-engineer 需要：
- 建立資料表前先檢視現有結構
- 測試新寫的 SQL 腳本
- 驗證資料遷移結果

**操作方式**：
```
你：請 database-engineer 檢視 branch 資料表的結構

Claude：[啟動 database-engineer agent]

Database-Engineer：[使用 get_table_info 工具]
                   資料表 branch 包含以下欄位：
                   - areacode (varchar, NOT NULL)
                   - areadesc (varchar, NOT NULL)
```

### 場景 3：整合到 Session 工作流程

在 `Claude, sess on.` 初始化時：
- MCP Server 會自動可用
- 所有 agents 都能使用相同的資料庫連線
- 不需要額外的初始化步驟

## 📖 操作規範（遵循 OPERATION_RULES.md）

### 模式 1：直接執行（查詢操作）

```
使用者：列出所有資料表

Claude/Agent：[直接執行 list_tables]
              找到 13 個資料表：
              - branch
              - cnc1
              - dept
              ...
```

**適用工具**：
- `sql_query` (SELECT)
- `get_table_info`
- `list_tables`
- `list_backups`
- `get_db_status`

### 模式 2：顯示 SQL 並確認（寫入操作）

```
使用者：在 user_vendor_taxid 新增一筆測試資料

Claude/Agent：我將執行以下 SQL：

              INSERT INTO user_vendor_taxid (user_id, vendor_taxid)
              VALUES ('test_user', '12345678')

              這個操作將會：
              - 新增 1 筆記錄到 user_vendor_taxid
              - 寫入前會自動建立備份

              請確認是否執行？

使用者：確認

Claude/Agent：[執行 sql_write]
              ✓ 成功新增 1 筆記錄
              ✓ 備份已建立：rolling_20251030_120000.bak
```

**適用工具**：
- `sql_write` (INSERT/UPDATE/DELETE)

### 模式 3：顯示操作並確認（關鍵操作）

```
使用者：還原到昨天的備份

Claude/Agent：[使用 list_backups 查看]
              找到備份：daily_20251029.bak (7.12 MB)

              還原操作將會：
              1. 關閉所有連線
              2. 將資料庫還原到 2025-10-29 的狀態
              3. 當前所有變更將遺失

              ⚠️ 警告：此操作無法復原！

              請確認是否執行？

使用者：確認

Claude/Agent：[執行 restore_backup]
              ✓ 資料庫已還原
```

**適用工具**：
- `restore_backup`
- `create_backup`

## 🔧 與 Agent-OS 的協作

### Database-Engineer 增強

**原有能力**：
- 建立資料庫遷移腳本
- 設計資料表結構
- 撰寫 SQL 查詢

**新增能力（透過 MCP Server）**：
- ✅ 即時查詢現有資料表結構
- ✅ 測試 SQL 腳本執行結果
- ✅ 驗證資料遷移前後的資料
- ✅ 自動備份保護（寫入前）

### 協作流程範例

#### 情境：建立新的統編資料表功能

**步驟 1：檢視現有結構**
```
你：請檢視 user_vendor_taxid 資料表

Claude：[使用 get_table_info]
        - user_id (int)
        - vendor_taxid (varchar)
        - created_at (datetime)
```

**步驟 2：設計新功能**
```
你：Claude, 請 database-engineer 設計一個新欄位用於記錄更新時間

Database-Engineer：[分析需求]
                   [產生遷移腳本到 shopfloor/Claude_TMP/SqlQuery/]

                   已產生：
                   - 01_AddColumn_user_vendor_taxid_updated_at.sql
```

**步驟 3：測試執行（你來執行，保持審查權）**
```
你：執行這個遷移腳本

Claude：[讀取 SQL 檔案]
        [使用 sql_write 執行]

        我將執行：
        ALTER TABLE user_vendor_taxid
        ADD updated_at DATETIME DEFAULT GETDATE()

        請確認？

你：確認

Claude：✓ 執行成功
        ✓ 備份已建立
```

**步驟 4：驗證結果**
```
你：驗證資料表結構

Claude：[使用 get_table_info]
        ✓ 確認 updated_at 欄位已新增
```

## 🎓 最佳實踐

### 1. 備份策略

**自動備份**：
- 所有 `sql_write` 操作前自動備份
- 流水備份：保留最近 30 個
- 每日備份：保留 7 天

**手動備份**：
```
重要操作前：
你：手動建立一個備份

Claude：[使用 create_backup]
        ✓ 備份已建立：rolling_20251030_150000.bak
```

### 2. 查詢與寫入分離

**只讀操作**：放心使用，不會影響資料
```
- sql_query
- get_table_info
- list_tables
- get_db_status
```

**寫入操作**：謹慎確認，會修改資料
```
- sql_write （需確認）
```

**關鍵操作**：特別小心，影響重大
```
- restore_backup （需確認 + 警告）
```

### 3. 與 Session 管理整合

**sess on（上班）**：
- MCP Server 自動可用
- 可以立即查詢資料庫狀態

**sess check（檢查進度）**：
- 可以查詢當前資料狀態
- 不會寫入資料

**sess wrap（階段存檔）**：
- 如果有重要的資料變更，可手動建立備份
- 記錄到 LastCheckPoint.log

**sess off（下班）**：
- MCP Server 會自動關閉連線
- 備份檔案保留在 `C:\SQLBackups\jtdb`

## 📊 實用範例

### 範例 1：檢視專案使用的所有資料表

```
你：列出這個專案使用的所有資料表，並說明用途

Claude：[使用 list_tables]
        找到 13 個資料表：

        核心資料表：
        - branch：分公司/區域資訊
        - dept：部門資料
        - User：使用者帳號
        - user_vendor_taxid：使用者常用統編

        簽核相關：
        - permissiongrp：權限群組
        - user_permissionnew：使用者權限

        ...
```

### 範例 2：查詢測試資料

```
你：查詢 user_vendor_taxid 中 user_id='U001' 的所有記錄

Claude：[使用 sql_query]
        SELECT * FROM user_vendor_taxid
        WHERE user_id = 'U001'

        結果：[顯示 3 筆記錄]
```

### 範例 3：新增測試資料（用於開發測試）

```
你：新增一筆測試統編資料

Claude：我將執行：
        INSERT INTO user_vendor_taxid (user_id, vendor_taxid)
        VALUES ('U999', '99999999')

        請確認？

你：確認

Claude：✓ 已新增
        ✓ 備份：rolling_20251030_160000.bak
```

### 範例 4：災難復原

```
情境：不小心刪除了重要資料

你：列出最近的備份

Claude：[使用 list_backups]
        流水備份（最近5個）：
        1. rolling_20251030_160000.bak (剛才新增資料前)
        2. rolling_20251030_155500.bak
        3. rolling_20251030_155000.bak
        ...

        每日備份：
        1. daily_20251030.bak (今天早上)
        2. daily_20251029.bak (昨天)

你：還原到 rolling_20251030_160000.bak

Claude：[執行 restore_backup]
        ⚠️ 這會回復到新增資料前的狀態
        請確認？

你：確認

Claude：✓ 已還原
```

## 🚨 注意事項

### 1. 資料庫環境

**當前配置**：
- 資料庫：`jtdb` (本機開發環境)
- 位置：`.\SQLEXPRESS2008R2`

**不包含**：
- JTTST1 (SAP 測試環境) - 需另外配置
- 生產環境 - 不建議使用 MCP Server

### 2. 安全性

- ✅ 所有寫入操作需要確認
- ✅ 自動備份保護
- ✅ 操作記錄到 `logs/mcp_server.log`
- ⚠️ 備份檔案在 Windows：`C:\SQLBackups\jtdb`

### 3. 效能考量

- 資料庫大小：7.12 MB（小型）
- 備份策略：保留 30 個流水備份 + 7 天每日備份
- 備份時間：約 1-2 秒
- 查詢速度：即時

## 📚 相關文檔

| 文檔 | 說明 |
|------|------|
| **OPERATION_RULES.md** | 完整操作規範（必讀） |
| **QUICK_REFERENCE.md** | 快速參考速查表 |
| **README.md** | 使用說明與安裝指南 |
| **.claude_rules** | Claude 自動讀取的規則 |
| **agent-os/SESSION_INIT.md** | Agent-OS 初始化規範 |

## 🎉 開始使用

### 測試 MCP Server 是否可用

```
你：列出所有資料表

如果成功顯示 13 個資料表，表示 MCP Server 已正確配置！
```

### 給 Database-Engineer 的第一個任務

```
你：Claude, 請 database-engineer 使用 MCP Server 檢視所有資料表，
    並產生一份資料字典文檔到 shopfloor/Claude_TMP/etc/
```

---

**最後更新**：2025-10-30
**版本**：1.0.0
**整合狀態**：✅ 完成

有任何問題，請參閱 OPERATION_RULES.md 或詢問 Claude！
