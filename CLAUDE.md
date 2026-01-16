# SapB1Solution - 專案配置

> 🔗 **全域規則**：`C:\Projects\claude-config\CLAUDE.md`
> 🔗 **技術 Skills**：`C:\Projects\claude-config\skills\`

---

## 專案概述

| 項目 | 說明 |
|------|------|
| 技術棧 | ASP.NET Web Forms (VB.NET) |
| 性質 | 財務系統，資料正確性極為重要 |
| 風險 | 錯誤可能導致帳務與法律問題 |

---

## 專案特定原則

除了全域的 POLA、WYSIWYG、Data Consistency 外，本專案額外要求：

| 原則 | 實踐方式 |
|------|----------|
| **Form State Integrity** | Sync 函數只讀取 UI，不重算 |

🔴 **禁止**：在 `SaveDocument` 或 `SyncToModel` 中重新計算金額/稅額

---

## 技術規則

| 規則 | 說明 |
|------|------|
| 控制項宣告 | `.aspx` 新增控制項時**必須同步更新** `.aspx.designer.vb` |
| Namespace | .aspx 的 Inherits **必須**加前綴 `MgmSP.` |
| 檔案編碼 | UTF-8 with BOM |
| 資料庫連線 | 本地 `jtdbConnectionString`、SAP `SapSQLConnection` |

### MCP SQL 查詢

🔴 **必讀 `skills/database-guide.md`**

| 表名特徵 | db 參數 | 範例 |
|----------|---------|------|
| **j 開頭** | `jtdb` | jOPRQ, jOPCH, jUser |
| **O 開頭** | `sapb1` | OITM, OCRD, OPRQ |

```python
mcp__sapb1-sql__sql_query(query="SELECT * FROM jOPRQ", db="jtdb")
mcp__sapb1-sql__sql_query(query="SELECT * FROM OITM", db="sapb1")
```

### 檔案編碼

🔴 **Claude Code 的 Write/Edit 工具不會保留 BOM！**

| 情境 | 必須執行 |
|------|----------|
| 創建/修改 `.aspx`/`.vb` | 驗證或轉換為 UTF-8 with BOM |

```powershell
# 驗證 BOM（應輸出 239,187,191）
[System.IO.File]::ReadAllBytes("檔案路徑")[0..2] -join ','

# 修復編碼
$content = Get-Content -Path $path -Raw -Encoding UTF8
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($path, $content, $utf8Bom)
```

---

## Agent 協作

| Agent | 職責 | 分支 |
|-------|------|------|
| Manager | 任務協調、代碼審查、日誌維護 | main |
| Backend | VB.NET、SAP 整合、資料庫 | agent/backend |
| UI-UX | 介面設計、CSS、響應式 | agent/ui-ux |
| Super | 全能模式（Manager + Backend + UI-UX） | main |

### 資源路徑

| 路徑 | 用途 |
|------|------|
| `.agent-workspace/{agent}/current.md` | Agent 狀態 |
| `.agent-workspace/handoff/{id}/` | 跨 Agent 任務交接 |
| `.claude/shared/` | 共享狀態 |
| `skills/INDEX.md` | 專案特定 Skills 索引 |

### 狀態規則

🔴 **回覆結束時**：更新 `.agent-workspace/{agent}/current.md` 的狀態為 `idle`

---

## 工作日誌

📁 **路徑**：`work-logs/daily/YYYY-MM/YYYY-MM-DD.md`

格式參見全域規則。本專案額外要求：

📚 **學習建議**（每日彙整，放在日誌最後）：
- 檢視用戶的錯誤、盲點、過時觀念
- 僅在有值得提出的觀察時記錄

🔄 **定期反思**：執行 `/reflect` 觸發

---

## 專案特定行為規則

🔴 **修正錯誤時**：先詢問是否檢查其他相似程式碼
🔴 **變更策略前**：先確認使用者目的與可接受的取捨
🔴 **禁止**：以「移除功能」規避錯誤（除非使用者同意）

---

## 版號管理

格式：`X.Y.Z`，位置：`VERSION` 檔案

規則參見全域配置。🔴 Z 不會自動進位到 Y。

---

## 封存

`specArchive/` 存放過時規格，除非使用者要求否則不讀取。
