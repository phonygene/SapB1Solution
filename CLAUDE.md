# SapB1Solution - 專案配置

> 🔗 **全域規則**：`C:\Projects\claude-config\CLAUDE.md`
> 🔗 **核心規則**：`C:\Projects\claude-config\core\` (財務原則、觸發式規則)

---

## 專案概述

| 項目 | 說明 |
|------|------|
| 技術棧 | ASP.NET Web Forms (VB.NET) |
| 性質 | 財務系統，資料正確性極為重要 |
| 風險 | 錯誤可能導致帳務與法律問題 |

---

## 專案特定規則

### 🔴 MCP SQL 查詢（必讀 `skills/database-guide.md`）

| 表名特徵 | db 參數 | 範例 |
|----------|---------|------|
| **j 開頭** | `jtdb` | jOPRQ, jOPCH, jUser |
| **O 開頭** | `sapb1` | OITM, OCRD, OPRQ |

```python
mcp__sapb1-sql__sql_query(query="SELECT * FROM jOPRQ", db="jtdb")
mcp__sapb1-sql__sql_query(query="SELECT * FROM OITM", db="sapb1")
```

### 🔴 ASP.NET 規則

| 規則 | 說明 |
|------|------|
| 控制項宣告 | `.aspx` 新增控制項時必須同步更新 `.aspx.designer.vb` |
| Namespace | .aspx 的 Inherits 必須加前綴 `MgmSP.` |
| 檔案編碼 | UTF-8 with BOM（修改後必須驗證） |

### 🔴 連線字串對照

| ConnectionString | db 參數 | 用途 |
|------------------|---------|------|
| `jtdbConnectionString` | `jtdb` | JET 自有資料 |
| `SapSQLConnection` | `sapb1` | SAP B1 |

---

## 專案 Skills

📁 **路徑**：`skills/INDEX.md`

| 觸發條件 | 資源檔案 |
|----------|----------|
| 執行 SQL 查詢（MCP 工具） | `database-guide.md` |
| 後端開發 | `backend-checklist.md` |
| SAP B1 整合 | `sap-checklist.md` |
| UI/前端開發 | `ui-checklist.md` |
| 所有開發任務 | `general-checklist.md` |

---

## 工作日誌

📁 **路徑**：`work-logs/daily/YYYY-MM/YYYY-MM-DD.md`

---

## 封存

`specArchive/` 存放過時規格，除非使用者要求否則不讀取。
