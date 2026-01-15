# 資料庫使用指南

> Agent 執行 SQL 查詢前**必須**先閱讀此文件，選擇正確的 MCP Server。

---

## 資料庫架構

本專案使用**兩個獨立的資料庫**，透過不同的 MCP Server 連接：

| MCP Server | 資料庫 | 位置 | 用途 |
|------------|--------|------|------|
| `mcp__jtdb__*` | jtdb | 本地 (.\SQLEXPRESS2008R2) | JET 自有資料 |
| `mcp__sapb1__*` | JTTST | 遠端 (192.168.1.31) | SAP Business One |

---

## 選擇規則（必須遵守）

### 使用 `jtdb` 的情況

**表名以 `j` 開頭** = JET 自有表，使用 `mcp__jtdb__*`

| 表名 | 說明 |
|------|------|
| jOPCH | 費用報支單表頭 |
| jPCH1 | 費用報支單明細 |
| jOPRQ | 請購單表頭 |
| jPRQ1 | 請購單明細 |
| jUser | 使用者資料 |
| jATTH | 附件表頭 |
| jATT1 | 附件明細 |

**範例**：
```
查詢 jOPRQ 狀態 → 使用 mcp__jtdb__sql_query
```

### 使用 `sapb1` 的情況

**表名以 `O` 開頭** = SAP 系統表，使用 `mcp__sapb1__*`

| 表名 | 說明 |
|------|------|
| OITM | 物料主檔 |
| OCRD | 業務夥伴主檔 |
| OPRQ | SAP 請購單 |
| OPOR | SAP 採購訂單 |
| OPCH | SAP 應付發票 |
| OSLP | 業務員主檔 |
| OHEM | 員工主檔 |
| OWHS | 倉庫主檔 |

**範例**：
```
查詢 OITM 品號資料 → 使用 mcp__sapb1__sql_query
```

### 快速判斷法

```
表名第一個字母：
  j → jtdb（JET 自有）
  O → sapb1（SAP 系統）
  其他 → 先查 jtdb，若不存在再查 sapb1
```

---

## 常見錯誤

### ❌ 錯誤：用 sapb1 查 jOPRQ

```
mcp__sapb1__sql_query("SELECT * FROM jOPRQ")
→ Error: Invalid object name 'jOPRQ'
```

### ✓ 正確：用 jtdb 查 jOPRQ

```
mcp__jtdb__sql_query("SELECT * FROM jOPRQ")
→ 成功返回資料
```

---

## 跨資料庫查詢

若需要同時查詢兩個資料庫的資料：

1. **分開查詢**：先查 jtdb，再查 sapb1，在程式碼中 JOIN
2. **Linked Server**：在 SQL 中使用 `[SAP-NEW-TST].[JTTST].dbo.OITM`（需 DBA 設定）

---

## MCP 工具對照

| 工具 | jtdb | sapb1 |
|------|------|-------|
| 查詢 | `mcp__jtdb__sql_query` | `mcp__sapb1__sql_query` |
| 寫入 | `mcp__jtdb__sql_write` | `mcp__sapb1__sql_write` |
| DDL | `mcp__jtdb__sql_ddl` | `mcp__sapb1__sql_ddl` |
| 表結構 | `mcp__jtdb__get_table_info` | `mcp__sapb1__get_table_info` |
| 列表 | `mcp__jtdb__list_tables` | `mcp__sapb1__list_tables` |

---

## Web.config 對照

程式碼中的連線字串與 MCP Server 的對應：

| ConnectionString | MCP Server |
|------------------|------------|
| `jtdbConnectionString` | `jtdb` |
| `SapSQLConnection` | `sapb1` |
| `MDRConnectionString` | （未設定 MCP） |

---

*最後更新：2026-01-15*
