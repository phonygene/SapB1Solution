# 資料庫使用指南

> 🔴 **`db` 參數為必填！每次調用都必須明確指定資料庫。**

---

## 資料庫架構

本專案有**兩個資料庫**，透過同一個 MCP Server 的 `db` 參數切換：

| db 參數 | 實際資料庫 | 位置 | 用途 |
|---------|-----------|------|------|
| `jtdb` | jtdb | 本地 (.\SQLEXPRESS2008R2) | JET 自有資料 |
| `sapb1` | JTTST | 遠端 (192.168.1.31) | SAP Business One |

⚠️ **沒有預設值**：未指定 `db` 參數將會報錯。

---

## 使用方式

### 查詢 JET 自有表（j 開頭）

```python
# 必須指定 db="jtdb"
mcp__sapb1-sql__sql_query(query="SELECT * FROM jOPRQ", db="jtdb")
```

### 查詢 SAP 表（O 開頭）

```python
# 必須指定 db="sapb1"
mcp__sapb1-sql__sql_query(query="SELECT * FROM OITM WHERE ItemCode = ?", params=["A001"], db="sapb1")
```

### 列出可用資料庫

```python
mcp__sapb1-sql__list_databases()
```

---

## 選擇規則

| 表名特徵 | db 參數 | 範例表 |
|----------|---------|--------|
| **j** 開頭 | `jtdb`（必須指定） | jOPRQ, jOPCH, jUser, jATTH |
| **O** 開頭 | `sapb1`（必須指定） | OITM, OCRD, OPRQ, OPCH |
| 其他 | 依實際所屬資料庫指定 | 自訂表 |

---

## 範例

### 正確用法

```python
# 查詢請購單（JET 自有）- 必須指定 db="jtdb"
mcp__sapb1-sql__sql_query(query="SELECT * FROM jOPRQ WHERE jID = ?", params=[123], db="jtdb")

# 查詢物料主檔（SAP）- 必須指定 db="sapb1"
mcp__sapb1-sql__sql_query(query="SELECT ItemCode, ItemName FROM OITM", db="sapb1")

# 取得表結構 - 都必須指定 db
mcp__sapb1-sql__get_table_info(table_name="jOPRQ", db="jtdb")  # JET 表
mcp__sapb1-sql__get_table_info(table_name="OITM", db="sapb1")  # SAP 表
```

### 錯誤用法

```python
# 錯誤：未指定 db 參數
mcp__sapb1-sql__sql_query(query="SELECT * FROM jOPRQ")
# → Error: 必須指定資料庫（db 參數）。可用選項: ['jtdb', 'sapb1']

# 正確：指定 db 參數
mcp__sapb1-sql__sql_query(query="SELECT * FROM jOPRQ", db="jtdb")
```

---

## Web.config 對照

| ConnectionString | db 參數 |
|------------------|---------|
| `jtdbConnectionString` | `jtdb` |
| `SapSQLConnection` | `sapb1` |

---

## 新增資料庫

如需新增其他資料庫，編輯 `mcp-sqlserver/config.json` 的 `databases` 區塊：

```json
{
  "databases": {
    "jtdb": { ... },
    "sapb1": { ... },
    "newdb": {
      "description": "新資料庫說明",
      "driver": "ODBC Driver 17 for SQL Server",
      "server": "server-address",
      "port": "1433",
      "database": "DatabaseName",
      "username": "user",
      "password": "pass",
      "backup_enabled": true,
      "backup_dir": "C:\\SQLBackups\\newdb"
    }
  }
}
```

---

*最後更新：2026-01-15*
