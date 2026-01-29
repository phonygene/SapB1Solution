# Skills 資源索引

> 🔗 **核心規則**：`claude-config/core/financial-rules.yaml`（觸發式規則）
> 本目錄僅保留專案特定的參考資料和錯誤案例

---

## 觸發條件對照表

| 觸發條件 | 資源檔案 | 說明 |
|----------|----------|------|
| **執行 SQL 查詢（MCP 工具）** | `database-guide.md` | 🔴 **必讀** - 選擇正確的資料庫 |
| 後端開發問題排查 | `backend-checklist.md` | PostBack、ViewState 問題模式 |
| SAP B1 整合 | `sap-checklist.md` | SAP 表格參考、欄位長度 |
| UTF-8 BOM 操作 | `general-checklist.md` | 驗證和修復命令 |
| MSBuild 執行 | `general-checklist.md` | Windows 環境 Bash 規範 |

---

## 核心規則（在 claude-config/core/）

| ID | 觸發時機 | 檔案 |
|----|---------|------|
| FIN-001~003 | 金額處理、Save/Sync | `financial-rules.yaml` |
| SAP-001~002 | SAP 整合 | `financial-rules.yaml` |
| DB-001~002 | 資料庫操作 | `financial-rules.yaml` |
| ASPX-001 | 控制項新增 | `financial-rules.yaml` |
| ENCODE-001 | 檔案編碼 | `financial-rules.yaml` |

---

## 檔案清單

| 檔案 | 內容 | 更新日期 |
|------|------|----------|
| `database-guide.md` | MCP SQL 使用指南 | 2026-01-15 |
| `backend-checklist.md` | 問題模式、錯誤案例 | 2026-01-29 |
| `sap-checklist.md` | SAP 表格參考 | 2026-01-29 |
| `general-checklist.md` | BOM 操作、MSBuild 規範 | 2026-01-29 |
| `ui-checklist.md` | UI 開發參考 | 2026-01-07 |
| `ui-design-system.md` | 設計規範 | 2026-01-07 |

---

## 使用方式

1. **核心規則**：觸發式規則會自動生效，不需手動查閱
2. **問題排查**：遇到特定問題時查閱對應檔案的「常見問題模式」
3. **參考資料**：需要 SAP 表格對照、欄位長度等時查閱
