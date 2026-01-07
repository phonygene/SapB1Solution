# JET Enterprise Platform - AI 協作規範

## 專案概述

- **技術棧**：ASP.NET Web Forms (VB.NET)
- **性質**：財務系統，資料正確性極為重要
- **風險**：錯誤可能導致帳務與法律問題

---

## 核心開發原則

| 原則 | 實踐方式 |
|------|----------|
| **POLA** | 不在背景修改用戶輸入的值 |
| **WYSIWYG** | 儲存前不重新計算已顯示的值 |
| **Data Consistency** | 避免在 Save 時修改 Model |
| **Form State Integrity** | Sync 函數只讀取 UI，不重算 |

🔴 **禁止**：在 `SaveDocument` 或 `SyncToModel` 中重新計算金額/稅額

---

## 技術規則

| 規則 | 說明 |
|------|------|
| 控制項宣告 | `.aspx` 新增控制項時**必須同時更新** `.aspx.designer.vb` |
| Namespace | .aspx 的 Inherits **必須**加前綴 `MgmSP.` |
| 檔案編碼 | UTF-8 with BOM |
| 資料庫連線 | 本地 `jtdbConnectionString`、SAP `SapSQLConnection` |

---

## Agent 協作

| Agent | 職責 | 分支 |
|-------|------|------|
| Manager | 任務協調、代碼審查、日誌維護 | main |
| Backend | VB.NET、SAP 整合、資料庫 | agent/backend |
| UI-UX | 介面設計、CSS、響應式 | agent/ui-ux |

- 使用者優先與 Manager 溝通
- 直接指派其他 Agent 時，該 Agent 必須同步回報 Manager

### 資源查找

🔵 **執行任務前**：讀取 `skills/INDEX.md` 查看是否有相關資源
🔴 **變更 skills/ 時**：必須同步更新 `skills/INDEX.md`

| 路徑模式 | 用途 |
|----------|------|
| `.agent-workspace/{agent}/current.md` | Agent 狀態（idle/thinking） |
| `.agent-workspace/handoff/{id}/` | 跨 Agent 任務交接 |
| `.claude/shared/` | 跨 Agent 共享狀態 |

### Slash Commands

- 位置：`.claude/commands/`
- `#name` 等同 `/name`

### 狀態規則

🔴 **回覆結束時**：更新 `.agent-workspace/{agent}/current.md` 的「## 狀態」為 `idle`

---

## 行為約束

🔴 **修正錯誤時**：先詢問是否檢查其他相似程式碼
🔴 **變更策略前**：先確認使用者目的與可接受的取捨
🔴 **禁止**：以「移除功能」規避錯誤（除非使用者同意）

---

## 版號管理

格式：`X.Y.Z`，位置：`VERSION` 檔案

| 位置 | 名稱 | 進位條件 |
|------|------|----------|
| X | Major | 重大架構變更、不相容更新 |
| Y | Minor | 新增功能 |
| Z | Patch | 維護、修復、次要變更 |

🔴 **Z 不會自動進位到 Y**：
- 1.1.9 + 維護 → 1.1.10 ✅
- 1.1.9 + 維護 → 1.2.0 ❌
- 只有「新增功能」才進位 Y

---

## 封存

`specArchive/` 存放過時規格，除非使用者要求否則不讀取。
