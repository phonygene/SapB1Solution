# JET Enterprise Platform - AI 協作規範

> 此檔案為所有 Agent 共用的核心原則。
> 專業領域規範請見 `.claude/agents/` 目錄。

## 專案概述

- **技術棧**：ASP.NET Web Forms (VB.NET)
- **性質**：財務相關系統，資料正確性極為重要
- **風險**：錯誤可能導致帳務與法律問題

---

## 核心開發原則

| 原則 | 說明 | 實踐方式 |
|------|------|----------|
| **POLA** | 最小驚訝原則 | 不在背景偷偷修改用戶輸入的值 |
| **WYSIWYG** | 所見即所得 | 儲存前不重新計算已顯示的值 |
| **Data Consistency** | 資料一致性 | 避免在 Save 時修改 Model 的值 |
| **Form State Integrity** | 表單狀態完整性 | Sync 函數只讀取 UI，不重算 |

**違反原則的常見錯誤**：
- 在 `SaveDocument` 或 `SyncToModel` 中重新計算金額/稅額
- 用戶手動修改的值被程式覆蓋
- PostBack 後數值被重算

---

## 關鍵技術規則

### ASP.NET 控制項宣告
在 `.aspx` 新增控制項時，**必須同時更新** `.aspx.designer.vb`。

### Namespace 規則
- Root Namespace：`MgmSP`
- .aspx 的 Inherits **必須**加前綴：`Inherits="MgmSP.ClassName"`

### 檔案編碼
所有檔案必須使用 **UTF-8 with BOM**。

### 資料庫連線
- 本地：`jtdbConnectionString`
- SAP：`SapSQLConnection`

---

## 多 Agent 協作規範

### Agent 角色

| Agent | 職責 | 分支前綴 |
|-------|------|----------|
| Backend | 功能開發、API、資料庫 | `feature/` |
| UI-UX | 介面設計、樣式、響應式 | `ui/` |
| QA | 代碼審查、測試、品質檢查 | `test/` |
| Manager | 任務協調、衝突檢測、排程 | - |

### 任務執行流程

1. 執行前讀取 `.claude/task-status.json` 確認無衝突
2. 更新狀態為 `in_progress`
3. 執行任務
4. 完成後更新狀態並記錄到 `work-logs/`

### 工作日誌規範

每個任務必須記錄到 `work-logs/daily/YYYY-MM/YYYY-MM-DD.md`：
- 新任務開始時建立記錄
- 遭遇問題時更新記錄
- 任務完成時填寫結果

詳細格式見 `.claude/agents/MANAGER.md`。

---

## 修正/變更規則

修正錯誤時，**必須先詢問**是否檢查其他相似程式碼。

---

## 版號管理

- 格式：`X.Y.Z`（超過 9 直接進位）
- 位置：`VERSION` 檔案
- 每次 Commit 前更新版號並產生 tag

---

## 專業領域規範

| 領域 | 檔案 |
|------|------|
| 後端/資料庫 | `.claude/agents/BACKEND.md` |
| UI/UX 設計 | `.claude/agents/UI-UX.md` |
| 品質保證 | `.claude/agents/QA.md` |
| 任務協調 | `.claude/agents/MANAGER.md` |
