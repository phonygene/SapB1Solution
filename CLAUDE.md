# JET Enterprise Platform - AI 協作規範

> 此檔案為所有 Agent 共用的核心原則。
> 詳細規範請見相關目錄。

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

## 多 Agent 協作架構（2+1 變體）

### Agent 角色

| Agent | 職責 | 分支 |
|-------|------|------|
| **Manager** | 任務協調、衝突檢測、代碼審查、日誌維護 | main |
| **Backend** | VB.NET 邏輯、SAP 整合、資料庫 | agent/backend |
| **UI-UX** | 介面設計、CSS 樣式、響應式 | agent/ui-ux |

### 檔案結構

```
.claude/
├── shared/              # 全局狀態（所有 Agent 可讀）
│   ├── project-status.md
│   └── active-tasks.json
├── handoff/             # 任務交接
│   └── {task-id}/
│       ├── spec.md
│       └── output.md
├── workspace/           # 私有工作區
│   ├── backend/
│   └── ui-ux/
├── agents/              # Agent 配置
│   ├── MANAGER.md
│   ├── BACKEND.md
│   └── UI-UX.md
└── commands/            # Session Skills
    ├── sess-on.md
    ├── sess-wrap.md
    ├── sess-off.md
    └── sess-check.md

skills/                  # 動態累積的經驗
├── general-checklist.md
├── backend-checklist.md
├── ui-checklist.md
├── sap-checklist.md
└── ui-design-system.md

work-logs/               # 工作紀錄
├── daily/
├── monthly/
├── yearly/
└── insights/
```

### Session Skills

| Skill | 用途 |
|-------|------|
| `/sess-on` | 開始工作，載入狀態 |
| `/sess-wrap` | 階段存檔，繼續工作 |
| `/sess-off` | 完整存檔，結束工作 |
| `/sess-check` | 查看進度（唯讀） |

### 任務執行流程

1. Manager 分析任務，建立 `handoff/{task-id}/spec.md`
2. Agent 讀取 spec，執行任務
3. Agent 完成後寫入 `handoff/{task-id}/output.md`
4. Manager 審查，記錄到 `work-logs/`

---

## 修正/變更規則

修正錯誤時，**必須先詢問**是否檢查其他相似程式碼。

---

## 版號管理

- 格式：`X.Y.Z`（語意版號）
- 位置：`VERSION` 檔案
- 每次 Commit 前更新版號並產生 tag

---

## 相關文件

| 類別 | 位置 |
|------|------|
| Agent 配置 | `.claude/agents/` |
| Session Skills | `.claude/commands/` |
| 經驗累積 | `skills/` |
| 工作紀錄 | `work-logs/` |
| 架構規格 | `.claude/MULTI_AGENT_ARCHITECTURE_SPEC.md` |
