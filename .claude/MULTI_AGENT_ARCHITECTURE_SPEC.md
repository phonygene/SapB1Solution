# 多 Agent 協作架構：最終實踐規格

> **版本**：1.4
> **日期**：2026-01-02
> **狀態**：定稿
> **整合來源**：原始 RFC + Claude/Gemini/GPT/Manus 建議
>
> **v1.4 變更**：
> - 廢棄 Session Skills（/sess-on, /sess-wrap, /sess-off, /sess-check）
> - 改用 Git 紀錄 + skills/ + 結構化 work-logs/ 取代
> - 新增 `/reflect` 指令用於手動觸發反思
> - 保留 `.claude/commands/` 機制供未來擴充
>
> **v1.3 變更**：
> - 移除舊的 `worklog/` checkpoint 機制
> - 簡化狀態管理，減少 Token 消耗
>
> **v1.2 變更**：
> - 採用 **2+1 變體**：Backend, UI-UX + Manager/QA
> - 採用**完整 Work Logs 系統**（日/月/年彙整 + insights）
> - Layer 0 採用 Gemini 替代方案（skills/ + [AI-Context] 註解）
> - fileConflicts 降級為「提示」
> - 加入檔案監聽推送機制

---

## 1. 架構總覽

### 1.1 設計原則

1. **不重造輪子**：用 Claude Code + Git Worktrees，不引入額外框架
2. **漸進式優化**：從錯誤中學習，而非預先定義所有規則
3. **最小必要結構**：只建立立即需要的東西
4. **Token 效率**：Agent 只載入自己需要的知識
5. **單一入口**：使用者優先與 Manager 溝通；若直接指派 Agent，必須同步回報 Manager

### 1.2 核心架構圖（2+1 變體）

```
┌─────────────────────────────────────────────────────────────────┐
│                    2+1 變體架構                                  │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│                         ┌───────────────┐                        │
│                         │     User      │                        │
│                         └───────┬───────┘                        │
│                                 │                                │
│                                 ▼                                │
│                    ┌────────────────────────┐                    │
│                    │    Manager + QA        │                    │
│                    │    ────────────────    │                    │
│                    │  • 任務分派與協調       │                    │
│                    │  • 衝突檢測            │                    │
│                    │  • 代碼審查協調         │   ← main branch    │
│                    │  • Work Logs 維護       │                    │
│                    │  • Skills 更新          │                    │
│                    └────────────┬───────────┘                    │
│                                 │                                │
│               ┌─────────────────┴─────────────────┐              │
│               │                                   │              │
│               ▼                                   ▼              │
│    ┌─────────────────────┐           ┌─────────────────────┐    │
│    │      Backend        │           │       UI-UX         │    │
│    │    ───────────      │           │     ─────────       │    │
│    │  • VB.NET 邏輯      │           │  • ASPX 介面        │    │
│    │  • SAP B1 整合      │           │  • CSS 樣式         │    │
│    │  • 資料庫查詢        │           │  • 響應式設計       │    │
│    │                     │           │                     │    │
│    │  ← agent/backend    │           │  ← agent/ui-ux      │    │
│    └─────────────────────┘           └─────────────────────┘    │
│                                                                  │
│    ═══════════════════════════════════════════════════════════  │
│                                                                  │
│    Token 節省：                                                  │
│    • Backend 不載入：色彩系統、CSS 架構、設計風格                 │
│    • UI-UX 不載入：SAP COM 物件、稅務計算、資料庫結構             │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
```

### 1.3 為什麼是 2+1 而非 3+1

| 考量 | 3+1 (獨立 QA) | 2+1 變體 (Manager 兼 QA) |
|------|--------------|-------------------------|
| **QA 工作量** | 可能不夠獨立 Agent | Manager 順手做 |
| **協調成本** | 多一個 Agent 要協調 | 減少一層 |
| **審查時機** | QA 可能成為瓶頸 | Manager 在 Merge 時審查 |
| **Token 成本** | 多一份 context | 節省 |

**結論**：對於這個專案規模，QA 功能由 Manager 在任務完成時執行更實際。

---

## 2. 自定義指令系統

### 2.1 概念

`.claude/commands/` 目錄用於定義 Slash Commands，讓 Agent 執行標準化流程。
**Sharp Commands 為通用替代語法（所有 Agent 一律支援）**：
- 輸入 `#manager` 視同 `/manager`
- 對應規則：`#{name}` -> `.claude/commands/{name}.md`
- 執行方式：讀取對應命令檔內容並依指示執行

### 2.2 目前可用指令

#### /reflect（手動觸發反思）

```markdown
# Reflect - 手動觸發反思

## 參數
- 無參數：反思最近 7 天
- `-w`：本週反思
- `-m`：本月反思

## 執行步驟
1. 讀取指定範圍的 `work-logs/daily/`
2. 識別重複問題模式
3. 更新 `work-logs/insights/patterns.md`
4. 提出 `skills/` 優化建議
5. 產出反思報告
```

### 2.3 擴充方式

在 `.claude/commands/` 新增 `{command-name}.md` 即可定義新指令。

---

## 3. Work Logs 完整系統

### 3.1 目錄結構

```
work-logs/
├── daily/                          # 每日原始紀錄
│   └── YYYY-MM/
│       └── YYYY-MM-DD.md
│
├── monthly/                        # 月度彙整
│   └── YYYY-MM.md
│
├── yearly/                         # 年度彙整
│   └── YYYY.md
│
└── insights/                       # 反思與優化
    ├── prompt-changelog.md         # Prompt 修改紀錄
    ├── patterns.md                 # 問題模式庫
    └── tools-evaluated.md          # 工具評估紀錄
```

### 3.2 每日紀錄格式

`work-logs/daily/YYYY-MM/YYYY-MM-DD.md`

```markdown
# YYYY-MM-DD 工作紀錄

## 2026-01-02-001 | feature | 成功
- **目標**: 實作費用申請單批次匯入功能
- **Agent**: backend
- **問題**: Excel 讀取在大檔案時記憶體不足
- **原因**: 一次載入整個檔案到記憶體
- **解法**: 改用串流讀取，分批處理
- **關鍵學習**: 大檔案處理務必使用串流模式
- **連結**: commit:abc1234, ExpenseClaimForm.aspx.vb
- **時間**: 09:00-12:30

## 2026-01-02-002 | ui | 成功
- **目標**: 調整報表頁面響應式佈局
- **Agent**: ui-ux
- **時間**: 14:00-15:30

---

## 當日摘要
- 完成任務數：2
- 進行中：0
- 主要問題：大檔案記憶體處理
- 更新的 Skills：backend-checklist.md
```

### 3.3 月度彙整格式

`work-logs/monthly/YYYY-MM.md`

```markdown
# YYYY-MM 月度彙整

## 統計
- 總任務數：25
- 成功：20 | 部分完成：3 | 失敗：1 | 待續：1
- 類型分布：feature(10), bug-fix(8), refactor(5), config(2)
- Agent 分布：backend(15), ui-ux(8), manager(2)

## 未完成任務
| Task ID | 類型 | 目標 | 狀態 |
|---------|------|------|------|
| 2026-01-15-003 | feature | 報表匯出功能 | 待續 |

## 失敗/部分完成案例
### 2026-01-10-002
[完整紀錄內容]

## 重複問題模式
- COM 物件未釋放：出現 3 次
  - 相關任務：2026-01-05-001, 2026-01-12-002, 2026-01-20-001

## 關鍵學習清單
- 大檔案處理務必使用串流模式 (from 2026-01-02-001)
- UpdatePanel 內的按鈕需在 Triggers 中註冊 (from 2026-01-08-003)

## 反思狀態
- reviewed: true
- reviewed_date: 2026-02-01
- insights_extracted_to: patterns.md, prompt-changelog.md
```

### 3.4 反思週期

| 週期 | 動作 | 觸發方式 |
|------|------|----------|
| 每日 | 只記錄，不反思 | 自然發生（Git commit + work-logs） |
| 每週 | 快速 review：識別重複錯誤 | `/reflect -w` 或手動 |
| 每月 | 正式反思：分析問題模式、更新 skills | `/reflect -m` 或手動 |
| 每季 | 大檢討：評估工具鏈、考慮新技術 | 手動 |

### 3.5 Insights 檔案

#### patterns.md（問題模式庫）

```markdown
# 問題模式庫

## [P001] COM 物件未釋放
- **症狀**：記憶體持續增長
- **原因**：未使用 Marshal.ReleaseComObject
- **解法**：使用後立即釋放並設為 Nothing
- **首次出現**：2026-01-05
- **出現次數**：5
- **相關任務**：2026-01-05-001, 2026-01-12-002...

## [P002] CSS 變數名稱不一致
- **症狀**：樣式套用失敗、主題切換無效
- **原因**：不同檔案使用不同的變數命名慣例
- **解法**：統一使用 jet-color-themes.css 定義的變數名稱
- **首次出現**：2026-01-02
- **出現次數**：1
```

#### prompt-changelog.md（Prompt 修改紀錄）

```markdown
# Prompt 修改紀錄

## v2.3 - 2026-02-05
- **來源**：2026-01 月度反思
- **變更**：
  - 新增：SAP B1 DI API 錯誤處理指引
  - 修改：代碼審查加入記憶體管理檢查
- **原因**：1 月份 COM 物件問題出現 5 次
```

---

## 4. Layer 0 替代方案（採納 Gemini 建議）

### 4.1 為什麼不需要預先建立 Layer 0

| 原因 | 說明 |
|------|------|
| **SAP 邏輯已穩定** | 現有 .vb 檔案就是最好的參考 |
| **邊際效益遞減** | 整理完整 sap-mapping.json 的時間成本過高 |
| **不影響協作流程** | 協作流程依賴 Layer 1 和 skills/ |

### 4.2 替代方案：用 skills/ 動態累積規則

```markdown
# skills/backend-checklist.md

## 代碼規範
- [ ] 變數命名用 camelCase
- [ ] 函數命名用 PascalCase

## 從錯誤中學習（持續新增）
- [ ] 2026-01-02：驗證邏輯要包含負數檢查
- [ ] 2026-01-05：COM 物件使用後必須釋放
```

### 4.3 用 [AI-Context] 註解取代 mapping 文件

```vb
' [AI-Context] SAP Table: OEXD (費用申請主表), 欄位: DocTotal
Dim totalAmount As Decimal = ...

' [AI-Context] 這個控制項對應 SAP 欄位 U_ExpType
Protected WithEvents ddlExpenseType As DropDownList
```

---

## 5. 推送機制（採納 Manus 建議）

### 5.1 推送 vs 輪詢

| 方式 | 問題 |
|------|------|
| **輪詢** | 浪費資源、延遲高、需人工提醒 |
| **推送** | 即時響應、零人工介入、低成本 |

### 5.2 推送機制流程

```
┌─────────────────────────────────────────────────────────────────┐
│                     推送機制運作流程                              │
├─────────────────────────────────────────────────────────────────┤
│                                                                  │
│   Manager                              Backend/UI-UX             │
│   ───────                              ──────────────            │
│       │                                     │                    │
│       │ 1. 寫入 spec.md                     │                    │
│       │    echo >> notifications.md ──────► │ ← inotifywait      │
│       │                                     │    監聽到變更       │
│       │                                     │                    │
│       │                                     │ 2. 執行任務        │
│       │                                     │                    │
│       │ ◄────────────────────────────────── │ 3. 寫入 output.md  │
│       │    inotifywait 監聽到               │                    │
│       │                                     │                    │
│       │ 4. 審查 & 更新狀態                   │                    │
│       ▼                                     ▼                    │
│                                                                  │
└─────────────────────────────────────────────────────────────────┘
```

### 5.3 監聽工具

| 平台 | 工具 | 安裝 |
|------|------|------|
| Linux/WSL | `inotifywait` | `apt-get install inotify-tools` |
| macOS | `fswatch` | `brew install fswatch` |
| Windows | PowerShell FileSystemWatcher 或 WSL | — |

### 5.4 Agent 狀態機（避免輸入中斷）

所有 Agent 使用 `.claude/workspace/{agent}/current.md` 的「## 狀態」欄位：
- `idle`：等候狀態（可接收新指令）
- `thinking`：處理中（不可被打斷）

規則：
1. littlebird 送出訊息前必須確認目標為 `idle`
2. littlebird 送出訊息後將狀態設為 `thinking`
3. Agent 完成處理後必須手動改回 `idle`

### 5.5 視窗聚焦備援（滑鼠點擊）

若 `SetForegroundWindow` 失敗，可在 `AGENT_WINDOWS` 設定 `click`：
```python
"backend": {
    "class_name": "CASCADIA_HOSTING_WINDOW_CLASS",
    "title": "Backend",
    "hotkey": ("ctrl", "alt", "1"),
    "click": {"monitor": 1, "x": 200, "y": 80}
},
```
`monitor` 為螢幕索引（從 `--list-windows` 輸出查看），`x/y` 為該螢幕內座標。

---

## 6. Agent 配置（2+1 變體）

### 6.1 MANAGER.md（兼 QA 協調）

```markdown
# Manager Agent（兼 QA 協調）

## 角色
你是專案的協調者，負責任務分析、拆分、追蹤、審查和團隊優化。

## 職責
1. **任務管理**：接收、分析、拆分、分派任務
2. **衝突檢測**：檢查檔案衝突，決定並行或等待
3. **代碼審查**：任務完成時執行簡易審查或觸發詳細審查
4. **日誌維護**：維護 Work Logs 系統
5. **經驗累積**：更新 skills/ 和 insights/

## 檔案監聽機制

### 啟動監聽
```bash
inotifywait -m -r -e close_write --format '%w%f' .claude/handoff/ | while read FILE
do
    if [[ "$FILE" == *"output.md" ]]; then
        TASK_ID=$(echo "$FILE" | grep -oP 'handoff/\K[^/]+')
        echo "🔔 Task $TASK_ID 完成！"
    fi
done &
```

## 工作流程

### 收到新任務時
1. 分析任務涉及的檔案
2. 判斷耦合度：
   - 高耦合（同頁面）→ 指派給單一 Agent
   - 低耦合（不同模組）→ 可並行指派
3. 建立 `handoff/{task-id}/spec.md`
4. 注入相關 skills
5. 更新 `active-tasks.json`
6. 推送通知到目標 Agent：
   ```bash
   echo "新任務 $TASK_ID" >> .claude/workspace/{agent}/notifications.md
   ```

### 任務完成時（審查流程）
1. 讀取 `handoff/{task-id}/output.md`
2. **簡易審查**（Manager 執行）：
   - 檢查是否符合 spec
   - 檢查 skills/ 中的 checklist
   - 確認無明顯問題
3. **如有疑慮**：要求 Agent 補充說明或修正
4. 審查通過後：
   - 更新 `active-tasks.json` 為 completed
   - 記錄到 `work-logs/daily/`
   - 協調 Git merge

### 定期反思
- **每週**：識別重複問題
- **每月**：正式反思，更新 skills/ 和 insights/

## 讀取範圍
- .claude/shared/*
- .claude/handoff/*
- .claude/workspace/*（監控用）
- skills/*
- work-logs/*

## 寫入範圍
- .claude/shared/*
- .claude/handoff/*/spec.md
- .claude/workspace/*/notifications.md
- work-logs/*
- skills/*
```

### 6.2 BACKEND.md

```markdown
# Backend Agent

## 角色
你是後端開發者，負責 VB.NET 業務邏輯和 SAP B1 整合。

## 不需要知道的事
- CSS 架構和主題系統
- 色彩設計規範
- 響應式佈局細節
- UI 設計原則

## 通知監聽機制

### 啟動監聽
```bash
inotifywait -m -e close_write --format '%w%f' .claude/workspace/backend/ | while read FILE
do
    if [[ "$FILE" == *"notifications.md" ]]; then
        echo "🔔 收到新通知！"
        tail -1 .claude/workspace/backend/notifications.md
    fi
done &
```

## 工作流程

### 開始任務前
1. 讀取 `handoff/{task-id}/spec.md`
2. 讀取 `skills/backend-checklist.md`
3. 讀取 `skills/sap-checklist.md`
4. 檢查代碼中的 [AI-Context] 註解

### 執行任務
1. 在 agent/backend 分支工作
2. 遵循 skills/ 中的檢查清單
3. 每個邏輯變更都要 commit（附 task-id）
4. 遇到新的 SAP 欄位，加上 [AI-Context] 註解

### 完成任務
1. 寫入 `handoff/{task-id}/output.md`
2. 列出修改的檔案
3. 列出新增的 [AI-Context] 註解
4. 標註任何發現的問題或風險

## 注意事項
- 修改 .aspx.vb 時檢查 .aspx 是否需要同步
- SAP Service Layer 呼叫前檢查 Session
- 金額用 Decimal，不用 Double
- 錯誤處理要完整（Try-Catch + 記錄）

## 讀取範圍
- .claude/handoff/{自己的任務}/*
- .claude/shared/active-tasks.json
- .claude/workspace/backend/notifications.md
- skills/backend-checklist.md
- skills/sap-checklist.md
- skills/general-checklist.md

## 寫入範圍
- .claude/handoff/{自己的任務}/output.md
- .claude/workspace/backend/*
- 專案代碼（在 agent/backend 分支）
```

### 6.3 UI-UX.md

```markdown
# UI-UX Agent

## 角色
你是前端開發者，負責 ASP.NET Web Forms 的介面和樣式。

## 不需要知道的事
- SAP B1 COM 物件管理
- 稅務計算邏輯
- 資料庫 Schema 細節
- 後端業務邏輯

## 通知監聽機制

### 啟動監聽
```bash
inotifywait -m -e close_write --format '%w%f' .claude/workspace/ui-ux/ | while read FILE
do
    if [[ "$FILE" == *"notifications.md" ]]; then
        echo "🔔 收到新通知！"
        tail -1 .claude/workspace/ui-ux/notifications.md
    fi
done &
```

## 工作流程

### 開始任務前
1. 讀取 `handoff/{task-id}/spec.md`
2. 讀取 `skills/ui-checklist.md`
3. 如果有 Backend 的 output.md，讀取了解 API 規格

### 執行任務
1. 在 agent/ui-ux 分支工作
2. 遵循 skills/ 中的檢查清單
3. 每個邏輯變更都要 commit（附 task-id）

### 完成任務
1. 寫入 `handoff/{task-id}/output.md`
2. 列出修改的檔案
3. 標註任何發現的問題或風險

## 注意事項
- 修改 .aspx 時檢查 .aspx.vb 的事件綁定
- 修改 .aspx 時更新 .aspx.designer.vb 控制項宣告
- PostBack 後要重新綁定下拉選單
- CSS 修改要考慮其他頁面的影響
- 使用 jet-color-themes.css 的變數名稱

## 讀取範圍
- .claude/handoff/{自己的任務}/*
- .claude/shared/active-tasks.json
- .claude/workspace/ui-ux/notifications.md
- skills/ui-checklist.md
- skills/general-checklist.md

## 寫入範圍
- .claude/handoff/{自己的任務}/output.md
- .claude/workspace/ui-ux/*
- 專案代碼（在 agent/ui-ux 分支）
```

---

## 7. 完整檔案結構

```
/your-project/
│
├── .claude/
│   ├── shared/                      # Layer 1: 全局狀態
│   │   ├── project-status.md
│   │   ├── active-tasks.json
│   │   └── blocked.json
│   │
│   ├── handoff/                     # Layer 2: 任務交接
│   │   └── {task-id}/
│   │       ├── spec.md
│   │       └── output.md            # ← 完成時寫入，觸發 Manager
│   │
│   ├── workspace/                   # Layer 3: 私有工作區
│   │   ├── backend/
│   │   │   ├── current.md
│   │   │   ├── notes.md
│   │   │   └── notifications.md     # ← Manager 推送通知
│   │   └── ui-ux/
│   │       ├── current.md
│   │       ├── notes.md
│   │       └── notifications.md     # ← Manager 推送通知
│   │   └── manager/
│   │       └── current.md            # ← 狀態機（idle/thinking）
│   │
│   ├── agents/                      # Agent 配置（2+1）
│   │   ├── MANAGER.md               # 兼 QA 協調
│   │   ├── BACKEND.md
│   │   └── UI-UX.md
│   │
│   └── commands/                    # 自定義指令
│       └── reflect.md               # 手動觸發反思
│
├── skills/                          # 動態累積的經驗
│   ├── general-checklist.md
│   ├── backend-checklist.md
│   ├── ui-checklist.md
│   ├── sap-checklist.md
│   └── examples/
│
├── work-logs/                       # 結構化工作紀錄
│   ├── daily/
│   │   └── YYYY-MM/
│   │       └── YYYY-MM-DD.md
│   ├── monthly/
│   │   └── YYYY-MM.md
│   ├── yearly/
│   │   └── YYYY.md
│   └── insights/
│       ├── prompt-changelog.md
│       ├── patterns.md
│       └── tools-evaluated.md
│
└── (專案代碼...)
```

---

## 8. 檔案格式規格

### 8.1 active-tasks.json

```json
{
  "lastUpdated": "2026-01-02T10:00:00+08:00",
  "tasks": [
    {
      "id": "2026-01-02-001",
      "title": "ExpenseClaimForm 金額驗證",
      "assignee": "backend",
      "status": "in_progress",
      "priority": "high",
      "blockedBy": [],
      "blocking": ["2026-01-02-002"],
      "affectedFiles": ["MgmSP/ExpenseClaimForm.aspx.vb"],
      "createdAt": "2026-01-02T09:00:00+08:00",
      "updatedAt": "2026-01-02T10:00:00+08:00"
    }
  ],
  "fileConflicts": {
    "MgmSP/ExpenseClaimForm.aspx.vb": {
      "lockedBy": "2026-01-02-001",
      "agent": "backend",
      "since": "2026-01-02T09:30:00+08:00",
      "type": "warning"
    }
  }
}
```

**fileConflicts.type = "warning"**：軟提示，Agent 可繼續工作，完成後用 Git merge 解決。

### 8.2 spec.md

```markdown
# Task: 2026-01-02-001

## 目標
為 ExpenseClaimForm 新增金額驗證功能

## 非目標
- 不修改 SAP 匯入邏輯
- 不調整 UI 佈局

## 涉及檔案
- MgmSP/ExpenseClaimForm.aspx.vb

## 接口契約
- 驗證函數：ValidateExpenseAmount(amount As Decimal) As Boolean
- 錯誤訊息顯示在 lblError 控制項

## 相關 Skills（Manager 注入）
- skills/backend-checklist.md
- skills/sap-checklist.md

## 過往相關問題（Manager 注入）
- 2025-12-15：金額驗證漏掉負數檢查
```

### 8.3 output.md

```markdown
# Task: 2026-01-02-001 - 完成報告

## 完成時間
2026-01-02 12:30

## 修改的檔案
- MgmSP/ExpenseClaimForm.aspx.vb（+45 行）

## 實作摘要
- 新增 ValidateExpenseAmount 函數
- 檢查負數、零、超過上限
- 錯誤訊息顯示在 lblError

## 新增的 [AI-Context] 註解
- Line 234: SAP 欄位 U_MaxAmount 對應

## 測試結果
- 輸入 -100 → 顯示錯誤 ✓
- 輸入 0 → 顯示錯誤 ✓
- 輸入 999999999 → 顯示錯誤 ✓
- 輸入 1000 → 通過 ✓

## 風險/備註
- 無
```

---

## 9. 實作步驟

### Step 1: 安裝監聽工具

```bash
# Ubuntu/Debian/WSL
sudo apt-get install inotify-tools

# macOS
brew install fswatch
```

### Step 2: 建立目錄結構

```bash
# 在專案根目錄執行
mkdir -p .claude/{shared,handoff,workspace/{backend,ui-ux},agents,commands}
mkdir -p skills/examples
mkdir -p work-logs/{daily,monthly,yearly,insights}

# 建立初始檔案
echo '{"lastUpdated":"","tasks":[],"fileConflicts":{}}' > .claude/shared/active-tasks.json
touch .claude/shared/project-status.md
touch .claude/shared/blocked.json
touch .claude/workspace/backend/notifications.md
touch .claude/workspace/ui-ux/notifications.md
touch skills/general-checklist.md
touch skills/backend-checklist.md
touch skills/ui-checklist.md
touch skills/sap-checklist.md
```

### Step 3: 建立 Git Worktrees

```bash
mkdir -p ../worktrees
git worktree add ../worktrees/backend -b agent/backend
git worktree add ../worktrees/ui-ux -b agent/ui-ux
git worktree list
```

### Step 4: 配置 Agent 定義

將 MANAGER.md、BACKEND.md、UI-UX.md 存到 `.claude/agents/`

### Step 5: 配置自定義指令

將 reflect.md 存到 `.claude/commands/`

### Step 6: 啟動 Agent Sessions

```bash
# Terminal 1: Manager（主專案）
cd /your/project
claude
# 告知 AI：你是 Manager Agent，請讀取 .claude/agents/MANAGER.md

# Terminal 2: Backend（worktree）
cd ../worktrees/backend
claude
# 告知 AI：你是 Backend Agent，請讀取 .claude/agents/BACKEND.md

# Terminal 3: UI-UX（如需並行）
cd ../worktrees/ui-ux
claude
# 告知 AI：你是 UI-UX Agent，請讀取 .claude/agents/UI-UX.md
```

---

## 10. 總結

### 採納的決策

| 項目 | 決策 |
|------|------|
| **Agent 架構** | 2+1 變體（Backend, UI-UX + Manager/QA） |
| **狀態追蹤** | Git 紀錄 + `.claude/shared/` |
| **工作紀錄** | 結構化 Work Logs（daily → monthly → yearly + insights） |
| **經驗累積** | skills/ 動態累積 + [AI-Context] 註解 |
| **反思機制** | `/reflect` 手動觸發 |
| **衝突處理** | fileConflicts 作為警告，用 Git merge 解決 |

### 執行優先順序

1. ✅ 確認架構設計
2. ✅ 建立 .claude/ 目錄結構
3. ⏳ 建立 Git Worktrees
4. ✅ 配置 Agent 定義
5. ✅ 配置自定義指令
6. ✅ 初始化 skills/
7. ⏳ 按需在代碼中加 [AI-Context] 註解
8. ⏳ 從實踐中累積經驗

---

## 變更紀錄

| 版本 | 日期 | 變更內容 |
|------|------|----------|
| 1.0 | 2026-01-02 | 初稿（4 Agent） |
| 1.1 | 2026-01-02 | Layer 0 替代方案、fileConflicts 降級、推送機制 |
| 1.2 | 2026-01-02 | 2+1 變體架構、完整 Work Logs |
| 1.3 | 2026-01-02 | 移除 worklog/ checkpoint |
| 1.4 | 2026-01-02 | 廢棄 Session Skills，改用 Git + skills + work-logs，新增 /reflect |
