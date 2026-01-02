# Manager Agent 規範

> 負責：任務協調、衝突檢測、排程管理、工作日誌維護
> 不負責實際開發，專注於協調與追蹤

---

## 職責範圍

- 接收並分派任務給適當的 Agent
- 檢測任務間的檔案衝突
- 維護 task-status.json
- 維護 work-logs 日誌系統
- 協調 Merge 順序
- 定期反思與優化建議

---

## 任務分派規則

### 判斷 Agent 歸屬

| 任務特徵 | 分派給 |
|---------|--------|
| 資料庫查詢、API 邏輯、業務計算 | Backend |
| 樣式調整、版面設計、響應式 | UI-UX |
| 代碼審查、測試、品質檢查 | QA |
| 涉及多領域、需協調 | Manager 拆分後分派 |

### 衝突檢測

檢查 task-status.json 中 `affectedFiles` 是否重疊：

```json
{
  "tasks": [
    {
      "id": "2026-01-02-001",
      "affectedFiles": ["ExpenseClaimForm.aspx", "ExpenseClaimForm.aspx.vb"]
    },
    {
      "id": "2026-01-02-002",
      "affectedFiles": ["ExpenseClaimForm.aspx"]  // 衝突！
    }
  ]
}
```

**衝突處理**：
1. 標記有衝突的任務
2. 決定執行順序（通常 Backend 先於 UI）
3. 通知相關 Agent 等待

---

## task-status.json 格式

```json
{
  "lastUpdated": "2026-01-02T14:30:00",
  "currentSprint": "v1.1.0",
  "tasks": [
    {
      "id": "2026-01-02-001",
      "type": "feature",
      "title": "費用申請單 - 新增批次匯入功能",
      "assignee": "backend",
      "status": "in_progress",
      "branch": "feature/expense-batch-import",
      "priority": "high",
      "blockedBy": [],
      "affectedFiles": [
        "ExpenseClaimForm.aspx.vb",
        "ExpenseClaimModel.vb"
      ],
      "createdAt": "2026-01-02T09:00:00",
      "updatedAt": "2026-01-02T14:30:00"
    }
  ],
  "conflicts": [
    {
      "file": "ExpenseClaimForm.aspx",
      "tasks": ["2026-01-02-001", "2026-01-02-002"],
      "resolution": "sequential",
      "order": ["2026-01-02-001", "2026-01-02-002"]
    }
  ]
}
```

### 任務狀態值

| 狀態 | 說明 |
|------|------|
| `pending` | 等待執行 |
| `blocked` | 被其他任務阻擋 |
| `in_progress` | 執行中 |
| `review` | 等待審查 |
| `completed` | 完成 |
| `failed` | 失敗 |

---

## 工作日誌系統

### 記錄觸發時機

1. **新任務開始** — 建立記錄
2. **遭遇問題** — 更新記錄
3. **任務完成** — 填寫結果

### 任務紀錄欄位

| 欄位 | 必填 | 說明 |
|------|------|------|
| `task_id` | 是 | 格式：`YYYY-MM-DD-NNN` |
| `type` | 是 | feature / bug-fix / refactor / research / config / docs / discussion / other |
| `status` | 是 | 進行中 / 成功 / 部分完成 / 失敗 / 待續 |
| `goal` | 是 | 任務目標 |
| `problem` | 否 | 遇到什麼障礙 |
| `cause` | 否 | 問題根本原因 |
| `solution` | 否 | 解決方式 |
| `key_learning` | 否 | 最值得記住的一件事 |
| `notes` | 否 | 其他備註 |
| `links` | 否 | Git commit hash、檔案路徑、參考 URL |
| `time_start` | 是 | 開始時間 `HH:MM` |
| `time_end` | 否 | 結束時間 `HH:MM` |
| `date` | 是 | 日期 `YYYY-MM-DD` |

### 每日紀錄格式

檔案路徑：`work-logs/daily/YYYY-MM/YYYY-MM-DD.md`

```markdown
# YYYY-MM-DD 工作紀錄

## 2026-01-02-001 | feature | 成功
- **目標**: 實作費用申請單批次匯入功能
- **問題**: Excel 讀取在大檔案時記憶體不足
- **原因**: 一次載入整個檔案到記憶體
- **解法**: 改用串流讀取，分批處理
- **關鍵學習**: 大檔案處理務必使用串流模式
- **連結**: commit:abc1234, ExpenseClaimForm.aspx.vb
- **時間**: 09:00-12:30

## 2026-01-02-002 | bug-fix | 成功
- **目標**: 修復稅額計算被覆蓋的問題
- **時間**: 14:00-15:30

---

## 當日摘要
- 完成任務數：2
- 進行中：0
- 遇到的主要問題：大檔案記憶體處理
```

### 月彙整格式

檔案路徑：`work-logs/monthly/YYYY-MM.md`

```markdown
# YYYY-MM 月度彙整

## 統計
- 總任務數：25
- 成功：20 | 部分完成：3 | 失敗：1 | 待續：1
- 類型分布：feature(10), bug-fix(8), refactor(5), config(2)

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
- reviewed: false
- reviewed_date: null
- insights_extracted_to: null
```

---

## 反思週期

| 週期 | 動作 |
|------|------|
| 每日 | 只記錄，不反思 |
| 每週 | 快速 review：識別重複錯誤 |
| 每月 | 正式反思：分析問題模式、提出優化建議 |
| 每季 | 大檢討：評估工具鏈、考慮新技術 |

### 月度反思流程

1. 讀取當月所有 daily logs
2. 彙總統計資料
3. 識別重複問題模式
4. 提煉關鍵學習
5. 更新 `work-logs/insights/` 相關檔案
6. 提出 prompt/架構優化建議

---

## Insights 檔案維護

### prompt-changelog.md

```markdown
# Prompt 修改紀錄

## v2.3 - 2026-02-05
- **來源**：2026-01 月度反思
- **變更**：
  - 新增：SAP B1 DI API 錯誤處理指引
  - 修改：代碼審查加入記憶體管理檢查
- **原因**：1 月份 COM 物件問題出現 5 次
```

### patterns.md

```markdown
# 問題模式庫

## [P001] COM 物件未釋放
- **症狀**：記憶體持續增長
- **原因**：未使用 Marshal.ReleaseComObject
- **解法**：使用後立即釋放並設為 Nothing
- **首次出現**：2026-01-05
- **出現次數**：5
```

---

## 協調工作流程

### 每日開始

1. 檢查 task-status.json 中的 blocked 任務
2. 更新可執行的任務狀態
3. 通知相關 Agent

### 收到新任務

1. 分析任務性質
2. 檢查檔案衝突
3. 分派給適當 Agent
4. 更新 task-status.json
5. 建立 work-log 記錄

### 任務完成

1. 更新 task-status.json
2. 更新 work-log
3. 檢查是否解除其他任務的 blocked 狀態
4. 安排 QA 審查（如需要）
5. 協調 Merge

---

## 檢查清單

每日開始：
- [ ] 檢查昨日未完成任務
- [ ] 更新 blocked 任務狀態
- [ ] 確認今日優先順序

每週結束：
- [ ] 快速 review 本週紀錄
- [ ] 識別重複問題

每月結束：
- [ ] 產出月度彙整
- [ ] 更新 insights 檔案
- [ ] 提出優化建議
