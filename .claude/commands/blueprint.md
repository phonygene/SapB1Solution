# 藍圖規劃 (Manager)

你現在是 **Manager**，負責分析任務並規劃藍圖。

## 執行步驟

### 1. 分析任務

根據用戶描述的任務，分析：
- 涉及哪些檔案？
- 需要哪些類型的工作？（backend / ui-ux / full-stack）
- 有沒有檔案可能被多個 Agent 同時修改？

### 2. 詢問 Agent 數量

根據分析結果，向用戶提出建議：

```
這個任務建議由 N 個 Agent 完成：

| Agent | 類別 | 負責內容 |
|-------|------|----------|
| A1 | backend | {描述} |
| A2 | ui-ux | {描述} |

原因：{為什麼建議這個數量}

要使用建議數量，還是調整？
```

### 3. 產出藍圖

用戶確認後，建立藍圖檔案：

**檔案路徑**：`.agent-workspace/blueprints/{task-id}.md`

**task-id 格式**：`YYYY-MM-DD-{簡短描述}`，例如 `2026-01-13-user-profile`

**藍圖內容**（參考 `_TEMPLATE.md`）：

```markdown
# 藍圖：{task-id}

> 建立時間：{當前時間}
> 狀態：in_progress

## 任務概述

{用戶需求描述}

## Agent 分配

**建議數量**：N 個 Agent
**建議原因**：{原因}
**用戶確認**：已確認

| Agent | 類別 | 任務摘要 | 涉及檔案 | 狀態 |
|-------|------|----------|----------|------|
| A1 | {類別} | {描述} | {檔案} | pending |
| A2 | {類別} | {描述} | {檔案} | pending |

## 衝突分析

（如果有多個 Agent 需要修改同一檔案，列出並標註需要複製副本）

| 檔案 | 涉及 Agent | 處理方式 |
|------|------------|----------|
| {檔案} | A1, A2 | 複製副本到各自 working/ |

## 副本追蹤

（如果有衝突，列出副本路徑）

| Agent | 副本路徑 | 原始檔案 | 狀態 |
|-------|----------|----------|------|
| A1 | .agent-workspace/working/A1/{task-id}/{檔案} | {原始路徑} | pending |

## 整合順序

1. {Agent} - {原因}
2. {Agent} - {原因}

## 驗收標準

- [ ] {標準}
```

### 4. 回報

完成後告訴用戶：

```
藍圖已建立：.agent-workspace/blueprints/{task-id}.md

可用以下命令讓 Agent 領取任務：
- /claim A1 {task-id}
- /claim A2 {task-id}

完成後用 /integrate {task-id} 整合。
```

## Agent 類別對照

| 類別 | 適用情境 |
|------|----------|
| `backend` | VB.NET 邏輯、SAP 整合、資料庫、.aspx.vb |
| `ui-ux` | ASPX 結構、CSS、JavaScript、介面設計 |
| `full-stack` | 同時涉及前後端的緊密耦合任務 |

## 注意事項

- 同一個 .aspx 和 .aspx.vb 通常高度耦合，建議給同一個 Agent
- 純 CSS 調整可以獨立給 ui-ux Agent
- 若不確定是否會衝突，寧可多分一個 Agent 並複製副本
