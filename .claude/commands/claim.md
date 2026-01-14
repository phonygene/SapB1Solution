# 領取任務

根據藍圖領取指定的 Agent 任務，並自動初始化對應角色。

## 參數

```
/claim {agent-id} {task-id}
```

- `agent-id`：藍圖中分配的 Agent 編號（如 A1, A2）
- `task-id`：藍圖的任務 ID（如 2026-01-13-user-profile）

## 執行步驟

### 1. 讀取藍圖

讀取 `.agent-workspace/blueprints/{task-id}.md`

### 2. 找到自己的任務

從藍圖的「Agent 分配」表格中，找到 `{agent-id}` 對應的：
- **類別**：backend / ui-ux / full-stack
- **任務摘要**：要做什麼
- **涉及檔案**：要修改哪些檔案

### 3. 載入角色配置

根據類別自動載入對應配置：

| 類別 | 配置檔 |
|------|--------|
| `backend` | `.claude/agents/BACKEND.md` |
| `ui-ux` | `.claude/agents/UI-UX.md` |
| `full-stack` | `.claude/agents/SUPER.md` |

讀取配置檔，了解：
- 核心原則
- 程式碼規範
- 檢查清單

### 4. 檢查是否需要複製副本

查看藍圖的「副本追蹤」區塊，如果有自己的副本任務：
1. 建立目錄：`.agent-workspace/working/{agent-id}/{task-id}/`
2. 複製原始檔案到該目錄
3. **在副本上進行修改，而非主代碼**

如果沒有副本任務，直接修改主代碼。

### 5. 更新藍圖狀態

將藍圖中自己的狀態從 `pending` 改為 `in_progress`

### 6. 回報已領取

```
已領取任務：{agent-id} @ {task-id}

角色：{類別}
任務：{任務摘要}
涉及檔案：{檔案列表}
工作模式：{直接修改主代碼 / 在副本上修改}

準備開始執行。
```

## 執行任務時

### 無衝突（直接修改主代碼）

1. 直接修改涉及的檔案
2. 完成後 commit，message 格式：
   ```
   [{agent-id}-{類別}] {task-id}: {簡短描述}
   ```

### 有衝突（在副本上修改）

1. 在 `.agent-workspace/working/{agent-id}/{task-id}/` 修改副本
2. 完成後 commit，message 格式：
   ```
   [{agent-id}-{類別}] {task-id}: {簡短描述} (副本)
   ```
3. 副本和主代碼一起 commit

## 完成任務時

1. 更新藍圖中自己的狀態為 `completed`
2. 回報：
   ```
   任務完成：{agent-id} @ {task-id}

   修改的檔案：
   - {檔案列表}

   Commit: {commit hash}

   如果所有 Agent 都完成，請執行 /integrate {task-id}
   ```

## 注意事項

- 嚴格按照藍圖分配的範圍工作，不要越界修改其他 Agent 負責的檔案
- 如果發現藍圖有遺漏或錯誤，先回報而非自行調整
- 遵循載入的角色配置中的核心原則和規範
