# Slash Command 開發規範

> 適用於所有 Agent（Manager, Backend, UI-UX）
> 確保跨模型一致性（Claude, GPT, Gemini 等）

---

## 1. 基本概念

Slash Command 是存放在 `.claude/commands/` 目錄下的 Markdown 檔案。
檔名（去掉 `.md`）即為命令名稱。

```
.claude/commands/my-command.md  →  /my-command
.claude/commands/backend/test.md  →  /test (顯示為 project:backend)
```

**Sharp Commands 為通用替代語法（所有 Agent 一律支援）**：
- 輸入 `#manager` 視同 `/manager`
- 對應規則：`#{name}` -> `.claude/commands/{name}.md`
- 執行方式：讀取對應命令檔內容並依指示執行

---

## 2. 檔案結構

### 2.1 基本結構

```markdown
---
description: 命令的簡短說明（顯示在 /help）
allowed-tools: Tool1, Tool2（可選，限制可用工具）
argument-hint: [arg1] [arg2]（可選，參數提示）
model: claude-opus-4-5（可選，指定模型）
---

# 命令標題

命令內容（作為 Prompt 注入）
```

### 2.2 Frontmatter 欄位說明

| 欄位 | 必填 | 說明 | 範例 |
|------|------|------|------|
| `description` | 建議 | 命令說明，顯示在 `/help` | `Backend Agent 初始化` |
| `allowed-tools` | 否 | 限制可用工具 | `Bash(git:*), Read` |
| `argument-hint` | 否 | 參數提示（自動補全用） | `[task-id] [priority]` |
| `model` | 否 | 指定執行模型 | `claude-opus-4-5` |
| `disable-model-invocation` | 否 | 禁止其他命令調用 | `true` |

---

## 3. 參數處理

### 3.1 全部參數：`$ARGUMENTS`

```markdown
# .claude/commands/search.md
---
description: 搜尋程式碼
---

在專案中搜尋：$ARGUMENTS
```

使用：
```
> /search validateAmount function
# $ARGUMENTS = "validateAmount function"
```

### 3.2 位置參數：`$1`, `$2`, `$3`...

```markdown
# .claude/commands/fix-issue.md
---
description: 修復 Issue
argument-hint: [issue-number] [priority]
---

修復 Issue #$1，優先級：$2
```

使用：
```
> /fix-issue 123 high
# $1 = "123", $2 = "high"
```

---

## 4. 進階功能

### 4.1 執行 Bash 命令

使用 `!` 前綴執行 shell 命令，結果會注入到 prompt：

```markdown
---
allowed-tools: Bash(git:*)
---

## 當前分支
!`git branch --show-current`

## 最近提交
!`git log --oneline -5`

請根據以上資訊...
```

### 4.2 引用檔案內容

使用 `@` 前綴引用檔案：

```markdown
請審查以下檔案：
@src/utils/validation.vb
@src/models/ExpenseModel.vb

確保符合 @skills/backend-checklist.md 的規範。
```

### 4.3 組合使用

```markdown
---
description: 完整代碼審查
allowed-tools: Bash(git:*), Read, Grep
argument-hint: [file-path]
---

## 檔案內容
@$1

## Git 歷史
!`git log --oneline -10 -- $1`

## 審查重點
參照 @skills/backend-checklist.md 進行審查。
```

---

## 5. 命令分類規範

### 5.1 Agent 初始化命令

**用途**：切換 Agent 角色

**命名**：`{agent-name}.md`

**必要內容**：
1. 角色聲明
2. 載入步驟（讀取配置、通知、任務）
3. 核心原則提醒
4. 完成後回報項目

**範例**：
```markdown
# Backend Agent 初始化

你現在是 **Backend Agent**，專注於 VB.NET 業務邏輯、SAP B1 整合。

## 角色確認

請執行以下初始化步驟：

1. **載入角色配置**：讀取 `.claude/agents/BACKEND.md`
2. **檢查任務通知**：讀取 `.claude/workspace/backend/notifications.md`
3. **確認當前任務**：讀取 `.claude/shared/active-tasks.json`
4. **載入經驗庫**：讀取相關 `skills/*.md`

## 核心原則（必須遵守）

- **POLA**：最小驚訝原則
- **WYSIWYG**：所見即所得
- **Data Consistency**：資料一致性

## 完成初始化後

回報：
1. 當前有無待處理任務
2. 是否準備好接受新任務
```

### 5.2 功能指令

**用途**：執行特定功能

**命名**：`{action}-{target}.md` 或 `{action}.md`

**必要內容**：
1. 參數說明（如有）
2. 執行步驟
3. 輸出格式

**範例**：
```markdown
---
description: 手動觸發反思
argument-hint: [-w|-m|-a]
---

# Reflect - 手動觸發反思

## 參數

- 無參數：反思最近 7 天
- `-w`：本週反思
- `-m`：本月反思
- `-a`：全部反思

## 執行步驟

1. 讀取指定範圍的 `work-logs/daily/`
2. 識別重複問題模式
3. 更新 `work-logs/insights/patterns.md`
4. 產出反思報告

## 報告格式

```markdown
# 反思報告 - YYYY-MM-DD

## 統計
- 分析範圍：...
- 任務總數：N

## 重複問題模式
| 模式 | 出現次數 |
|------|----------|
```
```

### 5.3 任務指令

**用途**：處理特定任務類型

**命名**：`{task-type}.md`

**範例**：
```markdown
---
description: 建立新任務
argument-hint: [title] [assignee]
allowed-tools: Read, Write
---

# 建立新任務

建立任務：$1
指派給：$2

## 執行步驟

1. 讀取 `.claude/shared/active-tasks.json`
2. 產生新的 task-id（格式：YYYY-MM-DD-NNN）
3. 建立 `.claude/handoff/{task-id}/spec.md`
4. 更新 `active-tasks.json`
5. 推送通知到指定 Agent
```

---

## 6. 跨模型相容性

### 6.1 設計原則

為確保不同 AI 模型都能正確執行 slash command：

1. **明確的動作指示**：使用「請執行」、「必須」等明確用語
2. **結構化步驟**：用編號列表說明步驟順序
3. **預期輸出格式**：明確定義輸出格式（Markdown、JSON 等）
4. **避免模型特定語法**：不使用只有特定模型理解的語法

### 6.2 範例對比

**不佳（模型依賴）**：
```markdown
用你的工具讀一下那個配置檔
```

**良好（跨模型相容）**：
```markdown
請執行以下步驟：

1. 使用 Read 工具讀取 `.claude/agents/BACKEND.md`
2. 確認檔案內容已載入
3. 回報載入狀態
```

### 6.3 必要的上下文注入

每個命令應包含足夠的上下文，不依賴模型的「記憶」：

```markdown
## 背景（每次執行時提供）

- 專案：JET Enterprise Platform
- 技術棧：ASP.NET Web Forms (VB.NET)
- 核心原則：POLA, WYSIWYG, Data Consistency

## 相關檔案

- 角色配置：`.claude/agents/{AGENT}.md`
- 經驗庫：`skills/*.md`
- 任務狀態：`.claude/shared/active-tasks.json`
```

---

## 7. 檢查清單

### 建立新 Slash Command 前

- [ ] 確認命令用途清晰
- [ ] 選擇適當的命令名稱
- [ ] 決定是否需要參數
- [ ] 確認所需的 allowed-tools

### 撰寫命令內容時

- [ ] 有 frontmatter（至少 description）
- [ ] 步驟有編號，順序明確
- [ ] 有預期輸出格式
- [ ] 避免模型特定語法
- [ ] 包含必要上下文

### 完成後

- [ ] 測試命令能正確執行
- [ ] 確認 `/help` 顯示正確
- [ ] 更新相關文件（如 CLAUDE.md）

---

## 8. 目前可用命令

| 命令 | 說明 | 類型 |
|------|------|------|
| `/backend` | Backend Agent 初始化 | Agent 角色 |
| `/manager` | Manager Agent 初始化 | Agent 角色 |
| `/ui-ux` | UI-UX Agent 初始化 | Agent 角色 |
| `/reflect` | 手動觸發反思 | 功能指令 |

---

## 變更紀錄

| 版本 | 日期 | 變更內容 |
|------|------|----------|
| 1.0 | 2026-01-05 | 初版，建立 slash command 開發規範 |
