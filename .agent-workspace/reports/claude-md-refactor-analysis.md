# CLAUDE.md 冗餘內容分析報告（修訂版）

> 修訂說明：採用「Glob 優先 + 關鍵用途說明」策略，平衡精簡與可理解性

## 總覽

| 分類 | 行數 | 佔比 | 處理方式 |
|------|------|------|----------|
| 🟢 核心保留 | ~60 行 | 21% | 保留或微調 |
| 🟡 需精簡 | ~40 行 | 14% | 縮減篇幅 |
| 🔴 冗餘移除 | ~186 行 | 65% | 移到參考文檔或刪除 |

---

## 策略說明

### 原本的問題

| 方式 | 問題 |
|------|------|
| 完整樹狀圖 | 佔用 42 行，易過時，每次都載入浪費 token |
| 完全移除 | AI 不知道資源存在，Glob 後也不懂用途 |

### 採用的策略：Glob 優先 + 關鍵用途說明

1. **加入 Glob 優先規則**：要求 AI 查找資源時直接 Glob 目錄，不要猜測路徑
2. **只保留用途說明**：對於檔名不自解釋的關鍵路徑，說明其用途
3. **移除完整結構圖**：檔名自解釋的資源不需要列出

### 效果

| 項目 | 原始 | 精簡後 |
|------|------|--------|
| 路徑相關行數 | 42 行 | ~10 行 |
| AI 能找到資源 | ✅ | ✅（靠 Glob） |
| AI 能理解用途 | ✅ | ✅（靠用途說明） |

---

## 冗餘區塊 1：檔案結構圖（第 63-104 行）

### 原文

```text
### 檔案結構

.claude/                 # 工作規範（需確認權限）
├── shared/              # 全局狀態（所有 Agent 可讀）
│   ├── project-status.md
│   └── active-tasks.json
├── agents/              # Agent 配置
│   ├── MANAGER.md
│   ├── BACKEND.md
│   └── UI-UX.md
└── commands/            # 自定義指令
    └── reflect.md       # 手動觸發反思

.agent-workspace/        # Agent 溝通區（免確認權限）
├── manager/
│   ├── current.md       # 狀態檔
│   └── notifications.md
├── backend/
│   ├── current.md
│   └── notifications.md
├── ui-ux/
│   ├── current.md
│   └── notifications.md
└── handoff/             # 任務交接
    └── {task-id}/
        ├── spec.md
        └── output.md

skills/                  # 動態累積的經驗
├── general-checklist.md
├── backend-checklist.md
├── ui-checklist.md
├── sap-checklist.md
└── ui-design-system.md

work-logs/               # 結構化工作紀錄
├── daily/
├── monthly/
├── yearly/
└── insights/
```

### 冗餘原因

1. **完整樹狀圖不需要**：大部分檔名已自解釋（如 `backend-checklist.md`）
2. **可用 Glob 取代**：AI 直接查看目錄即可
3. **但用途說明仍需保留**：如 `current.md` 的用途不明顯

### 新版本

```markdown
## 資源查找

🔵 查找配置或參考資料時，直接 Glob 相關目錄：
- `.claude/**/*.md` — 工作規範、Agent 配置、指令
- `skills/*.md` — 經驗累積、Checklist
- `work-logs/daily/*.md` — 工作日誌

### 關鍵路徑用途

| 路徑模式 | 用途 |
|----------|------|
| `.agent-workspace/{agent}/current.md` | Agent 狀態（idle/thinking） |
| `.agent-workspace/handoff/{id}/` | 跨 Agent 任務交接（spec.md → output.md） |
| `.claude/shared/` | 跨 Agent 共享狀態 |
```

**效果**：42 行 → 12 行

---

## 冗餘區塊 2：Slash Commands 完整說明（第 106-138 行）

### 原文

```markdown
### 自定義指令（Slash Commands / Sharp Commands）

Slash Command 是存放在 `.claude/commands/` 目錄下的 Markdown 檔案。
**Sharp Command 為通用替代語法（所有 Agent 一律支援）**：
- 輸入 `#manager` 視同 `/manager`
- 對應規則：`#{name}` -> `.claude/commands/{name}.md`
- 執行方式：讀取對應命令檔內容並依指示執行

詳細開發規範請見 `skills/slash-command-standards.md`。

#### 目前可用指令

| 指令 | 用途 | 類型 |
|------|------|------|
| `/manager` | Manager Agent 初始化 | Agent 角色 |
| `/backend` | Backend Agent 初始化 | Agent 角色 |
| `/ui-ux` | UI-UX Agent 初始化 | Agent 角色 |
| `/reflect` | 手動觸發反思，識別問題模式 | 功能指令 |

#### 快速建立新指令

# .claude/commands/{command-name}.md
---
description: 命令說明（顯示在 /help）
argument-hint: [arg1] [arg2]
---

命令內容...
使用 $ARGUMENTS 或 $1, $2 接收參數
使用 @filepath 引用檔案
使用 !`command` 執行 shell
```

### 冗餘原因

1. **教程性質**：建立新指令的說明不需要每次載入
2. **指令清單可 Glob 取得**：從 `.claude/commands/` 目錄讀取
3. **可用指令表格多餘**：Glob 後檔名已自解釋

### 新版本

```markdown
### Slash Commands
- 位置：`.claude/commands/`
- `#name` 等同 `/name`
- 查看可用指令：Glob `.claude/commands/*.md`
```

**效果**：33 行 → 4 行

---

## 冗餘區塊 3：狀態追蹤機制（第 140-145 行）

### 原文

```markdown
### 狀態追蹤機制

- **Git 紀錄**：工作歷史的自然追蹤
- **`.claude/shared/`**：全局狀態（active-tasks.json, project-status.md）
- **`work-logs/`**：結構化工作紀錄（daily → monthly → yearly + insights）
- **`skills/`**：動態經驗累積
```

### 冗餘原因

1. **與資源查找重複**：已整合到新的「資源查找」章節
2. **純描述性**：沒有可執行的規則
3. **Git 追蹤是常識**：不需要特別說明

### 新版本

**刪除整段**

**效果**：6 行 → 0 行

---

## 冗餘區塊 4：Agent 狀態機詳細說明（第 147-176 行）

### 原文

```markdown
### Agent 狀態機

所有 Agent 使用 `.agent-workspace/{agent}/current.md` 的「## 狀態」欄位：
- `idle`：等候狀態（可接收新指令）
- `thinking`：處理中（不可被打斷）

規則：
- littlebird 送出訊息前必須確認目標為 `idle`
- littlebird 送出訊息後將狀態設為 `thinking`
- **Agent 每次回覆結束前必須將狀態改回 `idle`**（由 AI 自行執行）

#### 強制規則（所有 Agent 必須遵守）

**無論回覆內容為何，每次回覆結束前必須執行：**

| Agent | 狀態檔案路徑 |
|-------|-------------|
| Manager | `.agent-workspace/manager/current.md` |
| Backend | `.agent-workspace/backend/current.md` |
| UI-UX | `.agent-workspace/ui-ux/current.md` |

將該檔案的「## 狀態」欄位改為 `idle`。

### 任務執行流程

1. Manager 分析任務，建立 `.agent-workspace/handoff/{task-id}/spec.md`
2. Agent 讀取 spec，執行任務
3. Agent 完成後寫入 `.agent-workspace/handoff/{task-id}/output.md`
4. Manager 審查，記錄到 `work-logs/`
```

### 冗餘原因

1. **規則重複三次**：狀態機說明、強制規則、路徑表格說的是同一件事
2. **littlebird 規則與 AI 無關**：這是外部腳本的邏輯
3. **任務執行流程是參考**：不是每次都需要的強制規則
4. **格式不利執行**：敘述性而非觸發式

### 新版本

```markdown
### Agent 狀態規則
🔴 **回覆結束時**：更新 `.agent-workspace/{agent}/current.md` 的「## 狀態」為 `idle`
```

**效果**：30 行 → 2 行

---

## 冗餘區塊 5：自動批准操作清單（第 198-253 行）

### 原文

（56 行的完整權限清單，此處省略）

### 冗餘原因

1. **系統層級控制**：權限由 `.claude/settings.json` 控制，Claude Code 自動執行
2. **AI 不需要知道**：AI 不會根據這個清單決定要不要問——系統會自動處理
3. **佔用 56 行**：完全是給人類看的文檔記錄

### 新版本

**完全刪除**

**效果**：56 行 → 0 行

---

## 冗餘區塊 6：封存政策詳細說明（第 257-270 行）

### 原文

```markdown
## 封存政策

### specArchive/ 目錄

存放過時的規格和歷史紀錄。**除非使用者明確要求，否則不要讀取此目錄**。

此目錄已在 `.claudeignore` 中設定為忽略，Claude Code 預設不會載入。

### .claudeignore

類似 `.gitignore`，列出 Claude Code 預設忽略的路徑：
- `specArchive/` - 封存的舊規格
- `bin/`, `obj/` - 編譯輸出
- `.vs/` - IDE 設定
```

### 冗餘原因

1. **`.claudeignore` 是系統層級**：Claude Code 自動讀取
2. **重複說明**：同一件事說了兩次

### 新版本

```markdown
## 封存
`specArchive/` 存放過時規格，除非使用者要求否則不讀取。
```

**效果**：14 行 → 2 行

---

## 冗餘區塊 7：相關文件索引（第 274-285 行）

### 原文

```markdown
## 相關文件

| 類別 | 位置 |
|------|------|
| Agent 配置 | `.claude/agents/` |
| 自定義指令 | `.claude/commands/` |
| **Slash Command 規範** | `skills/slash-command-standards.md` |
| 經驗累積 | `skills/` |
| 全局狀態 | `.claude/shared/` |
| 工作紀錄 | `work-logs/` |
| 架構規格 | `.claude/MULTI_AGENT_ARCHITECTURE_SPEC.md` |
| 封存 | `specArchive/`（預設忽略） |
```

### 冗餘原因

1. **已整合到資源查找**：新的「資源查找」章節已包含這些資訊
2. **Glob 可取代**：AI 可以直接查看目錄

### 新版本

**刪除整段**

**效果**：12 行 → 0 行

---

## 需精簡區塊：版號管理（第 190-194 行）

### 原文

```markdown
## 版號管理

- 格式：`X.Y.Z`（語意版號）
- 位置：`VERSION` 檔案
- 每次 Commit 前更新版號並產生 tag
```

### 問題

1. **進位規則不明確**：沒說明 X、Y、Z 各自的進位條件
2. **導致誤判**：Agent 把 1.1.9 + 維護 誤進位成 1.2.0（應為 1.1.10）

### 新版本

```markdown
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
```

**效果**：規則明確，避免誤判

---

## 需精簡區塊：修正/變更規則（第 179-186 行）

### 原文

```markdown
## 修正/變更規則

修正錯誤時，**必須先詢問**是否檢查其他相似程式碼。

## 需求確認規則

在調整行為或策略前，**必須先確認使用者目的與接受的取捨**，再執行修改。
不得以「移除功能」作為預設解法來規避錯誤，除非使用者明確同意。
```

### 問題

1. **觸發條件模糊**：「修正錯誤時」、「調整行為或策略前」不夠具體
2. **格式不統一**：兩個規則分成兩個章節

### 新版本

```markdown
## 行為約束

🔴 **修正錯誤時**：先詢問是否檢查其他相似程式碼
🔴 **變更策略前**：先確認使用者目的與可接受的取捨
🔴 **禁止**：以「移除功能」規避錯誤（除非使用者明確同意）
```

**效果**：格式統一，觸發點更明確

---

## 精簡後預估

| 項目 | 原始 | 精簡後 |
|------|------|--------|
| 總行數 | 286 行 | ~85 行 |
| Token 消耗 | ~2000 | ~550 |
| 核心規則佔比 | 14% | 80% |

---

## 新版 CLAUDE.md 預覽

```markdown
# JET Enterprise Platform - AI 協作規範

## 專案概述
- **技術棧**：ASP.NET Web Forms (VB.NET)
- **性質**：財務系統，資料正確性極為重要
- **風險**：錯誤可能導致帳務與法律問題

## 核心開發原則

| 原則 | 實踐方式 |
|------|----------|
| **POLA** | 不在背景修改用戶輸入的值 |
| **WYSIWYG** | 儲存前不重新計算已顯示的值 |
| **Data Consistency** | 避免在 Save 時修改 Model |
| **Form State Integrity** | Sync 函數只讀取 UI，不重算 |

🔴 **禁止**：在 `SaveDocument` 或 `SyncToModel` 中重新計算金額/稅額

## 技術規則

| 規則 | 說明 |
|------|------|
| 控制項宣告 | `.aspx` 新增控制項時**必須同時更新** `.aspx.designer.vb` |
| Namespace | .aspx 的 Inherits **必須**加前綴 `MgmSP.` |
| 檔案編碼 | UTF-8 with BOM |
| 資料庫連線 | 本地 `jtdbConnectionString`、SAP `SapSQLConnection` |

## Agent 協作

| Agent | 職責 | 分支 |
|-------|------|------|
| Manager | 任務協調、代碼審查、日誌維護 | main |
| Backend | VB.NET、SAP 整合、資料庫 | agent/backend |
| UI-UX | 介面設計、CSS、響應式 | agent/ui-ux |

- 使用者優先與 Manager 溝通
- 直接指派其他 Agent 時，該 Agent 必須同步回報 Manager

### 資源查找

🔵 查找配置或參考資料時，直接 Glob 相關目錄：
- `.claude/**/*.md` — 工作規範、Agent 配置、指令
- `skills/*.md` — 經驗累積、Checklist
- `work-logs/daily/*.md` — 工作日誌

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

## 行為約束

🔴 **修正錯誤時**：先詢問是否檢查其他相似程式碼
🔴 **變更策略前**：先確認使用者目的與可接受的取捨
🔴 **禁止**：以「移除功能」規避錯誤（除非使用者同意）

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

## 封存
`specArchive/` 存放過時規格，除非使用者要求否則不讀取。
```

---

---

## 新增機制：skills/INDEX.md 自維護索引

### 設計目標

讓 AI 明確知道「什麼情況該參考什麼資源」，並在變更 skills/ 時自動維護索引。

### 運作流程

```
AI 開始任務
    ↓
讀取 skills/INDEX.md
    ↓
比對觸發條件 → 找到匹配 → 讀取對應資源
    ↓
執行任務
    ↓
如果變更了 skills/ → 更新 INDEX.md
```

### INDEX.md 結構

```markdown
# Skills 資源索引

> 🔴 **維護規則**：變更 skills/ 任何檔案時，必須同步更新此索引

## 觸發條件對照表

| 觸發條件 | 資源檔案 | 說明 |
|----------|----------|------|
| 建立/修改 Slash Command | slash-command-standards.md | 指令格式規範 |
| 後端開發（VB.NET、資料庫） | backend-checklist.md | 後端開發檢查項目 |
| UI/前端開發 | ui-checklist.md | 前端開發檢查項目 |
| SAP B1 整合 | sap-checklist.md | SAP 整合注意事項 |
| UI 設計（顏色、樣式） | ui-design-system.md | 設計規範與元件 |
| 所有開發任務 | general-checklist.md | 通用檢查項目 |
```

### CLAUDE.md 中的引用

```markdown
### 資源查找

🔵 **執行任務前**：讀取 `skills/INDEX.md` 查看是否有相關資源
🔴 **變更 skills/ 時**：必須同步更新 `skills/INDEX.md`
```

---

## 執行計畫

1. ✅ 更新報告（本文件）
2. ⏳ 建立 `skills/INDEX.md`
3. ⏳ 重寫精簡版 `CLAUDE.md`
4. ⏳ 驗證並 commit push
