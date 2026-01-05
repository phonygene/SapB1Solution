# 多 Agent 協作架構建議方案

> **來源**：Claude Opus 4.5 分析
> **日期**：2026-01-02
> **針對**：MULTI_AGENT_ARCHITECTURE_RFC.md v0.1

---

## 1. RFC 評估摘要

### 1.1 設計優點（值得保留）

| 設計元素 | 評價 | 說明 |
|----------|------|------|
| **資訊分層架構** | ✅ 優秀 | Layer 1/2/3 的分層概念與業界最佳實踐吻合 |
| **衝突檢測機制** | ✅ 務實 | fileConflicts 追蹤是必要的 |
| **work-logs 系統** | ✅ 核心價值 | 經驗累積是持續優化的關鍵 |
| **明確不做的事項** | ✅ 正確判斷 | 避免 AI 開會、不要額外基礎設施 |
| **Agent 角色定義** | ✅ 合理 | Backend/UI-UX/QA/Manager 分工清晰 |

### 1.2 需要調整的部分

| 設計元素 | 問題 | 建議 |
|----------|------|------|
| **Handoff 機制** | 手動文件交接效率低 | 改用現有工具的自動化 handoff |
| **Manager 實現** | 未說明如何自動化執行 | 整合現有 orchestration 工具 |
| **並行執行方式** | 未明確技術實現 | 使用 Git Worktrees 隔離 |
| **狀態同步頻率** | 未確定具體方案 | 採用 event-driven + scheduled 混合 |

---

## 2. 核心發現：現有工具生態

### 2.1 專為 Claude Code 設計的 Orchestration 工具

經過搜尋，發現已有多個成熟工具可以實現 RFC 的設計目標：

#### 2.1.1 claude-flow（推薦，功能最完整）

**GitHub**: https://github.com/ruvnet/claude-flow

**特點**：
- 企業級 AI orchestration 平台
- Hive-mind swarm intelligence（Queen + Worker 架構）
- Persistent memory system
- 100+ MCP tools
- Dynamic Agent Architecture (DAA) 自組織 agents

**對應 RFC 設計**：
```
RFC 設計              → claude-flow 對應
─────────────────────────────────────────
Manager Agent        → Queen (自動協調)
Backend/UI-UX/QA     → Worker Agents
Layer 1 (全局狀態)    → Persistent Memory
Layer 2 (handoff)    → MCP Tools
work-logs            → Memory System
```

**快速開始**：
```bash
# 初始化 hive-mind 系統
npx claude-flow@alpha hive-mind wizard
npx claude-flow@alpha hive-mind spawn "build expense claim feature" --claude

# Session 管理
npx claude-flow@alpha hive-mind status
npx claude-flow@alpha hive-mind resume session-xxxxx

# Memory 查詢
npx claude-flow@alpha memory query "expense form" --recent
```

#### 2.1.2 ccswarm（Rust，高效能）

**GitHub**: https://github.com/nwiizo/ccswarm

**特點**：
- Rust-native patterns，零成本抽象
- ProactiveMaster 自動分析和委派
- Specialized Agent Pool（Frontend/Backend/DevOps/QA）
- Template System 支援 variable substitution
- WebSocket 即時通訊

**對應 RFC 設計**：
```
RFC 設計              → ccswarm 對應
─────────────────────────────────────────
Manager Agent        → ProactiveMaster
Backend Agent        → backend agent
UI-UX Agent          → frontend agent
QA Agent             → QA agent
handoff 文件格式      → Template System
```

**快速開始**：
```bash
# 初始化專案
ccswarm init --name "ExpenseClaim" --agents frontend,backend,qa

# 啟動系統
ccswarm start

# 新增任務（自動分派）
ccswarm task "Create expense form validation" --priority high --type feature

# 代理任務到特定 Agent
ccswarm delegate task "Add expense API endpoint" --agent backend --priority high

# 查看狀態
ccswarm status --detailed
```

#### 2.1.3 Tmux Orchestrator（輕量級）

**來源**: https://ktwu01.github.io/posts/2025/08/tmux-orchestrator/

**特點**：
- 輕量級，適合實驗
- Self-trigger：Agents 自己排程 check-ins
- Cross-project coordination
- 工作持續即使關閉 laptop

**快速開始**：
```bash
# 發送訊息給任何 Claude agent
./send-claude-message.sh backend:0 "What's your progress on the API?"
./send-claude-message.sh ui-ux:1 "The form layout is ready for styling"
./send-claude-message.sh project-manager:0 "Please coordinate with QA team"

# 自動排程 check-in（分鐘）
./schedule_with_note.sh 30 "Review implementation, assign next task"
./schedule_with_note.sh 60 "Check test coverage, merge if passing"
```

#### 2.1.4 其他相關工具

| 工具 | 特點 | 適用場景 |
|------|------|----------|
| **parallel-cc** | 自動 worktree 管理 | 快速啟動並行 session |
| **Crystal** | Desktop app，支援 Claude + Codex | 視覺化管理多 session |
| **claude-orchestrator** | Redis pub/sub 協調 | 大規模分散式 |
| **claude-code-by-agents** | @mentions 路由任務 | 簡單的任務分派 |

### 2.2 並行執行的基礎：Git Worktrees

所有工具的並行執行都基於 **Git Worktrees** 來避免衝突：

```bash
# 建立 worktrees
git worktree add ../project-backend feature/backend-api
git worktree add ../project-ui-ux feature/ui-redesign
git worktree add ../project-qa feature/test-coverage

# 每個 worktree 啟動獨立的 Claude Code session
cd ../project-backend && claude
cd ../project-ui-ux && claude
cd ../project-qa && claude

# 完成後合併
git worktree remove ../project-backend
```

**優點**：
- 每個 Agent 完全隔離的檔案狀態
- 共享同一個 Git history
- 不會互相覆蓋修改
- 合併時用標準 Git merge

---

## 3. 建議架構方案

### 3.1 方案 A：claude-flow 整合（推薦）

```
┌─────────────────────────────────────────────────────────────┐
│                    claude-flow hive-mind                    │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│   ┌─────────────┐                                           │
│   │   Queen     │  ← Manager Agent                          │
│   │ (自動協調)   │    • 任務分析和委派                        │
│   └──────┬──────┘    • 衝突檢測                             │
│          │           • 進度追蹤                             │
│          │           • 定期反思和優化                        │
│    ┌─────┴─────┬─────────────┐                              │
│    ▼           ▼             ▼                              │
│ ┌───────┐ ┌───────┐    ┌───────┐                           │
│ │Backend│ │ UI-UX │    │  QA   │  ← Worker Agents          │
│ │ Agent │ │ Agent │    │ Agent │    (git worktree 隔離)     │
│ └───────┘ └───────┘    └───────┘                           │
│     │          │            │                               │
│     └──────────┴────────────┘                               │
│              │                                              │
│              ▼                                              │
│     ┌─────────────────┐                                     │
│     │ Persistent      │  ← work-logs + 狀態管理             │
│     │ Memory System   │                                     │
│     └─────────────────┘                                     │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

**優點**：
- 功能最完整，已有 persistent memory
- MCP 整合，可擴展到其他工具
- 社群活躍，持續更新

**缺點**：
- 學習曲線較高
- 需要 Node.js 環境

### 3.2 方案 B：ccswarm 整合

```
┌─────────────────────────────────────────────────────────────┐
│                      ccswarm System                         │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│   ┌───────────────────┐                                     │
│   │  ProactiveMaster  │  ← Manager Agent                    │
│   │  (Goal-Driven)    │    • 任務分析                       │
│   └────────┬──────────┘    • 智能委派                       │
│            │               • 品質審查整合                    │
│     ┌──────┴──────┬────────────┐                            │
│     ▼             ▼            ▼                            │
│ ┌────────┐  ┌────────┐  ┌────────┐                         │
│ │Backend │  │Frontend│  │   QA   │                         │
│ │ Agent  │  │ Agent  │  │ Agent  │                         │
│ └────────┘  └────────┘  └────────┘                         │
│                                                             │
│   ┌───────────────────────────────────────┐                 │
│   │         Template System               │                 │
│   │  • Task Templates (變數替換)          │                 │
│   │  • Code Generation                    │                 │
│   │  • Documentation Templates            │                 │
│   └───────────────────────────────────────┘                 │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

**優點**：
- Rust 高效能
- 架構與你的 RFC 非常接近
- TUI 介面方便監控

**缺點**：
- 需要編譯 Rust
- 相對較新，文件較少

### 3.3 方案 C：保留 RFC 結構 + Tmux Orchestrator

如果你想保持對架構的完全控制，可以用你現有的 RFC 設計 + Tmux Orchestrator：

```
┌─────────────────────────────────────────────────────────────┐
│                    你的 RFC 架構                            │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  .claude/                                                   │
│  ├── shared/                 ← Layer 1: 全局狀態            │
│  │   ├── project-status.md                                  │
│  │   ├── active-tasks.json                                  │
│  │   └── blocked.json                                       │
│  │                                                          │
│  ├── handoff/                ← Layer 2: 任務交接            │
│  │   └── {task-id}/                                         │
│  │       ├── spec.md                                        │
│  │       └── output.md                                      │
│  │                                                          │
│  ├── workspace/              ← Layer 3: 私有工作區          │
│  │   ├── backend/                                           │
│  │   ├── ui-ux/                                             │
│  │   └── qa/                                                │
│  │                                                          │
│  └── agents/                 ← Agent 配置                   │
│      ├── MANAGER.md                                         │
│      ├── BACKEND.md                                         │
│      ├── UI-UX.md                                           │
│      └── QA.md                                              │
│                                                             │
│  + Tmux Orchestrator                                        │
│  ├── send-claude-message.sh  ← Agent 間通訊                 │
│  └── schedule_with_note.sh   ← 定期反思排程                 │
│                                                             │
│  + Git Worktrees                                            │
│  ├── ../project-backend/     ← Backend Agent 工作區         │
│  ├── ../project-ui-ux/       ← UI-UX Agent 工作區           │
│  └── ../project-qa/          ← QA Agent 工作區              │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

**優點**：
- 完全符合你的 RFC 設計
- 最大彈性
- 最小依賴

**缺點**：
- 需要自己寫協調邏輯
- Manager 自動化程度取決於你的實現

---

## 4. 關於 LangGraph / 其他 AI Flow 的建議

### 4.1 不建議使用 LangGraph 的理由

| 考量 | 說明 |
|------|------|
| **技術棧不匹配** | LangGraph 是 Python 生態，你的主力是 .NET |
| **已有專用工具** | claude-flow、ccswarm 專為 Claude Code 設計 |
| **過度抽象** | LangGraph 的 DAG 結構對你的需求來說過於複雜 |
| **生產環境問題** | API 經常變動、debugging 困難 |

### 4.2 如果真的需要 Orchestration Framework

| 生態系統 | 推薦工具 | 說明 |
|----------|----------|------|
| **Claude Code 專用** | claude-flow / ccswarm | ✅ 最推薦 |
| **.NET 生態** | Microsoft Agent Framework | Semantic Kernel + AutoGen 整合 |
| **Python 生態** | LangGraph | 如果你願意用 Python |

### 4.3 Microsoft Agent Framework（.NET 選項）

如果你想用 .NET 技術棧實現 Manager：

```csharp
// Semantic Kernel Handoff Orchestration 範例
using Microsoft.SemanticKernel.Agents.Orchestration.Handoff;

var handoffs = new OrchestrationHandoffs()
    .AddMany(
        sourceAgent: managerAgent.Name,
        targetAgents: new Dictionary<string, string>
        {
            { backendAgent.Name, "Transfer for backend API development" },
            { uiUxAgent.Name, "Transfer for UI/UX implementation" },
            { qaAgent.Name, "Transfer for quality assurance" }
        }
    );

var orchestration = new HandoffOrchestration(
    members: [managerAgent, backendAgent, uiUxAgent, qaAgent],
    handoffs: handoffs
);
```

**優點**：
- 與你的 .NET 技術棧整合
- 微軟官方支援
- 企業級功能

---

## 5. 實作建議

### 5.1 推薦的實作路徑

```
Phase 1: 驗證（1 週）
├── 選擇工具（claude-flow 或 ccswarm）
├── 用小任務測試多 Agent 並行
└── 驗證 Git Worktrees 工作流程

Phase 2: 整合 RFC 設計（2 週）
├── 建立 .claude/ 目錄結構
├── 配置 Agent prompts
├── 設定 work-logs 格式
└── 整合 Manager 自動化

Phase 3: 優化（持續）
├── 收集 work-logs 數據
├── 分析瓶頸和改進機會
├── 優化 Agent prompts
└── 調整工作流程
```

### 5.2 建議的 work-logs 格式（結構化）

```json
{
  "date": "2026-01-02",
  "session_id": "exp-claim-001",
  "tasks": [
    {
      "id": "TASK-001",
      "title": "ExpenseClaimForm 色彩系統",
      "agent": "ui-ux",
      "status": "completed",
      "duration_minutes": 45,
      "files_changed": [
        "MgmSP/ExpenseClaimForm.aspx",
        "css/Jet.css"
      ],
      "blockers": [],
      "handoff_to": "qa",
      "learnings": [
        "CSS 變數在 WebForms 中需要特殊處理",
        "IE11 不支援 CSS Grid，需要 fallback"
      ]
    }
  ],
  "agent_metrics": {
    "backend": { "tasks_completed": 2, "avg_duration": 35 },
    "ui-ux": { "tasks_completed": 1, "avg_duration": 45 },
    "qa": { "tasks_completed": 3, "avg_duration": 20 }
  },
  "reflections": {
    "what_went_well": "並行開發節省了等待時間",
    "what_could_improve": "handoff 文件需要更詳細的 API 規格",
    "prompt_adjustments": [
      { "agent": "backend", "change": "加入更多錯誤處理指引" }
    ],
    "next_priority": "整合測試"
  }
}
```

### 5.3 Manager Agent 的核心職責

```markdown
# MANAGER.md

## 角色定義
你是專案的 Manager Agent，負責協調 Backend、UI-UX、QA 三個 Agent 的工作。

## 核心職責

### 1. 任務分派
- 分析新任務，決定由哪個 Agent 執行
- 識別任務依賴關係
- 設定優先級

### 2. 衝突檢測
- 監控 active-tasks.json 中的 fileConflicts
- 當偵測到衝突時，暫停相關任務
- 協調執行順序

### 3. 進度追蹤
- 定期檢查每個 Agent 的 workspace/*/current.md
- 更新 project-status.md
- 識別阻塞項目

### 4. 定期反思（每 30 分鐘）
- 讀取 work-logs
- 分析完成率、阻塞原因、常見問題
- 生成改進建議
- 必要時調整 Agent prompts

### 5. 團隊優化
- 追蹤長期趨勢
- 識別重複出現的問題
- 建議流程改進

## 通訊協議
- 讀取：.claude/shared/* (全局狀態)
- 讀取：.claude/workspace/*/* (Agent 私有區，監控用)
- 寫入：.claude/shared/project-status.md
- 寫入：.claude/handoff/*/spec.md (任務規格)
- 寫入：work-logs/daily/*.json
```

---

## 6. 總結

### 6.1 RFC 評價

你的 RFC 設計**邏輯正確且完整**，特別是：
- 資訊分層架構
- Agent 角色定義
- work-logs 經驗累積

需要補充的是**具體實現方式**，特別是 Manager 的自動化執行機制。

### 6.2 工具推薦

| 優先級 | 工具 | 理由 |
|--------|------|------|
| 1️⃣ | **claude-flow** | 功能最完整，社群活躍 |
| 2️⃣ | **ccswarm** | 架構與 RFC 最接近 |
| 3️⃣ | **Tmux Orchestrator + 自建** | 最大彈性 |

### 6.3 不建議

- ❌ LangGraph（技術棧不匹配）
- ❌ 純 Claude Code subagent（無法實現自主 Manager）
- ❌ 從零自建所有功能（重造輪子）

---

## 附錄：參考資源

### 工具文件
- claude-flow: https://github.com/ruvnet/claude-flow
- ccswarm: https://github.com/nwiizo/ccswarm
- parallel-cc: https://github.com/frankbria/parallel-cc
- Tmux Orchestrator: https://ktwu01.github.io/posts/2025/08/tmux-orchestrator/

### 最佳實踐
- Anthropic Claude Code Best Practices: https://www.anthropic.com/engineering/claude-code-best-practices
- Multi-agent Research System: https://www.anthropic.com/engineering/multi-agent-research-system
- Microsoft Agent Framework: https://learn.microsoft.com/en-us/agent-framework/overview/agent-framework-overview

### 相關討論
- Git Worktrees with Claude Code: https://incident.io/blog/shipping-faster-with-claude-code-and-git-worktrees
- Multi-Agent Orchestration Patterns: https://sjramblings.io/multi-agent-orchestration-claude-code-when-ai-teams-beat-solo-acts/
