# Claude Code 協作流程與架構規範包

> [DEPRECATED] 本規範包為歷史文件，已由專案內 `C:\Projects\SapB1Solution\.claude\MULTI_AGENT_ARCHITECTURE_SPEC.md` 取代。僅供參考，不再作為現行規範。

**版本**：2.0
**更新日期**：2025-11-14
**適用對象**：使用 Claude Code 進行專案開發的開發者

---

## 📋 簡介

這是一套針對 Claude Code 開發的**協作流程與架構規範**，專為提升 AI 輔助開發效率而設計。

### 核心特色

✅ **Shopfloor 協作模式** - 檔案輸出優先，避免對話中貼大段程式碼，節省 Token
✅ **Session 管理機制** - 完整的工作狀態追蹤與恢復，支援長期專案開發
✅ **責任歸屬規範** - 清晰的錯誤處理與溝通規範，提升協作品質
✅ **靈活可配置** - 支援各種程式語言、框架與專案類型
✅ **完整的文件架構** - 從初始化到日常協作的完整流程

### 適用場景

- 需要與 Claude Code 進行長期專案開發
- 希望建立結構化的 AI 協作流程
- 需要追蹤工作進度與狀態
- 追求高效、低 Token 消耗的協作方式
- 多人團隊希望統一 AI 協作規範

---

## 🚀 快速開始（5 分鐘設定）

### 步驟 1：複製檔案到你的專案

```bash
# 進入你的專案目錄
cd /path/to/your/project

# 複製協作規範目錄
cp -r /path/to/collaboration-package/agent-os ./

# 複製 Slash Commands（選用）
cp -r /path/to/collaboration-package/examples/.claude ./

# 建立工作目錄
mkdir -p shopfloor/Claude_TMP/{sql,code,etc}
mkdir -p worklog
```

### 步驟 2：初始化你的專案配置

執行初始化助手（推薦）：

```bash
# 在 Claude Code 中執行
/init-project
```

或手動編輯 `agent-os/SESSION_INIT.md`，填寫：
- 專案名稱
- 技術環境（語言、框架、資料庫等）
- 專案結構
- 語言偏好與時區

### 步驟 3：開始協作

在 Claude Code 中執行：

```
/sess-on
```

Claude 會讀取你的配置並開始協作！

---

## 📦 套件內容

```
collaboration-package/
├── README.md                          # 本檔案
├── QUICK_START.md                     # 詳細安裝指南
├── CONFIGURATION_GUIDE.md             # 配置引導手冊
├── agent-os/                          # 協作規範核心
│   ├── SESSION_INIT.md                # Session 初始化清單（需配置）
│   └── standards/global/              # 全域標準規範
│       ├── session-management.md      # Session 管理機制
│       ├── workflow-standards.md      # Shopfloor 協作模式
│       ├── communication-standards.md # 溝通與責任歸屬
│       ├── localization.md            # 語言與術語標準
│       └── coding-style.md            # 程式碼風格規範
├── examples/                          # 範例檔案
│   ├── .claude/commands/              # Slash Commands 範例
│   ├── shopfloor/Claude_TMP/          # 檔案輸出範例
│   ├── worklog/                       # 工作日誌範例
│   └── configs/                       # 各種語言/框架配置範例
│       ├── python-fastapi/            # Python + FastAPI
│       ├── python-django/             # Python + Django
│       ├── javascript-nodejs/         # JavaScript + Node.js
│       └── typescript-react/          # TypeScript + React
└── docs/                              # 詳細文件
    ├── shopfloor-workflow.md          # Shopfloor 協作模式詳解
    ├── session-guide.md               # Session 管理指南
    └── faq.md                         # 常見問題
```

---

## 🎯 核心概念

### 1. Shopfloor 協作模式

**問題**：傳統方式中，Claude 在對話中貼大段程式碼，造成：
- Token 消耗大
- 對話視覺疲勞
- 複製貼上困難

**解決方案**：所有程式碼產出先輸出到 `shopfloor/Claude_TMP/`

**工作流程**：
```
1. Claude 產生檔案到 shopfloor/Claude_TMP/
2. Claude 在對話中簡要說明（檔案清單、使用順序）
3. 你在 IDE 中開啟檔案並執行
4. 簡短回報結果（"執行成功" / "報錯：..."）
5. Claude 繼續下一步
```

**優點**：
- 節省 Token（不在對話中顯示大量代碼）
- 檔案可直接在 IDE 使用（語法高亮、程式碼提示）
- 保留完整歷史記錄
- 易於版本控制

詳細說明：`docs/shopfloor-workflow.md`

### 2. Session 管理機制

**四個核心指令**：
- `/sess-on` - 上班/開始工作（讀取上次狀態）
- `/sess-check` - 檢查進度（不寫檔案）
- `/sess-wrap` - 階段存檔，繼續工作
- `/sess-off` - 完整存檔並下班

**狀態追蹤**：
- `worklog/LastCheckPoint.log` - 最新工作狀態
- `TODO.md` - 未完成的待辦事項

**使用場景**：
```
早上開工：
你: /sess-on
Claude: [讀取昨天的進度，報告待辦事項]

工作中斷：
你: /sess-wrap
Claude: [存檔當前進度，可隨時離開]

下班前：
你: /sess-off
Claude: [完整存檔，總結今天工作]
```

詳細說明：`docs/session-guide.md`

### 3. 溝通與責任歸屬

**正確的責任歸屬**：
- ✅ AI 提供的代碼有錯 → AI 承認並修正
- ✅ 友善但明確地指出使用者錯誤
- ❌ 推卸責任或含糊其辭

**方案呈現原則**：
- 主方案詳述（完整程式碼 + 優點）
- 其他方案簡述（一句話 + 優缺點對比）
- 節省時間，聚焦推薦方案

詳細規範：`agent-os/standards/global/communication-standards.md`

---

## 🔧 配置指南

### 方式 1：使用初始化助手（推薦）

複製 Slash Command 後執行：

```
/init-project
```

Claude 會引導你完成：
1. 專案名稱與描述
2. 技術環境選擇（從範例中選擇或自訂）
3. 目錄結構配置
4. 語言與時區偏好
5. 協作模式參數

完成後自動生成 `agent-os/SESSION_INIT.md`

### 方式 2：從範例複製（快速）

```bash
# 選擇適合你的範例配置
cp examples/configs/python-fastapi/SESSION_INIT.md agent-os/

# 編輯並調整為你的專案
vim agent-os/SESSION_INIT.md
```

### 方式 3：手動配置（進階）

參考 `CONFIGURATION_GUIDE.md` 手動編輯各個規範檔案。

---

## 📚 範例配置

我們提供了多種語言/框架的配置範例：

### Python + FastAPI
```bash
cp -r examples/configs/python-fastapi/* ./
```
適合：RESTful API、微服務、後端開發

### Python + Django
```bash
cp -r examples/configs/python-django/* ./
```
適合：全端 Web 應用、內容管理系統

### JavaScript + Node.js
```bash
cp -r examples/configs/javascript-nodejs/* ./
```
適合：後端 API、即時應用

### TypeScript + React
```bash
cp -r examples/configs/typescript-react/* ./
```
適合：前端 SPA、React 應用

**沒有你的語言？** 參考 `CONFIGURATION_GUIDE.md` 建立自訂配置

---

## 💡 使用範例

### 範例 1：開始新的工作 Session

```
你: /sess-on

Claude: 📋 Session 初始化完成

專案：MyAPI (FastAPI 專案)
上次工作時間：2025-11-13 18:30 UTC

未完成的待辦事項：
1. [高] 完成使用者認證 API
2. [中] 撰寫單元測試
3. [低] 更新 API 文件

⚠️ 協作模式提醒：
本專案採用 Shopfloor 協作模式，所有程式碼會先輸出到 shopfloor/Claude_TMP/

建議下一步：繼續完成使用者認證 API

請問要從哪裡開始？
```

### 範例 2：Shopfloor 協作流程

```
你: 請建立使用者註冊 API

Claude: 📁 已產生檔案到 shopfloor/Claude_TMP/：

SQL 腳本（1 個）：
1. sql/01_create_users_table.sql - 建立使用者資料表

Python 檔案（3 個）：
1. code/user_model.py - SQLAlchemy 資料模型
2. code/user_schema.py - Pydantic 驗證綱要
3. code/user_router.py - FastAPI 路由

📋 使用順序：
1. 執行 SQL 腳本
2. 將 Python 檔案加入對應目錄
3. 在 main.py 中註冊路由

請執行後回報結果。

你: SQL 執行成功，檔案已加入

Claude: 很好！接下來產生單元測試...
```

### 範例 3：中斷與恢復工作

```
你: /sess-wrap

Claude: 📝 階段存檔完成

當前進度：
✅ 使用者註冊 API - 已完成
⏳ 單元測試 - 進行中（60%）

已更新 worklog/LastCheckPoint.log

[稍後...]

你: /sess-on

Claude: 歡迎回來！
上次進度：單元測試進行中（60%）
建議繼續：完成剩餘的測試案例
```

---

## ❓ 常見問題

### Q1: 我的專案不是 Python，可以使用嗎？

**A:** 可以！這套規範與程式語言無關。請：
1. 參考 `examples/configs/` 中其他語言的範例
2. 或參考 `CONFIGURATION_GUIDE.md` 建立自訂配置
3. 調整 `shopfloor/Claude_TMP/` 的子目錄名稱（如 `javascript/`、`csharp/` 等）

### Q2: 一定要用 Slash Commands 嗎？

**A:** 不一定。你可以：
- 使用 Slash Commands：`/sess-on`
- 使用純文字指令：`Claude, sess on.`
- 直接請 Claude：`請讀取 agent-os/SESSION_INIT.md 並開始協作`

### Q3: Shopfloor 目錄可以改名嗎？

**A:** 可以！只需要：
1. 修改目錄名稱（如改為 `claude-output/`）
2. 在 `agent-os/SESSION_INIT.md` 中更新路徑
3. 在協作規範中同步更新

### Q4: 如何在團隊中使用？

**A:** 建議：
1. 將 `agent-os/` 目錄加入版本控制
2. 團隊成員都複製相同的配置
3. 在專案 README 中說明協作流程
4. 定期檢視和更新規範

### Q5: 需要全部採用嗎？

**A:** 不需要！你可以：
- 只使用 Shopfloor 協作模式
- 只使用 Session 管理
- 只使用溝通規範
- 或全部採用

彈性選擇適合你的部分。

### Q6: 與其他 AI 工具相容嗎？

**A:** 部分相容。核心概念（Shopfloor、Session 管理）可以適用於其他 AI 工具，但 Slash Commands 是 Claude Code 專屬功能。

---

## 📖 進階主題

### 自訂協作規範

參考 `CONFIGURATION_GUIDE.md` 了解：
- 如何調整 Shopfloor 目錄結構
- 如何自訂 Session 指令
- 如何擴充溝通規範
- 如何整合到 CI/CD

### 整合 MCP Server（選用）

如果你的專案需要直接操作資料庫，可以整合 MCP Server。

參考範例：`examples/mcp-server/`（提供 SQL Server 範例）

### 多語言專案

如果專案包含多種語言（如前後端分離），建議：
- 在 `shopfloor/Claude_TMP/` 建立語言子目錄
- 在 `SESSION_INIT.md` 中說明各語言的規範
- 參考 `examples/configs/fullstack/` 範例

---

## 🤝 貢獻與回饋

這套協作規範是開源的，歡迎：
- ⭐ Star 與分享
- 🐛 回報問題
- 💡 提出改進建議
- 🔧 貢獻新的語言/框架範例
- 📝 分享使用心得

---

## 📜 授權

本套件採用 **MIT License**，你可以自由使用、修改和分享。

---

## 🙏 致謝

感謝所有使用並提供回饋的開發者。

特別感謝 Claude (Anthropic) 在實際專案協作中的持續優化。

---

## 📞 資源連結

- **完整文件**：`docs/`
- **配置指南**：`CONFIGURATION_GUIDE.md`
- **快速開始**：`QUICK_START.md`
- **範例專案**：`examples/configs/`
- **常見問題**：`docs/faq.md`

---

**最後更新**：2025-11-14
**版本**：2.0
**維護者**：開源社群

祝你與 Claude 的協作愉快！🎉
