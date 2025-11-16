# 快速開始指南

**預計時間**：10-15 分鐘
**適用對象**：第一次使用本協作規範的開發者

---

## 🎯 目標

完成本指南後，你將：
- ✅ 在你的專案中建立協作規範結構
- ✅ 配置適合你專案的設定
- ✅ 能夠使用 `/sess-on` 開始與 Claude 協作
- ✅ 了解 Shopfloor 協作模式的基本流程

---

## 📋 前置需求

- 已安裝 Claude Code
- 有一個正在開發的專案（或準備開始新專案）
- 5-15 分鐘的時間

---

## 🚀 安裝步驟

### 步驟 1：複製檔案結構（2 分鐘）

在你的專案根目錄執行：

```bash
# 進入你的專案目錄
cd /path/to/your/project

# 複製協作規範核心目錄
cp -r /path/to/collaboration-package/agent-os ./

# 複製 Slash Commands（選用但推薦）
cp -r /path/to/collaboration-package/examples/.claude ./

# 建立工作目錄
mkdir -p shopfloor/Claude_TMP/{sql,code,etc}
mkdir -p worklog

# 建立待辦事項檔案
touch TODO.md
```

**Windows 使用者**：
```cmd
xcopy /E /I collaboration-package\agent-os your-project\agent-os
xcopy /E /I collaboration-package\examples\.claude your-project\.claude
mkdir your-project\shopfloor\Claude_TMP\sql
mkdir your-project\shopfloor\Claude_TMP\code
mkdir your-project\shopfloor\Claude_TMP\etc
mkdir your-project\worklog
type nul > your-project\TODO.md
```

完成後你的專案結構應該是：
```
your-project/
├── agent-os/
│   ├── SESSION_INIT.md
│   └── standards/global/
├── .claude/
│   └── commands/
├── shopfloor/Claude_TMP/
│   ├── sql/
│   ├── code/
│   └── etc/
├── worklog/
└── TODO.md
```

---

### 步驟 2：選擇配置方式（選一個）

#### 方式 A：從範例快速配置（推薦，3 分鐘）

我們提供了多種語言/框架的配置範例，選擇一個最接近你專案的：

**Python + FastAPI**（RESTful API、微服務）
```bash
cp examples/configs/python-fastapi/SESSION_INIT.md agent-os/
```

**Python + Django**（全端 Web 應用）
```bash
cp examples/configs/python-django/SESSION_INIT.md agent-os/
```

**JavaScript + Node.js**（後端 API）
```bash
cp examples/configs/javascript-nodejs/SESSION_INIT.md agent-os/
```

**TypeScript + React**（前端 SPA）
```bash
cp examples/configs/typescript-react/SESSION_INIT.md agent-os/
```

然後編輯 `agent-os/SESSION_INIT.md`，修改：
- 專案名稱
- 當前階段
- 資料庫資訊（如果有）
- 其他專案特定資訊

#### 方式 B：使用初始化助手（自動化，5 分鐘）

在 Claude Code 中執行：

```
/init-project
```

Claude 會引導你完成配置，包括：
1. 詢問專案名稱
2. 選擇技術環境（從範例中選擇或自訂）
3. 配置目錄結構
4. 設定語言偏好
5. 自動生成 `agent-os/SESSION_INIT.md`

#### 方式 C：手動配置（進階使用者，10 分鐘）

參考 `CONFIGURATION_GUIDE.md` 逐步配置所有設定。

---

### 步驟 3：初始化 TODO 與工作日誌（2 分鐘）

**初始化 TODO.md**：

```bash
# 複製範例
cp examples/TODO.md ./

# 或自己建立
cat > TODO.md << 'EOF'
# 專案待辦事項

**專案名稱**：[你的專案名稱]
**最後更新**：[日期]

## 🔥 高優先級
- [ ] [待辦事項1]

## ⚡ 中優先級
- [ ] [待辦事項2]

## 💡 低優先級
- [ ] [待辦事項3]
EOF
```

**初始化工作日誌**：

```bash
cat > worklog/LastCheckPoint.log << 'EOF'
# 工作檢查點日誌

**最後更新**：[日期]
**Session 狀態**：未開始

---

## 當前專案狀態

**專案名稱**：[你的專案名稱]
**當前階段**：[開發階段]
**完成度**：0%

---

## 待辦事項

請參考 TODO.md
EOF
```

---

### 步驟 4：測試配置（2 分鐘）

在 Claude Code 中執行：

```
/sess-on
```

Claude 應該會：
1. 讀取 `agent-os/SESSION_INIT.md`
2. 讀取 `worklog/LastCheckPoint.log` 和 `TODO.md`
3. 報告專案狀態
4. 提醒你 Shopfloor 協作模式
5. 詢問要從哪裡開始

**如果成功**：✅ 配置完成！跳到「下一步」

**如果失敗**：參考「疑難排解」

---

## 🎓 第一個 Shopfloor 協作範例（3 分鐘）

現在試試 Shopfloor 協作模式：

```
你: 請幫我建立一個簡單的 Hello World API

Claude: 我將為你建立 Hello World API。

📁 已產生檔案到 shopfloor/Claude_TMP/：

code/hello_router.py - FastAPI 路由

請將檔案加入你的專案並回報結果。

你: [開啟 shopfloor/Claude_TMP/code/hello_router.py]
    [複製到專案 app/api/ 目錄]

你: 已加入專案

Claude: 很好！接下來...
```

**觀察**：
- ✅ Claude 沒有在對話中貼大段程式碼
- ✅ 程式碼在檔案中，有完整的語法高亮
- ✅ 檔案開頭有使用說明
- ✅ 對話保持簡潔

---

## ✅ 下一步

### 學習 Session 管理

試試其他 Session 指令：

```
/sess-check   # 查看當前進度
/sess-wrap    # 階段性存檔
/sess-off     # 下班存檔
```

### 閱讀詳細文件

- `docs/shopfloor-workflow.md` - Shopfloor 協作模式詳解
- `docs/session-guide.md` - Session 管理指南
- `agent-os/standards/global/workflow-standards.md` - 工作流程規範

### 自訂配置

參考 `CONFIGURATION_GUIDE.md` 調整：
- Shopfloor 目錄結構
- 檔案命名規範
- 溝通方式
- 程式碼風格

---

## 🔧 疑難排解

### 問題 1：/sess-on 沒有反應

**可能原因**：Slash Commands 未正確設定

**解決方案**：
1. 檢查 `.claude/commands/` 目錄是否存在
2. 檢查 `sess-on.md` 是否存在
3. 試試純文字指令：`Claude, sess on.`

### 問題 2：Claude 沒有讀取配置

**可能原因**：SESSION_INIT.md 路徑錯誤

**解決方案**：
1. 確認 `agent-os/SESSION_INIT.md` 存在
2. 檢查檔案內容是否正確
3. 直接請 Claude：`請讀取 agent-os/SESSION_INIT.md`

### 問題 3：Claude 仍在對話中貼程式碼

**可能原因**：Claude 還沒理解 Shopfloor 模式

**解決方案**：
1. 提醒 Claude：`請遵循 Shopfloor 協作模式，將程式碼輸出到檔案`
2. 確認 `workflow-standards.md` 已被讀取
3. 在 SESSION_INIT.md 中強調 Shopfloor 模式

### 問題 4：不知道如何配置技術環境

**解決方案**：
1. 查看 `examples/configs/` 找最接近的範例
2. 參考 `CONFIGURATION_GUIDE.md`
3. 直接問 Claude：`請幫我配置 [你的技術棧] 的協作環境`

---

## 💡 小技巧

### 技巧 1：善用 sess-check

工作中斷前，養成習慣：
```
/sess-wrap    # 存檔
```

回來後：
```
/sess-check   # 快速查看進度
```

### 技巧 2：定期清理 shopfloor

已加入專案的檔案可以刪除：
```bash
rm -rf shopfloor/Claude_TMP/code/*
```

或移到 archive：
```bash
mkdir -p shopfloor/archive/$(date +%Y-%m)
mv shopfloor/Claude_TMP/code/* shopfloor/archive/$(date +%Y-%m)/
```

### 技巧 3：自訂 Slash Commands

在 `.claude/commands/` 建立新檔案：

```bash
cat > .claude/commands/my-review.md << 'EOF'
請審閱最近的程式碼變更，並提供改進建議。
遵循專案的程式碼風格規範。
EOF
```

使用：`/my-review`

### 技巧 4：團隊協作

將協作規範加入版本控制：
```bash
git add agent-os/ .claude/ shopfloor/ worklog/ TODO.md
git commit -m "Add Claude collaboration standards"
```

團隊成員 clone 後即可使用相同規範。

---

## 📚 延伸閱讀

### 核心概念
- [Shopfloor 協作模式詳解](docs/shopfloor-workflow.md)
- [Session 管理指南](docs/session-guide.md)
- [溝通規範說明](agent-os/standards/global/communication-standards.md)

### 進階主題
- [配置指南](CONFIGURATION_GUIDE.md)
- [自訂協作規範](docs/customization.md)
- [多語言專案](docs/multilang-projects.md)

### 範例專案
- [Python FastAPI 完整範例](examples/projects/fastapi-full/)
- [JavaScript Node.js 範例](examples/projects/nodejs-api/)
- [全端專案範例](examples/projects/fullstack/)

---

## ❓ 還有問題？

- 查看 [常見問題](docs/faq.md)
- 閱讀 [配置指南](CONFIGURATION_GUIDE.md)
- 直接問 Claude：`關於協作規範，我想了解...`

---

**恭喜！你已經完成快速開始指南！** 🎉

現在你可以：
- 使用 `/sess-on` 開始每天的工作
- 享受 Shopfloor 協作模式的便利
- 讓 Claude 幫助你更有效率地開發

祝協作愉快！
