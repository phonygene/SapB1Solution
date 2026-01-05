# 配置指南

> [DEPRECATED] 本指南為歷史文件，現行協作規範以 `C:\Projects\SapB1Solution\.claude\MULTI_AGENT_ARCHITECTURE_SPEC.md` 為準。

**預計時間**：5-15 分鐘
**目標**：完成專案協作規範的初始化配置

---

## 🚀 快速開始（推薦）

### 步驟 1：複製檔案到專案

```bash
# 進入你的專案目錄
cd /path/to/your/project

# 複製協作規範目錄
cp -r /path/to/collaboration-package/agent-os ./

# 複製 Slash Commands
cp -r /path/to/collaboration-package/examples/.claude ./

# 建立工作目錄
mkdir -p shopfloor/Claude_TMP/{sql,code,etc}
mkdir -p worklog
```

### 步驟 2：執行初始化助手

在 Claude Code 中執行：

```
/init-project
```

Claude 會引導你完成配置！

---

## 📋 初始化流程說明

### 方式 A：從範例快速開始（推薦新手）

**優點**：
- ✅ 快速完成（5 分鐘）
- ✅ 基於實際案例
- ✅ 減少出錯機會

**流程**：
1. 選擇最接近你專案的範例
2. 回答 3-5 個關鍵問題
3. 自動產生配置檔
4. 確認並開始使用

**適合**：
- 專案技術棧與範例相近
- 希望快速開始
- 初次使用協作規範

### 方式 B：互動式逐步配置（適合自訂需求）

**優點**：
- ✅ 完全符合專案需求
- ✅ 一步一步引導
- ✅ 了解每個設定的用途

**流程**：
1. 回答必要問題（6 個）
   - 專案名稱
   - 程式語言
   - 框架
   - 資料庫
   - 語言偏好
   - 時區
2. 選擇是否設定可選項目
3. 自動產生配置檔
4. 確認並開始使用

**適合**：
- 專案技術棧較特殊
- 有特定協作需求
- 想要完全自訂

### 方式 C：手動編輯（適合進階使用者）

**優點**：
- ✅ 完全控制
- ✅ 可以參考多個範例
- ✅ 適合複雜專案

**流程**：
1. 開啟 `agent-os/SESSION_INIT.md`
2. 參考 `examples/configs/` 中的範例
3. 手動填寫所有 `[請填寫...]` 標記
4. 儲存並測試

**適合**：
- 熟悉協作規範
- 有複雜的專案結構
- 需要整合多種技術

---

## 🎯 配置檢查清單

### 必要配置（必須完成）

- [ ] 專案名稱
- [ ] 程式語言與框架
- [ ] 溝通語言（繁中/英文/其他）
- [ ] 時區設定

### 推薦配置（建議完成）

- [ ] 資料庫類型（如果有）
- [ ] 套件管理工具
- [ ] 主要依賴套件
- [ ] 專案結構說明

### 可選配置（依需求）

- [ ] 當前開發階段
- [ ] 程式碼風格設定
- [ ] 自訂 Shopfloor 目錄
- [ ] 時間標籤格式
- [ ] 特殊協作規則

---

## 📖 詳細說明

### 1. 專案名稱

**用途**：在 Session 初始化時顯示，幫助識別專案

**範例**：
- `MyAPI - RESTful API 專案`
- `WebShop - 電商網站`
- `DataPipeline - 資料處理系統`

**填寫位置**：`agent-os/SESSION_INIT.md` → 專案名稱區塊

---

### 2. 程式語言與框架

**用途**：
- Claude 會根據語言提供適合的程式碼範例
- 自動套用對應的程式碼風格
- 使用正確的術語

**常見組合**：
| 語言 | 框架 | 適合 |
|------|------|------|
| Python | FastAPI | RESTful API、微服務 |
| Python | Django | 全端 Web、CMS |
| JavaScript | Express | 後端 API |
| TypeScript | React | 前端 SPA |
| Java | Spring Boot | 企業級應用 |
| C# | .NET | Windows 應用、Web API |
| Go | Gin | 高效能 API |

**填寫位置**：`agent-os/SESSION_INIT.md` → 技術環境區塊

---

### 3. 資料庫

**用途**：
- 產生正確的 SQL 語法
- 提供資料庫遷移建議
- 使用適合的 ORM 範例

**選擇指南**：
- **PostgreSQL**：功能完整、開源、適合生產環境
- **MySQL**：流行、易用、適合 Web 應用
- **SQLite**：輕量、無需安裝、適合開發/小型應用
- **MongoDB**：NoSQL、適合文件型資料
- **SQL Server**：微軟生態系、適合企業環境

**填寫位置**：`agent-os/SESSION_INIT.md` → 技術環境區塊

---

### 4. 溝通語言

**用途**：
- Claude 的回應語言
- 程式碼註解語言（可另外設定）
- 文件語言

**選項**：
- 繁體中文
- 簡體中文
- English
- 日本語
- 其他

**術語對照**：
如果使用中文，建議設定術語對照表，例如：
- Database = 資料庫
- API Endpoint = API 端點

**填寫位置**：`agent-os/SESSION_INIT.md` → 溝通語言與術語區塊

---

### 5. 時區

**用途**：
- Session 工作日誌的時間記錄
- 時間標籤的時區
- 與團隊協作時的時間對齊

**常見時區**：
- UTC（協調世界時，推薦用於國際團隊）
- UTC+8（台灣、香港、新加坡、中國）
- UTC+9（日本、韓國）
- UTC-5（美國東部 EST）
- UTC-8（美國西部 PST）

**填寫位置**：`agent-os/SESSION_INIT.md` → 溝通語言與術語區塊

---

### 6. Shopfloor 目錄結構

**預設結構**：
```
shopfloor/Claude_TMP/
├── sql/            # SQL 腳本
├── code/           # 程式碼檔案
└── etc/            # 配置與文件
```

**自訂範例**：

依語言分類：
```
shopfloor/Claude_TMP/
├── python/
├── javascript/
└── sql/
```

依功能分類：
```
shopfloor/Claude_TMP/
├── api/
├── models/
├── tests/
└── migrations/
```

**填寫位置**：`agent-os/SESSION_INIT.md` → 協作模式區塊

---

## 🔍 驗證配置

### 配置完成後的檢查

1. **檢查必填項目**
```bash
# 開啟配置檔
cat agent-os/SESSION_INIT.md | grep "\[請填寫"
# 如果有輸出，表示還有未填寫的項目
```

2. **測試 Session 初始化**
```
/sess-on
```

應該看到：
- ✅ 專案名稱正確顯示
- ✅ 沒有 `[請填寫...]` 的錯誤訊息
- ✅ Shopfloor 協作模式提醒

3. **測試檔案產生**
```
請幫我建立一個簡單的 Hello World 範例
```

應該看到：
- ✅ 檔案產生到 `shopfloor/Claude_TMP/`
- ✅ 程式碼使用正確的語言和框架
- ✅ 註解語言符合設定

---

## ❓ 常見問題

### Q1: 初始化助手沒有反應？

**檢查**：
- [ ] `.claude/commands/init-project.md` 是否存在
- [ ] 是否在專案根目錄執行

**解決**：
```bash
# 確認檔案存在
ls .claude/commands/init-project.md

# 如果不存在，複製範例
cp /path/to/collaboration-package/examples/.claude/commands/init-project.md .claude/commands/
```

### Q2: 我的技術棧不在範例中怎麼辦？

**解決方案**：
1. 選擇「互動式逐步配置」
2. 手動輸入你的語言和框架
3. 或參考最接近的範例，手動修改

### Q3: 可以修改已完成的配置嗎？

**可以！**
1. 直接編輯 `agent-os/SESSION_INIT.md`
2. 或重新執行 `/init-project`（會保留已填寫的內容）

### Q4: 團隊成員需要各自配置嗎？

**不需要！**
- 將 `agent-os/` 加入 git
- 團隊成員 clone 後直接使用
- 個人偏好（如語言、時區）可以單獨調整

### Q5: 配置錯誤會影響使用嗎？

**影響有限**：
- 必填項目錯誤：Session 初始化會顯示錯誤，但仍可使用
- 可選項目錯誤：不影響基本功能
- 隨時可以修正

---

## 📚 參考資源

### 範例配置

詳細範例請查看：
- `examples/configs/python-fastapi/` - Python + FastAPI 完整範例
- `examples/configs/python-django/` - Python + Django（待補充）
- `examples/configs/javascript-nodejs/` - Node.js（待補充）

### 規範文件

深入了解請閱讀：
- `agent-os/standards/global/session-management.md` - Session 管理
- `agent-os/standards/global/workflow-standards.md` - Shopfloor 模式
- `agent-os/standards/global/localization.md` - 語言設定詳解

### 快速開始

- `README.md` - 完整說明
- `QUICK_START.md` - 快速開始指南

---

## 💡 最佳實踐

### 1. 從簡單開始

- 先完成必要配置
- 使用一段時間後再調整可選項目
- 不要一開始就過度自訂

### 2. 參考範例

- 選擇最接近的範例作為起點
- 逐步調整為符合專案需求
- 保留範例的良好結構

### 3. 定期更新

- 專案進入新階段時更新「當前階段」
- 技術棧改變時更新「技術環境」
- 新增重要協作規則時記錄下來

### 4. 團隊協作

- 與團隊討論協作規範
- 達成共識後才正式採用
- 定期回顧和改進

---

## 🎉 配置完成後

恭喜完成配置！現在你可以：

1. **開始第一個 Session**
```
/sess-on
```

2. **試試 Shopfloor 協作模式**
```
請幫我建立一個 [功能描述]
```

3. **建立待辦事項**
編輯 `TODO.md`，列出專案待辦事項

4. **定期存檔**
```
/sess-wrap    # 階段存檔
/sess-off     # 下班存檔
```

---

**需要協助？** 隨時詢問 Claude！

---

**最後更新**：2025-11-14
**版本**：2.0
