# Claude 協作初始化清單

**版本**：2.0
**更新日期**：2025-11-14

---

## 說明

本檔案是 Session 初始化的**執行清單**。
詳細的 Session 管理規範（指令定義、檔案格式）請參考 `agent-os/standards/global/session-management.md`。

> **⚠️ 重要提示**：
> 這是模板檔案，請根據你的專案需求填寫「專案基本資訊」區塊。
> 你可以參考 `examples/configs/` 中的範例配置。

---

## Session On 初始化步驟

當觸發 `/sess-on` 指令時，依序執行：

### 第一步：讀取核心規範
按優先順序讀取以下檔案：

1. `agent-os/standards/global/session-management.md` - Session 管理機制
2. `agent-os/standards/global/workflow-standards.md` - 工作流程與 Shopfloor 協作模式
3. `agent-os/standards/global/communication-standards.md` - 溝通與責任歸屬規範
4. `agent-os/standards/global/localization.md` - 語言與術語標準
5. `agent-os/standards/global/coding-style.md` - 程式碼風格規範

### 第二步：讀取專案狀態
1. `worklog/LastCheckPoint.log` - 最新工作狀態與待辦事項
2. `TODO.md` - 保留功能與未完成項目

### 第三步：報告與確認
向使用者報告：
- 上次工作的時間點
- 當前專案狀態摘要（完成度百分比）
- 未完成的待辦事項（按優先順序）
- **立即需要確認的事項**（如：等待使用者執行的任務）
- 建議的下一步工作
- **⚠️ 協作模式提醒**：提醒使用者本專案採用 Shopfloor 協作模式

### 第四步：等待使用者回應
- 使用者可能回報上次待辦事項的執行結果
- 使用者可能提出新的任務
- 根據回應調整工作計畫

---

## 專案基本資訊（請填寫此區塊）

> **📝 填寫指引**：
> - 將所有 `[請填寫...]` 替換為你的專案資訊
> - 參考 `examples/configs/` 中的範例配置
> - 不需要的區塊可以刪除或保持空白

### 專案名稱

**[請填寫你的專案名稱]**

範例：
- `MyAPI - RESTful API 專案`
- `WebApp - 電商網站`
- `DataPipeline - 資料處理系統`

---

### 當前階段

**[請填寫當前開發階段與目標]**

範例：
```
Sprint 1（2025-11-14 至 2025-11-28）
- 目標：完成使用者認證與授權功能
- 架構：[你的架構]
- 功能：[主要功能列表]
```

---

### 技術環境

> **選擇方式**：
> 1. 從 `examples/configs/` 複製最接近的範例
> 2. 或參考下方格式自行填寫

**程式語言與框架**：
- **語言**：[請填寫，例如：Python 3.10+ / JavaScript ES6+ / TypeScript / Java / C# / Go / Rust]
- **框架**：[請填寫，例如：FastAPI / Django / Express / React / Vue / Spring Boot / .NET]
- **IDE**：[請填寫，例如：VS Code / PyCharm / IntelliJ / Visual Studio]

**資料庫**：
- **類型**：[請填寫，例如：PostgreSQL / MySQL / MongoDB / SQLite / SQL Server]
- **版本**：[請填寫版本號]
- **連線方式**：[使用環境變數配置，不要在此寫入帳號密碼]

**套件管理**：
- [請填寫，例如：pip / npm / yarn / maven / gradle / nuget]

**主要套件/依賴**：
```
[請列出主要套件，例如：]
fastapi==0.104.1
sqlalchemy==2.0.23
[或]
express: ^4.18.0
mongoose: ^7.0.0
```

**開發工具**：
```
[請列出 linter、formatter、測試工具等，例如：]
black, flake8, pytest
[或]
eslint, prettier, jest
```

---

### 專案結構

> **填寫提示**：
> - 描述你的專案目錄結構
> - 包含 shopfloor/Claude_TMP/ 和 worklog/ 目錄
> - 可參考 `examples/configs/python-fastapi/SESSION_INIT.md`

```
[請填寫你的專案結構，例如：]

your-project/
├── src/                          # 原始碼目錄
│   ├── [你的模組結構]
│   └── ...
├── tests/                        # 測試檔案
├── shopfloor/Claude_TMP/         # Claude 產出檔案
│   ├── sql/                      # SQL 腳本（如果有資料庫）
│   ├── code/                     # 程式碼檔案
│   └── etc/                      # 文件與配置
├── worklog/                      # 工作日誌
│   └── LastCheckPoint.log
├── agent-os/                     # 協作規範
│   ├── standards/global/
│   └── SESSION_INIT.md           # 本檔案
├── .env                          # 環境變數（不要提交到 git）
├── .env.example                  # 環境變數範例
├── TODO.md                       # 待辦事項
└── README.md                     # 專案說明
```

---

### 溝通語言與術語

> **語言選擇**：
> - 繁體中文 / 簡體中文 / English / 日本語 / 其他
> - 參考 `agent-os/standards/global/localization.md`

**語言**：[請選擇，例如：繁體中文 / English]

**術語對照**（選用）：
```
[如果使用中文，可以定義技術術語對照，例如：]
- Database = 資料庫
- Table = 資料表
- Row = 列
- Column = 欄
- API Endpoint = API 端點
- Middleware = 中介軟體

[如果使用英文，可以省略此區塊]
```

**時區**：[請填寫，例如：UTC / UTC+8 / UTC-5 / 其他]

**回應時間標籤**（選用）：
- 格式：`[YYYY-MM-DD HH:mm TIMEZONE]`
- 範例：`[2025-11-14 15:30 UTC]`

---

### 協作模式

**檔案輸出**：
- 產生檔案到 `shopfloor/Claude_TMP/`
- 不在對話中貼大段程式碼
- 節省 Token，提升效率

**Shopfloor 目錄結構**（可自訂）：
```
[預設結構：]
shopfloor/Claude_TMP/
├── sql/            # SQL 腳本、資料庫遷移
├── code/           # 程式碼檔案
└── etc/            # 配置檔、文件

[你也可以自訂，例如：]
shopfloor/Claude_TMP/
├── python/
├── javascript/
├── sql/
└── docs/
```

**溝通方式**：
- Claude 簡要說明產生的檔案
- 使用者在 IDE 中檢視並執行
- 簡短回報結果（"執行成功" / "報錯：..."）
- Claude 繼續下一步

---

### 程式碼風格（選用）

> **風格指南**：
> - 參考 `agent-os/standards/global/coding-style.md`
> - 可以指定使用的 formatter 和 linter

**風格指南**：[請填寫，例如：PEP 8 / Airbnb Style Guide / Google Style Guide]

**自動化工具**：
- Formatter: [請填寫，例如：Black / Prettier / gofmt]
- Linter: [請填寫，例如：Flake8 / ESLint / Pylint]
- Type Checker: [請填寫，例如：mypy / TypeScript]

**註解語言**：[請填寫，例如：繁體中文 / English / 雙語]

---

## 使用方式

### Slash Commands（推薦）

使用 Claude Code 的 slash command 功能：
- `/sess-on` - 上班/開始工作
- `/sess-check` - 查看進度（不寫檔案）
- `/sess-wrap` - 階段存檔，繼續工作
- `/sess-off` - 完整存檔並下班

### 純文字指令（備用）

如果 slash command 無法使用，可以輸入：
- `Claude, sess on.`
- `Claude, sess check.`
- `Claude, sess wrap.`
- `Claude, sess off.`

---

## 快速開始範例

### 選項 1：從範例複製配置

```bash
# 選擇最接近你專案的範例
cp examples/configs/python-fastapi/SESSION_INIT.md agent-os/

# 或
cp examples/configs/python-django/SESSION_INIT.md agent-os/

# 或
cp examples/configs/javascript-nodejs/SESSION_INIT.md agent-os/

# 然後編輯並調整為你的專案
```

### 選項 2：手動填寫

1. 開啟本檔案
2. 找到所有 `[請填寫...]` 標記
3. 替換為你的專案資訊
4. 刪除不需要的區塊
5. 儲存檔案

### 選項 3：使用初始化助手（如果可用）

```
/init-project
```

---

## 維護建議

### 定期更新

- **專案進入新階段**：更新「當前階段」區塊
- **技術棧改變**：更新「技術環境」區塊
- **協作規範調整**：更新相應區塊

### 版本控制

建議將此檔案加入 git：
```bash
git add agent-os/SESSION_INIT.md
git commit -m "Update session initialization config"
```

團隊成員可共用相同配置。

### 保持簡潔

- 本檔案只記錄基本資訊和執行流程
- 詳細規範保持在 `agent-os/standards/global/` 中
- 避免在此檔案中重複規範內容

---

## 參考資源

### 範例配置

- `examples/configs/python-fastapi/` - Python + FastAPI 專案
- `examples/configs/python-django/` - Python + Django 專案
- `examples/configs/javascript-nodejs/` - JavaScript + Node.js 專案
- `examples/configs/typescript-react/` - TypeScript + React 專案

### 規範文件

- `agent-os/standards/global/session-management.md` - Session 管理詳細規範
- `agent-os/standards/global/workflow-standards.md` - Shopfloor 協作模式
- `agent-os/standards/global/communication-standards.md` - 溝通規範
- `agent-os/standards/global/localization.md` - 語言與時區設定
- `agent-os/standards/global/coding-style.md` - 程式碼風格

### 快速開始指南

- `QUICK_START.md` - 10-15 分鐘快速開始
- `CONFIGURATION_GUIDE.md` - 詳細配置指南
- `README.md` - 完整說明文件

---

## 常見問題

### Q: 一定要填寫所有區塊嗎？

A: 不一定。必填項目：
- ✅ 專案名稱
- ✅ 技術環境（至少語言和框架）
- ✅ 溝通語言

選填項目：
- 當前階段（如果專案剛開始可以省略）
- 程式碼風格（可以使用預設）
- 其他自訂區塊

### Q: 可以修改 Shopfloor 目錄名稱嗎？

A: 可以！只要在「協作模式」區塊中說明即可。
例如改為 `claude-output/` 或 `ai-generated/`。

### Q: 支援多語言專案嗎？

A: 支援！可以在「溝通語言與術語」區塊中說明：
- 程式碼註解使用英文
- 文件提供雙語版本
- 或其他自訂規則

### Q: 如何讓團隊成員使用相同配置？

A: 將 `agent-os/` 目錄加入版本控制（git），
團隊成員 clone 後即可使用相同規範。

---

**最後更新**：2025-11-14
**版本**：2.0（通用模板）
**維護者**：開源社群

---

## 💡 提示

首次使用？建議：
1. 參考 `examples/configs/` 選擇最接近的範例
2. 複製到 `agent-os/SESSION_INIT.md`
3. 根據你的專案調整
4. 執行 `/sess-on` 測試

不確定如何填寫？
- 閱讀 `CONFIGURATION_GUIDE.md`（詳細配置指南）
- 查看 `QUICK_START.md`（快速開始）
- 或直接詢問 Claude！
