# Claude 協作初始化清單

**版本**：2.0
**更新日期**：2025-11-14
**範例類型**：Python + FastAPI

---

## 說明

本檔案是 Session 初始化的**執行清單**（Python + FastAPI 範例）。
詳細的 Session 管理規範請參考 `agent-os/standards/global/session-management.md`。

> **提示**：這是範例配置，請根據你的專案需求調整。

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

## 專案基本資訊（請修改此區塊）

### 專案名稱
**[請填寫你的專案名稱]**

範例：`MyAPI - RESTful API 專案`

### 當前階段
**[請填寫當前開發階段]**

範例：
```
Sprint 1（2025-11-14 至 2025-11-28）
- 目標：完成使用者認證與授權功能
- 架構：FastAPI + PostgreSQL + Redis
- 功能：JWT 認證、角色權限管理、API 速率限制
```

### 技術環境

- **語言**：Python 3.10+
- **框架**：FastAPI 0.104+
- **IDE**：VS Code / PyCharm
- **資料庫**：PostgreSQL 14+ / MySQL 8+ / SQLite（開發環境）
- **套件管理**：pip / poetry / uv
- **虛擬環境**：venv / virtualenv / conda

**主要套件**：
```
fastapi==0.104.1
uvicorn[standard]==0.24.0
sqlalchemy==2.0.23
pydantic==2.5.0
python-jose[cryptography]==3.3.0
passlib[bcrypt]==1.7.4
python-multipart==0.0.6
alembic==1.12.1
```

**開發工具**：
```
black==23.11.0
flake8==6.1.0
mypy==1.7.0
pytest==7.4.3
pytest-cov==4.1.0
```

### 專案結構

```
your-project/
├── app/                          # 主要應用程式
│   ├── api/                      # API 路由
│   │   ├── v1/                   # API 版本 1
│   │   │   ├── endpoints/        # 端點
│   │   │   │   ├── auth.py       # 認證相關
│   │   │   │   ├── users.py      # 使用者相關
│   │   │   │   └── ...
│   │   │   └── deps.py           # 依賴注入
│   │   └── router.py             # 主路由
│   ├── core/                     # 核心配置
│   │   ├── config.py             # 配置管理
│   │   ├── security.py           # 安全相關
│   │   └── database.py           # 資料庫連線
│   ├── models/                   # SQLAlchemy 資料模型
│   │   ├── __init__.py
│   │   ├── user.py
│   │   └── ...
│   ├── schemas/                  # Pydantic schemas
│   │   ├── __init__.py
│   │   ├── user.py
│   │   └── ...
│   ├── crud/                     # CRUD 操作
│   │   ├── __init__.py
│   │   ├── user.py
│   │   └── ...
│   ├── middleware/               # 中介軟體
│   │   ├── __init__.py
│   │   └── ...
│   ├── utils/                    # 工具函式
│   │   └── ...
│   └── main.py                   # 應用程式入口
├── tests/                        # 測試檔案
│   ├── api/
│   ├── crud/
│   └── conftest.py
├── alembic/                      # 資料庫遷移
│   ├── versions/
│   └── env.py
├── shopfloor/Claude_TMP/         # Claude 產出檔案
│   ├── sql/                      # SQL 腳本
│   ├── code/                     # Python 檔案
│   └── etc/                      # 文件與配置
├── worklog/                      # 工作日誌
│   └── LastCheckPoint.log
├── agent-os/                     # 協作規範
│   ├── standards/global/
│   └── SESSION_INIT.md           # 本檔案
├── .env                          # 環境變數（不要提交）
├── .env.example                  # 環境變數範例
├── requirements.txt              # Python 套件清單
├── pyproject.toml                # 專案配置（如果使用 poetry）
├── alembic.ini                   # Alembic 配置
├── TODO.md                       # 待辦事項
└── README.md                     # 專案說明
```

### 溝通語言與術語

> **請根據你的偏好調整**

- **語言**：繁體中文（或 English / 其他語言）
- **術語對照**：
  - Database = 資料庫
  - Table = 資料表
  - Row = 列
  - Column = 欄
  - API Endpoint = API 端點
  - Middleware = 中介軟體
  - Schema = 資料綱要
  - Migration = 資料庫遷移
- **時區**：UTC（或 UTC+8 / 其他時區）
- **回應時間標籤**：`[YYYY-MM-DD HH:mm UTC]`

### 協作模式

- **檔案輸出**：產生檔案到 `shopfloor/Claude_TMP/`，不在對話中貼大段程式碼
- **溝通方式**：簡要說明 + 等待使用者回報結果
- **回應結尾**：加上時間標籤（選用）

**Shopfloor 目錄結構**：
```
shopfloor/Claude_TMP/
├── sql/            # SQL 腳本、Alembic migrations
├── code/           # Python 檔案（models, schemas, routers, etc.）
└── etc/            # 配置檔、文件、測試資料
```

### 程式碼風格

- **風格指南**：PEP 8
- **Formatter**：Black（line-length: 88）
- **Linter**：Flake8
- **Type Checker**：mypy
- **Import 排序**：isort
- **註解語言**：[繁體中文 / English]
- **Docstring 格式**：Google Style

**配置檔案**：

`.flake8`:
```ini
[flake8]
max-line-length = 88
extend-ignore = E203, W503
exclude = .git,__pycache__,venv,.venv,alembic
```

`pyproject.toml` (Black + isort):
```toml
[tool.black]
line-length = 88
target-version = ['py310']

[tool.isort]
profile = "black"
```

### API 設計規範

**RESTful API 慣例**：
- GET /users - 列表
- GET /users/{id} - 詳情
- POST /users - 建立
- PUT /users/{id} - 完整更新
- PATCH /users/{id} - 部分更新
- DELETE /users/{id} - 刪除

**回應格式**：
```json
{
  "success": true,
  "data": { ... },
  "message": "操作成功"
}
```

**錯誤格式**：
```json
{
  "success": false,
  "error": {
    "code": "USER_NOT_FOUND",
    "message": "使用者不存在",
    "details": { ... }
  }
}
```

---

## 使用方式

### Slash Commands（推薦）

- `/sess-on` - 上班/開始工作
- `/sess-check` - 查看進度（不寫檔案）
- `/sess-wrap` - 階段存檔，繼續工作
- `/sess-off` - 完整存檔並下班

### 純文字指令（備用）

- `Claude, sess on.`
- `Claude, sess check.`
- `Claude, sess wrap.`
- `Claude, sess off.`

---

## 維護建議

- 當協作規範有重大變更時，更新本檔案
- 當專案進入新階段時，更新「當前階段」資訊
- 當技術棧改變時，更新「技術環境」區塊
- 保持本檔案簡潔，詳細規範仍在 `agent-os/standards/global/` 中

---

## 範例：初始化新專案

如果這是全新的 FastAPI 專案，可以請 Claude 幫你初始化：

```
你：請幫我初始化一個 FastAPI 專案，包含使用者認證功能

Claude 會：
1. 產生專案結構到 shopfloor/Claude_TMP/
2. 產生必要的配置檔（.env.example, requirements.txt 等）
3. 產生基本的 models, schemas, routers
4. 產生 Alembic migrations
5. 產生 README 和使用說明

然後你：
1. 檢視產生的檔案
2. 複製到專案對應位置
3. 執行初始化指令（創建虛擬環境、安裝套件、執行遷移）
4. 回報結果
```

---

**最後更新**：2025-11-14
**版本**：2.0（Python + FastAPI 範例）
