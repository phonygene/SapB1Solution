# 語言與術語標準 (Localization)

**版本**：2.0
**更新日期**：2025-11-14

---

## 概述

本規範定義如何在專案中設定語言偏好、術語對照和時區設定。

---

## 語言設定

### 在 SESSION_INIT.md 中配置

在 `agent-os/SESSION_INIT.md` 的「溝通語言與術語」區塊設定：

```markdown
### 溝通語言與術語

- **語言**：[你的偏好語言]
- **術語對照**：[技術術語的翻譯對照]
- **時區**：[你的時區]
```

### 支援的語言範例

#### 繁體中文
```markdown
- **語言**：繁體中文 (Traditional Chinese)
- **術語對照**：
  - Row = 列
  - Column = 欄
  - Database = 資料庫
  - Table = 資料表
  - Query = 查詢
- **時區**：UTC+8 (台灣、香港、新加坡)
```

#### 簡體中文
```markdown
- **語言**：簡體中文 (Simplified Chinese)
- **術語對照**：
  - Row = 行
  - Column = 列
  - Database = 數據庫
  - Table = 數據表
  - Query = 查詢
- **時區**：UTC+8 (中國)
```

#### English
```markdown
- **語言**：English
- **術語對照**：
  - Use standard technical English terminology
- **時區**：UTC / UTC-5 (EST) / UTC-8 (PST) / [your timezone]
```

#### 日本語
```markdown
- **語言**：日本語 (Japanese)
- **術語對照**：
  - Row = 行
  - Column = 列
  - Database = データベース
  - Table = テーブル
  - Query = クエリ
- **時區**：UTC+9 (日本)
```

---

## 技術術語對照

### 資料庫相關

| English | 繁體中文 | 簡體中文 | 日本語 |
|---------|---------|---------|--------|
| Database | 資料庫 | 數據庫 | データベース |
| Table | 資料表 | 數據表 | テーブル |
| Row | 列 | 行 | 行 |
| Column | 欄 | 列 | 列 |
| Query | 查詢 | 查詢 | クエリ |
| Index | 索引 | 索引 | インデックス |
| Primary Key | 主鍵 | 主鍵 | 主キー |
| Foreign Key | 外鍵 | 外鍵 | 外部キー |

### Web 開發相關

| English | 繁體中文 | 簡體中文 | 日本語 |
|---------|---------|---------|--------|
| Endpoint | API 端點 | API 端點 | エンドポイント |
| Route | 路由 | 路由 | ルート |
| Middleware | 中介軟體 | 中間件 | ミドルウェア |
| Request | 請求 | 請求 | リクエスト |
| Response | 回應 | 響應 | レスポンス |
| Authentication | 認證 | 認證 | 認証 |
| Authorization | 授權 | 授權 | 認可 |

### 程式設計相關

| English | 繁體中文 | 簡體中文 | 日本語 |
|---------|---------|---------|--------|
| Function | 函式 | 函數 | 関数 |
| Class | 類別 | 類 | クラス |
| Method | 方法 | 方法 | メソッド |
| Variable | 變數 | 變量 | 変数 |
| Array | 陣列 | 數組 | 配列 |
| Object | 物件 | 對象 | オブジェクト |
| Interface | 介面 | 接口 | インターフェース |

---

## 時區設定

### 常見時區

| 時區代碼 | 說明 | UTC 偏移 |
|---------|------|----------|
| UTC | 協調世界時 | +0 |
| GMT | 格林威治標準時間 | +0 |
| EST | 美國東部標準時間 | -5 |
| PST | 美國太平洋標準時間 | -8 |
| CST | 中國標準時間 | +8 |
| JST | 日本標準時間 | +9 |
| IST | 印度標準時間 | +5:30 |
| AEST | 澳洲東部標準時間 | +10 |

### 時間格式

**ISO 8601 格式（推薦）**：
```
2025-11-14T15:30:00Z          # UTC
2025-11-14T15:30:00+08:00     # UTC+8
2025-11-14T15:30:00-05:00     # UTC-5 (EST)
```

**人類可讀格式**：
```
2025-11-14 15:30 UTC
2025-11-14 15:30 UTC+8
2025-11-14 15:30 EST
```

### 在程式碼中處理時區

**Python 範例**：
```python
from datetime import datetime, timezone
import pytz

# UTC 時間
utc_now = datetime.now(timezone.utc)

# 轉換到特定時區
tw_tz = pytz.timezone('Asia/Taipei')
tw_now = utc_now.astimezone(tw_tz)

# 格式化輸出
formatted = tw_now.strftime('%Y-%m-%d %H:%M %Z')
# 輸出：2025-11-14 15:30 CST
```

**JavaScript 範例**：
```javascript
// UTC 時間
const utcNow = new Date();

// 轉換到特定時區
const options = {
  timeZone: 'Asia/Taipei',
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit',
  timeZoneName: 'short'
};

const formatted = utcNow.toLocaleString('en-US', options);
// 輸出：11/14/2025, 03:30 PM GMT+8
```

---

## 回應時間標籤

如果專案要求在回應結尾加上時間標籤：

### 配置

在 `SESSION_INIT.md` 中設定：

```markdown
### 協作模式
- **回應結尾**：加上時間標籤 `[YYYY-MM-DD HH:mm TIMEZONE]`
```

### 範例

```
我已經完成使用者認證功能，請檢視檔案。

[2025-11-14 15:30 UTC+8]
```

### 何時使用

**✅ 應該使用**：
- 產生檔案後的報告
- Session 存檔（sess-wrap, sess-off）
- 重要的階段性報告
- 工作日誌更新

**❌ 不需要使用**：
- 簡短的確認訊息
- 錯誤提示
- 一般對話
- 問題詢問

---

## 多語言專案

如果專案包含多種語言的程式碼（如前後端分離）：

### 目錄結構建議

```
shopfloor/Claude_TMP/
├── frontend/          # 前端相關（可能用英文註解）
├── backend/           # 後端相關（可能用中文註解）
└── docs/              # 文件（可能雙語）
```

### 註解語言建議

**方案 A：統一使用英文**（推薦）
```python
# Validate user authentication token
def validate_token(token: str) -> bool:
    # Check token expiration
    if is_expired(token):
        return False
    return True
```

**方案 B：使用母語**
```python
# 驗證使用者認證 token
def validate_token(token: str) -> bool:
    # 檢查 token 是否過期
    if is_expired(token):
        return False
    return True
```

**方案 C：雙語註解**（適合團隊有多國成員）
```python
# Validate user authentication token / 驗證使用者認證 token
def validate_token(token: str) -> bool:
    # Check token expiration / 檢查 token 是否過期
    if is_expired(token):
        return False
    return True
```

**建議**：
- 開源專案：使用英文
- 私有專案：團隊內部決定
- 文件：可以提供多語言版本

---

## 文件語言

### README 和文件

**方案 A：只有一種語言**
```
README.md                 # 英文
docs/                     # 英文文件
```

**方案 B：多語言版本**
```
README.md                 # 英文（主要版本）
README.zh-TW.md           # 繁體中文
README.zh-CN.md           # 簡體中文
README.ja.md              # 日文
docs/
  ├── en/                 # 英文文件
  ├── zh-TW/              # 繁體中文文件
  └── zh-CN/              # 簡體中文文件
```

### Git Commit 訊息

**建議使用英文**（業界慣例）：
```bash
git commit -m "feat: add user authentication"
git commit -m "fix: resolve token expiration issue"
git commit -m "docs: update API documentation"
```

**如果團隊決定使用母語**：
```bash
git commit -m "功能: 新增使用者認證"
git commit -m "修正: 解決 token 過期問題"
git commit -m "文件: 更新 API 文件"
```

---

## 自訂術語對照表

如果專案有特定領域的術語，建議建立自訂對照表：

### 範例：金融領域

```markdown
### 專案術語對照

| English | 繁體中文 | 說明 |
|---------|---------|------|
| Account | 帳戶 | 使用者帳戶 |
| Transaction | 交易 | 金融交易記錄 |
| Balance | 餘額 | 帳戶餘額 |
| Deposit | 存款 | 存入金額 |
| Withdrawal | 提款 | 提出金額 |
| Transfer | 轉帳 | 帳戶間轉移 |
```

### 使用方式

1. 將對照表加入 `SESSION_INIT.md`
2. 或建立獨立檔案 `docs/terminology.md`
3. 在 Session 初始化時讓 Claude 讀取

---

## 配置範例

### 繁體中文專案

```markdown
### 溝通語言與術語

- **語言**：繁體中文
- **術語對照**：
  - Database = 資料庫
  - Table = 資料表
  - Row = 列
  - Column = 欄
  - API Endpoint = API 端點
  - Middleware = 中介軟體
- **時區**：UTC+8
- **回應時間標籤**：`[YYYY-MM-DD HH:mm UTC+8]`
- **註解語言**：繁體中文
- **Git Commit**：英文
```

### 英文專案

```markdown
### Communication Language and Terminology

- **Language**: English
- **Terminology**: Standard technical English
- **Timezone**: UTC
- **Response Timestamp**: `[YYYY-MM-DD HH:mm UTC]`
- **Code Comments**: English
- **Git Commit**: English
```

### 多語言團隊

```markdown
### 溝通語言與術語 / Communication Settings

- **主要語言 / Primary Language**: English
- **次要語言 / Secondary Language**: 繁體中文
- **術語對照 / Terminology**:
  - 程式碼註解使用英文
  - 文件提供雙語版本
  - API 文件使用英文
- **時區 / Timezone**: UTC
```

---

## 參考資料

- [ISO 8601 時間格式標準](https://en.wikipedia.org/wiki/ISO_8601)
- [IANA 時區資料庫](https://www.iana.org/time-zones)
- [Unicode CLDR](http://cldr.unicode.org/)

---

**最後更新**：2025-11-14
**版本**：2.0（通用版本）
**維護者**：開源社群
