# 程式碼風格規範

**版本**：2.0
**更新日期**：2025-11-14

---

## 概述

本規範提供通用的程式碼風格建議。具體專案應根據使用的程式語言和框架調整。

---

## 通用原則

### 1. 一致性優先

**最重要的原則**：保持專案內部風格一致。

- 如果專案已有既定風格，遵循現有風格
- 如果是新專案，選擇一種風格並堅持
- 使用自動化工具（linter, formatter）確保一致性

### 2. 可讀性優先

**程式碼是寫給人看的**：

- ✅ 使用有意義的變數名稱
- ✅ 適當的註解說明複雜邏輯
- ✅ 保持函式簡短（單一職責）
- ✅ 適當的空白行分隔邏輯區塊

### 3. 遵循語言慣例

每種語言都有自己的最佳實踐：

- Python: PEP 8
- JavaScript: Airbnb Style Guide / StandardJS
- TypeScript: TypeScript Style Guide
- Java: Google Java Style Guide
- C#: Microsoft C# Coding Conventions
- Go: Effective Go
- Rust: Rust Style Guide

---

## Python 風格規範

### 基本規則

**PEP 8 標準**：

```python
# 縮排：4 個空格
def calculate_total(items: list) -> float:
    total = 0.0
    for item in items:
        total += item.price
    return total

# 命名慣例
class UserAccount:  # 類別：PascalCase
    def get_balance(self):  # 方法：snake_case
        return self._balance  # 私有變數：_開頭

MAX_RETRY_COUNT = 3  # 常數：UPPER_SNAKE_CASE
user_name = "John"   # 變數：snake_case
```

### Import 順序

```python
# 1. 標準庫
import os
import sys
from datetime import datetime

# 2. 第三方套件
import numpy as np
from fastapi import FastAPI
from sqlalchemy import Column, Integer

# 3. 本地模組
from app.models import User
from app.core.config import settings
```

### 型別提示（推薦）

```python
from typing import List, Dict, Optional

def get_users(
    limit: int = 10,
    offset: int = 0
) -> List[Dict[str, any]]:
    """取得使用者列表

    Args:
        limit: 限制數量
        offset: 偏移量

    Returns:
        使用者資料列表
    """
    pass
```

---

## JavaScript / TypeScript 風格規範

### 基本規則

```javascript
// 縮排：2 個空格
function calculateTotal(items) {
  let total = 0;
  for (const item of items) {
    total += item.price;
  }
  return total;
}

// 命名慣例
class UserAccount {  // 類別：PascalCase
  getBalance() {     // 方法：camelCase
    return this._balance;  // 私有變數：_開頭
  }
}

const MAX_RETRY_COUNT = 3;  // 常數：UPPER_SNAKE_CASE
const userName = 'John';    // 變數：camelCase
```

### TypeScript 型別

```typescript
interface User {
  id: number;
  name: string;
  email: string;
  isActive?: boolean;  // 可選屬性
}

function getUsers(
  limit: number = 10,
  offset: number = 0
): Promise<User[]> {
  // ...
}
```

### 現代 JavaScript

```javascript
// 使用 const/let，不用 var
const apiUrl = 'https://api.example.com';
let counter = 0;

// 箭頭函式
const add = (a, b) => a + b;

// 解構賦值
const { name, email } = user;
const [first, second] = array;

// 模板字串
const message = `Hello, ${name}!`;

// 展開運算子
const newArray = [...oldArray, newItem];
const newObject = { ...oldObject, newProp: value };
```

---

## 註解規範

### 何時需要註解

**✅ 應該註解**：
- 複雜的業務邏輯
- 非顯而易見的演算法
- 臨時的 workaround（並標註 TODO）
- 公開的 API / 函式
- 重要的決策原因

**❌ 不需要註解**：
- 顯而易見的程式碼
- 重複程式碼內容的註解

```python
# ❌ 不好的註解
# 設定 x 為 1
x = 1

# 迴圈遍歷使用者
for user in users:
    print(user)

# ✅ 好的註解
# 使用二分搜尋以提升效能（資料量 > 10000 時）
index = binary_search(sorted_list, target)

# FIXME: 暫時解法，應該改用 Redis cache
# TODO: Issue #123
cache = {}
```

### 文件字串（Docstring）

**Python**：
```python
def calculate_discount(
    price: float,
    discount_rate: float,
    is_member: bool = False
) -> float:
    """計算折扣後價格

    根據折扣率和會員狀態計算最終價格。
    會員可享有額外 5% 折扣。

    Args:
        price: 原始價格
        discount_rate: 折扣率（0.0 ~ 1.0）
        is_member: 是否為會員

    Returns:
        折扣後價格

    Raises:
        ValueError: 當 discount_rate 超出範圍時

    Examples:
        >>> calculate_discount(100, 0.1)
        90.0
        >>> calculate_discount(100, 0.1, is_member=True)
        85.5
    """
    if not 0 <= discount_rate <= 1:
        raise ValueError("discount_rate must be between 0 and 1")

    final_price = price * (1 - discount_rate)
    if is_member:
        final_price *= 0.95

    return final_price
```

**JavaScript / JSDoc**：
```javascript
/**
 * 計算折扣後價格
 *
 * @param {number} price - 原始價格
 * @param {number} discountRate - 折扣率 (0.0 ~ 1.0)
 * @param {boolean} [isMember=false] - 是否為會員
 * @returns {number} 折扣後價格
 * @throws {Error} 當 discountRate 超出範圍時
 *
 * @example
 * calculateDiscount(100, 0.1);  // 90
 * calculateDiscount(100, 0.1, true);  // 85.5
 */
function calculateDiscount(price, discountRate, isMember = false) {
  if (discountRate < 0 || discountRate > 1) {
    throw new Error('discountRate must be between 0 and 1');
  }

  let finalPrice = price * (1 - discountRate);
  if (isMember) {
    finalPrice *= 0.95;
  }

  return finalPrice;
}
```

---

## 檔案組織

### Python 模組結構

```python
"""
模組說明：使用者認證服務

提供使用者登入、登出、token 驗證等功能。
"""

# Imports
from typing import Optional
from datetime import datetime

# Constants
TOKEN_EXPIRY_HOURS = 24
MAX_LOGIN_ATTEMPTS = 5

# Classes
class AuthService:
    """認證服務"""

    def __init__(self):
        pass

    def login(self, username: str, password: str) -> Optional[str]:
        """使用者登入"""
        pass

# Functions
def validate_password(password: str) -> bool:
    """驗證密碼強度"""
    pass

# Main
if __name__ == "__main__":
    # 測試程式碼
    pass
```

### JavaScript 模組結構

```javascript
/**
 * 使用者認證服務
 *
 * 提供使用者登入、登出、token 驗證等功能
 */

// Imports
import bcrypt from 'bcrypt';
import jwt from 'jsonwebtoken';

// Constants
const TOKEN_EXPIRY_HOURS = 24;
const MAX_LOGIN_ATTEMPTS = 5;

// Classes
class AuthService {
  /**
   * 認證服務
   */
  constructor() {
    // ...
  }

  /**
   * 使用者登入
   */
  async login(username, password) {
    // ...
  }
}

// Functions
function validatePassword(password) {
  /**
   * 驗證密碼強度
   */
  // ...
}

// Exports
export { AuthService, validatePassword };
```

---

## 錯誤處理

### Python

```python
# ✅ 具體的例外類型
try:
    user = get_user(user_id)
except UserNotFoundError as e:
    logger.error(f"User {user_id} not found: {e}")
    raise
except DatabaseError as e:
    logger.error(f"Database error: {e}")
    # 轉換為業務例外
    raise ServiceUnavailableError("Service temporarily unavailable")

# ❌ 避免過於廣泛的例外捕捉
try:
    do_something()
except Exception:  # 太廣泛
    pass  # 吃掉錯誤，難以除錯
```

### JavaScript

```javascript
// ✅ 適當的錯誤處理
async function getUser(userId) {
  try {
    const user = await database.findUser(userId);
    if (!user) {
      throw new UserNotFoundError(`User ${userId} not found`);
    }
    return user;
  } catch (error) {
    if (error instanceof UserNotFoundError) {
      logger.error(error.message);
      throw error;
    }
    // 轉換為業務例外
    logger.error('Database error:', error);
    throw new ServiceUnavailableError('Service temporarily unavailable');
  }
}

// ❌ 避免空的 catch
try {
  doSomething();
} catch (error) {
  // 什麼都不做，難以除錯
}
```

---

## 安全性考量

### 防止 SQL Injection

```python
# ❌ 危險：SQL Injection
query = f"SELECT * FROM users WHERE username = '{username}'"

# ✅ 安全：參數化查詢
query = "SELECT * FROM users WHERE username = ?"
cursor.execute(query, (username,))

# ✅ 使用 ORM
user = session.query(User).filter(User.username == username).first()
```

### 防止 XSS

```javascript
// ❌ 危險：XSS
element.innerHTML = userInput;

// ✅ 安全：使用 textContent
element.textContent = userInput;

// ✅ 使用框架的安全機制（React）
<div>{userInput}</div>  // React 自動 escape
```

### 敏感資訊

```python
# ❌ 不要硬編碼
API_KEY = "sk-1234567890abcdef"
DATABASE_URL = "postgresql://user:password@localhost/db"

# ✅ 使用環境變數
import os
API_KEY = os.getenv("API_KEY")
DATABASE_URL = os.getenv("DATABASE_URL")

# ❌ 不要記錄敏感資訊
logger.info(f"User password: {password}")

# ✅ 遮蔽敏感資訊
logger.info(f"User ID: {user_id}")
```

---

## 測試

### 測試命名

```python
# 清楚描述測試內容
def test_user_login_with_valid_credentials_should_return_token():
    # Arrange
    username = "testuser"
    password = "password123"

    # Act
    token = auth_service.login(username, password)

    # Assert
    assert token is not None
    assert len(token) > 0
```

### 測試結構（AAA 模式）

```python
def test_calculate_discount_for_member():
    # Arrange - 準備測試資料
    price = 100
    discount_rate = 0.1
    is_member = True

    # Act - 執行被測試的功能
    result = calculate_discount(price, discount_rate, is_member)

    # Assert - 驗證結果
    assert result == 85.5
```

---

## 自動化工具

### Python

**Formatter**：
```bash
# Black（推薦）
pip install black
black your_file.py

# autopep8
pip install autopep8
autopep8 --in-place your_file.py
```

**Linter**：
```bash
# Flake8（推薦）
pip install flake8
flake8 your_file.py

# Pylint
pip install pylint
pylint your_file.py
```

**Type Checker**：
```bash
# mypy
pip install mypy
mypy your_file.py
```

### JavaScript / TypeScript

**Formatter**：
```bash
# Prettier（推薦）
npm install --save-dev prettier
npx prettier --write your_file.js
```

**Linter**：
```bash
# ESLint（推薦）
npm install --save-dev eslint
npx eslint your_file.js
```

### 配置範例

**.prettierrc** (JavaScript):
```json
{
  "semi": true,
  "singleQuote": true,
  "tabWidth": 2,
  "trailingComma": "es5"
}
```

**.flake8** (Python):
```ini
[flake8]
max-line-length = 88
extend-ignore = E203, W503
exclude = .git,__pycache__,venv
```

---

## 配置到專案

在 `SESSION_INIT.md` 中加入：

```markdown
### 程式碼風格

- **Python**: PEP 8, Black formatter
- **JavaScript**: Airbnb Style Guide, Prettier
- **Formatter**: 使用自動化工具（Black / Prettier）
- **Linter**: Flake8 / ESLint
- **註解語言**: [繁體中文 / English]
```

---

**最後更新**：2025-11-14
**版本**：2.0（通用版本）
**維護者**：開源社群
