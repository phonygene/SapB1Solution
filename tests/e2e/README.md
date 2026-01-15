# JET E2E 測試

使用 Playwright 進行端對端測試。

## 快速開始

```bash
# 1. 安裝依賴
cd tests/e2e
npm install

# 2. 安裝 Playwright 瀏覽器
npx playwright install chromium

# 3. 執行測試
npm test
```

## 執行方式

| 命令 | 說明 |
|------|------|
| `npm test` | 執行所有測試（無頭模式） |
| `npm run test:ui` | 開啟 Playwright UI 模式 |
| `npm run test:headed` | 顯示瀏覽器執行 |
| `npm run test:debug` | 除錯模式 |
| `npm run test:pr` | 只跑請購單測試 |
| `npm run report` | 查看測試報告 |

## 設定

### 環境變數

| 變數 | 說明 | 預設值 |
|------|------|--------|
| `BASE_URL` | 測試目標網址 | `http://localhost:51062` |
| `TEST_USER` | 測試帳號 | `testuser` |
| `TEST_PASSWORD` | 測試密碼 | `test123` |

### 修改設定

編輯 `playwright.config.ts`：

```typescript
use: {
  baseURL: 'http://your-server/MgmSP',
  // ...
}
```

## 目錄結構

```
tests/e2e/
├── package.json           # 依賴
├── playwright.config.ts   # Playwright 設定
├── pages/                 # Page Object Model
│   ├── login.page.ts
│   └── purchase-request.page.ts
└── tests/                 # 測試案例
    └── purchase-request.spec.ts
```

## 撰寫新測試

1. 在 `pages/` 建立頁面物件
2. 在 `tests/` 建立測試檔案
3. 使用 Page Object 模式

範例：

```typescript
import { test } from '@playwright/test';
import { LoginPage } from '../pages/login.page';

test('我的測試', async ({ page }) => {
  const loginPage = new LoginPage(page);
  await loginPage.goto();
  // ...
});
```

## 注意事項

1. **測試環境隔離**：建議使用獨立測試資料庫
2. **測試資料清理**：測試後清理建立的資料
3. **平行執行**：財務系統預設關閉平行（避免資料衝突）
4. **SAP 測試**：SAP 相關測試預設跳過，需要 SAP 環境才能執行
