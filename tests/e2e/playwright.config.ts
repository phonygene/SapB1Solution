import { defineConfig, devices } from '@playwright/test';

/**
 * JET Enterprise Platform - Playwright 設定
 *
 * 執行方式：
 *   npm test              - 執行所有測試
 *   npm run test:ui       - 開啟 UI 模式
 *   npm run test:headed   - 顯示瀏覽器執行
 *   npm run test:pr       - 只跑請購單測試
 */

export default defineConfig({
  // 測試目錄
  testDir: './tests',

  // 測試檔案匹配模式
  testMatch: '**/*.spec.ts',

  // 平行執行設定
  fullyParallel: false,  // 財務系統不建議平行（可能有資料衝突）

  // 失敗重試
  retries: process.env.CI ? 2 : 0,

  // 報告格式
  reporter: [
    ['html', { outputFolder: 'playwright-report' }],
    ['list']
  ],

  // 全域設定
  use: {
    // 基礎 URL（本機開發）
    baseURL: process.env.BASE_URL || 'http://localhost:51062',

    // 追蹤設定（失敗時記錄）
    trace: 'on-first-retry',

    // 截圖設定
    screenshot: 'only-on-failure',

    // 影片設定
    video: 'on-first-retry',

    // 超時設定
    actionTimeout: 15000,
    navigationTimeout: 30000,
  },

  // 全域超時
  timeout: 60000,

  // 測試專案（瀏覽器）
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
    // 如需測試其他瀏覽器可取消註解
    // {
    //   name: 'firefox',
    //   use: { ...devices['Desktop Firefox'] },
    // },
    // {
    //   name: 'webkit',
    //   use: { ...devices['Desktop Safari'] },
    // },
  ],

  // 本機開發伺服器（如果需要自動啟動）
  // webServer: {
  //   command: 'dotnet run',
  //   url: 'http://localhost:51062',
  //   reuseExistingServer: !process.env.CI,
  // },
});
