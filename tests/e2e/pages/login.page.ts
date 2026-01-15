import { Page, expect } from '@playwright/test';

/**
 * 登入頁面物件
 */
export class LoginPage {
  readonly page: Page;

  // 元素定位器
  readonly userInput;
  readonly passwordInput;
  readonly loginButton;
  readonly errorMessage;

  constructor(page: Page) {
    this.page = page;
    this.userInput = page.locator('#txtUserName');
    this.passwordInput = page.locator('#txtPassword');
    this.loginButton = page.locator('#btnLogin');
    this.errorMessage = page.locator('.error-message, .alert-danger');
  }

  /**
   * 前往登入頁面
   */
  async goto() {
    await this.page.goto('/Login.aspx');
  }

  /**
   * 執行登入
   */
  async login(username: string, password: string) {
    await this.userInput.fill(username);
    await this.passwordInput.fill(password);
    await this.loginButton.click();

    // 等待頁面跳轉或錯誤訊息
    await Promise.race([
      this.page.waitForURL('**/Home.aspx', { timeout: 10000 }),
      this.errorMessage.waitFor({ state: 'visible', timeout: 10000 }).catch(() => {}),
    ]);
  }

  /**
   * 驗證登入成功
   */
  async expectLoginSuccess() {
    await expect(this.page).toHaveURL(/Home\.aspx/);
  }

  /**
   * 驗證登入失敗
   */
  async expectLoginFailed() {
    await expect(this.errorMessage).toBeVisible();
  }
}
