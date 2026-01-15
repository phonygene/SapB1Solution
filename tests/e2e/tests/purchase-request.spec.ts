import { test, expect } from '@playwright/test';
import { LoginPage } from '../pages/login.page';
import { PurchaseRequestPage } from '../pages/purchase-request.page';

/**
 * 請購單 E2E 測試
 *
 * 測試涵蓋：
 * 1. 基本 CRUD 操作
 * 2. 金額計算正確性
 * 3. 審核流程
 * 4. PDF 匯出
 * 5. SAP 過帳（需要 SAP 環境）
 */

// 測試資料（可從環境變數覆寫）
const TEST_USER = process.env.TEST_USER || 'testuser';
const TEST_PASSWORD = process.env.TEST_PASSWORD || 'test123';

test.describe('請購單功能測試', () => {
  let loginPage: LoginPage;
  let prPage: PurchaseRequestPage;

  test.beforeEach(async ({ page }) => {
    loginPage = new LoginPage(page);
    prPage = new PurchaseRequestPage(page);

    // 登入
    await loginPage.goto();
    await loginPage.login(TEST_USER, TEST_PASSWORD);
    await loginPage.expectLoginSuccess();
  });

  test.describe('基本操作', () => {
    test('應該能新增請購單', async ({ page }) => {
      // 前往新增頁面
      await prPage.gotoNew();

      // 填寫表頭
      await prPage.fillHeader({
        reqName: '王',  // 搜尋姓氏
        docDate: new Date().toISOString().split('T')[0],
      });

      // 新增明細
      await prPage.addLine({
        itemCode: 'A001',
        quantity: 10,
        price: 100,
      });

      // 儲存
      await prPage.save();

      // 驗證
      await prPage.expectSaveSuccess();
      const jID = await prPage.getJID();
      expect(jID).toBeGreaterThan(0);
    });

    test('應該能查詢並開啟已存在的請購單', async ({ page }) => {
      // 假設 jID=1 存在（實際測試可能需要先建立或用 fixture）
      await prPage.gotoEdit(1);

      // 驗證頁面載入成功
      await expect(prPage.txtReqName).not.toBeEmpty();
      await expect(prPage.gvLines.locator('tr')).toHaveCount.greaterThan(1);
    });

    test('應該能更新請購單', async ({ page }) => {
      await prPage.gotoEdit(1);

      // 修改備註
      await page.locator('#MainContent_txtComments').fill('E2E 測試更新 - ' + new Date().toISOString());

      // 更新
      await prPage.update();

      // 驗證
      await prPage.expectSaveSuccess();
    });
  });

  test.describe('金額計算', () => {
    test('明細金額應該正確計算（數量 × 單價）', async ({ page }) => {
      await prPage.gotoNew();

      await prPage.fillHeader({
        reqName: '王',
        docDate: new Date().toISOString().split('T')[0],
      });

      // 新增明細：10 × 100 = 1000
      await prPage.addLine({
        itemCode: 'A001',
        quantity: 10,
        price: 100,
      });

      // 檢查明細行的 LineTotal
      const lineTotal = await page.locator('#MainContent_gvLines tr:last-child [id*="txtLineTotal"]').inputValue();
      expect(parseFloat(lineTotal.replace(/,/g, ''))).toBe(1000);
    });

    test('表頭總金額應該等於明細加總', async ({ page }) => {
      await prPage.gotoNew();

      await prPage.fillHeader({
        reqName: '王',
        docDate: new Date().toISOString().split('T')[0],
      });

      // 新增多筆明細
      await prPage.addLine({ itemCode: 'A001', quantity: 10, price: 100 });  // 1000
      await prPage.addLine({ itemCode: 'A002', quantity: 5, price: 200 });   // 1000

      // 檢查總金額
      const docTotal = await prPage.getDocTotal();
      expect(docTotal).toBe(2000);
    });
  });

  test.describe('驗證規則', () => {
    test('缺少請購人應該顯示錯誤', async ({ page }) => {
      await prPage.gotoNew();

      // 只填日期，不填請購人
      await prPage.fillHeader({
        docDate: new Date().toISOString().split('T')[0],
      });

      await prPage.addLine({
        itemCode: 'A001',
        quantity: 1,
        price: 100,
      });

      await prPage.save();

      // 應該顯示錯誤
      await prPage.expectError('請購人');
    });

    test('沒有明細應該顯示錯誤', async ({ page }) => {
      await prPage.gotoNew();

      await prPage.fillHeader({
        reqName: '王',
        docDate: new Date().toISOString().split('T')[0],
      });

      // 不新增明細直接儲存
      await prPage.save();

      await prPage.expectError('明細');
    });
  });

  test.describe('審核流程', () => {
    test('應該能審核通過', async ({ page }) => {
      // 假設 jID=2 是待審核狀態
      await prPage.gotoEdit(2);

      await prPage.approve('E2E 測試審核通過');

      await expect(page.locator('#MainContent_lblApprovalStatus')).toContainText('已核准');
    });

    test('應該能駁回', async ({ page }) => {
      // 假設 jID=3 是待審核狀態
      await prPage.gotoEdit(3);

      await prPage.reject('E2E 測試駁回原因');

      await expect(page.locator('#MainContent_lblApprovalStatus')).toContainText('已駁回');
    });
  });

  test.describe('PDF 匯出', () => {
    test('應該能匯出 PDF', async ({ page }) => {
      await prPage.gotoEdit(1);

      const download = await prPage.exportPDF();

      // 驗證下載的檔案
      expect(download.suggestedFilename()).toMatch(/\.pdf$/i);
    });
  });

  // SAP 整合測試（需要 SAP 環境，預設跳過）
  test.describe('SAP 整合', () => {
    test.skip('應該能過帳到 SAP', async ({ page }) => {
      // 假設 jID=4 是已審核但未過帳
      await prPage.gotoEdit(4);

      await prPage.postToSAP();

      // 驗證過帳成功
      await expect(page.locator('#MainContent_lblB1PostStatus')).toContainText('已過帳');
      await expect(page.locator('#MainContent_txtDocEntry')).not.toBeEmpty();
    });
  });
});

// 冒煙測試：快速驗證核心功能
test.describe('冒煙測試', () => {
  test('登入 → 新增請購單 → 儲存 → PDF', async ({ page }) => {
    const loginPage = new LoginPage(page);
    const prPage = new PurchaseRequestPage(page);

    // 1. 登入
    await loginPage.goto();
    await loginPage.login(TEST_USER, TEST_PASSWORD);
    await loginPage.expectLoginSuccess();

    // 2. 新增請購單
    await prPage.gotoNew();
    await prPage.fillHeader({
      reqName: '王',
      docDate: new Date().toISOString().split('T')[0],
    });
    await prPage.addLine({
      itemCode: 'A001',
      quantity: 1,
      price: 100,
    });

    // 3. 儲存
    await prPage.save();
    await prPage.expectSaveSuccess();

    // 4. PDF（可選）
    // const download = await prPage.exportPDF();
    // expect(download).toBeTruthy();
  });
});
