import { Page, expect } from '@playwright/test';

/**
 * 請購單頁面物件
 */
export class PurchaseRequestPage {
  readonly page: Page;

  // 表頭元素
  readonly txtReqCode;
  readonly txtReqName;
  readonly ddlReqDept;
  readonly txtDocDate;
  readonly txtReqDate;
  readonly ddlCardCode;
  readonly txtCardName;
  readonly ddlDocCurrency;
  readonly ddlPurchaser;

  // 明細 GridView
  readonly gvLines;

  // 按鈕
  readonly btnSave;
  readonly btnUpdate;
  readonly btnNewDocument;
  readonly btnAddLine;
  readonly btnExportPDF;
  readonly btnApprove;
  readonly btnReject;
  readonly btnPostToSAP;

  // 訊息
  readonly successMessage;
  readonly errorMessage;
  readonly warningMessage;

  // 搜尋彈窗
  readonly mpeReqName;
  readonly mpeItemSearch;
  readonly mpeVendorSearch;

  constructor(page: Page) {
    this.page = page;

    // 表頭
    this.txtReqCode = page.locator('#MainContent_txtReqCode');
    this.txtReqName = page.locator('#MainContent_txtReqName');
    this.ddlReqDept = page.locator('#MainContent_ddlReqDept');
    this.txtDocDate = page.locator('#MainContent_txtDocDate');
    this.txtReqDate = page.locator('#MainContent_txtReqDate');
    this.ddlCardCode = page.locator('#MainContent_ddlCardCode');
    this.txtCardName = page.locator('#MainContent_txtCardName');
    this.ddlDocCurrency = page.locator('#MainContent_ddlDocCurrency');
    this.ddlPurchaser = page.locator('#MainContent_ddlPurchaser');

    // 明細
    this.gvLines = page.locator('#MainContent_gvLines');

    // 按鈕
    this.btnSave = page.locator('#MainContent_btnSave');
    this.btnUpdate = page.locator('#MainContent_btnUpdate');
    this.btnNewDocument = page.locator('#MainContent_btnNewDocument');
    this.btnAddLine = page.locator('#MainContent_btnAddLine');
    this.btnExportPDF = page.locator('#MainContent_btnExportPDF');
    this.btnApprove = page.locator('#MainContent_btnApprove');
    this.btnReject = page.locator('#MainContent_btnReject');
    this.btnPostToSAP = page.locator('#MainContent_btnPostToSAP');

    // 訊息
    this.successMessage = page.locator('.alert-success, .success-message');
    this.errorMessage = page.locator('.alert-danger, .error-message');
    this.warningMessage = page.locator('.alert-warning, .warning-message');

    // 彈窗
    this.mpeReqName = page.locator('#MainContent_pnlReqNameSearch');
    this.mpeItemSearch = page.locator('#MainContent_pnlItemSearch');
    this.mpeVendorSearch = page.locator('#MainContent_pnlVendorSearch');
  }

  /**
   * 前往請購單頁面（新增模式）
   */
  async gotoNew() {
    await this.page.goto('/PurchaseRequestForm.aspx');
  }

  /**
   * 開啟指定的請購單
   */
  async gotoEdit(jID: number) {
    await this.page.goto(`/PurchaseRequestForm.aspx?jID=${jID}`);
  }

  /**
   * 填寫表頭基本資料
   */
  async fillHeader(data: {
    reqName?: string;
    reqDept?: string;
    docDate?: string;
    reqDate?: string;
    vendor?: string;
    currency?: string;
  }) {
    if (data.reqName) {
      // 點擊搜尋按鈕開啟彈窗
      await this.page.locator('#MainContent_btnSearchReqName').click();
      await this.mpeReqName.waitFor({ state: 'visible' });

      // 在彈窗中搜尋
      await this.page.locator('#MainContent_txtReqNameKeyword').fill(data.reqName);
      await this.page.locator('#MainContent_btnDoSearchReqName').click();

      // 等待結果並選擇第一筆
      await this.page.locator('#MainContent_gvReqNameSearch tr:nth-child(2)').waitFor();
      await this.page.locator('#MainContent_gvReqNameSearch tr:nth-child(2) a').first().click();
    }

    if (data.reqDept) {
      await this.ddlReqDept.selectOption({ label: data.reqDept });
    }

    if (data.docDate) {
      await this.txtDocDate.fill(data.docDate);
    }

    if (data.reqDate) {
      await this.txtReqDate.fill(data.reqDate);
    }

    if (data.currency) {
      await this.ddlDocCurrency.selectOption({ label: data.currency });
    }
  }

  /**
   * 新增明細行
   */
  async addLine(data: {
    itemCode: string;
    quantity: number;
    price: number;
    warehouse?: string;
  }) {
    // 點擊新增明細按鈕
    await this.btnAddLine.click();

    // 等待新行出現
    const lastRow = this.gvLines.locator('tr').last();

    // 填寫品號（開啟搜尋彈窗）
    await lastRow.locator('[id*="btnSearchItem"]').click();
    await this.mpeItemSearch.waitFor({ state: 'visible' });

    await this.page.locator('#MainContent_txtItemKeyword').fill(data.itemCode);
    await this.page.locator('#MainContent_btnDoSearchItem').click();

    // 選擇搜尋結果
    await this.page.locator('#MainContent_gvItemSearch tr:nth-child(2)').waitFor();
    await this.page.locator('#MainContent_gvItemSearch tr:nth-child(2) a').first().click();

    // 填寫數量和單價
    await lastRow.locator('[id*="txtQuantity"]').fill(data.quantity.toString());
    await lastRow.locator('[id*="txtPrice"]').fill(data.price.toString());

    if (data.warehouse) {
      await lastRow.locator('[id*="ddlWarehouse"]').selectOption({ label: data.warehouse });
    }
  }

  /**
   * 儲存單據
   */
  async save() {
    await this.btnSave.click();
    // 等待 PostBack 完成
    await this.page.waitForLoadState('networkidle');
  }

  /**
   * 更新單據
   */
  async update() {
    await this.btnUpdate.click();
    await this.page.waitForLoadState('networkidle');
  }

  /**
   * 驗證儲存成功
   */
  async expectSaveSuccess() {
    await expect(this.successMessage).toBeVisible({ timeout: 10000 });
    // 確認 URL 包含 jID（代表已儲存）
    await expect(this.page).toHaveURL(/jID=\d+/);
  }

  /**
   * 驗證錯誤訊息
   */
  async expectError(message?: string) {
    await expect(this.errorMessage).toBeVisible();
    if (message) {
      await expect(this.errorMessage).toContainText(message);
    }
  }

  /**
   * 取得當前 jID
   */
  async getJID(): Promise<number | null> {
    const url = this.page.url();
    const match = url.match(/jID=(\d+)/);
    return match ? parseInt(match[1]) : null;
  }

  /**
   * 取得表頭總金額
   */
  async getDocTotal(): Promise<number> {
    const text = await this.page.locator('#MainContent_txtDocTotal').inputValue();
    return parseFloat(text.replace(/,/g, '')) || 0;
  }

  /**
   * 匯出 PDF
   */
  async exportPDF() {
    // 開啟新視窗/下載
    const [download] = await Promise.all([
      this.page.waitForEvent('download'),
      this.btnExportPDF.click(),
    ]);
    return download;
  }

  /**
   * 審核通過
   */
  async approve(comments?: string) {
    if (comments) {
      await this.page.locator('#MainContent_txtApprovalComments').fill(comments);
    }
    await this.btnApprove.click();
    await this.page.waitForLoadState('networkidle');
  }

  /**
   * 駁回
   */
  async reject(comments: string) {
    await this.page.locator('#MainContent_txtApprovalComments').fill(comments);
    await this.btnReject.click();
    await this.page.waitForLoadState('networkidle');
  }

  /**
   * 過帳到 SAP
   */
  async postToSAP() {
    await this.btnPostToSAP.click();
    await this.page.waitForLoadState('networkidle');
  }
}
