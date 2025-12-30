/*
=============================================================================
AI 輔助功能 - 費用申請單欄位說明
目標資料庫：jtdb
頁面：/MgmSP/ExpenseClaimForm.aspx
=============================================================================
*/

DECLARE @PageURL NVARCHAR(500) = '/MgmSP/ExpenseClaimForm.aspx'
DECLARE @UpdateBy VARCHAR(50) = 'SYSTEM'

-- ============================================================================
-- 清除此頁面的舊說明（如需重新匯入）
-- ============================================================================
-- DELETE FROM [dbo].[AI_FieldHelp] WHERE [PageURL] = @PageURL

-- ============================================================================
-- 單頭區域 (HEADER)
-- ============================================================================

-- 供應商代碼
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtCardCode', N'供應商代碼',
N'輸入或搜尋供應商代碼。此為必填欄位。
• 可直接輸入已知的供應商代碼
• 或點擊搜尋按鈕從清單中選擇
• 選擇後系統會自動帶入供應商名稱及相關資訊', 1, GETDATE(), @UpdateBy)

-- 供應商名稱
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtCardName', N'供應商名稱',
N'顯示供應商的完整名稱。
• 可直接輸入名稱進行搜尋
• 或由供應商代碼自動帶入
• 此名稱會顯示在列印文件上', 1, GETDATE(), @UpdateBy)

-- 供應商參考號
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtNumAtCard', N'供應商參考號',
N'供應商的單據編號或參考資訊。
• 可填入供應商的發票號碼
• 或其他供應商提供的參考編號
• 此欄位為選填', 1, GETDATE(), @UpdateBy)

-- 文件幣別
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'ddlDocCurrency', N'文件幣別',
N'選擇此筆費用申請的幣別。
• 預設為本國幣別（TWD）
• 選擇外幣時系統會自動帶入當日匯率
• 幣別變更會影響所有明細的金額計算', 1, GETDATE(), @UpdateBy)

-- 匯率
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtDocRate', N'匯率',
N'外幣對本國幣的匯率。
• 系統會根據過帳日期自動帶入
• 可手動修改匯率
• 點擊更新按鈕可重新取得系統匯率', 1, GETDATE(), @UpdateBy)

-- 收貨地址名稱
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'ddlDeliveryAddr', N'收貨地址名稱',
N'選擇收貨地址。
• 從公司已設定的地址清單中選擇
• 選擇後會自動帶入完整地址', 1, GETDATE(), @UpdateBy)

-- 收貨地址
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtAddress', N'收貨地址',
N'完整的收貨地址。
• 由收貨地址名稱自動帶入
• 可手動修改地址內容', 1, GETDATE(), @UpdateBy)

-- 付款條件
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'ddlGroupNum', N'付款條件',
N'選擇付款條件。
• 會影響到期日的計算
• 常見選項如：月結30天、月結60天等
• 選擇後系統會自動計算到期日', 1, GETDATE(), @UpdateBy)

-- 付款條件(列印)
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtPymntGroup', N'付款條件(列印)',
N'列印文件時顯示的付款條件名稱。
• 可自訂列印時顯示的文字
• 若不填寫則使用系統預設名稱', 1, GETDATE(), @UpdateBy)

-- 簽核系統 PID
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtUPID', N'簽核系統 PID',
N'簽核流程的識別碼。
• 送審後由系統自動產生
• 可用於追蹤簽核進度', 1, GETDATE(), @UpdateBy)

-- 過帳日期
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtDocDate', N'過帳日期',
N'會計過帳的日期。此為必填欄位。
• 決定此筆費用入帳的會計期間
• 會影響匯率的取得
• 通常填寫實際發生費用的日期', 1, GETDATE(), @UpdateBy)

-- 到期日
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtDocDueDate', N'到期日',
N'付款到期日。此為必填欄位。
• 依付款條件自動計算
• 可手動修改
• 用於應付帳款的付款排程', 1, GETDATE(), @UpdateBy)

-- 文件日期
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, 'txtTaxDate', N'文件日期',
N'單據的文件日期。此為必填欄位。
• 通常為發票或收據上的日期
• 會影響營業稅申報期間', 1, GETDATE(), @UpdateBy)

-- ============================================================================
-- 單身區域 - 費用申請明細頁籤 (DETAIL - expense)
-- ============================================================================

-- 費用類別
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'ddlExpCategory', N'費用類別',
N'選擇費用的分類。
• 不同類別會自動帶入對應的會計科目
• 請依實際費用性質選擇適當類別
• 常見類別：交通費、餐費、文具用品等', 1, GETDATE(), @UpdateBy)

-- 說明
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'txtDescription', N'說明',
N'費用的詳細說明。
• 說明此筆費用的用途或原因
• 例如：拜訪客戶交通費、部門聚餐等
• 清楚的說明有助於審核', 1, GETDATE(), @UpdateBy)

-- 會計科目
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'txtAcctCode', N'會計科目',
N'費用入帳的會計科目。
• 通常由費用類別自動帶入
• 可點擊搜尋按鈕手動選擇
• 請確認科目正確以利會計作帳', 1, GETDATE(), @UpdateBy)

-- 未稅金額
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'txtLineTotal', N'未稅金額',
N'不含稅的費用金額。
• 輸入後系統會自動計算稅額
• 與含稅金額連動計算
• 請確認金額正確', 1, GETDATE(), @UpdateBy)

-- 稅別
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'ddlVatGroup', N'稅別',
N'選擇適用的稅率。
• 應稅：一般 5% 營業稅
• 零稅：稅率 0%
• 免稅：免徵營業稅
• 選擇後會自動重算稅額', 1, GETDATE(), @UpdateBy)

-- 稅額
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'txtVatSum', N'稅額',
N'營業稅金額。
• 由系統根據稅別自動計算
• 可手動修改（會影響含稅金額）
• 應與發票上的稅額一致', 1, GETDATE(), @UpdateBy)

-- 含稅金額
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'txtPriceAfterVat', N'含稅金額',
N'包含稅額的總金額。
• 等於未稅金額加上稅額
• 可直接輸入含稅金額，系統會反算未稅及稅額
• 應與實際支付金額一致', 1, GETDATE(), @UpdateBy)

-- 產品
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'ddlCostingCode', N'產品',
N'費用歸屬的產品線或專案。
• 用於成本分攤
• 請選擇此費用相關的產品', 1, GETDATE(), @UpdateBy)

-- 部門
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', 'ddlCostingCode2', N'部門',
N'費用歸屬的部門。
• 用於部門成本分析
• 預設為您所屬的部門
• 必要時可選擇其他部門', 1, GETDATE(), @UpdateBy)

-- ============================================================================
-- 單身區域 - 憑證明細頁籤 (DETAIL - mdr)
-- ============================================================================

-- 統一編號
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'txtSTCEG', N'統一編號',
N'開立發票之廠商的統一編號。
• 請填入發票上的賣方統一編號
• 用於營業稅申報勾稽', 1, GETDATE(), @UpdateBy)

-- 憑證號碼
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'txtXBLNR', N'憑證號碼',
N'發票號碼或收據編號。
• 請填入完整的發票號碼（如：AB-12345678）
• 電子發票請填入完整字軌號碼
• 此為重要的稅務憑證依據', 1, GETDATE(), @UpdateBy)

-- 憑證類型
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'ddlZFORM_CODE', N'憑證類型',
N'發票或憑證的類型。
• 21-三聯手開發票：傳統三聯式發票
• 22-高鐵/二聯收銀機：長條型收據
• 25-電子發票/公營事業：電子發票、水電費等
• 28-海關代徵營業稅：進口貨物
• 99-其他：其他類型憑證', 1, GETDATE(), @UpdateBy)

-- 憑證日期
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'txtBLDAT', N'憑證日期',
N'發票或憑證上的日期。
• 請填入發票上的開立日期
• 此日期用於稅務申報', 1, GETDATE(), @UpdateBy)

-- 營業稅日期
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'txtVATDATE', N'營業稅日期',
N'營業稅申報的所屬期間。
• 決定此發票要在哪個期間申報
• 通常與憑證日期相同
• 跨月發票可能需要調整', 1, GETDATE(), @UpdateBy)

-- 未稅金額 (憑證)
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'txtHWBAS', N'未稅金額',
N'憑證上的未稅銷售額。
• 請填入發票上的銷售額
• 系統會根據稅別自動計算稅額', 1, GETDATE(), @UpdateBy)

-- 稅別 (憑證)
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'ddlTAX_TYPE', N'稅別',
N'憑證適用的稅率類型。
• 1-應稅：一般 5% 營業稅
• 2-零稅：適用零稅率
• 3-免稅：免徵營業稅', 1, GETDATE(), @UpdateBy)

-- 稅額 (憑證)
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', 'txtHWSTE', N'稅額',
N'憑證上的營業稅金額。
• 請填入發票上的稅額
• 應與發票上金額一致
• 可手動調整以符合實際發票', 1, GETDATE(), @UpdateBy)

-- ============================================================================
-- 單尾區域 (FOOTER)
-- ============================================================================

-- 採購人員
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'FOOTER', NULL, 'ddlPurchaser', N'採購人員',
N'負責此筆費用的採購人員。
• 從人員清單中選擇
• 用於追蹤費用負責人', 1, GETDATE(), @UpdateBy)

-- 備註
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'FOOTER', NULL, 'txtRemarks', N'備註',
N'單據的補充說明。
• 可填寫額外需要說明的事項
• 審核人員會看到此備註
• 此欄位為選填', 1, GETDATE(), @UpdateBy)

-- ============================================================================
-- 區域層級說明 (ElementID 為 NULL)
-- ============================================================================

-- 單頭區域說明
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'HEADER', NULL, NULL, N'表頭資訊',
N'費用申請單的基本資訊區域。
• 包含供應商資訊、日期、幣別等
• 紅色星號 (*) 標示為必填欄位
• 請先完成表頭資訊再填寫明細', 1, GETDATE(), @UpdateBy)

-- 費用申請明細頁籤說明
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'費用申請明細', NULL, N'費用申請明細',
N'填寫費用的明細項目。
• 點擊「新增明細」新增費用項目
• 每行填寫一筆費用
• 可勾選後點擊「刪除選取」移除
• 點擊「產生憑證明細」可自動產生對應的憑證資料', 1, GETDATE(), @UpdateBy)

-- 憑證明細頁籤說明
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'DETAIL', N'憑證明細', NULL, N'憑證明細',
N'填寫發票或收據的資訊。
• 此區塊用於營業稅申報
• 可從費用明細自動產生
• 請確認統一編號、憑證號碼、金額正確
• 資料會傳送至稅務系統', 1, GETDATE(), @UpdateBy)

-- 單尾區域說明
INSERT INTO [dbo].[AI_FieldHelp] ([PageURL], [AreaCode], [TabName], [ElementID], [HelpTitle], [HelpContent], [IsActive], [UpdateTime], [UpdateBy])
VALUES (@PageURL, 'FOOTER', NULL, NULL, N'單據資訊',
N'單據的彙總資訊與送審。
• 顯示單據總金額
• 填寫完成後點擊「儲存並送審」送出
• 或點擊「暫存」儲存草稿稍後繼續編輯', 1, GETDATE(), @UpdateBy)

-- ============================================================================
-- 頁面區域定義
-- ============================================================================

-- 先檢查是否已有資料
IF NOT EXISTS (SELECT 1 FROM [dbo].[AI_PageArea] WHERE [PageURL] = @PageURL)
BEGIN
    INSERT INTO [dbo].[AI_PageArea] ([PageURL], [AreaCode], [AreaName], [ContainerID], [SortOrder], [IsActive])
    VALUES
        (@PageURL, 'HEADER', N'表頭資訊', 'divHeader', 1, 1),
        (@PageURL, 'DETAIL', N'明細資料', 'divDetail', 2, 1),
        (@PageURL, 'FOOTER', N'單據資訊', 'divFooter', 3, 1)
    PRINT 'Inserted AI_PageArea for ExpenseClaimForm'
END

PRINT ''
PRINT '============================================================================='
PRINT '費用申請單欄位說明匯入完成'
PRINT '總計: 約 30 筆欄位說明'
PRINT '============================================================================='
