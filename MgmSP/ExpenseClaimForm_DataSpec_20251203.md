# 費用申請單 (Expense Claim Form) 資料規格與對應檢查

**日期**: 2025-12-03
**目的**: 確保程式碼寫入與資料庫結構一致，並標示潛在問題。

## ⚠️ 重大異常 (Blockers)

在開始測試寫入前，必須解決以下結構性問題，否則 SQL 寫入必定失敗：

1.  **jPCH1 (費用明細) 寫入失敗**:
    *   **問題**: 資料庫中 `jPCH1` 的主鍵為 `jID` + `LineNum`，且 `jID` 不允許 NULL 且**不是 Identity (自動遞增)**。
    *   **程式碼**: 目前的 INSERT 語法僅寫入 `DocEntry`，未寫入 `jID`。
    *   **影響**: 違反 Not Null 限制，寫入失敗。
    *   **建議**: 確認 `jPCH1.jID` 是否應填入表頭的 ID (即 `DocEntry`)。若是，需修正 SQL 語句：`INSERT INTO jPCH1 (jID, DocEntry, ...)`。
A：此欄位應為自動產生，自動遞增才對。

2.  **jOPCH (表頭) 缺少 Attachment 欄位**:
    *   **問題**: 程式碼嘗試寫入 `@Attachment` 到 `jOPCH`。
    *   **現況**: 資料庫 `jOPCH` 表中**不存在** `Attachment` 欄位。
    *   **影響**: `Invalid column name 'Attachment'` 錯誤。
    *   **建議**: 於資料庫新增欄位，或暫時移除程式碼中的寫入動作。
A：請新增Attachment，且我希望修改程式，讓使用者不只可以Attach一個文件。
請新增一個Table為jAttach，AttachFile會上傳到SapB1Solution\MgmSP\AttachFile\ExpenseClaimForm
jAttach中則紀錄與jOPCH對應的jID,AP DocEntry,LineNum,FilePath,Uploader,UploadDate,UploadTime，
請先給我這個表格的規格建議，包括欄位類型與長度，還有鍵值，我確認無誤後再請你新增。

3.  **MDR 明細 (jMGUIAPDetail) 結構確認**:
    *   **現況**: 程式碼將 MDR 表頭的 ID (`jMGUIAP.ID`) 寫入明細的 `jID` 欄位。這在邏輯上看似正確 (Parent-Child)，但需確認 `jMGUIAPDetail` 的 PK 設計是否確為 (`jID`, `LineNum`)。目前的 DB Schema 顯示確實如此。

---

## 1. AP 表頭 (Header)

*   **對應資料表**: `jOPCH`
*   **主鍵**: `jID` (Identity)

| UI 欄位 | 程式碼變數/參數 | 資料庫欄位 (`jOPCH`) | 類型 | 備註/檢查結果 |
| :--- | :--- | :--- | :--- | :--- |
| (系統自動) | (Identity) | `jID` | int | PK, Identity (OK) |
| AP 單號 (B1) | `@DocEntry` | `DocEntry` | int | |
| (關聯單號) | `@DocNum` | `DocNum` | int | |
| 供應商代碼 | `@CardCode` | `CardCode` | nvarchar(50) | |
| 供應商名稱 | `@CardName` | `CardName` | nvarchar(100) | |
| 供應商參考號 | `@NumAtCard` | `NumAtCard` | nvarchar(100) | |
| 文件日期 | `@DocDate` | `DocDate` | date | |
| 到期日 | `@DocDueDate` | `DocDueDate` | date | |
| 過帳日期 | `@TaxDate` | `TaxDate` | date | |
| 文件幣別 | `@DocCurrency` | `DocCurrency` | nvarchar(3) | |
| 匯率 | `@DocRate` | `DocRate` | decimal | |
| 單據總額 (含稅) | `@DocTotal` | `DocTotal` | decimal | **注意**: DB 定義為含稅總額 |
| 稅額 | `@VatSum` | `VatSum` | decimal | |
| 付款條件 | `@GroupNum` | `GroupNum` | int | |
| 備註 | `@Comments` | `Comments` | nvarchar(254) | |
| 附件 | `@Attachment` | **(缺失)** | - | **⚠️ 資料表無此欄位，寫入將失敗** |
| 文件狀態 | `@Status` | `ApprovalStatus` | nvarchar(20) | P/W/A/R |
| 建立者 | `@User` | `CreateBy` | nvarchar(50) | |
| 簽核 PID | `@UPID` | `U_PID` | int | |
| 收貨地址名稱 | `@AddressName` | `AddressName` | nvarchar(100) | |
| 收貨地址 | `@Address` | `Address` | nvarchar(254) | |

## 2. AP 明細 (Expense Lines)

*   **對應資料表**: `jPCH1`
*   **主鍵**: `jID`, `LineNum`

| UI 欄位 | 程式碼變數/參數 | 資料庫欄位 (`jPCH1`) | 類型 | 備註/檢查結果 |
| :--- | :--- | :--- | :--- | :--- |
| (關聯 ID) | **(缺失)** | `jID` | int | **⚠️ PK, Not Null。程式碼未寫入** |
| 列號 | `@LineNum` | `LineNum` | int | PK |
| AP 單號 | `@DocEntry` | `DocEntry` | int | |
| 費用類別 | `@ItemCode` | `ItemCode` | nvarchar(50) | |
| 說明 | `@Dscription` | `Dscription` | nvarchar(200) | |
| 會計科目 | `@AcctCode` | `AcctCode` | nvarchar(50) | |
| 未稅金額 | `@LineTotal` | `LineTotal` | decimal | |
| 稅別 | `@VatGroup` | `VatGroup` | nvarchar(20) | |
| 稅率 | `@VatPrcnt` | `VatPrcnt` | decimal | |
| 稅額 | `@LineVat` | `LineVat` | decimal | |
| 含稅金額 | `@GTotal` | `GTotal` | decimal | |
| 產品 | `@CostingCode` | `CostingCode` | nvarchar(50) | |
| 部門 | `@CostingCode2` | `CostingCode2` | nvarchar(50) | |
| (附件) | - | `Attachment` | nvarchar(500) | DB 有欄位但程式碼未寫入 (可選) |

## 3. MDR 發票明細 (MDR Details)

此部分包含表頭彙總 (`jMGUIAP`) 與明細 (`jMGUIAPDetail`)。

### 3.1 MDR 表頭 (`jMGUIAP`)

*   **對應資料表**: `jMGUIAP`
*   **主鍵**: `ID` (Identity)

| 欄位 | 程式碼變數 | 資料庫欄位 | 類型 | 備註/檢查結果 |
| :--- | :--- | :--- | :--- | :--- |
| ID (PK) | (Identity) | `ID` | int | PK, Identity (OK) |
| (關聯 ID) | **(缺失)** | `jID` | int | **⚠️ Not Null, No Default。程式碼未寫入** (建議填入 0 或 Header ID) |
| AP 單號 | `@DocEntry` | `DocEntry` | int | |
| (AP 單號) | `@DocNum` | `DocNum` | int | |
| 未稅總額 | `@DocTotal` | `DocTotal` | decimal | Code 邏輯為未稅 (Sum of HWBAS) |
| 稅額總計 | `@VatSum` | `VatSum` | decimal | |
| 建立者 | `@User` | `CreateBy` | nvarchar(50) | |

### 3.2 MDR 明細 (`jMGUIAPDetail`)

*   **對應資料表**: `jMGUIAPDetail`
*   **主鍵**: `jID`, `LineNum`

| UI 欄位 | 程式碼變數 | 資料庫欄位 | 類型 | 備註/檢查結果 |
| :--- | :--- | :--- | :--- | :--- |
| MDR ID | `@jID` | `jID` | int | PK, 關聯至 `jMGUIAP.ID` (OK) |
| 列號 | `@LineNum` | `LineNum` | int | PK |
| AP 單號 | `@DocEntry` | `DocEntry` | int | |
| 供應商代碼 | `@LIFNR` | `U_LIFNR` | nvarchar(50) | |
| 統一編號 | `@STCEG` | `U_STCEG` | nvarchar(20) | |
| 發票號碼 | `@XBLNR` | `U_XBLNR` | nvarchar(50) | |
| 發票類型 | `@ZFORM` | `U_ZFORM_CODE`| nvarchar(10) | |
| 憑證日期 | `@BLDAT` | `U_BLDAT` | date | |
| 營業稅日期 | `@VATDATE` | `U_VATDATE` | date | |
| 未稅金額 | `@HWBAS` | `U_HWBAS` | decimal | |
| 稅額 | `@HWSTE` | `U_HWSTE` | decimal | |
| 稅別 | `@TAXTYPE` | `U_TAX_TYPE` | nvarchar(10) | |