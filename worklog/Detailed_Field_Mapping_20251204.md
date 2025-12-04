# 詳細欄位對應分析報告 (Detailed Field Mapping Analysis)

**日期**: 2025-12-04
**目的**: 確認 UI 所有欄位在資料庫中皆有對應，並指出程式碼寫入錯誤。

## ⚠️ 發現的重大缺失 (Critical Gaps)

經過逐一比對 ASPX UI 控制項、VB 後端程式碼與 SQL 資料表結構，發現以下具體問題：

### 1. 表頭 (jOPCH) 寫入失敗與資料遺失
*   **DeliveryAddrID (收貨地址 ID)**
    *   **UI**: `ddlDeliveryAddr` (DropDownList) 存在。
    *   **VB**: `INSERT` 語句包含 `DeliveryAddrID` 欄位 (`VALUES (..., @DeliveryAddrID, ...)`).
    *   **DB**: `jOPCH` 資料表 **沒有** `DeliveryAddrID` 欄位。
    *   **結果**: **寫入時會發生 SQL Error** (`Invalid column name 'DeliveryAddrID'`)。
*   **Purchaser (採購人員)**
    *   **UI**: `ddlPurchaser` (DropDownList) 存在，使用者可選擇採購人員。
    *   **VB**: 存檔時 (`SetHeaderParameters`) **完全忽略** 此欄位，僅將當前登入者 (`currentUserId`) 寫入 `CreateBy`。
    *   **DB**: `jOPCH` 資料表 **沒有** `SlpCode` (業務/採購員代碼) 欄位。
    *   **結果**: 使用者選擇的採購人員無法儲存，下次開啟單據會變成預設值或建立者。

### 2. MDR 表頭 (jMGUIAP) 寫入失敗
*   **jID (關聯 ID)**
    *   **DB**: `jMGUIAP` 有 `jID` 欄位，且為 `NOT NULL`。
    *   **VB**: `INSERT` 語句 **未包含** `jID`。
    *   **結果**: **寫入時會發生 SQL Error** (`Cannot insert the value NULL into column 'jID'`).

### 3. UI 欄位完整性檢查表

| 區塊 | UI 控制項 | VB 參數 (@Param) | DB 欄位 | 狀態 | 備註 |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **表頭** | `txtCardCode` | `@CardCode` | `CardCode` | ✅ OK | |
| | `txtCardName` | `@CardName` | `CardName` | ✅ OK | |
| | `txtNumAtCard` | `@NumAtCard` | `NumAtCard` | ✅ OK | |
| | `ddlDocCurrency`| `@DocCurrency` | `DocCurrency`| ✅ OK | |
| | `txtDocRate` | `@DocRate` | `DocRate` | ✅ OK | |
| | `ddlDeliveryAddr`| `@DeliveryAddrID`| **(缺失)** | ❌ **Error** | DB 缺欄位 |
| | `txtAddress` | `@Address` | `Address` | ✅ OK | |
| | `ddlGroupNum` | `@GroupNum` | `GroupNum` | ✅ OK | |
| | `ddlPurchaser` | **(無)** | **(缺失)** | ❌ **Lost** | 程式未讀取，DB 缺 `SlpCode` |
| | `txtRemarks` | `@Comments` | `Comments` | ✅ OK | |
| | `txtJID` | (Auto) | `jID` | ✅ OK | |
| | `txtB1DocEntry` | (Updated) | `DocEntry` | ✅ OK | |
| | `txtUPID` | `@UPID` | `U_PID` | ✅ OK | |
| | `txtTaxDate` | `@TaxDate` | `TaxDate` | ✅ OK | |
| | `txtDocDueDate` | `@DocDueDate` | `DocDueDate` | ✅ OK | |
| | `txtDocDate` | `@DocDate` | `DocDate` | ✅ OK | |
| **明細** | `ddlExpCategory`| `@ItemCode` | `ItemCode` | ✅ OK | |
| | `txtDescription`| `@Dscription` | `Dscription` | ✅ OK | |
| | `txtAcctCode` | `@AcctCode` | `AcctCode` | ✅ OK | |
| | `txtLineTotal` | `@LineTotal` | `LineTotal` | ✅ OK | |
| | `ddlVatGroup` | `@VatGroup` | `VatGroup` | ✅ OK | |
| | `txtVatSum` | `@LineVat` | `LineVat` | ✅ OK | |
| | `txtPriceAfterVat`| `@GTotal` | `GTotal` | ✅ OK | |
| | `ddlCostingCode`| `@CostingCode` | `CostingCode`| ✅ OK | |
| | `ddlCostingCode2`| `@CostingCode2`| `CostingCode2`| ✅ OK | |
| **MDR** | (Header) | **(缺失)** | `jID` | ❌ **Error** | VB 未寫入 |
| | `txtLIFNR` | `@LIFNR` | `U_LIFNR` | ✅ OK | |
| | `txtSTCEG` | `@STCEG` | `U_STCEG` | ✅ OK | |
| | `txtXBLNR` | `@XBLNR` | `U_XBLNR` | ✅ OK | |

## 建議修正方案

1.  **DB 修改**:
    *   `jOPCH` 新增 `DeliveryAddrID` (nvarchar(50))。
    *   `jOPCH` 新增 `SlpCode` (int) 或 `Purchaser` (nvarchar(50)) 以儲存採購人員。
2.  **VB 修改**:
    *   修正 `jMGUIAP` 的 INSERT 語法，補上 `jID`。
    *   修正 `SetHeaderParameters`，將 `ddlPurchaser.SelectedValue` 寫入新的 `SlpCode` 欄位。
