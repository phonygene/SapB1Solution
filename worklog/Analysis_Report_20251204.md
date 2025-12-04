# 費用申請單 (Expense Claim) 分析報告

**日期**: 2025-12-04
**狀態**: 分析完成

## 1. 資料庫結構分析 (Schema Analysis)

經檢查，目前資料庫結構如下：

### jOPCH (表頭)
*   **PK**: `jID` (int) - 應為 Identity (自動遞增)。
*   **Key**: `DocEntry` (int) - 目前程式碼將其設為與 `jID` 相同。
*   **狀態**: 欄位齊全，支援 `jAttach` (附件) 改為獨立 Table 關聯。

### jPCH1 (費用明細)
*   **PK**: `jID` + `LineNum`。
*   **關係**: `jID` 為 Foreign Key，對應 `jOPCH.jID`。
*   **寫入邏輯**: 程式碼正確將 `jOPCH` 產生的 `jID` 寫入此欄位。
*   **注意**: 規格文件中提到的 "jID 應為自動遞增" 可能是指表頭的 `jID`，若指明細的 `jID` 則與複合主鍵結構衝突。依目前 Schema，維持 FK 寫入是正確的。

### jMGUIAP (MDR 表頭)
*   **PK**: `ID` (int) - Identity。
*   **問題**: **資料表包含 `jID` (int, NOT NULL) 欄位**，但目前的 `INSERT` 語句**未包含此欄位**。這會導致寫入失敗。
*   **修正**: 寫入時必須將 `jOPCH.jID` 填入 `jMGUIAP.jID`，建立兩者關聯。

### jMGUIAPDetail (MDR 明細)
*   **PK**: `jID` + `LineNum`。
*   **關係**: 這裡的 `jID` 對應 `jMGUIAP.ID` (MDR Header ID)。
*   **邏輯**: 程式碼使用 `SCOPE_IDENTITY()` 取得 MDR ID 並寫入，邏輯正確。

### jAttach (附件)
*   **PK**: `ID` (int)。
*   **欄位**: `jID`, `DocEntry`, `LineNum`, `FilePath`, `FileName` 等。
*   **邏輯**: 支援多檔上傳與獨立儲存，符合需求。

---

## 2. 修正計畫 (Fix Plan)

### A. 修正寫入程式碼 (ExpenseClaimForm.aspx.vb)
1.  **MDR 寫入修正**: 修改 `jMGUIAP` 的 INSERT 語句，加入 `jID` 欄位 (值為 `jOPCH` 的 ID)。
2.  **交易完整性**: 確保 Header, Lines, MDR, Attachments 在同一 Transaction 中寫入。

### B. 實作查詢功能 (ExpenseClaimList.aspx)
*   **檔案**: 新增 `ExpenseClaimList.aspx`。
*   **篩選條件**: jID, AP單號, 簽核PID, 文件狀態, 日期(過帳/到期/文件), 放行狀態, 供應商(代碼/名稱), 備註。
*   **列表欄位**: 顯示上述主要資訊。
*   **跳轉**: 點擊 `jID` 跳轉至 `ExpenseClaimForm.aspx?DocEntry={DocEntry}`。

### C. 表單讀取模式調整
*   目前的 `ExpenseClaimForm.aspx` 已具備 `LoadDocument` 功能。
*   確認 `Update` / `Delete` 按鈕在讀取模式下的顯示邏輯 (目前已有 `currentDocEntry > 0` 的判斷)。
