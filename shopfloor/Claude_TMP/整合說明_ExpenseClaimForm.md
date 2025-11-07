# 費用申請單 - 程式碼整合說明

**建立日期**: 2025-11-05
**版本**: v1.0
**說明**: 本文件說明如何整合分階段產生的費用申請單程式碼

---

## 📁 檔案清單

### SQL 腳本（需先執行）

1. `05_CreateTable_addr.sql` - 建立收貨地址表
2. `06_CreateTable_expense_category.sql` - 建立費用類別表
3. `07_AlterTable_jOPCH_Add_ApprovalComments.sql` - 新增審核意見欄位
4. `08_AlterTable_User_Add_CanApproveExpense.sql` - 新增審核權限欄位

### 介面檔案（需整合）

1. `ExpenseClaimForm_Part1_Header.aspx` - 表頭架構（完整檔案）
2. `ExpenseClaimForm_Part1_Header.aspx.vb` - 表頭 CodeBehind（完整檔案）
3. `ExpenseClaimForm_Part2_GridView.aspx.snippet` - 費用明細 GridView（片段）
4. `ExpenseClaimForm_Part2_GridView.aspx.vb.snippet` - 費用明細 CodeBehind（片段）
5. `ExpenseClaimForm_Part3_MDR_Tab.aspx.snippet` - MDR Tab（片段）
6. `ExpenseClaimForm_Part3_MDR_Tab.aspx.vb.snippet` - MDR Tab CodeBehind（片段）

---

## 🔧 整合步驟

### 步驟 1：執行 SQL 腳本

依序執行 SQL 腳本建立資料表：

```bash
# 方式一：使用 MCP SQL 工具執行（推薦）
# Claude 可以協助執行這些腳本

# 方式二：使用 SSMS 手動執行
# 開啟 SSMS 連線到 jtdb 資料庫，依序執行 4 個 SQL 檔案
```

執行完畢後，請手動設定有審核權限的使用者：

```sql
-- 範例：設定特定使用者為審核者
UPDATE [User] SET CanApproveExpense = 1 WHERE id = 'admin'
UPDATE [User] SET CanApproveExpense = 1 WHERE id = 'finance_manager'
```

---

### 步驟 2：整合 ASPX 介面檔案

#### 2.1 複製基礎檔案

將 `ExpenseClaimForm_Part1_Header.aspx` 複製到專案目錄：

```bash
cp shopfloor/Claude_TMP/ExpenseClaimForm_Part1_Header.aspx shopfloor/ExpenseClaimForm.aspx
```

#### 2.2 插入費用明細 GridView

開啟 `shopfloor/ExpenseClaimForm.aspx`，找到：

```html
<!-- 費用明細 GridView - 下一階段實作 -->
<div class="section-title">費用明細</div>
<div>
    <p>[明細 GridView 將在第二階段實作]</p>
</div>
```

**刪除** `<div><p>[明細 GridView 將在第二階段實作]</p></div>`，
並插入 `ExpenseClaimForm_Part2_GridView.aspx.snippet` 的完整內容。

#### 2.3 插入 MDR Tab

找到：

```html
<!-- MDR Tab - 下一階段實作 -->
<div id="mdr-content" class="tab-content">
    <div class="section-title">MDR 發票明細（唯讀，自動同步）</div>
    <div>
        <p>[MDR Tab 內容將在第二階段實作]</p>
    </div>
</div>
```

**保留** `<div id="mdr-content" class="tab-content">` 和 `</div>`（最外層），
**刪除** 內部的 `<div class="section-title">...</div>` 和 `<div><p>...</p></div>`，
並插入 `ExpenseClaimForm_Part3_MDR_Tab.aspx.snippet` 的完整內容。

---

### 步驟 3：整合 CodeBehind 檔案

#### 3.1 複製基礎檔案

將 `ExpenseClaimForm_Part1_Header.aspx.vb` 複製到專案目錄：

```bash
cp shopfloor/Claude_TMP/ExpenseClaimForm_Part1_Header.aspx.vb shopfloor/ExpenseClaimForm.aspx.vb
```

#### 3.2 插入費用明細 GridView 程式碼

開啟 `shopfloor/ExpenseClaimForm.aspx.vb`，找到：

```vb
#End Region

End Class
```

在 `#End Region` **之後**、`End Class` **之前**，插入 `ExpenseClaimForm_Part2_GridView.aspx.vb.snippet` 的內容。

#### 3.3 插入 MDR Tab 程式碼

繼續在 `End Class` **之前**，插入 `ExpenseClaimForm_Part3_MDR_Tab.aspx.vb.snippet` 的內容。

#### 3.4 修改現有函式（重要！）

根據 snippet 檔案中的註解，修改以下函式：

##### A. 修改 `Page_Load` 函式

在 `SetDefaultValues()` 或 `LoadDocument()` **之後**，加入：

```vb
' 初始化 GridView（新增模式會在 SetDefaultValues 後執行）
If Not IsPostBack Then
    InitializeGridView()
    InitializeMDRGridView()
End If
```

##### B. 修改 `btnSave_Click` 和 `btnSubmit_Click` 函式

在 `ValidateBasicFields()` **之後**，加入：

```vb
' 收集 GridView 資料
CollectGridViewData()
If Not ValidateDetailLines() Then
    Return
End If

' 收集 MDR GridView 資料
CollectMDRGridViewData()
If Not ValidateMDRLines() Then
    Return
End If
```

##### C. 修改 `InsertNewDocument` 函式

在 `trans.Commit()` **之前**，加入：

```vb
' 儲存明細
SaveDetailLinesToDB(newDocEntry, conn, trans)
SaveMDRLinesToDB(newDocEntry, conn, trans)
```

##### D. 修改 `UpdateDocument` 函式

在 `trans.Commit()` **之前**，加入：

```vb
' 更新明細
SaveDetailLinesToDB(currentDocEntry, conn, trans)
SaveMDRLinesToDB(currentDocEntry, conn, trans)
```

##### E. 修改 `LoadDocument` 函式

在函式最後（載入表頭後），加入：

```vb
' 載入明細
LoadDetailLinesFromDB(docEntry, conn)
LoadMDRLinesFromDB(docEntry, conn)
```

##### F. 在表頭變更事件中同步 MDR 資訊

在以下事件的結尾加入 `SyncMDRHeaderInfo()`：

- `ddlCardCode_SelectedIndexChanged` 結尾
- `ddlDocCurrency_SelectedIndexChanged` 結尾

範例：

```vb
Protected Sub ddlCardCode_SelectedIndexChanged(sender As Object, e As EventArgs)
    ' ... 原有程式碼 ...

    ' 同步 MDR Tab 表頭資訊
    SyncMDRHeaderInfo()
End Sub
```

---

### 步驟 4：配置 Web.config

確認 `Web.config` 包含必要的資料庫連線字串：

```xml
<connectionStrings>
    <!-- 本地資料庫 (jtdb) -->
    <add name="LocalDB"
         connectionString="Data Source=YOUR_SERVER;Initial Catalog=jtdb;User ID=YOUR_USER;Password=YOUR_PASSWORD"
         providerName="System.Data.SqlClient" />

    <!-- SAP B1 資料庫 (JTTST1) -->
    <add name="SAPB1DB"
         connectionString="Data Source=YOUR_SERVER;Initial Catalog=JTTST1;User ID=YOUR_USER;Password=YOUR_PASSWORD"
         providerName="System.Data.SqlClient" />
</connectionStrings>
```

請替換：

- `YOUR_SERVER`: SQL Server 位址
- `YOUR_USER`: 資料庫使用者
- `YOUR_PASSWORD`: 資料庫密碼

---

### 步驟 5：建立上傳資料夾

在專案根目錄建立檔案上傳目錄：

```bash
mkdir -p shopfloor/Uploads/ExpenseClaims
```

設定資料夾權限（IIS 需要寫入權限）。

---

### 步驟 6：編譯與測試

#### 6.1 編譯專案

在 Visual Studio 中：

1. 清理方案（Clean Solution）
2. 重建方案（Rebuild Solution）
3. 確認無編譯錯誤

#### 6.2 測試流程

1. **測試新增單據**
   
   - 選擇供應商（必填）
   - 選擇收貨地址（必填）
   - 選擇請款日期、到期日（必填）
   - 新增費用明細至少 1 筆
   - 點選「儲存」或「送出」

2. **測試檔案上傳**
   
   - 選擇檔案並上傳
   - 確認檔案顯示在「已上傳檔案」
   - 測試下載功能

3. **測試費用明細**
   
   - 新增明細行
   - 選擇費用類別（自動填入總帳科目）
   - 輸入數量、單價（自動計算金額）
   - 驗證合計是否正確
   - 測試刪除功能

4. **測試 MDR Tab**
   
   - 切換到 MDR Tab
   - 驗證表頭資訊是否同步
   - 新增發票明細
   - 輸入發票號碼、日期、金額
   - 點選「驗證金額總和」

5. **測試審核功能（需有審核權限）**
   
   - 以有審核權限的使用者登入
   - 載入待審核單據
   - 輸入審核意見
   - 測試「放行」和「駁回」按鈕
   - 測試「發送通知」郵件

6. **測試編輯模式**
   
   - 建立單據並儲存
   - 使用 `ExpenseClaimForm.aspx?DocEntry=1` 載入
   - 驗證所有欄位是否正確載入
   - 修改資料並儲存

---

## ⚠️ 重要注意事項

### 1. 資料庫表結構

確認以下資料表已建立：

- `jOPCH` (費用申請單表頭)
- `jOPC1` (費用申請單明細)
- `jMDR1` (MDR 發票明細)
- `addr` (收貨地址)
- `expense_category` (費用類別)
- `User` 表已新增 `CanApproveExpense` 欄位

### 2. SAP B1 相關表

程式會讀取 SAP B1 資料庫的以下表：

- `OCRD` (供應商主檔)
- `OCPR` (聯絡人)
- `OOCR` (產品/部門)
- `OCRN` (幣別)
- `ORTT` (匯率)

請確認 SAP B1 資料庫連線正常。

### 3. 郵件發送

郵件發送依賴 `CommUtil.SendMail()` 函式，確認：

- SMTP 伺服器設定正確（smg.jettech.com.tw）
- User 表包含有效的 email 欄位
- 測試環境可能需要調整 SMTP 設定

### 4. 檔案命名規則

檔案上傳採用**方案 C（混合方式）**：

- 清理原始檔名（移除不安全字元）
- 加上時間戳記：`原檔名_20251105_143022.ext`
- 避免中文檔名問題

### 5. 權限控制

- 審核區塊：僅有 `CanApproveExpense = 1` 的使用者可見
- 審核意見：非審核人員為唯讀
- 建議實作 Session 檢查避免未授權存取

---

## 🐛 常見問題排查

### 問題 1：GridView 無法顯示

**症狀**：頁面載入但 GridView 是空的

**解決方案**：

1. 確認 `Page_Load` 中呼叫了 `InitializeGridView()` 和 `InitializeMDRGridView()`
2. 確認 `If Not IsPostBack Then` 區塊內有初始化程式碼
3. 檢查 `BindGridView()` 是否被正確呼叫

### 問題 2：儲存時明細沒有寫入資料庫

**症狀**：表頭儲存成功，但明細表是空的

**解決方案**：

1. 確認 `CollectGridViewData()` 在儲存前被呼叫
2. 確認 `SaveDetailLinesToDB()` 在 Transaction 內被呼叫
3. 檢查是否有 SQL Exception（查看錯誤日誌）

### 問題 3：MDR Tab 表頭資訊沒有同步

**症狀**：切換到 MDR Tab 時，表頭資訊是空的

**解決方案**：

1. 確認 `SyncMDRHeaderInfo()` 在適當的事件中被呼叫
2. 檢查控制項 ID 是否正確（`lblMDR_CardCode` 等）
3. 確認 `Page_Load` 編輯模式時有呼叫 `SyncMDRHeaderInfo()`

### 問題 4：檔案上傳失敗

**症狀**：點選上傳後出現錯誤

**解決方案**：

1. 確認 `~/Uploads/ExpenseClaims/` 資料夾存在
2. 檢查 IIS 應用程式集區的身分是否有寫入權限
3. 檢查檔案大小是否超過 10MB
4. 確認 `web.config` 的 `maxRequestLength` 設定

### 問題 5：審核按鈕不顯示

**症狀**：有審核權限的使用者看不到審核按鈕

**解決方案**：

1. 確認 User 表的 `CanApproveExpense` 欄位已更新為 1
2. 確認 `CheckApprovalPermission()` 在 `Page_Load` 中被呼叫
3. 檢查 `pnlApproval.Visible` 的設定邏輯

---

## 📝 後續開發建議

### 1. 安全性強化

- 實作輸入驗證防止 SQL Injection（目前已使用參數化查詢）
- 加強 Session 驗證機制
- 實作 HTTPS 強制導向
- 加密敏感資料

### 2. 功能擴充

- 實作列表頁面 `ExpenseClaimList.aspx` 顯示所有單據
- 加入搜尋和篩選功能
- 實作 SAP B1 DI API 整合（提交到 SAP B1）
- 加入批次核准功能
- 實作報表功能（匯出 Excel）

### 3. 使用者體驗優化

- 加入 AJAX 避免整頁 PostBack
- 實作 JavaScript 前端驗證
- 加入載入動畫（Spinner）
- 優化手機版 RWD 顯示

### 4. 維護性改善

- 將資料庫操作抽離到 DAL 層
- 建立統一的錯誤處理機制
- 實作詳細的操作日誌
- 加入單元測試

---

## 📞 聯絡資訊

如有問題或需要協助，請聯絡：

- **開發者**: Claude (Anthropic)
- **建立日期**: 2025-11-05
- **文件版本**: v1.0

---

**完成整合後，請記得：**

✅ 備份資料庫
✅ 測試所有功能
✅ 記錄測試結果
✅ 部署到測試環境
✅ 請使用者驗收

祝開發順利！🎉
