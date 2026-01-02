# Backend Agent 規範

> 負責：功能開發、API 邏輯、資料庫操作、業務邏輯
> 分支前綴：`feature/`

---

## 職責範圍

- 新功能的後端邏輯實作
- 資料庫查詢與資料操作
- API 端點設計與實作
- SAP B1 整合 (DI API / Service Layer)
- 資料驗證與業務規則

---

## 程式碼規範

### ASP.NET Web Forms

1. **控制項宣告**：新增控制項時必須同步更新 `.aspx.designer.vb`

   ```vb
   ' designer.vb 範例
   Protected WithEvents btnSave As Global.System.Web.UI.WebControls.Button
   Protected WithEvents gvData As Global.System.Web.UI.WebControls.GridView
   ```

2. **事件處理**：使用 `Handles` 關鍵字明確綁定

   ```vb
   Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
   ```

3. **ViewState 管理**：大型資料使用 Session，避免 ViewState 膨脹

### 資料庫操作

1. **使用參數化查詢**，禁止字串拼接 SQL

   ```vb
   ' 正確
   Dim sql As String = "SELECT * FROM OITM WHERE ItemCode = @code"
   cmd.Parameters.AddWithValue("@code", itemCode)

   ' 錯誤
   Dim sql As String = "SELECT * FROM OITM WHERE ItemCode = '" & itemCode & "'"
   ```

2. **連線管理**：使用 `Using` 確保資源釋放

   ```vb
   Using conn As New SqlConnection(connStr)
       conn.Open()
       ' ...
   End Using
   ```

3. **交易處理**：多表操作使用 Transaction

### SAP B1 整合

1. **DI API 物件釋放**

   ```vb
   ' 必須釋放 COM 物件
   If oDoc IsNot Nothing Then
       System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc)
       oDoc = Nothing
   End If
   ```

2. **錯誤處理**：捕獲 SAP 錯誤碼

   ```vb
   If oCompany.GetLastErrorCode() <> 0 Then
       Dim errMsg As String = oCompany.GetLastErrorDescription()
       ' 記錄錯誤
   End If
   ```

---

## 計算邏輯原則

### 稅額計算（財務系統核心）

1. **只在值變更事件計算**，不在 Save 時重算

   ```vb
   ' 正確：在輸入變更時計算
   Protected Sub txtAmount_TextChanged(...)
       CalculateTax()
   End Sub

   ' 錯誤：在儲存時重算
   Protected Sub btnSave_Click(...)
       CalculateTax()  ' 這會覆蓋用戶修改的值！
       Save()
   End Sub
   ```

2. **用戶手動修改的值優先保留**

   ```vb
   ' 計算前檢查是否已被手動修改
   If Not IsUserModified("TaxAmount") Then
       txtTaxAmount.Text = CalculatedTax.ToString()
   End If
   ```

3. **Sync 函數只讀取，不計算**

   ```vb
   ' 正確
   Private Sub SyncToModel()
       model.Amount = Decimal.Parse(txtAmount.Text)
       model.TaxAmount = Decimal.Parse(txtTaxAmount.Text)  ' 直接讀取
   End Sub

   ' 錯誤
   Private Sub SyncToModel()
       model.Amount = Decimal.Parse(txtAmount.Text)
       model.TaxAmount = model.Amount * 0.05  ' 重新計算！
   End Sub
   ```

---

## 常見問題模式

### [P001] COM 物件未釋放
- **症狀**：記憶體持續增長、SAP 連線數過多
- **解法**：使用 `Marshal.ReleaseComObject` 並設為 Nothing

### [P002] PostBack 後資料遺失
- **症狀**：頁面刷新後用戶輸入消失
- **解法**：檢查 `IsPostBack` 邏輯，正確使用 ViewState/Session

### [P003] UpdatePanel 事件不觸發
- **症狀**：按鈕點擊無反應
- **解法**：在 Triggers 中明確註冊 AsyncPostBackTrigger

---

## 協作注意事項

### 提供給 UI-UX Agent 的資訊

當完成後端功能時，需在 task-status 中記錄：
- 可用的資料欄位與型別
- API 端點與參數
- 資料驗證規則
- 錯誤代碼與訊息

### 提供給 QA Agent 的資訊

- 測試所需的前置資料
- 邊界條件與例外情況
- 效能考量點

---

## 檢查清單

執行任務前：
- [ ] 讀取 `.claude/task-status.json` 確認無衝突
- [ ] 確認影響的檔案列表

執行任務後：
- [ ] 更新 designer.vb（如有新增控制項）
- [ ] 確認 COM 物件有釋放
- [ ] 確認 SQL 使用參數化查詢
- [ ] 確認沒有在 Save 時重新計算用戶輸入
- [ ] 更新 task-status.json
- [ ] 記錄到 work-logs
