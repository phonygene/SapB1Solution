# Backend 任務規格：使用者必填資訊 - 後端模組

> 任務 ID: 2026-01-08-user-required-fields-backend
> 指派: Backend Agent
> 優先級: High
> 建立時間: 2026-01-08

---

## 目標

建立使用者資料存取模組，支援工號欄位，並修改費用部門檢查邏輯。

---

## 任務清單

### B1: DDL - 新增 EmpSeries 欄位

```sql
ALTER TABLE [User] ADD EmpSeries NVARCHAR(50) NULL
```

### B2: 建立/更新 UserProfileHelper.vb

位置：`MgmSP/Modules/UserProfileHelper.vb`

封裝以下功能：

```vb
Public Class UserProfileHelper
    ' 取得使用者資料
    Public Shared Function GetUserProfile(userId As String) As UserProfileModel

    ' 更新使用者資料（密碼、工號、費用部門、Email）
    Public Shared Function UpdateUserProfile(userId As String, model As UserProfileModel) As Boolean

    ' 檢查必填欄位是否完整（ExpDept + EmpSeries）
    Public Shared Function CheckRequiredFields(userId As String) As RequiredFieldsResult
End Class

Public Class UserProfileModel
    Public Property UserId As String
    Public Property UserName As String
    Public Property Password As String
    Public Property EmpSeries As String    ' 工號
    Public Property ExpDept As String      ' 費用部門代碼
    Public Property ExpDeptName As String  ' 費用部門名稱（唯讀）
    Public Property Email As String
End Class

Public Class RequiredFieldsResult
    Public Property IsComplete As Boolean
    Public Property MissingExpDept As Boolean
    Public Property MissingEmpSeries As Boolean
End Class
```

### B3: 修改 ExpenseClaimForm.aspx.vb - CheckUserExpDept

修改 `CheckUserExpDept()` 方法：

**現行邏輯**：只檢查 expDEPT 是否為空
**新邏輯**：檢查 expDEPT 和 EmpSeries 兩者是否都有值

```vb
Private Sub CheckUserRequiredFields()
    Dim result = UserProfileHelper.CheckRequiredFields(currentUserId)

    If Not result.IsComplete Then
        ' 載入費用部門下拉選單
        LoadExpDeptDropDown()
        ' 設定現有值（如果有）
        ' 顯示彈窗
        mpeExpDept.Show()
    Else
        ' 設定表頭費用部門
        SetHeaderExpDept()
    End If
End Sub
```

### B4: 修改 btnExpDeptConfirm_Click

**現行邏輯**：只存 expDEPT
**新邏輯**：同時存 expDEPT 和 EmpSeries

```vb
Protected Sub btnExpDeptConfirm_Click(sender As Object, e As EventArgs)
    ' 驗證兩個欄位都必填
    If String.IsNullOrEmpty(ddlExpDeptSelect.SelectedValue) Then
        ShowAlert("請選擇費用部門")
        mpeExpDept.Show()
        Return
    End If

    If String.IsNullOrEmpty(txtEmpSeries.Text.Trim()) Then
        ShowAlert("請輸入工號")
        mpeExpDept.Show()
        Return
    End If

    ' 更新資料庫
    Dim sql As String = "UPDATE [User] SET expDEPT = @ExpDept, EmpSeries = @EmpSeries WHERE id = @UserId"
    ' ... 執行更新
End Sub
```

---

## 注意事項

1. **密碼處理**：前端顯示用 `type="password"`（黑色圓點），後端存儲維持現有方式
2. **向後相容**：EmpSeries 允許 NULL，舊使用者進入時會觸發彈窗補填
3. **重用性**：UserProfileHelper 需供 Home.aspx 呼叫，注意 Public Shared

---

## 影響檔案

- `[User]` 資料表（DDL）
- `MgmSP/Modules/UserProfileHelper.vb`（新建）
- `MgmSP/ExpenseClaimForm.aspx.vb`（修改）

---

## 驗收標準

1. [ ] EmpSeries 欄位已建立
2. [ ] UserProfileHelper 可正常讀寫使用者資料
3. [ ] 進入費用申請單時，若 ExpDept 或 EmpSeries 任一為空則彈窗
4. [ ] 彈窗確認後兩欄位都正確存入資料庫

---

## 完成後

1. 更新 `.agent-workspace/handoff/2026-01-08-user-required-fields-backend/output.md`
2. 通知 Manager 進行審查
