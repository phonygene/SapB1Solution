# 待完成工作規格彙整

> 匯出時間：2026-01-09
> 用途：交接給 Codex 或其他 AI 工具繼續開發

---

## 整體功能目標

**v1.2.0 - 使用者必填欄位功能**

讓使用者在進入系統時，必須完成「費用部門」和「工號」的設定。同時提供帳號設定介面供使用者修改個人資料。

---

## 已完成項目

| 項目 | 檔案 | 狀態 |
|------|------|------|
| DDL：新增 EmpSeries 欄位 | `ALTER TABLE [User] ADD EmpSeries NVARCHAR(50) NULL` | 完成 |
| UserProfileHelper.vb | `MgmSP/commcode/UserProfileHelper.vb` | 完成 |
| Home.aspx 帳號設定 UI | `MgmSP/Home.aspx` | 完成 (Modal 已建立) |

---

## 待完成項目

### 任務 1：ExpenseClaimForm.aspx 彈窗 UI 修改

**檔案**：
- `MgmSP/ExpenseClaimForm.aspx`
- `MgmSP/ExpenseClaimForm.aspx.designer.vb`

**需求**：在現有的費用部門彈窗 `pnlExpDept` (約第 1726 行) 中新增「工號」輸入欄位

**新增控制項**：

```html
<!-- 在 pnlExpDept 的 modalBody 內，費用部門欄位之前新增 -->
<div class="form-group">
    <label class="form-label" style="width:100px;">
        <span class="required">*</span>工號:
    </label>
    <div class="form-control">
        <asp:TextBox ID="txtEmpSeriesPopup" runat="server"
            placeholder="請輸入您的工號"></asp:TextBox>
    </div>
</div>
```

**designer.vb 宣告**：

```vb
Protected WithEvents txtEmpSeriesPopup As Global.System.Web.UI.WebControls.TextBox
```

**其他修改**：
- 彈窗標題從「所屬的費用部門」改為「設定必填資訊」
- 說明文字改為「請完成以下必填資訊設定：」

---

### 任務 2：ExpenseClaimForm.aspx.vb 邏輯修改

**檔案**：`MgmSP/ExpenseClaimForm.aspx.vb`

#### 2.1 修改 CheckUserExpDept() 方法

找到現有的 `CheckUserExpDept()` 方法，改為使用 `UserProfileHelper.CheckRequiredFields()`

**修改後邏輯**：

```vb
Private Sub CheckUserRequiredFields()
    ' 取得當前使用者 ID (通常從 Session)
    Dim userId As String = Session("userid")?.ToString()
    If String.IsNullOrEmpty(userId) Then Return

    ' 檢查必填欄位
    Dim result = UserProfileHelper.CheckRequiredFields(userId)

    If Not result.IsComplete Then
        ' 載入費用部門下拉選單
        LoadExpDeptDropDown()

        ' 預填已有的值
        Dim profile = UserProfileHelper.GetUserProfile(userId)
        If profile IsNot Nothing Then
            If Not String.IsNullOrEmpty(profile.ExpDept) Then
                ddlExpDeptSelect.SelectedValue = profile.ExpDept
            End If
            If Not String.IsNullOrEmpty(profile.EmpSeries) Then
                txtEmpSeriesPopup.Text = profile.EmpSeries
            End If
        End If

        ' 顯示彈窗
        mpeExpDept.Show()
    Else
        ' 設定表頭費用部門
        SetHeaderExpDept()
    End If
End Sub
```

#### 2.2 修改 btnExpDeptConfirm_Click

找到現有的 `btnExpDeptConfirm_Click` 方法，改為同時儲存兩個欄位

**修改後邏輯**：

```vb
Protected Sub btnExpDeptConfirm_Click(sender As Object, e As EventArgs)
    Dim userId As String = Session("userid")?.ToString()
    If String.IsNullOrEmpty(userId) Then Return

    ' 驗證費用部門
    If String.IsNullOrEmpty(ddlExpDeptSelect.SelectedValue) Then
        ' 顯示錯誤訊息 (使用現有的 ShowAlert 或 ScriptManager)
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alert", "alert('請選擇費用部門');", True)
        mpeExpDept.Show()
        Return
    End If

    ' 驗證工號
    If String.IsNullOrEmpty(txtEmpSeriesPopup.Text.Trim()) Then
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alert", "alert('請輸入工號');", True)
        mpeExpDept.Show()
        Return
    End If

    ' 使用 UserProfileHelper 更新
    Dim success As Boolean = UserProfileHelper.UpdateExpDeptAndEmpSeries(
        userId,
        ddlExpDeptSelect.SelectedValue,
        txtEmpSeriesPopup.Text.Trim()
    )

    If success Then
        ' 更新表頭費用部門顯示
        SetHeaderExpDept()
    Else
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alert", "alert('儲存失敗，請重試');", True)
        mpeExpDept.Show()
    End If
End Sub
```

---

### 任務 3：Home.aspx.vb 邏輯實作

**檔案**：`MgmSP/Home.aspx.vb`

**需要實作的方法**：

#### 3.1 Page_Load

```vb
Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    If Not IsPostBack Then
        LoadUserDisplay()
        LoadDropDownLists()
    End If
End Sub

Private Sub LoadUserDisplay()
    Dim userId As String = Session("userid")?.ToString()
    If String.IsNullOrEmpty(userId) Then
        Response.Redirect("~/usermgm/login.aspx")
        Return
    End If

    Dim profile = UserProfileHelper.GetUserProfile(userId)
    If profile IsNot Nothing Then
        lblUserDisplay.Text = profile.UserName
        lblUserName.Text = profile.UserName
    End If
End Sub

Private Sub LoadDropDownLists()
    ' 載入費用部門
    LoadExpDeptDropDown()
    ' 載入員工編號 (如果需要從 SAP 取得)
    LoadEmpSeriesDropDown()
End Sub
```

#### 3.2 費用部門下拉選單載入

```vb
Private Sub LoadExpDeptDropDown()
    Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Using conn As New SqlConnection(connStr)
        conn.Open()
        Dim sql As String = "SELECT EDeptID, EDeptName FROM jDEPT ORDER BY EDeptName"
        Using cmd As New SqlCommand(sql, conn)
            Using dr As SqlDataReader = cmd.ExecuteReader()
                ddlExpDept.Items.Clear()
                ddlExpDept.Items.Add(New ListItem("-- 請選擇 --", ""))
                While dr.Read()
                    ddlExpDept.Items.Add(New ListItem(
                        dr("EDeptName").ToString(),
                        dr("EDeptID").ToString()
                    ))
                End While
            End Using
        End Using
    End Using
End Sub
```

#### 3.3 點擊使用者名稱開啟設定

```vb
Protected Sub lnkUserSettings_Click(sender As Object, e As EventArgs)
    LoadUserSettingsForm()
    mpeUserSettings.Show()
End Sub

Private Sub LoadUserSettingsForm()
    Dim userId As String = Session("userid")?.ToString()
    If String.IsNullOrEmpty(userId) Then Return

    Dim profile = UserProfileHelper.GetUserProfile(userId)
    If profile IsNot Nothing Then
        txtUserId.Text = profile.UserId
        txtUserName.Text = profile.UserName
        txtEmail.Text = profile.Email

        ' 設定費用部門下拉選單
        If Not String.IsNullOrEmpty(profile.ExpDept) Then
            If ddlExpDept.Items.FindByValue(profile.ExpDept) IsNot Nothing Then
                ddlExpDept.SelectedValue = profile.ExpDept
            End If
        End If

        ' 設定員工編號 (如果是下拉選單)
        If Not String.IsNullOrEmpty(profile.EmpSeries) Then
            If ddlEmpSeries.Items.FindByValue(profile.EmpSeries) IsNot Nothing Then
                ddlEmpSeries.SelectedValue = profile.EmpSeries
            End If
        End If
    End If

    ' 清除訊息
    pnlMessage.Visible = False
End Sub
```

#### 3.4 儲存設定

```vb
Protected Sub btnSaveSettings_Click(sender As Object, e As EventArgs)
    Dim userId As String = Session("userid")?.ToString()
    If String.IsNullOrEmpty(userId) Then Return

    ' 驗證必填欄位
    If String.IsNullOrEmpty(txtUserName.Text.Trim()) Then
        ShowMessage("請輸入姓名", False)
        mpeUserSettings.Show()
        Return
    End If

    If String.IsNullOrEmpty(ddlExpDept.SelectedValue) Then
        ShowMessage("請選擇費用部門", False)
        mpeUserSettings.Show()
        Return
    End If

    ' 建立更新模型
    Dim model As New UserProfileModel()
    model.UserName = txtUserName.Text.Trim()
    model.Email = txtEmail.Text.Trim()
    model.ExpDept = ddlExpDept.SelectedValue
    model.EmpSeries = ddlEmpSeries.SelectedValue

    ' 注意：密碼欄位在 Home.aspx 目前沒有，若需要可另外加

    ' 執行更新 (需要在 UserProfileHelper 新增 UpdateUserProfileFull 方法)
    ' 或使用現有方法分開更新
    Dim success As Boolean = UpdateUserSettings(userId, model)

    If success Then
        ShowMessage("設定已儲存", True)
        ' 更新顯示
        lblUserDisplay.Text = model.UserName
        lblUserName.Text = model.UserName
    Else
        ShowMessage("儲存失敗，請重試", False)
    End If

    mpeUserSettings.Show()
End Sub

Private Function UpdateUserSettings(userId As String, model As UserProfileModel) As Boolean
    ' 這裡可以直接寫 SQL 或擴充 UserProfileHelper
    Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Try
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Dim sql As String = "UPDATE [User] SET name = @Name, email = @Email, expDEPT = @ExpDept, EmpSeries = @EmpSeries WHERE id = @UserId"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Name", model.UserName)
                cmd.Parameters.AddWithValue("@Email", If(String.IsNullOrEmpty(model.Email), DBNull.Value, model.Email))
                cmd.Parameters.AddWithValue("@ExpDept", If(String.IsNullOrEmpty(model.ExpDept), DBNull.Value, model.ExpDept))
                cmd.Parameters.AddWithValue("@EmpSeries", If(String.IsNullOrEmpty(model.EmpSeries), DBNull.Value, model.EmpSeries))
                cmd.Parameters.AddWithValue("@UserId", userId)
                Return cmd.ExecuteNonQuery() > 0
            End Using
        End Using
    Catch ex As Exception
        Return False
    End Try
End Function

Private Sub ShowMessage(msg As String, isSuccess As Boolean)
    lblMessage.Text = msg
    pnlMessage.CssClass = If(isSuccess, "save-message success", "save-message error")
    pnlMessage.Visible = True
End Sub
```

#### 3.5 關閉/取消按鈕

```vb
Protected Sub btnCloseSettings_Click(sender As Object, e As EventArgs)
    mpeUserSettings.Hide()
End Sub

Protected Sub btnCancelSettings_Click(sender As Object, e As EventArgs)
    mpeUserSettings.Hide()
End Sub
```

---

## 資料庫相關

### User 表結構（相關欄位）

| 欄位 | 類型 | 說明 |
|------|------|------|
| id | NVARCHAR | 帳號 (PK) |
| name | NVARCHAR | 姓名 |
| pwd | NVARCHAR | 密碼 |
| email | NVARCHAR | Email |
| expDEPT | NVARCHAR | 費用部門代碼 |
| EmpSeries | NVARCHAR(50) | 工號 (**已新增**) |

### 費用部門來源

```sql
SELECT EDeptID, EDeptName FROM jDEPT ORDER BY EDeptName
```

### 員工編號來源（如需從 SAP 取得）

```sql
-- 連線字串：SapSQLConnection
-- 可能的查詢（需確認實際表結構）
SELECT empID, firstName + ' ' + lastName AS empName FROM OHEM ORDER BY firstName
```

---

## UserProfileHelper.vb 現有方法

位置：`MgmSP/commcode/UserProfileHelper.vb`

| 方法 | 說明 |
|------|------|
| `GetUserProfile(userId)` | 取得使用者資料，回傳 UserProfileModel |
| `UpdateUserProfile(userId, model)` | 更新使用者資料（密碼、工號、費用部門、Email） |
| `CheckRequiredFields(userId)` | 檢查必填欄位，回傳 RequiredFieldsResult |
| `UpdateExpDeptAndEmpSeries(userId, expDept, empSeries)` | 更新費用部門和工號 |

---

## 驗收標準

1. [ ] ExpenseClaimForm：進入時若 expDEPT 或 EmpSeries 任一為空，彈出設定視窗
2. [ ] ExpenseClaimForm：彈窗可輸入工號和選擇費用部門
3. [ ] ExpenseClaimForm：確認後兩欄位正確存入資料庫
4. [ ] Home.aspx：右上角顯示使用者名稱，可點擊
5. [ ] Home.aspx：帳號設定 Modal 可正常開啟/關閉
6. [ ] Home.aspx：帳號設定可正常儲存
7. [ ] 所有新控制項已在 `.designer.vb` 中宣告

---

## 關鍵檔案清單

| 檔案 | 待辦 |
|------|------|
| `MgmSP/ExpenseClaimForm.aspx` | 新增 `txtEmpSeriesPopup` 控制項、修改彈窗標題 |
| `MgmSP/ExpenseClaimForm.aspx.designer.vb` | 宣告 `txtEmpSeriesPopup` |
| `MgmSP/ExpenseClaimForm.aspx.vb` | 修改 `CheckUserExpDept` 和 `btnExpDeptConfirm_Click` |
| `MgmSP/Home.aspx.vb` | 實作 Page_Load、lnkUserSettings_Click、btnSaveSettings_Click 等 |
| `MgmSP/Home.aspx.designer.vb` | 確認所有控制項已宣告 |

---

## 技術注意事項

1. **Namespace**：aspx 的 Inherits 必須加前綴 `MgmSP.`
2. **檔案編碼**：UTF-8 with BOM
3. **連線字串**：
   - 本地資料庫：`jtdbConnectionString`
   - SAP 資料庫：`SapSQLConnection`
4. **控制項宣告**：在 `.aspx` 新增控制項時，必須同時更新 `.aspx.designer.vb`
5. **AjaxToolkit**：Modal 使用 `ModalPopupExtender`，需要一個隱藏的 `TargetControlID`

---

## 參考檔案

- 彈窗樣式參考：`ExpenseClaimForm.aspx` 第 1720-1744 行
- 使用者資訊樣式參考：`ExpenseClaimForm.aspx` 第 905-909 行
- UserProfileHelper 完整程式碼：`MgmSP/commcode/UserProfileHelper.vb`
