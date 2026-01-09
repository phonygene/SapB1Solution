Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Public Class Home
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private ReadOnly sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim timeout As Integer
        Dim act As String
        Dim masterPage As MySite1 = TryCast(Me.Master, MySite1)
        If masterPage IsNot Nothing Then
            masterPage.SetThemeClass("theme-dark")
        End If

        If Not IsPostBack Then
            LoadUserDisplay()
            LoadDropDownLists()
        End If

        timeout = Request.QueryString("timeout")
        act = Request.QueryString("act")
        If (timeout = 1) Then
            CommUtil.ShowMsg(Me, "閒置時間太久,請重新登錄")
        End If
        If (act = "signfinish") Then
            CommUtil.ShowMsg(Me, "簽核已全部完成")
        End If
        If (act = "setsap") Then
            CommUtil.ShowMsg(Me, "設定Sap帳號密碼成功")
        End If
        If (act = "modifypwd") Then
            CommUtil.ShowMsg(Me, "修改密碼成功")
        End If
    End Sub

    Private Function GetCurrentUserId() As String
        Dim userId As String = TryCast(Session("userid"), String)
        If String.IsNullOrEmpty(userId) AndAlso Session("s_id") IsNot Nothing Then
            userId = Session("s_id").ToString()
        End If
        Return userId
    End Function

    Private Sub LoadUserDisplay()
        Dim userId As String = GetCurrentUserId()
        If String.IsNullOrEmpty(userId) Then
            Response.Redirect("~/usermgm/login.aspx")
            Return
        End If

        Dim profile = UserProfileHelper.GetUserProfile(userId)
        If profile IsNot Nothing AndAlso Not String.IsNullOrEmpty(profile.UserName) Then
            lblUserDisplay.Text = profile.UserName
            lblUserName.Text = profile.UserName
        Else
            lblUserDisplay.Text = userId
            lblUserName.Text = userId
        End If
    End Sub

    Private Sub LoadDropDownLists()
        LoadExpDeptDropDown()
        LoadEmpSeriesDropDown()
    End Sub

#Region "下拉選單初始化"
    ''' <summary>
    ''' 載入＊費用部門下拉選單
    ''' </summary>
    Private Sub LoadExpDeptDropDown()
        ddlExpDept.Items.Clear()
        ddlExpDept.Items.Add(New ListItem("-- 請選擇 --", ""))
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT EDeptID, EDeptName FROM jDEPT ORDER BY EDeptName"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddlExpDept.Items.Add(New ListItem(dr("EDeptName").ToString(), dr("EDeptID").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 靜默處理
        End Try
    End Sub

    ''' <summary>
    ''' 載入員工編號下拉選單 (從 SAP OHEM)
    ''' </summary>
    Private Sub LoadEmpSeriesDropDown()
        ddlEmpSeries.Items.Clear()
        ddlEmpSeries.Items.Add(New ListItem("- 請選擇 -", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT Code, lastName + firstName AS EmpName FROM OHEM WHERE Active = 'Y' ORDER BY Code"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            Dim code As String = dr("Code").ToString()
                            Dim empName As String = If(IsDBNull(dr("EmpName")), "", dr("EmpName").ToString())
                            ddlEmpSeries.Items.Add(New ListItem(code & " - " & empName, code))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 靜默處理
        End Try
    End Sub
#End Region

#Region "帳號設定 Modal"
    ''' <summary>
    ''' 點擊使用者名稱 - 開啟帳號設定 Modal
    ''' </summary>
    Protected Sub lnkUserSettings_Click(sender As Object, e As EventArgs)
        ' 載入使用者資料
        LoadUserSettingsForm()
        ' 清除訊息
        pnlMessage.Visible = False
        ClearErrors()
        ' 顯示 Modal
        mpeUserSettings.Show()
    End Sub

    ''' <summary>
    ''' 載入使用者資料到 Modal
    ''' </summary>
    Private Sub LoadUserSettingsForm()
        Dim userId As String = GetCurrentUserId()
        If String.IsNullOrEmpty(userId) Then Return

        Dim profile = UserProfileHelper.GetUserProfile(userId)
        If profile IsNot Nothing Then
            txtUserId.Text = profile.UserId
            txtUserName.Text = profile.UserName
            txtEmail.Text = profile.Email

            If Not String.IsNullOrEmpty(profile.ExpDept) Then
                If ddlExpDept.Items.FindByValue(profile.ExpDept) IsNot Nothing Then
                    ddlExpDept.SelectedValue = profile.ExpDept
                End If
            End If

            If Not String.IsNullOrEmpty(profile.EmpSeries) Then
                If ddlEmpSeries.Items.FindByValue(profile.EmpSeries) IsNot Nothing Then
                    ddlEmpSeries.SelectedValue = profile.EmpSeries
                End If
            End If
        End If

        pnlMessage.Visible = False
    End Sub

    ''' <summary>
    ''' 儲存帳號設定
    ''' </summary>
    Protected Sub btnSaveSettings_Click(sender As Object, e As EventArgs)
        Dim userId As String = GetCurrentUserId()
        If String.IsNullOrEmpty(userId) Then Return

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

        If Not String.IsNullOrEmpty(txtEmail.Text.Trim()) AndAlso Not IsValidEmail(txtEmail.Text.Trim()) Then
            ShowMessage("Email 格式不正確", False)
            mpeUserSettings.Show()
            Return
        End If

        Dim model As New UserProfileModel()
        model.UserName = txtUserName.Text.Trim()
        model.Email = txtEmail.Text.Trim()
        model.ExpDept = ddlExpDept.SelectedValue
        model.EmpSeries = ddlEmpSeries.SelectedValue

        Dim success As Boolean = UpdateUserSettings(userId, model)

        If success Then
            ShowMessage("設定已儲存", True)
            lblUserDisplay.Text = model.UserName
            lblUserName.Text = model.UserName
        Else
            ShowMessage("儲存失敗，請重試", False)
        End If

        mpeUserSettings.Show()
    End Sub

    Private Function UpdateUserSettings(userId As String, model As UserProfileModel) As Boolean
        If String.IsNullOrEmpty(userId) OrElse model Is Nothing Then Return False

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

    ''' <summary>
    ''' 取消 - 關閉 Modal
    ''' </summary>
    Protected Sub btnCancelSettings_Click(sender As Object, e As EventArgs)
        mpeUserSettings.Hide()
    End Sub

    ''' <summary>
    ''' 關閉按鈕
    ''' </summary>
    Protected Sub btnCloseSettings_Click(sender As Object, e As EventArgs)
        mpeUserSettings.Hide()
    End Sub

    ''' <summary>
    ''' 驗證表單
    ''' </summary>
    Private Function ValidateForm() As Boolean
        ClearErrors()
        Dim isValid As Boolean = True

        ' 驗證姓名
        If String.IsNullOrEmpty(txtUserName.Text.Trim()) Then
            lblNameError.Text = "姓名為必填欄位"
            lblNameError.Visible = True
            isValid = False
        End If

        ' 驗證 Email 格式 (如果有填)
        If Not String.IsNullOrEmpty(txtEmail.Text.Trim()) Then
            If Not IsValidEmail(txtEmail.Text.Trim()) Then
                lblEmailError.Text = "Email 格式不正確"
                lblEmailError.Visible = True
                isValid = False
            End If
        End If

        ' 驗證＊費用部門
        If String.IsNullOrEmpty(ddlExpDept.SelectedValue) Then
            lblExpDeptError.Text = "請選擇＊費用部門"
            lblExpDeptError.Visible = True
            isValid = False
        End If

        Return isValid
    End Function

    ''' <summary>
    ''' 清除錯誤訊息
    ''' </summary>
    Private Sub ClearErrors()
        lblNameError.Visible = False
        lblEmailError.Visible = False
        lblExpDeptError.Visible = False
    End Sub

    ''' <summary>
    ''' 顯示訊息
    ''' </summary>
    Private Sub ShowMessage(message As String, isSuccess As Boolean)
        lblMessage.Text = message
        pnlMessage.CssClass = If(isSuccess, "save-message success", "save-message error")
        pnlMessage.Visible = True
    End Sub

    ''' <summary>
    ''' 驗證 Email 格式
    ''' </summary>
    Private Function IsValidEmail(email As String) As Boolean
        Try
            Dim addr = New System.Net.Mail.MailAddress(email)
            Return addr.Address = email
        Catch
            Return False
        End Try
    End Function
#End Region

End Class
