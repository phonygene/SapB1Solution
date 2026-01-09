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

        ' 設定用戶顯示資訊 - 顯示用戶ID（首字母大寫）
        If Not IsPostBack Then
            If Session("s_id") IsNot Nothing AndAlso Session("s_id").ToString() <> "" Then
                Dim userId As String = Session("s_id").ToString()
                ' 首字母大寫
                If userId.Length > 0 Then
                    Dim displayName As String = userId.Substring(0, 1).ToUpper() & userId.Substring(1).ToLower()
                    lblUserName.Text = displayName
                    lblUserDisplay.Text = displayName
                Else
                    lblUserName.Text = userId
                    lblUserDisplay.Text = userId
                End If
            End If

            ' 初始化下拉選單
            LoadExpDeptDropDown()
            LoadEmpSeriesDropDown()
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

#Region "下拉選單初始化"
    ''' <summary>
    ''' 載入費用部門下拉選單
    ''' </summary>
    Private Sub LoadExpDeptDropDown()
        ddlExpDept.Items.Clear()
        ddlExpDept.Items.Add(New ListItem("- 請選擇 -", ""))
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT EDeptID, EDeptName FROM jDEPT ORDER BY EDeptID"
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
        LoadUserProfile()
        ' 清除訊息
        pnlMessage.Visible = False
        ClearErrors()
        ' 顯示 Modal
        mpeUserSettings.Show()
    End Sub

    ''' <summary>
    ''' 載入使用者資料到 Modal
    ''' </summary>
    Private Sub LoadUserProfile()
        If Session("s_id") Is Nothing Then Return

        Dim userId As String = Session("s_id").ToString()
        txtUserId.Text = userId

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT name, email, expDEPT, EmpSeries FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            txtUserName.Text = If(IsDBNull(dr("name")), "", dr("name").ToString())
                            txtEmail.Text = If(IsDBNull(dr("email")), "", dr("email").ToString())

                            ' 設定費用部門
                            Dim expDept As String = If(IsDBNull(dr("expDEPT")), "", dr("expDEPT").ToString())
                            If ddlExpDept.Items.FindByValue(expDept) IsNot Nothing Then
                                ddlExpDept.SelectedValue = expDept
                            End If

                            ' 設定員工編號 (如果欄位存在)
                            Try
                                Dim empSeries As String = If(IsDBNull(dr("EmpSeries")), "", dr("EmpSeries").ToString())
                                If ddlEmpSeries.Items.FindByValue(empSeries) IsNot Nothing Then
                                    ddlEmpSeries.SelectedValue = empSeries
                                End If
                            Catch
                                ' 欄位不存在，忽略
                            End Try
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowMessage("載入資料失敗: " & ex.Message, False)
        End Try
    End Sub

    ''' <summary>
    ''' 儲存帳號設定
    ''' </summary>
    Protected Sub btnSaveSettings_Click(sender As Object, e As EventArgs)
        ' 前端驗證
        If Not ValidateForm() Then
            mpeUserSettings.Show()
            Return
        End If

        If Session("s_id") Is Nothing Then Return
        Dim userId As String = Session("s_id").ToString()

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 檢查 EmpSeries 欄位是否存在
                Dim hasEmpSeries As Boolean = False
                Try
                    Dim checkSql As String = "SELECT COL_LENGTH('User', 'EmpSeries')"
                    Using checkCmd As New SqlCommand(checkSql, conn)
                        Dim result = checkCmd.ExecuteScalar()
                        hasEmpSeries = (result IsNot Nothing AndAlso Not IsDBNull(result))
                    End Using
                Catch
                    hasEmpSeries = False
                End Try

                ' 建立更新 SQL
                Dim sql As String
                If hasEmpSeries Then
                    sql = "UPDATE [User] SET name = @Name, email = @Email, expDEPT = @ExpDept, EmpSeries = @EmpSeries WHERE id = @UserId"
                Else
                    sql = "UPDATE [User] SET name = @Name, email = @Email, expDEPT = @ExpDept WHERE id = @UserId"
                End If

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    cmd.Parameters.AddWithValue("@Name", txtUserName.Text.Trim())
                    cmd.Parameters.AddWithValue("@Email", If(String.IsNullOrEmpty(txtEmail.Text.Trim()), DBNull.Value, txtEmail.Text.Trim()))
                    cmd.Parameters.AddWithValue("@ExpDept", If(String.IsNullOrEmpty(ddlExpDept.SelectedValue), DBNull.Value, ddlExpDept.SelectedValue))

                    If hasEmpSeries Then
                        cmd.Parameters.AddWithValue("@EmpSeries", If(String.IsNullOrEmpty(ddlEmpSeries.SelectedValue), DBNull.Value, ddlEmpSeries.SelectedValue))
                    End If

                    cmd.ExecuteNonQuery()
                End Using
            End Using

            ShowMessage("儲存成功！", True)
            mpeUserSettings.Show()

        Catch ex As Exception
            ShowMessage("儲存失敗: " & ex.Message, False)
            mpeUserSettings.Show()
        End Try
    End Sub

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

        ' 驗證費用部門
        If String.IsNullOrEmpty(ddlExpDept.SelectedValue) Then
            lblExpDeptError.Text = "請選擇費用部門"
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
