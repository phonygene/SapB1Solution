Imports System.Data.SqlClient
Imports System.Web.Configuration

''' <summary>
''' 系統管理員登入頁面
''' 維護模式中，只有 Admin=1 的帳號可以登入
''' </summary>
Partial Public Class AdminLogin
    Inherits System.Web.UI.Page

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' 不做任何維護檢查，讓管理員可以登入
    End Sub

    Protected Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Dim userId As String = txtUserId.Text.Trim()
        Dim password As String = txtPassword.Text

        ' 基本驗證
        If String.IsNullOrEmpty(userId) Then
            lblError.Text = "請輸入帳號"
            Return
        End If

        If String.IsNullOrEmpty(password) Then
            lblError.Text = "請輸入密碼"
            Return
        End If

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT id, name, pwd, admin, sapid, sappwd, grp, branch FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            ' 驗證密碼
                            If dr("pwd").ToString() <> password Then
                                lblError.Text = "密碼錯誤"
                                Return
                            End If

                            ' 檢查是否為管理員
                            Dim isAdmin As Boolean = (Convert.ToInt32(If(IsDBNull(dr("admin")), 0, dr("admin"))) = 1)
                            If Not isAdmin Then
                                lblError.Text = "網站維護中只有系統管理員可以登入"
                                Return
                            End If

                            ' 登入成功，設定 Session
                            Session("s_id") = dr("id")
                            Session("s_name") = dr("name")
                            Session("sapid") = If(IsDBNull(dr("sapid")), "", dr("sapid"))
                            Session("sappwd") = If(IsDBNull(dr("sappwd")), "", dr("sappwd"))
                            Session("grp") = If(IsDBNull(dr("grp")), "", dr("grp"))
                            Session("branch") = If(IsDBNull(dr("branch")), "", dr("branch"))
                            Session("usingserver") = "127.0.0.1"
                            Session("usingdb") = "JTSTD"
                            Session("usingwhsfull") = "JET"
                            Session("usingwhs") = "JET"

                            ' 導向首頁
                            Response.Redirect("~/Home.aspx?smid=index")
                        Else
                            lblError.Text = "無此帳號"
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            lblError.Text = "登入失敗：" & ex.Message
        End Try
    End Sub

End Class
