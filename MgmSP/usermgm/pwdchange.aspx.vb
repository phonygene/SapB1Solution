Imports System.Data
Imports System.Data.SqlClient
Partial Public Class pwdchange
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public dr As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
    End Sub

    Protected Sub modifybtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles modifybtn.Click
        If (oldpwdtxt.Text <> "" And newpwdtxt.Text <> "" And cnewpwdtxt.Text <> "") Then
            Dim SqlCmd As String
            Dim sqlresult As Boolean
            'InitLocalSQLConnection()
            SqlCmd = "Select pwd From dbo.[User] where dbo.[User].id='" & Session("s_id") & "'"

            'myCommand = New SqlCommand(SqlCmd, conn)
            'dr = myCommand.ExecuteReader()
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                If (oldpwdtxt.Text <> dr("pwd")) Then
                    errmsg.Text = "舊密碼key錯"
                    dr.Close()
                    conn.Close()
                Else
                    dr.Close()
                    conn.Close()
                    SqlCmd = "Update dbo.[User]  set dbo.[User].pwd= '" & newpwdtxt.Text & "'where dbo.[User].id='" & Session("s_id") & "'"
                    'myCommand = New SqlCommand(SqlCmd, conn)
                    'count = myCommand.ExecuteNonQuery()
                    sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                    If (sqlresult) Then
                        Response.Redirect("~\index.aspx?smid=index&smode=0&act=modifypwd")
                    Else
                        CommUtil.ShowMsg(Me, "更新失敗")
                    End If
                    conn.Close()
                End If
            Else
                errmsg.Text = "沒找到相應id"
                conn.Close()
            End If
        Else
            errmsg.Text = "欄位不能空白"
        End If
    End Sub
End Class