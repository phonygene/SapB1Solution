Imports System.Data
Imports System.Data.SqlClient
Partial Public Class addsapid
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public SqlCmd As String
    Public dr As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        If (Not IsPostBack) Then
            'errmsg.Visible = False
            SqlCmd = "Select sapid,sappwd From dbo.[User] where dbo.[User].id='" & Session("s_id") & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                sapidtxt.Text = dr(0)
                sappwdtxt.Text = dr(1)
            End If
            dr.Close()
            conn.Close()
        End If
    End Sub

    Protected Sub addbtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles addbtn.Click
        If (sappwdtxt.Text <> "" And sapidtxt.Text <> "") Then
            'MsgBox(sapidtxt.Text)
            Dim SqlCmd As String
            Dim sqlresult As Boolean
            'InitLocalSQLConnection()
            SqlCmd = "Update dbo.[User]  set dbo.[User].sappwd= '" & sappwdtxt.Text & "', dbo.[User].sapid='" & sapidtxt.Text & "'where dbo.[User].id='" & Session("s_id") & "'"
            'myCommand = New SqlCommand(SqlCmd, conn)
            'count = myCommand.ExecuteNonQuery()
            sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
            If (sqlresult) Then
                Response.Redirect("~\index.aspx?smid=index&smode=0&act=setsap")
            Else
                errmsg.Text = "設定失敗"
                'errmsg.Visible = True
            End If
        Else
            errmsg.Text = "欄位不能空白"
            'errmsg.Visible = True
        End If
        conn.Close()
        'Response.Redirect("~/usermgm/addsapid.aspx?smid=index&smode=2")
    End Sub
End Class