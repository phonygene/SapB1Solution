Partial Public Class Home
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim timeout As Integer
        Dim act As String

        ' 設定用戶顯示資訊
        If Not IsPostBack Then
            If Session("s_name") IsNot Nothing AndAlso Session("s_name").ToString() <> "" Then
                lblUserName.Text = Session("s_name").ToString()
            ElseIf Session("s_id") IsNot Nothing Then
                lblUserName.Text = Session("s_id").ToString()
            End If
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
End Class
