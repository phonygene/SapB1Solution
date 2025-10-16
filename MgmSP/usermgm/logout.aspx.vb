Public Partial Class logout
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim i As Integer
        Dim act As String
        Session.Clear()
        Session.Abandon()
        act = Request.QueryString("act")
        'For i = 0 To Response.Cookies.Count - 1
        '    Response.Cookies(i).Expires = DateTime.Now '//cookie將馬上過期
        'Next
        If (act <> "") Then
            Response.Redirect("~/index.aspx?act=" & act)
        Else
            Response.Redirect("~/index.aspx")
        End If
        'Response.Write("<script>window.open();</script>")
        'Response.Write("<script>window.close();</script>")
        'Response.Write("<Script language='JavaScript'>alert('喔喔!這裡寫入你的訊息喔!');</Script>")
        'Response.Write("<script language='JavaScript'>window.opener=null;window.close();</script>")
        'Dim JScript As New System.Text.StringBuilder("")
        'JScript.Append("window.close()")
        'Me.Page.ClientScript.RegisterStartupScript(Me.GetType(), "Return Value", JScript.ToString(), True)
    End Sub

End Class