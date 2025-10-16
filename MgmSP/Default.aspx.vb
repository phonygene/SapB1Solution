Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Partial Public Class _Default
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public ScriptManager1 As New ScriptManager
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim dde As CalendarExtender
        'CommUtil.ConfirmBox(Me, "主頁為index.php,按確定轉至index")
        Page.Form.Controls.Add(ScriptManager1)
        Response.Redirect("~/index.aspx?smid=index")
        dde = New CalendarExtender
        dde.TargetControlID = TextBox1.ID
        dde.ID = "ce_shipdate"


        'MsgBox1("主頁為index.php,按確定轉至index")
        'MsgBox("KK")
        'Test()
        'If (Not IsPostBack) Then
        'Button1.OnClientClick = "return confirm('要複製嗎')"
        'MsgBox("KK")
        'End If
        'ConfirmBox("KKKKK")
        ' ConfirmBox("bbbbbb")
        'MsgBox("JJ")
    End Sub
    Public Sub ConfirmBox(ByVal Message As String)
        Dim sScript As String
        Dim sMessage As String
        sMessage = Strings.Replace(Message, "'", "\'") '處理單引號
        sMessage = Strings.Replace(sMessage, vbNewLine, "\n") '處理換行
        sScript = String.Format("confirm('{0}');", sMessage)
        'sScript = "confirm('KK')"
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "confirm", sScript, True)

    End Sub
    Sub Test()
        Dim str As String
        Dim text As String
        text = "你確定嗎"
        str = "<script type=application/javascript> confirm('show') </script>"
        Response.Write(str)
        'MsgBox a
    End Sub
    Sub MsgBox1(ByVal text As String)
        Dim scriptstr As String
        scriptstr = "<script language=javascript>" + Chr(10) _
        + "confirm(""" + text + """)" + Chr(10) _
        + "</script>"
        Response.Write(scriptstr)
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox2.Text = TextBox1.Text
    End Sub
End Class