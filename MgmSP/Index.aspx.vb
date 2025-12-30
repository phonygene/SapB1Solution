Partial Public Class WebForm1
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
        ' [移除] 生產管理簽核功能，此版本不使用 SAP 簽核資料表 (@XSPMT, @XSDET, @XSFTT)
        ' perm = CommUtil.GetAssignRight("sg300", Session("s_id"))
        ' If (InStr(perm, "m") Or InStr(perm, "n") Or InStr(perm, "d")) Then
        '     str = Split(CommSignOff.ArchiveCheck(), "-")
        '     If (str(0) <> 0 And str(1) <> 0) Then
        '         SqlCmd = "select T0.sfname from [dbo].[@XSFTT] T0 " & 'XSFTT 簽核表單種類
        '                 "where T0.sfid=" & str(0)
        '         drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        '         If (drL.HasRows) Then
        '             drL.Read()
        '             CommUtil.ShowMsg(Me, "簽核表單(" & drL(0) & ")設定歸檔屬性有" & str(2) & "個,但發現有" & str(1) & "個未設定部門排除,請前往簽核管理設定")
        '         End If
        '         drL.Close()
        '         connL.Close()
        '     End If
        ' End If
        'SqlCmd = "update [dbo].[@UPSP] set " &
        '"u_shiploc='台北發貨'"
        'CommUtil.SqlSapExecute("upd", SqlCmd, connL)
        'connL.Close()
        'CommUtil.SendMail("ron@jettech.com.tw", "mail test", "test")
        'MsgBox(System.Web.HttpContext.Current.Server.MapPath("~/"))
        'Dim dd As DateTime
        'dd = Now()
        'If (dd.DayOfWeek = 2) Then
        'MsgBox(dd.Year & "-" & dd.Month & "-" & dd.Day & "-" & dd.Hour & "-" & dd.Minute & " " & dd.DayOfWeek)
        'End If
        'Dim dd, dd1 As DateTime
        'Dim ts As TimeSpan
        'dd = Now()
        'dd1 = "2023/12/20 01:00:00"
        'ts = dd - dd1
        'MsgBox(ts.Hours & "-" & ts.Minutes & "-" & ts.Seconds)
        'Dim agnid As String
        'agnid = CommUtil.AgencySet("ron")
        'MsgBox(agnid)
    End Sub

End Class