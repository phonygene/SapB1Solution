Imports System.Web.SessionState
Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Windows
Imports System.Windows.Interop

Public Class Global_asax
    Inherits System.Web.HttpApplication
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public objTimer As New System.Timers.Timer 'With {
    '.Interval = 60000, '1 小時//这个时间单位毫秒,比如1秒，就写1000
    '.Enabled = True
    '   }

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        'Application("http") = "http://210.61.165.199/:8080/"
        'Application("http") = "http://59.124.82.189:8080/"
        'Application("http") = "http://192.168.100.112:8081/"
        Application("http") = "http://localhost:50601/"

        Application("user_sessions") = 0

        Application("localdir") = "C:\sapupload\"
        ' 在應用程式啟動時引發
        'Dim objTimer As New System.Timers.Timer With {
        '    .Interval = 3600000, '1 小時//这个时间单位毫秒,比如1秒，就写1000
        '    .Enabled = True
        '}
        Dim targetDir As String
        targetDir = Application("localdir") & "CLFormFile\"
        If (Not System.IO.Directory.Exists(targetDir)) Then
            Directory.CreateDirectory(targetDir)
        End If
        targetDir = Application("localdir") & "SignOffsFormFiles\"
        If (Not System.IO.Directory.Exists(targetDir)) Then
            Directory.CreateDirectory(targetDir)
        End If
        targetDir = Application("localdir") & "QC\DW\"
        If (Not System.IO.Directory.Exists(targetDir)) Then
            Directory.CreateDirectory(targetDir)
        End If
        targetDir = HttpContext.Current.Server.MapPath("~/") & "\AttachFile\QC\DW\"
        If (Not System.IO.Directory.Exists(targetDir)) Then
            Directory.CreateDirectory(targetDir)
        End If
        targetDir = Application("localdir") & "FileTemp\"
        If (Not System.IO.Directory.Exists(targetDir)) Then
            Directory.CreateDirectory(targetDir)
        End If
        'targetDir = HttpContext.Current.Server.MapPath("~/") & "\AttachFile\AnsFiles\"
        'If (Not System.IO.Directory.Exists(targetDir)) Then
        '    Directory.CreateDirectory(targetDir)
        'End If
        'targetDir = Application("localdir") & "AnsFiles\"
        'If (Not System.IO.Directory.Exists(targetDir)) Then
        '    Directory.CreateDirectory(targetDir)
        'End If
        Dim dd As DateTime
        Dim nowTo00diff As Long
        dd = Now()
        ''訂每日 01:00 發送未簽核mail
        'If (dd.Hour <> 0) Then
        '    nowTo00diff = (24 - dd.Hour) * 3600000 + (59 - dd.Minute) * 60000 + (60 - dd.Second) * 1000
        'Else
        '    nowTo00diff = (59 - dd.Minute) * 60000 + (60 - dd.Second) * 1000
        'End If

        '訂每日 6:00 發送未簽核mail
        If (dd.Hour >= 6) Then
            nowTo00diff = (24 - dd.Hour + 5) * 3600000 + (59 - dd.Minute) * 60000 + (60 - dd.Second) * 1000
        Else
            nowTo00diff = (5 - dd.Hour) * 3600000 + (59 - dd.Minute) * 60000 + (60 - dd.Second) * 1000
        End If

        'nowTo00diff = (59 - dd.Minute) * 60000 + (60 - dd.Second + 10) * 1000 'test
        'objTimer.Enabled = False

        'If (dd.Minute < 30) Then
        '    nowTo00diff = (30 - dd.Minute) * 60000 + (60 - dd.Second + 5) * 1000 'test
        'Else
        '    nowTo00diff = (59 - dd.Minute) * 60000 + (60 - dd.Second + 5) * 1000 'test
        'End If
        objTimer.Interval = nowTo00diff
        'objTimer.Interval = 10000
        objTimer.Enabled = True
        AddHandler objTimer.Elapsed, AddressOf objTimer_Elapsed

        'Dim delayTime As TimeSpan = New TimeSpan(0, 59 - dd.Minute, 60 - dd.Second) ' // 應用程式起動後多久開始執行
        ' Dim intervalTime As TimeSpan = New TimeSpan(0, 30, 0) ' // 應用程式起動後間隔多久重複執行
        'Dim timerDelegate As TimerCallback = New TimerCallback(BatchMethod) '  // 委派呼叫方法
        'AddHandler timerDelegate., AddressOf timerDelegate_Elapsed
        'Dim Timer As New Timer(timerDelegate, Nothing, delayTime, intervalTime) '  // 產生 timer
    End Sub

    Protected Sub objTimer_Elapsed(ByVal sender As Object, ByVal e As EventArgs)
        Dim dd As DateTime
        Dim onedayms As Long = 86400000
        dd = Now()
        objTimer.Interval = onedayms
        'objTimer.Interval = 3600000
        'CommUtil.SendMail("ron@jettech.com.tw", "timer test from local server", dd.Year & "/" & dd.Month & "/" & dd.Day & " " & dd.Hour & ":" & dd.Minute & ":" & dd.Second & " " & dd.DayOfWeek)
        If (dd.DayOfWeek <> 6 And dd.DayOfWeek <> 0) Then
            CommSignOff.SignOffPush(Application("http"), 1) 'Application("http") 不能出現在單獨 .vb 之程式,否則會執行到那沒動作
            CommSignOff.ToDoListPush(Application("http"))
        End If

        'Dim targetDir As String
        'targetDir = HttpContext.Current.Server.MapPath("~/") & "SignOffsFormFiles\"
        'File.Move(targetDir & CStr(Application("timer_count")) & ".pdf", targetDir & CStr(Application("timer_count") + 1) & ".pdf")
        'Application.Lock()
        'Application("timer_count") = Application("timer_count") + 1
        'Application.UnLock()
        'MsgBox(targetDir & CStr(Application("timer_count") - 1) & ".pdf" & "   " & targetDir & CStr(Application("timer_count")) & ".pdf")


        'CommSignOff.RecordPushSignFlowHistoty("ron", 2)
        'CommUtil.SendMail("ron@jettech.com.tw", "timer test from local server", dd.Year & "/" & dd.Month & "/" & dd.Day & " " & dd.Hour & ":" & dd.Minute & ":" & dd.Second & " " & dd.DayOfWeek)
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 在工作階段啟動時引發
        Application.Lock()
        Application("user_sessions") = CInt(Application("user_sessions")) + 1
        Application.UnLock()
        'MsgBox("session start:" & Application("user_sessions"))
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 在各個要求開始時引發
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 在嘗試驗證使用時引發
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' 在錯誤發生時引發
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 在工作階段結束時引發
        Application.Lock()
        Application("user_sessions") = CInt(Application("user_sessions")) - 1
        Application.UnLock()
        'MsgBox("session end:" & Application("user_sessions"))
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 在應用程式結束時引發
    End Sub


End Class