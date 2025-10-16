Imports System.Data
Imports System.Data.SqlClient
Partial Public Class rolemodify
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public tTxt As TextBox
    Public dr As SqlDataReader
    Public idstr As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        's100 :每日工作
        'p100 :帳戶管理
        'p200 :製造管理  p201:開工單  p202:發料通知  p203:領料操作
        Dim tBtn As Button
        Dim label1 As Label
        Dim tRow As TableRow
        Dim tCell As TableCell
        Dim i, j, k As Integer
        tRow = New TableRow()
        i = 0
        k = 8
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = k + 1
        Table1.Rows(i).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).BackColor = Drawing.Color.DeepSkyBlue
        Table1.Rows(i).Font.Bold = True
        Table1.Rows(i).Cells(0).Text = "權限修改中..."

        tRow = New TableRow()
        i = i + 1
        For j = 0 To k
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).BackColor = Drawing.Color.DeepSkyBlue
        Table1.Rows(i).Font.Bold = True
        Table1.Rows(i).Cells(0).Text = "使用人"
        Table1.Rows(i).Cells(1).Text = "每日工作"
        Table1.Rows(i).Cells(2).Text = "帳戶管理"
        Table1.Rows(i).Cells(3).Text = "製造管理"
        Table1.Rows(i).Cells(4).Text = "工單開立"
        Table1.Rows(i).Cells(5).Text = "領料通知"
        Table1.Rows(i).Cells(6).Text = "領料操作"
        Table1.Rows(i).Cells(7).Text = "工單總表"
        Table1.Rows(i).Cells(8).Text = "工單狀態"
        tRow = New TableRow()
        i = i + 1
        For j = 0 To k
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).HorizontalAlign = HorizontalAlign.Left
        Table1.Rows(i).BackColor = Drawing.Color.LightGreen
        idstr = Request.QueryString("id")
        label1 = New Label()
        label1.ID = "label_save"
        label1.Text = idstr & "<br>"
        Table1.Rows(i).Cells(0).Controls.Add(label1)
        tBtn = New Button()
        tBtn.ID = "save"
        tBtn.Text = "儲存"
        Me.Table1.Rows(i).Cells(0).Controls.Add(tBtn)
        AddHandler tBtn.Click, AddressOf tBtn_Click

        Dim SqlCmd As String
        'InitLocalSQLConnection()
        If (Not IsPostBack) Then
            SqlCmd = "Select s100,p100,p200,p201,p202,p203,p204,p205 From dbo.[user_permission] where dbo.[user_permission].id='" & idstr & "'"
            'myCommand = New SqlCommand(SqlCmd, conn)
            'dr = myCommand.ExecuteReader()
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                '增加權限項目步驟 :
                '1:在此依序增加:Page_Load之Title加入 ==>cell 增加(上述之 K 值要增加)=>上述的SqlCmd ==> 下述的Show_Permission
                '2:tBtn_Click 也要填入或修改此enmd變化
                '3:在user_permission 資料表加入此權限代號欄位
                Show_Permission("s100", dr(0), "enmd", 1, 0) '
                Show_Permission("p100", dr(1), "enmd", 2, 0) '帳號管理
                Show_Permission("p200", dr(2), "enmd", 3, 0) '製造管理
                Show_Permission("p201", dr(3), "enmd", 4, 0) '開工單
                Show_Permission("p202", dr(4), "enmd", 5, 0) '發料通知
                Show_Permission("p203", dr(5), "enmd", 6, 0) '發料操作
                Show_Permission("p204", dr(6), "enmd", 7, 0) '加工總表
                Show_Permission("p205", dr(7), "enmd", 8, 0) '工單狀態
                dr.Close()
            End If
            conn.Close()
        Else
            Show_Permission("s100", "", "enmd", 1, 1)
            Show_Permission("p100", "", "enmd", 2, 1)
            Show_Permission("p200", "", "enmd", 3, 1)
            Show_Permission("p201", "", "enmd", 4, 1)
            Show_Permission("p202", "", "enmd", 5, 1) '發料通知
            Show_Permission("p203", "", "enmd", 6, 1) '發料操作
            Show_Permission("p204", "", "enmd", 7, 1) '加工總表
            Show_Permission("p205", "", "enmd", 8, 1) '工單狀態
        End If
    End Sub

    Sub Show_Permission(ByVal sysid As String, ByVal permission As String, ByVal defall As String, ByVal col As Integer, ByVal showtype As Integer)
        Dim cChkBox As CheckBox

        If (InStr(defall, "e")) Then
            cChkBox = New CheckBox
            cChkBox.ID = sysid & "_e"
            cChkBox.Text = "進入<br>"
            If (showtype = 0) Then
                If (InStr(permission, "e")) Then
                    cChkBox.Checked = True
                End If
            End If
            Me.Table1.Rows(2).Cells(col).Controls.Add(cChkBox)
        End If
        If (InStr(defall, "n")) Then
            cChkBox = New CheckBox
            cChkBox.ID = sysid & "_n"
            cChkBox.Text = "新增<br>"
            If (showtype = 0) Then
                If (InStr(permission, "n")) Then
                    cChkBox.Checked = True
                End If
            End If
            Me.Table1.Rows(2).Cells(col).Controls.Add(cChkBox)
        End If
        If (InStr(defall, "m")) Then
            cChkBox = New CheckBox
            cChkBox.ID = sysid & "_m"
            cChkBox.Text = "修改<br>"
            If (showtype = 0) Then
                If (InStr(permission, "m")) Then
                    cChkBox.Checked = True
                End If
            End If
            Me.Table1.Rows(2).Cells(col).Controls.Add(cChkBox)
        End If
        If (InStr(defall, "d")) Then
            cChkBox = New CheckBox
            cChkBox.ID = sysid & "_d"
            cChkBox.Text = "刪除<br>"
            If (showtype = 0) Then
                If (InStr(permission, "d")) Then
                    cChkBox.Checked = True
                End If
            End If
            Me.Table1.Rows(2).Cells(col).Controls.Add(cChkBox)
        End If
        'If (InStr(defall, "p")) Then
        '    cChkBox = New CheckBox
        '    cChkBox.ID = sysid & "_p"
        '    cChkBox.Text = "開工單<br>"
        '    If (showtype = 0) Then
        '        If (InStr(permission, "p")) Then
        '            cChkBox.Checked = True
        '        End If
        '    End If
        '    Me.Table1.Rows(2).Cells(col).Controls.Add(cChkBox)
        'End If
        'If (InStr(defall, "q")) Then
        '    cChkBox = New CheckBox
        '    cChkBox.ID = sysid & "_q"
        '    cChkBox.Text = "領料通知<br>"
        '    If (showtype = 0) Then
        '        If (InStr(permission, "q")) Then
        '            cChkBox.Checked = True
        '        End If
        '    End If
        '    Me.Table1.Rows(2).Cells(col).Controls.Add(cChkBox)
        'End If
        'If (InStr(defall, "r")) Then
        '    cChkBox = New CheckBox
        '    cChkBox.ID = sysid & "_r"
        '    cChkBox.Text = "領料操作<br>"
        '    If (showtype = 0) Then
        '        If (InStr(permission, "r")) Then
        '            cChkBox.Checked = True
        '        End If
        '    End If
        '    Me.Table1.Rows(2).Cells(col).Controls.Add(cChkBox)
        'End If
    End Sub

    Protected Sub tBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim cur_permission, cur_system As String
        Dim all_system() As String = {"s100_enmd", "p100_enmd", "p200_enmd", "p201_enmd", "p202_enmd", "p203_enmd", "p204_enmd",
                                      "p205_enmd"}
        'Dim myCommand As SqlCommand
        Dim SqlCmd As String
        Dim sqlresult As Boolean
        Dim str() As String
        Dim pstr As String
        Dim sysid As String
        Dim ok As Boolean
        ok = True
        CommUtil.InitLocalSQLConnection(conn)
        For Each cur_system In all_system
            str = Split(cur_system, "_")
            sysid = str(0)
            pstr = str(1)
            cur_permission = GetPermision(sysid, pstr)
            SqlCmd = "Update dbo.[user_permission]  set " & sysid & "='" & cur_permission & "'where dbo.[user_permission].id='" & idstr & "'"
            'myCommand = New SqlCommand(SqlCmd, conn)
            'count = myCommand.ExecuteNonQuery()
            sqlresult = CommUtil.SqlExecute("upd", SqlCmd, conn)
            If (sqlresult = False) Then
                CommUtil.ShowMsg(Me, "更新失敗")
                ok = False
            End If
        Next
        If (ok) Then
            Me.Table1.Rows(0).Cells(0).Text = "權限修改中==>成功"
        End If
        conn.Close()
    End Sub

    Function GetPermision(ByVal sysid As String, pstr As String)
        Dim out As String
        Dim cChkBox As CheckBox
        out = ""
        'CommUtil.ShowMsg(Me,sysid & "**" & pstr)
        If (InStr(pstr, "e")) Then
            cChkBox = Table1.FindControl(sysid & "_e")
            If (cChkBox.Checked) Then
                out = out & "e"
            End If
        End If
        If (InStr(pstr, "n")) Then
            cChkBox = Table1.FindControl(sysid & "_n")
            If (cChkBox.Checked) Then
                out = out & "n"
            End If
        End If
        If (InStr(pstr, "m")) Then
            cChkBox = Table1.FindControl(sysid & "_m")
            If (cChkBox.Checked) Then
                out = out & "m"
            End If
        End If
        If (InStr(pstr, "d")) Then
            cChkBox = Table1.FindControl(sysid & "_d")
            If (cChkBox.Checked) Then
                out = out & "d"
            End If
        End If
        'If (InStr(pstr, "p")) Then
        '    cChkBox = Table1.FindControl(sysid & "_p")
        '    If (cChkBox.Checked) Then
        '        out = out & "p"
        '    End If
        '    'CommUtil.ShowMsg(Me,out)
        'End If
        'If (InStr(pstr, "q")) Then
        '    cChkBox = Table1.FindControl(sysid & "_q")
        '    If (cChkBox.Checked) Then
        '        out = out & "q"
        '    End If
        'End If
        'If (InStr(pstr, "r")) Then
        '    cChkBox = Table1.FindControl(sysid & "_r")
        '    If (cChkBox.Checked) Then
        '        out = out & "r"
        '    End If
        'End If
        Return out
    End Function
End Class