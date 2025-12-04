Imports System.Data
Imports System.Data.SqlClient
Public Class rolesetup
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn, conn1, conn2 As New SqlConnection
    Public tTxt As TextBox
    Public dr, dr1, dr2 As SqlDataReader
    Public idstr As String
    Public SqlCmd As String
    Public ds As New DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        idstr = Request.QueryString("id")
        SqlCmd = "Select distinct T0.pgrpid from dbo.[permissiongrp] T0"
        dr2 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn2)
        If (dr2.HasRows) Then
            Do While (dr2.Read())
                RoleSet(dr2(0))
                'RoleSet("ac")
                'RoleSet("mf")
                'RoleSet("qc")
                'RoleSet("sp")
            Loop
        End If
        dr2.Close()
        conn2.Close()
    End Sub
    Sub RoleSet(pgrpid As String)
        Dim tRow As TableRow
        Dim tCell As TableCell
        Dim i, count As Integer
        Dim label1 As Label
        Dim tBtn As Button
        Dim pstr, defall As String

        SqlCmd = "Select T0.num,T0.pid,T0.pgrpid,T0.pgrpdesc,T0.defall from dbo.[permissiongrp] T0 where T0.pgrpid='" & pgrpid & "' order by T0.pid"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        count = ds.Tables(0).Rows.Count
        If (Session("s_id") = "ron") Then
            tRow = New TableRow()
            tRow.HorizontalAlign = HorizontalAlign.Center
            tRow.BackColor = Drawing.Color.DeepSkyBlue
            'tRow.Font.Bold = True
            tCell = New TableCell()
            tCell.Text = "代號"
            tRow.Cells.Add(tCell)
            For i = 0 To count - 1
                tCell = New TableCell()
                tCell.Text = ds.Tables(0).Rows(i)("pid")
                tRow.Cells.Add(tCell)
            Next
            RT.Rows.Add(tRow)
        End If
        tRow = New TableRow()
        tRow.HorizontalAlign = HorizontalAlign.Center
        tRow.BackColor = Drawing.Color.DeepSkyBlue
        tRow.Font.Bold = True
        tCell = New TableCell()
        tCell.Text = "使用者"
        tRow.Cells.Add(tCell)
        For i = 0 To count - 1
            tCell = New TableCell()
            tCell.Text = ds.Tables(0).Rows(i)("pgrpdesc")
            tRow.Cells.Add(tCell)
        Next
        RT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.HorizontalAlign = HorizontalAlign.Center
        tRow.BackColor = Drawing.Color.LightGreen
        'tRow.Font.Bold = True
        tCell = New TableCell()
        label1 = New Label()
        label1.ID = "label_" & ds.Tables(0).Rows(0)("pgrpid")
        label1.Text = idstr & "<br>"
        tCell.Controls.Add(label1)
        tBtn = New Button()
        tBtn.ID = "save_" & ds.Tables(0).Rows(0)("pgrpid") & "_" & idstr
        tBtn.Text = "儲存"
        tCell.Controls.Add(tBtn)
        AddHandler tBtn.Click, AddressOf tBtn_Click
        Dim syspid As String
        tRow.Cells.Add(tCell)
        For i = 0 To count - 1
            syspid = ds.Tables(0).Rows(i)("pid")
            SqlCmd = "Select T0.num,T0.permission,T0.pid from dbo.[user_permissionnew] T0 where T0.id='" & idstr & "' and T0.pid='" & syspid & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                pstr = dr(1)
            Else
                pstr = ""
            End If
            defall = ds.Tables(0).Rows(i)("defall")
            tCell = New TableCell()
            Show_Permission(tCell, pstr, defall, ds.Tables(0).Rows(i)("pgrpid"), syspid, idstr, 0)
            tRow.Cells.Add(tCell)
            dr.Close()
            conn.Close()
        Next
        RT.Rows.Add(tRow)
        ds.Reset()
    End Sub

    Function Show_Permission(ByVal tCell As TableCell, ByVal permission As String, ByVal defall As String, pgrpid As String, syspid As String, uid As String, showtype As Integer)
        Dim cChkBox As CheckBox
        Dim i As Integer
        Dim pch As String
        For i = 1 To defall.Length
            pch = Mid(defall, i, 1)
            cChkBox = New CheckBox
            cChkBox.ID = "chk_" & pch & "_" & pgrpid & "_" & syspid
            If (pch = "e") Then
                cChkBox.Text = "進入<br>"
            ElseIf (pch = "n") Then
                cChkBox.Text = "新增<br>"
            ElseIf (pch = "m") Then
                cChkBox.Text = "修改<br>"
            ElseIf (pch = "d") Then
                cChkBox.Text = "刪除<br>"
            ElseIf (pch = "a") Then
                cChkBox.Text = "審核<br>"
            ElseIf (pch = "p") Then
                cChkBox.Text = "金額<br>"
            ElseIf (pch = "s") Then
                cChkBox.Text = "統計<br>"
            End If
            If (showtype = 0) Then
                If (InStr(permission, pch)) Then
                    cChkBox.Checked = True
                End If
            End If
            tCell.Controls.Add(cChkBox)
        Next
        Return tCell
    End Function

    Protected Sub tBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Dim saveid As String
        Dim pgrpid, syspid, chkid, pch, defall As String
        Dim str() As String
        Dim count As Integer
        Dim InsertFlag As Boolean
        Dim i, j As Integer
        Dim out, uid As String
        str = Split(sender.ID, "_")
        pgrpid = str(1)
        uid = str(2)
        SqlCmd = "Select T0.pid,T0.defall from dbo.[permissiongrp] T0 where T0.pgrpid='" & pgrpid & "' order by T0.pid"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        count = ds.Tables(0).Rows.Count
        For i = 0 To count - 1
            syspid = ds.Tables(0).Rows(i)("pid")
            defall = ds.Tables(0).Rows(i)("defall")
            SqlCmd = "Select T0.num from dbo.[user_permissionnew] T0 where T0.id='" & idstr & "' and T0.pid='" & syspid & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            out = ""
            For j = 1 To defall.Length
                pch = Mid(defall, j, 1)
                chkid = "chk_" & pch & "_" & pgrpid & "_" & syspid
                If (CType(RT.FindControl(chkid), CheckBox).Checked) Then
                    out = out & pch
                End If
            Next
            If (dr.HasRows) Then
                InsertFlag = False
            Else
                InsertFlag = True
            End If
            dr.Close()
            conn.Close()
            If (InsertFlag) Then
                SqlCmd = "insert into dbo.[user_permissionnew] (id,pid,permission) " &
                         "values('" & uid & "','" & syspid & "','" & out & "')"
                If (CommUtil.SqlLocalExecute("ins", SqlCmd, conn)) Then
                    CommUtil.ShowMsg(Me, "新增成功")
                Else
                    CommUtil.ShowMsg(Me, "新增失敗")
                End If
                conn.Close()
            Else
                SqlCmd = "update dbo.[user_permissionnew] set permission='" & out & "' " &
                        "where id='" & uid & "' and pid='" & syspid & "'"
                If (CommUtil.SqlLocalExecute("upd", SqlCmd, conn)) Then
                    CommUtil.ShowMsg(Me, "修改成功")
                Else
                    CommUtil.ShowMsg(Me, "修改失敗")
                End If
                conn.Close()
            End If
        Next
        ds.Reset()
    End Sub
End Class