Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Partial Public Class WebForm5
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public tTxt, tTxtName, tTxtPwd, tTxtTtl, tTxtEmail, tTxtId, tTxtPrice, tTxtInDate, tTxtEmpId As TextBox
    Public tDDLGrp, tDDLArea, tDDLPosition, tDDLSignLevel, tDDLAnencyPerson As DropDownList
    Public tRBLArea, tRBLDenyf, tRBLTopSignOffs As RadioButtonList
    Public tCBDel, tCBAgency As CheckBox
    Public tBtn As Button
    Public conn As New SqlConnection
    Public Reader As SqlDataReader
    Public SqlCmd As String
    Public dr As SqlDataReader
    Public TxtSDate, TxtEDate As TextBox
    Public ScriptManager1 As New ScriptManager

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ce As CalendarExtender
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer

        For i = 0 To 16
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To 1
                If (i = 16 And j = 1) Then
                    Exit For
                End If
                tCell = New TableCell()
                tCell.BorderWidth = 1
                'tCell.Text = ""
                tRow.Cells.Add(tCell)
            Next
            Me.Table1.Rows.Add(tRow)
        Next
        Me.Table1.Rows(16).Cells(0).ColumnSpan = 2
        Me.Table1.Rows(16).Cells(0).HorizontalAlign = HorizontalAlign.Center


        Me.Table1.Rows(0).Cells(0).Text = "帳號"
        tTxtId = New TextBox()
        tTxtId.ID = "id"
        tTxtId.Text = ""
        Me.Table1.Rows(0).Cells(1).Controls.Add(tTxtId)

        Me.Table1.Rows(1).Cells(0).Text = "姓名"
        tTxtName = New TextBox()
        tTxtName.ID = "name"
        Me.Table1.Rows(1).Cells(1).Controls.Add(tTxtName)

        Me.Table1.Rows(2).Cells(0).Text = "密碼"
        tTxtPwd = New TextBox()
        tTxtPwd.ID = "pwd"
        tTxtPwd.TextMode = TextBoxMode.Password
        Me.Table1.Rows(2).Cells(1).Controls.Add(tTxtPwd)

        SqlCmd = "Select T0.areacode,T0.areadesc From dbo.[branch] T0"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        Me.Table1.Rows(3).Cells(0).Text = "區域"
        tDDLArea = New DropDownList()
        tDDLArea.ID = "branch_ddlist"
        If (dr.HasRows) Then
            tDDLArea.Items.Clear()
            tDDLArea.Items.Add("請選擇分區")
            Do While (dr.Read())
                tDDLArea.Items.Add(dr(0) & " " & dr(1))
            Loop
        End If
        Me.Table1.Rows(3).Cells(1).Controls.Add(tDDLArea)
        dr.Close()
        conn.Close()

        SqlCmd = "Select T0.deptcode,T0.deptdesc From dbo.[dept] T0 order by T0.deptcode"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        Me.Table1.Rows(4).Cells(0).Text = "群組"
        tDDLGrp = New DropDownList()
        tDDLGrp.ID = "grp_ddlist"
        If (dr.HasRows) Then
            tDDLGrp.Items.Clear()
            tDDLGrp.Items.Add("請選擇群組")
            Do While (dr.Read())
                tDDLGrp.Items.Add(dr(0) & " " & dr(1))
            Loop
        End If
        Me.Table1.Rows(4).Cells(1).Controls.Add(tDDLGrp)
        dr.Close()
        conn.Close()

        Me.Table1.Rows(5).Cells(0).Text = "有效期"
        tTxtTtl = New TextBox()
        tTxtTtl.ID = "ttl"
        Me.Table1.Rows(5).Cells(1).Controls.Add(tTxtTtl)

        Me.Table1.Rows(6).Cells(0).Text = "電子郵件"
        tTxtEmail = New TextBox()
        tTxtEmail.ID = "email"
        Me.Table1.Rows(6).Cells(1).Controls.Add(tTxtEmail)

        Me.Table1.Rows(7).Cells(0).Text = "外部登錄"
        tRBLArea = New RadioButtonList()
        tRBLArea.ID = "area_rblist"
        tRBLArea.Items.Add("不可外部登錄")
        tRBLArea.Items.Add("可外部登錄")
        tRBLArea.Items(0).Value = 0
        tRBLArea.Items(1).Value = 1
        tRBLArea.SelectedValue = 0
        tRBLArea.RepeatDirection = RepeatDirection.Horizontal
        Me.Table1.Rows(7).Cells(1).Controls.Add(tRBLArea)

        Me.Table1.Rows(8).Cells(0).Text = "在職狀況"
        tRBLDenyf = New RadioButtonList()
        tRBLDenyf.ID = "denyf_rblist"
        tRBLDenyf.Items.Add("已離職")
        tRBLDenyf.Items.Add("在職中")
        tRBLDenyf.Items(0).Value = 1
        tRBLDenyf.Items(1).Value = 0
        tRBLDenyf.SelectedValue = 0
        tRBLDenyf.RepeatDirection = RepeatDirection.Horizontal
        Me.Table1.Rows(8).Cells(1).Controls.Add(tRBLDenyf)

        Me.Table1.Rows(9).Cells(0).Text = "主管職稱"
        tDDLPosition = New DropDownList()
        tDDLPosition.ID = "position_ddlist"
        tDDLPosition.Items.Add("NA")
        tDDLPosition.Items.Add("董事長")
        tDDLPosition.Items.Add("總經理")
        tDDLPosition.Items.Add("副總")
        tDDLPosition.Items.Add("協理")
        tDDLPosition.Items.Add("處長")
        tDDLPosition.Items.Add("經理")
        tDDLPosition.Items.Add("副理")
        tDDLPosition.Items.Add("主任")
        Me.Table1.Rows(9).Cells(1).Controls.Add(tDDLPosition)

        Me.Table1.Rows(10).Cells(0).Text = "上層簽核"
        tDDLSignLevel = New DropDownList()
        tDDLSignLevel.ID = "signlevel_ddlist"
        tDDLSignLevel.Items.Add("NA")
        SqlCmd = "Select T0.id,T0.name,T0.position From dbo.[User] T0 where T0.position<>'NA' and denyf<>1"
        Reader = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (Reader.HasRows) Then
            Do While (Reader.Read())
                tDDLSignLevel.Items.Add(Reader(0) & " " & Reader(1) & " " & Reader(2))
            Loop
        End If
        Reader.Close()
        conn.Close()
        Me.Table1.Rows(10).Cells(1).Controls.Add(tDDLSignLevel)

        Me.Table1.Rows(11).Cells(0).Text = "簽核金額"
        tTxtPrice = New TextBox()
        tTxtPrice.ID = "price"
        tTxtPrice.Text = 0
        Me.Table1.Rows(11).Cells(1).Controls.Add(tTxtPrice)

        Me.Table1.Rows(12).Cells(0).Text = "簽核層級"
        tRBLTopSignOffs = New RadioButtonList()
        tRBLTopSignOffs.Items.Add("部門群組")
        tRBLTopSignOffs.Items.Add("跨部門")
        tRBLTopSignOffs.Items(0).Value = 0
        tRBLTopSignOffs.Items(1).Value = 1
        tRBLTopSignOffs.ID = "topsignoffs_rblist"
        tRBLTopSignOffs.RepeatDirection = RepeatDirection.Horizontal
        tRBLTopSignOffs.SelectedValue = 0
        Me.Table1.Rows(12).Cells(1).Controls.Add(tRBLTopSignOffs)

        Me.Table1.Rows(13).Cells(0).Text = "入職日期"
        tTxtInDate = New TextBox()
        tTxtInDate.ID = "indate"
        Me.Table1.Rows(13).Cells(1).Controls.Add(tTxtInDate)

        Me.Table1.Rows(14).Cells(0).Text = "Sap編號"
        tTxtEmpId = New TextBox()
        tTxtEmpId.ID = "empid"
        Me.Table1.Rows(14).Cells(1).Controls.Add(tTxtEmpId)

        Dim Labelx As Label
        Me.Table1.Rows(15).Cells(0).Text = "代理人"
        tCBAgency = New CheckBox()
        tCBAgency.ID = "CBAgency"
        tCBAgency.Text = "勾選打開代理功能<br>"
        'tCBAgency.ForeColor = Drawing.Color.Red
        Me.Table1.Rows(15).Cells(1).Controls.Add(tCBAgency)
        tDDLAnencyPerson = New DropDownList()
        tDDLAnencyPerson.ID = "AnencyPerson_ddlist"
        tDDLAnencyPerson.Items.Add("代理人選擇")
        SqlCmd = "Select T0.id,T0.name,T0.position From dbo.[User] T0 where T0.position<>'NA' and denyf<>1"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            Do While (dr.Read())
                tDDLAnencyPerson.Items.Add(dr(0) & " " & dr(1) & " " & dr(2))
            Loop
        End If
        dr.Close()
        conn.Close()
        Me.Table1.Rows(15).Cells(1).Controls.Add(tDDLAnencyPerson)
        Labelx = New Label
        Labelx.ID = "agencyfrom_label"
        Labelx.Text = "<br>代理日期 From:"
        Me.Table1.Rows(15).Cells(1).Controls.Add(Labelx)
        TxtSDate = New TextBox()
        TxtSDate.ID = "txt_sdate"
        TxtSDate.Width = 100
        'tCell.Controls.Add(TxtSDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtSDate.ID
        ce.ID = "ce_begindate"
        ce.Format = "yyyy/MM/dd"
        Me.Table1.Rows(15).Cells(1).Controls.Add(ce)
        Me.Table1.Rows(15).Cells(1).Controls.Add(TxtSDate)
        Labelx = New Label
        Labelx.ID = "agencyto_label"
        Labelx.Text = "<br>代理日期 To&nbsp&nbsp&nbsp&nbsp:"
        Me.Table1.Rows(15).Cells(1).Controls.Add(Labelx)
        TxtEDate = New TextBox()
        TxtEDate.ID = "txt_edate"
        TxtEDate.Width = 100
        'tCell.Controls.Add(TxtEDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtEDate.ID
        ce.ID = "ce_enddate"
        ce.Format = "yyyy/MM/dd"
        Me.Table1.Rows(15).Cells(1).Controls.Add(ce)
        Me.Table1.Rows(15).Cells(1).Controls.Add(TxtEDate)

        tBtn = New Button()
        tBtn.ID = "ok"
        tBtn.Text = "新增"
        tBtn.PostBackUrl = "useradd.aspx"
        Me.Table1.Rows(16).Cells(0).Controls.Add(tBtn)
        AddHandler tBtn.Click, AddressOf tBtn_Click
    End Sub

    Protected Sub tBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim SqlCmd, str(), agnid As String
        Dim empid As Integer
        Dim id, pwd, name, grp, ttl, email, makeby, position, signlevel, topsignoffs, indate, branch, area, denyf As String
        Dim makedate As Date
        Dim signprice As Long
        Dim agnen As Integer
        makeby = Session("s_id")
        'makedate = Format(Now(), "yyyy-mmm-dd")
        makedate = FormatDateTime(Now(), DateFormat.ShortDate)
        id = tTxtId.Text
        pwd = tTxtPwd.Text
        name = tTxtName.Text
        str = Split(tDDLGrp.SelectedValue, " ")
        grp = str(0)
        str = Split(tDDLArea.SelectedValue, " ")
        branch = str(0)
        ttl = tTxtTtl.Text
        email = tTxtEmail.Text
        area = tRBLArea.SelectedValue
        denyf = tRBLDenyf.SelectedValue
        position = tDDLPosition.SelectedValue
        str = Split(tDDLSignLevel.SelectedValue, " ")
        signlevel = str(0)
        signprice = CLng(tTxtPrice.Text)
        topsignoffs = tRBLTopSignOffs.SelectedValue
        indate = tTxtInDate.Text
        If (tTxtEmpId.Text = "") Then
            empid = 0
        Else
            empid = CInt(tTxtEmpId.Text)
        End If
        If (tDDLAnencyPerson.SelectedIndex = 0) Then
            agnid = ""
        Else
            str = Split(tDDLAnencyPerson.SelectedValue, " ")
            agnid = str(0)
        End If
        If (tCBAgency.Checked = False) Then
            agnen = 0
        Else
            agnen = 1
        End If
        If (id = "" Or pwd = "" Or name = "" Or grp = "" Or ttl = "" Or area = "" Or denyf = "") Then
            CommUtil.ShowMsg(Me, "除email 外 , 其他欄位不能空白")
            Exit Sub
        End If
        If (tDDLSignLevel.SelectedValue = "NA" And position <> "董事長") Then
            CommUtil.ShowMsg(Me, "上層簽核欄位不能空白")
            Exit Sub
        End If
        SqlCmd = "Select id From dbo.[User] where dbo.[User].id='" & id & "'"
        Reader = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (Not Reader.HasRows) Then
            Reader.Close()
            conn.Close()
            SqlCmd = "Insert Into dbo.[User] (dbo.[User].id,dbo.[User].pwd,dbo.[User].name,dbo.[User].grp,dbo.[User].ttl," &
                    "dbo.[User].email,dbo.[User].area,dbo.[User].denyf,makeby,makedate,position,signlevel,signprice,topsignoffs,indate,empidsap,branch," &
                    "agnen,agnid,agndatefrom,agndateto) " &
                    "Values ('" & id & "', '" & pwd & "','" & name & "','" & grp &
                    "','" & ttl & "','" & email & "', '" & area & "','" & denyf &
                    "','" & makeby & "','" & makedate & "','" & position & "','" & signlevel & "'," & signprice & ",'" & topsignoffs &
                    "','" & indate & "'," & empid & ",'" & branch & "'," & agnen & ",'" & agnid & "','" & TxtSDate.Text & "','" & TxtEDate.Text & "')"

            CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
            conn.Close()
            '寫入default permission
            'sg
            'Dim perm, pid As String
            'perm = ""
            'pid = ""
            'For i = 0 To 4
            '    If (i = 0) Then
            '        pid = "sg000"
            '        perm = "e"
            '    ElseIf (i = 1) Then
            '        pid = "sg100"
            '        perm = "e"
            '    ElseIf (i = 2) Then
            '        pid = "sg200"
            '        perm = "e"
            '    ElseIf (i = 3) Then
            '        pid = "sg300"
            '        perm = "e"
            '    ElseIf (i = 4) Then
            '        pid = "sg400"
            '        perm = "e"
            '    End If
            '    SqlCmd = "Insert Into dbo.[user_permissionnew] (id,pid,permission) " &
            '                    "Values ('" & id & "','" & pid & "','" & perm & "')"
            '    CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
            '    conn.Close()
            'Next
            'If (sqlresult) Then
            Response.Redirect("userlist.aspx?smid=userlist")
            'Else
            'CommUtil.ShowMsg(Me,"新增失敗")
            'End If
        Else
            'myCommand.Cancel()
            Reader.Close()
            conn.Close()
            CommUtil.ShowMsg(Me,"帳號:" & id & "已存在")
        End If
    End Sub
End Class