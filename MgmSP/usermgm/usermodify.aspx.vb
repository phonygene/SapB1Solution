Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Partial Public Class WebForm4
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public tTxt, tTxtName, tTxtPwd, tTxtTtl, tTxtEmail, tTxtPrice, tTxtInDate, tTxtEmpId As TextBox
    Public tDDLGrp, tDDLArea, tDDLPosition, tDDLSignLevel, tDDLAnencyPerson As DropDownList
    Public tRBLArea, tRBLDenyf, tRBLTopSignOffs As RadioButtonList
    Public tCBDel, tCBAgency As CheckBox
    Public tBtn As Button
    Public conn As New SqlConnection
    Public ds As New DataSet
    Public SqlCmd As String
    Public dr As SqlDataReader
    Public TxtSDate, TxtEDate As TextBox
    Public ScriptManager1 As New ScriptManager

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        TableCreate()
        If (IsPostBack) Then

        Else
            PutData()
        End If
    End Sub
    Sub SetText(ByVal s As Object)
        Dim txt As TextBox = s
        Dim text, id As String
        text = txt.Text
        id = txt.ID
        ViewState(id) = text
    End Sub

    Protected Sub tBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim empid As Integer
        Dim indate, branch, grp As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim candelete As Boolean
        candelete = True
        If (tTxtPwd.Text = "" Or tTxtName.Text = "" Or tDDLGrp.SelectedValue = "" Or tTxtTtl.Text = "" Or tRBLArea.SelectedValue = "" Or tRBLDenyf.SelectedValue = "") Then
            CommUtil.ShowMsg(Me, "除email 外 , 其他欄位不能空白")
            Exit Sub
        End If
        Dim SqlCmd As String
        Dim sqlresult As Boolean
        'InitLocalSQLConnection()
        Dim idstr As String
        Dim signlevel, str(), agnid As String
        Dim agnen As Integer
        idstr = ViewState("id")
        str = Split(tDDLGrp.SelectedValue, " ")
        grp = str(0)
        str = Split(tDDLArea.SelectedValue, " ")
        branch = str(0)
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

        If (tCBDel.Checked = False) Then
            If (tDDLSignLevel.SelectedValue = "NA" And tDDLPosition.SelectedValue <> "董事長") Then
                CommUtil.ShowMsg(Me, "上層簽核欄位不能空白")
                Exit Sub
            End If
            str = Split(tDDLSignLevel.SelectedValue, " ")
            signlevel = str(0)
            SqlCmd = "Update dbo.[User]  set dbo.[User].pwd= '" & tTxtPwd.Text & "' , dbo.[User].name= '" & tTxtName.Text & "' , " &
                                                                           "dbo.[User].grp= '" & grp & "' , dbo.[User].ttl= '" & FormatDateTime(tTxtTtl.Text, DateFormat.ShortDate) & "' , " &
                                                                           "dbo.[User].email= '" & tTxtEmail.Text & "' , dbo.[User].area= '" & tRBLArea.SelectedValue & "' , " &
                                                                           "dbo.[User].denyf= '" & tRBLDenyf.SelectedValue & "',dbo.[User].position='" & tDDLPosition.SelectedValue & "' , " &
                                                                           "dbo.[User].signlevel='" & signlevel & "'," &
                                                                           "dbo.[User].signprice=" & CLng(tTxtPrice.Text) & " ," &
                                                                           "dbo.[User].topsignoffs='" & tRBLTopSignOffs.SelectedValue & "'," &
                                                                           "dbo.[User].indate='" & indate & "'," &
                                                                           "dbo.[User].empidsap=" & empid & ", " &
                                                                           "dbo.[User].branch='" & branch & "', " &
                                                                           "dbo.[User].agnen=" & agnen & ", " &
                                                                           "dbo.[User].agnid='" & agnid & "', " &
                                                                           "dbo.[User].agndatefrom='" & TxtSDate.Text & "', " &
                                                                           "dbo.[User].agndateto='" & TxtEDate.Text & "' " &
                                                                           "where dbo.[User].id='" & idstr & "'"
            'myCommand = New SqlCommand(SqlCmd, conn)
            'count = myCommand.ExecuteNonQuery()
            sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
            If (sqlresult) Then
                If (CommSignOff.AgencySet(idstr) <> "") Then
                    CommSignOff.SignOffPush(Application("http"), 2)
                End If
                Response.Redirect("userlist.aspx")
            Else
                CommUtil.ShowMsg(Me, "更新失敗")
            End If
            conn.Close()
        Else
            SqlCmd = "select count(*) from [@XSPWT] T0 where uid='" & Request.QueryString("id") & "'"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            drL.Read()
            If (drL(0) > 0) Then
                candelete = False
            End If
            drL.Close()
            connL.Close()

            SqlCmd = "select count(*) from [@XSPMT] T0 where uid='" & Request.QueryString("id") & "'"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            drL.Read()
            If (drL(0) > 0) Then
                candelete = False
            End If
            drL.Close()
            connL.Close()

            SqlCmd = "select count(*) from [@XSPAT] T0 where uid='" & Request.QueryString("id") & "'"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            drL.Read()
            If (drL(0) > 0) Then
                candelete = False
            End If
            drL.Close()
            connL.Close()

            SqlCmd = "select count(*) from [@XASCH] T0 where sid='" & Request.QueryString("id") & "'"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL(0) > 0) Then
                candelete = False
            End If
            drL.Close()
            connL.Close()

            SqlCmd = "select count(*) from [@XRSCT] T0 where id='" & Request.QueryString("id") & "' or createid='" & Request.QueryString("id") & "'"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            drL.Read()
            If (drL(0) > 0) Then
                candelete = False
            End If
            drL.Close()
            connL.Close()

            SqlCmd = "select count(*) from [@XSTDT] T0 where incharge='" & Request.QueryString("id") & "' or traceperson='" & Request.QueryString("id") & "'"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            drL.Read()
            If (drL(0) > 0) Then
                candelete = False
            End If
            drL.Close()
            connL.Close()

            If (candelete) Then
                SqlCmd = "Delete From dbo.[User]  where dbo.[User].id='" & idstr & "'"
                CommUtil.SqlLocalExecute("del", SqlCmd, conn)
                conn.Close()
                SqlCmd = "Delete From dbo.[user_permissionnew]  where dbo.[user_permissionnew].id='" & idstr & "'"
                sqlresult = CommUtil.SqlLocalExecute("del", SqlCmd, conn)
                If (sqlresult) Then
                    Response.Redirect("userlist.aspx?smid=userlist")
                Else
                    CommUtil.ShowMsg(Me, "刪除User_permission失敗")
                End If
                conn.Close()
            Else
                CommUtil.ShowMsg(Me, "在系統已使用過此user之單據存在, 故無法刪除")
            End If
        End If
    End Sub

    Sub TableCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim ce As CalendarExtender
        Dim i, j As Integer
        For i = 0 To 17
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To 1
                If (i = 17 And j = 1) Then
                    Exit For
                End If
                tCell = New TableCell()
                tCell.BorderWidth = 1
                'If (j = 1) Then
                '    tCell.Width = 400
                'End If
                tRow.Cells.Add(tCell)
            Next
            Me.Table1.Rows.Add(tRow)
        Next
        Me.Table1.Rows(17).Cells(0).ColumnSpan = 2
        Me.Table1.Rows(17).Cells(0).HorizontalAlign = HorizontalAlign.Center

        Me.Table1.Rows(0).Cells(0).Text = "帳號"

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
            tDDLArea.Items.Add("請選擇區域")
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
        'tTxtTtl.AutoPostBack = True
        Me.Table1.Rows(5).Cells(1).Controls.Add(tTxtTtl)

        Me.Table1.Rows(6).Cells(0).Text = "電子郵件"
        tTxtEmail = New TextBox()
        tTxtEmail.ID = "email"
        tTxtEmail.EnableViewState = True
        Me.Table1.Rows(6).Cells(1).Controls.Add(tTxtEmail)

        Me.Table1.Rows(7).Cells(0).Text = "外部登錄"
        tRBLArea = New RadioButtonList()
        tRBLArea.Items.Add("不可外部登錄")
        tRBLArea.Items.Add("可外部登錄")
        tRBLArea.Items(0).Value = 0
        tRBLArea.Items(1).Value = 1
        tRBLArea.ID = "area_rblist"
        tRBLArea.RepeatDirection = RepeatDirection.Horizontal
        Me.Table1.Rows(7).Cells(1).Controls.Add(tRBLArea)

        Me.Table1.Rows(8).Cells(0).Text = "在職狀況"
        tRBLDenyf = New RadioButtonList()
        tRBLDenyf.Items.Add("已離職")
        tRBLDenyf.Items.Add("在職中")
        tRBLDenyf.Items(0).Value = 1
        tRBLDenyf.Items(1).Value = 0
        tRBLDenyf.ID = "denyf_rblist"
        tRBLDenyf.RepeatDirection = RepeatDirection.Horizontal
        AddHandler tRBLDenyf.SelectedIndexChanged, AddressOf tRBLDenyf_SelectedIndexChanged
        tRBLDenyf.AutoPostBack = True
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
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            Do While (dr.Read())
                tDDLSignLevel.Items.Add(dr(0) & " " & dr(1) & " " & dr(2))
            Loop
        End If
        dr.Close()
        conn.Close()
        Me.Table1.Rows(10).Cells(1).Controls.Add(tDDLSignLevel)

        Me.Table1.Rows(11).Cells(0).Text = "簽核金額"
        tTxtPrice = New TextBox()
        tTxtPrice.ID = "price"
        Me.Table1.Rows(11).Cells(1).Controls.Add(tTxtPrice)

        Me.Table1.Rows(12).Cells(0).Text = "簽核層級"
        tRBLTopSignOffs = New RadioButtonList()
        tRBLTopSignOffs.Items.Add("部門群組")
        tRBLTopSignOffs.Items.Add("跨部門")
        tRBLTopSignOffs.Items(0).Value = 0
        tRBLTopSignOffs.Items(1).Value = 1
        tRBLTopSignOffs.ID = "topsignoffs_rblist"
        tRBLTopSignOffs.RepeatDirection = RepeatDirection.Horizontal
        Me.Table1.Rows(12).Cells(1).Controls.Add(tRBLTopSignOffs)
        'Me.Table1.Rows(12).Cells(0).Visible = False
        'Me.Table1.Rows(12).Cells(1).Visible = False

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

        Me.Table1.Rows(16).Cells(0).Text = "刪除帳號"
        tCBDel = New CheckBox()
        tCBDel.AutoPostBack = True
        tCBDel.ID = "cb1"
        tCBDel.Text = "勾選後此帳號將被刪除"
        tCBDel.ForeColor = Drawing.Color.Red
        Me.Table1.Rows(16).Cells(1).Controls.Add(tCBDel)
        AddHandler tCBDel.CheckedChanged, AddressOf tCBDel_CheckedChanged

        tBtn = New Button()
        tBtn.ID = "ok"
        tBtn.Text = "儲存"
        Me.Table1.Rows(17).Cells(0).Controls.Add(tBtn)
        AddHandler tBtn.Click, AddressOf tBtn_Click
    End Sub
    Protected Sub tCBDel_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)

        If (tCBDel.Checked) Then
            tBtn.Text = "刪除"
        Else
            tBtn.Text = "儲存"
        End If
    End Sub
    Protected Sub tRBLDenyf_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim idlist As String
        Dim first As Boolean
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim mes As String
        Dim count As Integer
        count = 1
        first = True
        idlist = ""
        mes = ""
        If (tRBLDenyf.SelectedValue = 1) Then
            '找出上層主管帳號為此離職者之帳號,需更正
            SqlCmd = "select id from dbo.[user] where signlevel='" & Request.QueryString("id") & "' order by branch,grp"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                Do While (dr.Read())
                    If (first) Then
                        idlist = dr(0)
                        first = False
                    Else
                        idlist = idlist & "," & dr(0)
                    End If
                Loop
                ' CommUtil.ShowMsg(Me, idlist & " 帳號存在上層簽核為" & Request.QueryString("id") & " 需先更正")
                mes = count & ". " & idlist & " 帳號存在上層簽核為" & Request.QueryString("id") & " 需先更正"
                count = count + 1
                'tRBLDenyf.SelectedValue = 0
            End If
            dr.Close()
            conn.Close()

            SqlCmd = "select uid from [@XSPWT] T0 where uid='" & Request.QueryString("id") & "' and (status = 0 or status = 1)"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                'drL.Read()
                'CommUtil.ShowMsg(Me, "還有正在簽核的單據存在此人員,請先至簽核管理之設定工具處理")
                'drL.Close()
                'connL.Close()
                'tRBLDenyf.SelectedValue = 0
                mes = mes & "\n" & count & ". 還有正在簽核的單據存在此人員,請先至簽核管理之設定工具處理"
                count = count + 1
                'Exit Sub
            End If
            drL.Close()
            connL.Close()

            SqlCmd = "select uid from [@XSPMT] T0 where uid='" & Request.QueryString("id") & "'"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                'drL.Read()
                'CommUtil.ShowMsg(Me, "還有簽核設定表存在此人員,請先至簽核管理之設定工具處理")
                'drL.Close()
                'connL.Close()
                'tRBLDenyf.SelectedValue = 0
                mes = mes & "\n" & count & ". 還有簽核設定表存在此人員,請先至簽核管理之設定工具處理"
                count = count + 1
                'Exit Sub
            End If
            drL.Close()
            connL.Close()

            SqlCmd = "select sid from [@XASCH] T0 where sid='" & Request.QueryString("id") & "' and (status = 'A' or status = 'E' or status = 'D')"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                'drL.Read()
                'CommUtil.ShowMsg(Me, "還有未發出簽核單據存在此人員,請先至簽核管理之設定工具處理")
                'drL.Close()
                'connL.Close()
                'tRBLDenyf.SelectedValue = 0
                mes = mes & "\n" & count & ". 還有未發出簽核單據存在此人員,請先至簽核管理之設定工具處理"
                count = count + 1
                'Exit Sub
            End If
            drL.Close()
            connL.Close()

            If (mes <> "") Then
                CommUtil.ShowMsg(Me, mes)
                tRBLDenyf.SelectedValue = 0
            End If
        End If
    End Sub
    Sub PutData()
        Dim idstr As String
        Dim SelectCmd As String
        idstr = Request.QueryString("id")
        SelectCmd = "SELECT T0.id,T0.name,T0.grp,T0.email ,convert(varchar(11),T0.ttl,120) as ttl ,T0.area,T0.denyf,T0.pwd,T0.branch, " &
                    "T0.position,T0.signlevel,T0.signprice,T0.topsignoffs,T0.indate,T0.empidsap,T0.agnen,T0.agndatefrom,T0.agndateto,T0.agnid " &
                    "FROM dbo.[User] T0 " &
                    "where T0.id='" & idstr & "'"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
        conn.Close()
        Me.Table1.Rows(0).Cells(1).Text = ds.Tables(0).Rows(0)("id")

        tTxtName.Text = ds.Tables(0).Rows(0)("name")

        tTxtPwd.Attributes("value") = ds.Tables(0).Rows(0)("pwd")

        SqlCmd = "Select T0.areadesc From dbo.[branch] T0 where T0.areacode='" & ds.Tables(0).Rows(0)("branch") & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        If (dr.HasRows) Then
            tDDLArea.SelectedValue = ds.Tables(0).Rows(0)("branch") & " " & dr(0)
        Else
            CommUtil.ShowMsg(Me, "在分公司(區域)資料表中找不到 " & ds.Tables(0).Rows(0)("branch") & " 此區域")
        End If
        dr.Close()
        conn.Close()

        SqlCmd = "Select T0.deptdesc From dbo.[dept] T0 where T0.deptcode='" & ds.Tables(0).Rows(0)("grp") & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        If (dr.HasRows) Then
            tDDLGrp.SelectedValue = ds.Tables(0).Rows(0)("grp") & " " & dr(0)
        Else
            CommUtil.ShowMsg(Me, "在群組資料表中找不到 " & ds.Tables(0).Rows(0)("grp") & " 此群組")
        End If
        dr.Close()
        conn.Close()

        tTxtTtl.Text = ds.Tables(0).Rows(0)("ttl")

        tTxtEmail.Text = ds.Tables(0).Rows(0)("email")

        If (ds.Tables(0).Rows(0)("area") = 1) Then
            tRBLArea.Items(1).Selected = True
        Else
            tRBLArea.Items(0).Selected = True
        End If

        If (ds.Tables(0).Rows(0)("denyf") = 1) Then
            tRBLDenyf.Items(0).Selected = True
        Else
            tRBLDenyf.Items(1).Selected = True
        End If

        tDDLPosition.SelectedValue = ds.Tables(0).Rows(0)("position")

        SqlCmd = "Select T0.id,T0.name,T0.position From dbo.[User] T0 where T0.position<>'NA'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        Dim selectedvalue As String
        selectedvalue = "NA"
        If (dr.HasRows) Then
            Do While (dr.Read())
                If (dr(0) = ds.Tables(0).Rows(0)("signlevel")) Then
                    selectedvalue = dr(0) & " " & dr(1) & " " & dr(2)
                End If
            Loop
        End If
        dr.Close()
        conn.Close()
        tDDLSignLevel.SelectedValue = selectedvalue

        tTxtPrice.Text = ds.Tables(0).Rows(0)("signprice")

        tRBLTopSignOffs.SelectedValue = ds.Tables(0).Rows(0)("topsignoffs")
        tTxtInDate.Text = ds.Tables(0).Rows(0)("indate")
        If (ds.Tables(0).Rows(0)("empidsap") <> 0) Then
            tTxtEmpId.Text = CStr(ds.Tables(0).Rows(0)("empidsap"))
        End If
        If (ds.Tables(0).Rows(0)("agnen") <> 0) Then
            tCBAgency.Checked = True
        Else
            tCBAgency.Checked = False
        End If

        If (ds.Tables(0).Rows(0)("agnid") <> "") Then
            SqlCmd = "Select T0.id,T0.name,T0.position From dbo.[User] T0 where T0.position<>'NA'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                Do While (dr.Read())
                    If (dr(0) = ds.Tables(0).Rows(0)("agnid")) Then
                        tDDLAnencyPerson.SelectedValue = dr(0) & " " & dr(1) & " " & dr(2)
                        Exit Do
                    End If
                Loop
            End If
            dr.Close()
            conn.Close()
        End If
        TxtSDate.Text = ds.Tables(0).Rows(0)("agndatefrom")
        TxtEDate.Text = ds.Tables(0).Rows(0)("agndateto")
        ViewState("id") = idstr
    End Sub
End Class