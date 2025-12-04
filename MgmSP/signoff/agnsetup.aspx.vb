Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class agnsetup
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public tTxtPwd, tTxtEmail, tTxtInDate, tTxtEmpId As TextBox
    Public tDDLAnencyPerson As DropDownList
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
        If (Request.QueryString("act") = "save") Then
            CommUtil.ShowMsg(Me, "更新成功")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        TableCreate()
        If (IsPostBack) Then

        Else
            PutData()
        End If
    End Sub
    Sub TableCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim ce As CalendarExtender
        Dim i, j As Integer
        For i = 0 To 12
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To 1
                If (i = 12 And j = 1) Then
                    Exit For
                End If
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tRow.Cells.Add(tCell)
            Next
            Me.Table1.Rows.Add(tRow)
        Next
        Me.Table1.Rows(12).Cells(0).ColumnSpan = 2
        Me.Table1.Rows(12).Cells(0).HorizontalAlign = HorizontalAlign.Center

        Me.Table1.Rows(0).Cells(0).Text = "帳號"

        Me.Table1.Rows(1).Cells(0).Text = "姓名"

        Me.Table1.Rows(2).Cells(0).Text = "密碼"
        tTxtPwd = New TextBox()
        tTxtPwd.ID = "pwd"
        tTxtPwd.TextMode = TextBoxMode.Password
        Me.Table1.Rows(2).Cells(1).Controls.Add(tTxtPwd)

        Me.Table1.Rows(3).Cells(0).Text = "區域"

        Me.Table1.Rows(4).Cells(0).Text = "群組"

        Me.Table1.Rows(5).Cells(0).Text = "有效期"

        Me.Table1.Rows(6).Cells(0).Text = "電子郵件"
        tTxtEmail = New TextBox()
        tTxtEmail.ID = "email"
        tTxtEmail.EnableViewState = True
        Me.Table1.Rows(6).Cells(1).Controls.Add(tTxtEmail)

        Me.Table1.Rows(7).Cells(0).Text = "主管職稱"

        Me.Table1.Rows(8).Cells(0).Text = "簽核金額"

        Me.Table1.Rows(9).Cells(0).Text = "入職日期"
        tTxtInDate = New TextBox()
        tTxtInDate.ID = "indate"
        Me.Table1.Rows(9).Cells(1).Controls.Add(tTxtInDate)

        Me.Table1.Rows(10).Cells(0).Text = "Sap編號"
        tTxtEmpId = New TextBox()
        tTxtEmpId.ID = "empid"
        Me.Table1.Rows(10).Cells(1).Controls.Add(tTxtEmpId)

        Dim Labelx As Label
        Me.Table1.Rows(11).Cells(0).Text = "代理人"
        tCBAgency = New CheckBox()
        tCBAgency.ID = "CBAgency"
        tCBAgency.Text = "勾選打開代理功能<br>"
        'tCBAgency.ForeColor = Drawing.Color.Red
        Me.Table1.Rows(11).Cells(1).Controls.Add(tCBAgency)
        tDDLAnencyPerson = New DropDownList()
        tDDLAnencyPerson.ID = "AnencyPerson_ddlist"
        tDDLAnencyPerson.Items.Add("代理人選擇")
        SqlCmd = "Select T0.id,T0.name,T0.position From dbo.[User] T0 where T0.position<>'NA'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            Do While (dr.Read())
                tDDLAnencyPerson.Items.Add(dr(0) & " " & dr(1) & " " & dr(2))
            Loop
        End If
        dr.Close()
        conn.Close()
        Me.Table1.Rows(11).Cells(1).Controls.Add(tDDLAnencyPerson)
        Labelx = New Label
        Labelx.ID = "agencyfrom_label"
        Labelx.Text = "<br>代理日期 From:"
        Me.Table1.Rows(11).Cells(1).Controls.Add(Labelx)
        TxtSDate = New TextBox()
        TxtSDate.ID = "txt_sdate"
        TxtSDate.Width = 100
        'tCell.Controls.Add(TxtSDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtSDate.ID
        ce.ID = "ce_begindate"
        ce.Format = "yyyy/MM/dd"
        Me.Table1.Rows(11).Cells(1).Controls.Add(ce)
        Me.Table1.Rows(11).Cells(1).Controls.Add(TxtSDate)
        Labelx = New Label
        Labelx.ID = "agencyto_label"
        Labelx.Text = "<br>代理日期 To&nbsp&nbsp&nbsp&nbsp:"
        Me.Table1.Rows(11).Cells(1).Controls.Add(Labelx)
        TxtEDate = New TextBox()
        TxtEDate.ID = "txt_edate"
        TxtEDate.Width = 100
        'tCell.Controls.Add(TxtEDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtEDate.ID
        ce.ID = "ce_enddate"
        ce.Format = "yyyy/MM/dd"
        Me.Table1.Rows(11).Cells(1).Controls.Add(ce)
        Me.Table1.Rows(11).Cells(1).Controls.Add(TxtEDate)

        tBtn = New Button()
        tBtn.ID = "ok"
        tBtn.Text = "儲存"
        Me.Table1.Rows(12).Cells(0).Controls.Add(tBtn)
        AddHandler tBtn.Click, AddressOf tBtn_Click
    End Sub
    Sub PutData()
        Dim idstr As String
        Dim SelectCmd As String
        idstr = Session("s_id")
        SelectCmd = "SELECT T0.id,T0.name,T0.grp,T0.email ,convert(varchar(11),T0.ttl,120) as ttl ,T0.area,T0.denyf,T0.pwd,T0.branch, " &
                    "T0.position,T0.signlevel,T0.signprice,T0.topsignoffs,T0.indate,T0.empidsap,T0.agnen,T0.agndatefrom,T0.agndateto,T0.agnid " &
                    "FROM dbo.[User] T0 " &
                    "where T0.id='" & idstr & "'"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
        conn.Close()
        Me.Table1.Rows(0).Cells(1).Text = ds.Tables(0).Rows(0)("id")

        Me.Table1.Rows(1).Cells(1).Text = ds.Tables(0).Rows(0)("name")

        tTxtPwd.Attributes("value") = ds.Tables(0).Rows(0)("pwd")

        SqlCmd = "Select T0.areadesc From dbo.[branch] T0 where T0.areacode='" & ds.Tables(0).Rows(0)("branch") & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        If (dr.HasRows) Then
            Me.Table1.Rows(3).Cells(1).Text = ds.Tables(0).Rows(0)("branch") & " " & dr(0)
        Else
            CommUtil.ShowMsg(Me, "在分公司(區域)資料表中找不到 " & ds.Tables(0).Rows(0)("branch") & " 此區域")
        End If
        dr.Close()
        conn.Close()

        SqlCmd = "Select T0.deptdesc From dbo.[dept] T0 where T0.deptcode='" & ds.Tables(0).Rows(0)("grp") & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        If (dr.HasRows) Then
            Me.Table1.Rows(4).Cells(1).Text = ds.Tables(0).Rows(0)("grp") & " " & dr(0)
        Else
            CommUtil.ShowMsg(Me, "在群組資料表中找不到 " & ds.Tables(0).Rows(0)("grp") & " 此群組")
        End If
        dr.Close()
        conn.Close()

        Me.Table1.Rows(5).Cells(1).Text = ds.Tables(0).Rows(0)("ttl")

        tTxtEmail.Text = ds.Tables(0).Rows(0)("email")

        Me.Table1.Rows(7).Cells(1).Text = ds.Tables(0).Rows(0)("position")

        Me.Table1.Rows(8).Cells(1).Text = ds.Tables(0).Rows(0)("signprice")

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
    End Sub
    Protected Sub tBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim empid As Integer
        Dim indate As String
        Dim SqlCmd As String
        Dim sqlresult As Boolean
        Dim str(), agnid As String
        Dim agnen As Integer

        If (tTxtPwd.Text = "") Then
            CommUtil.ShowMsg(Me, "密碼欄位不能空白")
            Exit Sub
        End If
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
        SqlCmd = "Update dbo.[User]  set dbo.[User].pwd= '" & tTxtPwd.Text & "' , " &
                                        "dbo.[User].email= '" & tTxtEmail.Text & "' , " &
                                        "dbo.[User].indate='" & indate & "'," &
                                        "dbo.[User].empidsap=" & empid & ", " &
                                        "dbo.[User].agnen=" & agnen & ", " &
                                        "dbo.[User].agnid='" & agnid & "', " &
                                        "dbo.[User].agndatefrom='" & TxtSDate.Text & "', " &
                                        "dbo.[User].agndateto='" & TxtEDate.Text & "' " &
                                        "where dbo.[User].id='" & Session("s_id") & "'"
        sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        If (CommSignOff.AgencySet(Session("s_id")) <> "") Then
            CommSignOff.SignOffPush(Application("http"), 2)
        End If
        If (sqlresult) Then
            Response.Redirect("~/signoff/agnsetup.aspx?smid=sg&smode=4&act=save")
        Else
            CommUtil.ShowMsg(Me, "更新失敗")
        End If
        conn.Close()
    End Sub
End Class