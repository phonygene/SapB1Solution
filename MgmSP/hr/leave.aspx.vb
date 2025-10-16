Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class leave
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap As SqlDataReader
    Public TxtSDate, TxtEDate As TextBox
    Public DDLID As DropDownList
    Public BtnFilter, BtnAdd As Button
    Public ds As New DataSet
    Public permshr100 As String
    Public ScriptManager1 As New ScriptManager
    Public fmode As String
    'Public LBUserList, LBALBHour, LBALEHour, LBALBMin, LBALEMin As New ListBox
    Public RBSearchType As RadioButtonList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim rule, dstr, str() As String
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        permshr100 = CommUtil.GetAssignRight("hr100", Session("s_id"))
        fmode = Request.QueryString("fmode")
        FTCreate() '要置放在判斷controlobj之前 , 先重新create物件後 , 才能catch到
        CreateUpdItem() '同上
        Dim controlobj As Control
        controlobj = CommUtil.GetPostBackControl(Page) ' or CommUtil.GetPostBackControl(sender)
        If (controlobj IsNot Nothing) Then
            str = Split(controlobj.ID, "_")
            If (controlobj.ID = "btn_filter" Or controlobj.ID = "btn_action" Or controlobj.ID = "btn_cancel" Or controlobj.ID = "rbl_searchtype") Then
                fmode = "show"
            ElseIf (str(0) = "lb") Then
                If (FTDDL.SelectedIndex = 3) Then
                    fmode = "add"
                Else
                    fmode = "modify"
                End If
            ElseIf (controlobj.ID = "FTDDL") Then
                If (FTDDL.SelectedIndex = 3) Then
                    fmode = "add"
                Else
                    fmode = "show"
                End If
            Else
                If (FTDDL.SelectedIndex = 3) Then
                    fmode = "add"
                Else
                    fmode = "modify"
                End If
            End If
        End If
        'MsgBox(fmode)
        If (FTDDL.SelectedIndex = 2) Then
            FilterT.Enabled = True
        Else
            FilterT.Enabled = False
        End If

        If (fmode = "show") Then
            FTDDL.Visible = True
            AddT.Visible = False
            FilterT.Visible = True
            UpdT.Visible = False
            gv1.Visible = True
        Else
            FTDDL.Visible = False
            FilterT.Visible = False
            AddT.Visible = False
            gv1.Visible = False
            UpdT.Visible = True
        End If
        WriteFilterCombo()
        If (Not IsPostBack) Then
            InsertUserList()
            dstr = Format(Now(), "yyyy/MM/dd")
            rule = " where ((albdate<='" & dstr & "' and aledate>='" & dstr & "') or (" &
                        "aledate>='" & dstr & "' and aledate <='" & dstr & "') or (" &
                        "albdate>='" & dstr & "' and albdate <='" & dstr & "'))"
            WriteSelectItemToGridView(rule)
        Else

        End If
        If (fmode = "modify" Or fmode = "add") Then
            ShowUpdItemData()
        End If
    End Sub
    Sub WriteFilterCombo()
        Dim i As Integer
        Dim LBx As ListBox

        SqlCmd = "Select T0.id,T0.name From dbo.[User] T0 order by branch,grp"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            LBx = UpdT.FindControl("lb_user")
            LBx.Items.Clear()
            LBx.Items.Add("人員")
            Do While (dr.Read())
                LBx.Items.Add(dr(0) & " " & dr(1))
            Loop
        End If
        dr.Close()
        conn.Close()

        LBx = UpdT.FindControl("lb_altype")
        LBx.Items.Clear()
        LBx.Items.Add("選假別")
        LBx.Items.Add("事假")
        LBx.Items.Add("特休")
        LBx.Items.Add("病假")
        LBx.Items.Add("喪假")
        LBx.Items.Add("婚假")
        LBx.Items.Add("公假")
        LBx.Items.Add("其他")
        LBx.SelectedIndex = 0

        LBx = UpdT.FindControl("lb_albhour")
        LBx.Items.Clear()
        LBx.Items.Add("填幾點")
        For i = 0 To 23
            LBx.Items.Add(i)
        Next
        LBx.SelectedIndex = 0

        LBx = UpdT.FindControl("lb_albmin")
        LBx.Items.Clear()
        LBx.Items.Add("填幾分")
        LBx.Items.Add("0")
        LBx.Items.Add("30")
        LBx.SelectedIndex = 0

        LBx = UpdT.FindControl("lb_alehour")
        LBx.Items.Clear()
        LBx.Items.Add("填幾點")
        For i = 0 To 23
            LBx.Items.Add(i)
        Next
        LBx.SelectedIndex = 0

        LBx = UpdT.FindControl("lb_alemin")
        LBx.Items.Clear()
        LBx.Items.Add("填幾分")
        LBx.Items.Add("0")
        LBx.Items.Add("30")
        LBx.SelectedIndex = 0

    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        'Dim Hyper As New HyperLink
        Dim Hyper As New LinkButton
        If (e.Row.RowType = DataControlRowType.DataRow) Then

            'MsgBox(ds.Tables(0).Rows(realindex)("altype"))
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            e.Row.Cells(2).Text = TransALSymbolToReal(ds.Tables(0).Rows(realindex)("altype"))
            Hyper.Text = e.Row.Cells(0).Text
            Hyper.PostBackUrl = "~/hr/leave.aspx?smid=hr&smode=1&fmode=modify&num=" & e.Row.Cells(0).Text
            'Hyper.NavigateUrl = "~/hr/leave.aspx?smid=hr&smode=1&fmode=modify&num=" & e.Row.Cells(0).Text
            e.Row.Cells(0).Controls.Add(Hyper)
        End If
    End Sub
    Sub FTCreate()
        FilterTCreate()
    End Sub
    Protected Sub RBSearchType_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        'GetActionPara()
        If (RBSearchType.SelectedIndex = 0) Then
            TxtSDate.Enabled = False
            TxtEDate.Enabled = False
            TxtSDate.Text = ""
            TxtEDate.Text = ""
        Else
            TxtSDate.Enabled = True
            TxtEDate.Enabled = True
        End If
    End Sub
    Sub FilterTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim ce As CalendarExtender
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Left
        'Labelx = New Label()
        'Labelx.ID = "label_rblist"
        'Labelx.Text = "&nbsp&nbsp&nbsp&nbsp:"
        'tCell.Controls.Add(Labelx)
        RBSearchType = New RadioButtonList()
        RBSearchType.ID = "rbl_searchtype"
        RBSearchType.Items.Add("已請未休")
        RBSearchType.Items.Add("全部")
        RBSearchType.Items(0).Value = 1
        RBSearchType.Items(1).Value = 2
        RBSearchType.Font.Size = 10
        'RBSearchType.Width = 100
        RBSearchType.RepeatDirection = RepeatDirection.Vertical
        RBSearchType.SelectedIndex = 0
        RBSearchType.AutoPostBack = True
        'CommUtil.DisableObjectByPermission(RBSearchType, permshr100, "m")
        AddHandler RBSearchType.SelectedIndexChanged, AddressOf RBSearchType_SelectedIndexChanged
        tCell.Controls.Add(RBSearchType)
        tRow.Cells.Add(tCell)
        '-----------------------------------------
        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Left
        DDLID = New DropDownList()
        DDLID.ID = "ddl_id"
        'DDLID.Width = 600
        tCell.Controls.Add(DDLID)

        Labelx = New Label()
        Labelx.ID = "label_sdate"
        Labelx.Text = "起始日期:"
        tCell.Controls.Add(Labelx)
        TxtSDate = New TextBox()
        TxtSDate.ID = "txt_sdate"
        TxtSDate.Width = 100
        'tCell.Controls.Add(TxtSDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtSDate.ID
        ce.ID = "ce_begindate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(TxtSDate)

        Labelx = New Label()
        Labelx.ID = "label_edate"
        Labelx.Text = "結束日期:"
        tCell.Controls.Add(Labelx)
        TxtEDate = New TextBox()
        TxtEDate.ID = "txt_edate"
        TxtEDate.Width = 100
        'tCell.Controls.Add(TxtEDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtEDate.ID
        ce.ID = "ce_enddate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(TxtEDate)
        'tRow.Cells.Add(tCell)
        '-----------------------------------------
        'tCell = New TableCell()
        'tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label()
        Labelx.ID = "label_adds1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnFilter = New Button()
        BtnFilter.ID = "btn_filter"
        BtnFilter.Text = "篩選"
        AddHandler BtnFilter.Click, AddressOf BtnFilter_Click
        tCell.Controls.Add(BtnFilter)
        tRow.Cells.Add(tCell)
        FilterT.Rows.Add(tRow)

    End Sub

    Sub InsertUserList()
        SqlCmd = "Select T0.id,T0.name From dbo.[User] T0"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            DDLID.Items.Clear()
            DDLID.Items.Add("請選擇員工")
            Do While (dr.Read())
                DDLID.Items.Add(dr(0) & " " & dr(1))
            Loop
        End If
        dr.Close()
        conn.Close()
    End Sub
    Protected Sub BtnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (RBSearchType.SelectedIndex <> 0) Then
            If ((TxtSDate.Text <> "" And TxtEDate.Text <> "") Or DDLID.SelectedIndex <> 0) Then
                'gv1.Visible = True
                FilterALSearch()
            Else
                CommUtil.ShowMsg(Me, "需設定日期區間或選擇人員")
            End If
        Else
            FilterALSearch() '以今天日期為主 , 不須設定日期範圍
        End If
    End Sub

    Sub FilterALSearch()
        Dim rule As String
        Dim id, begindate, enddate As String
        Dim filterflag As Boolean
        Dim str() As String
        filterflag = False
        rule = " where "
        If (DDLID.SelectedIndex <> 0) Then
            str = Split(DDLID.SelectedValue, " ")
            id = str(0)
            If (filterflag = True) Then
                rule = rule & " and id='" & id & "' "
            Else
                rule = rule & " id='" & id & "' "
            End If
            filterflag = True
        End If
        'MsgBox(RBSearchType.SelectedIndex)
        If (RBSearchType.SelectedIndex = 0) Then
            begindate = Format(Now(), "yyyy/MM/dd")
            If (filterflag = True) Then
                rule = rule & " and aledate>='" & begindate & "'"
            Else
                rule = rule & " aledate>='" & begindate & "'"
            End If
            filterflag = True
        Else
            If (TxtSDate.Text <> "" And TxtEDate.Text <> "") Then
                begindate = TxtSDate.Text
                enddate = TxtEDate.Text
                If (filterflag = True) Then
                    rule = rule & " and ((albdate<='" & begindate & "' and aledate>='" & enddate & "') or (" &
                        "aledate>='" & begindate & "' and aledate <='" & enddate & "') or (" &
                        "albdate>='" & begindate & "' and albdate <='" & enddate & "'))"
                Else
                    rule = rule & " ((albdate<='" & begindate & "' and aledate>='" & enddate & "') or (" &
                        "aledate>='" & begindate & "' and aledate <='" & enddate & "') or (" &
                        "albdate>='" & begindate & "' and albdate <='" & enddate & "'))"
                End If
                filterflag = True
            End If
        End If

        If (filterflag = False) Then
            rule = ""
        End If
        rule = rule & " order by albdate"
        WriteSelectItemToGridView(rule)
    End Sub

    Sub WriteSelectItemToGridView(rule As String) ' not finish
        'MsgBox(rule)
        'gv1.Visible = True
        ds.Reset()
        SetGridViewStyle()
        If (FTDDL.SelectedIndex = 0 Or FTDDL.SelectedIndex = 1 Or FTDDL.SelectedIndex = 2) Then
            SetShowALGridViewFields()
            SqlCmd = "SELECT T0.num,T0.idname,T0.altype,convert(char(12),T0.cdate,111) as cdate,convert(char(12),T0.albdate,111) as albdate, " &
                     "T0.albhour,T0.albmin,convert(char(12),T0.aledate,111) as aledate, " &
                     "T0.alehour,T0.alemin,T0.sname,T0.alreason " &
            "FROM dbo.[@HALRT] T0 " & rule
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()
            gv1.DataSource = ds.Tables(0)
            gv1.DataBind()
        ElseIf (FTDDL.SelectedIndex = 3) Then

        End If
    End Sub
    Sub SetGridViewStyle() ' not finish
        gv1.AutoGenerateColumns = False
        'gv1.ShowHeader = True
        gv1.AllowPaging = True
        'gv1.AllowSorting = True
        gv1.PageSize = 14
        'gv1.Font.Size = FontSize.Smaller
        'gv1.ForeColor =
        gv1.GridLines = GridLines.Both

        gv1.PagerStyle.HorizontalAlign = HorizontalAlign.Center

        'gv1.HeaderStyle.BackColor =
        'gv1.RowStyle.BackColor
        'gv1.AlternatingRowStyle.BackColor
        'gv1.HeaderStyle.ForeColor
    End Sub

    Sub SetShowALGridViewFields() ' not finish
        Dim oBoundField As BoundField
        'Dim oHyperLinkField As HyperLinkField
        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "序號"
        'oHyperLinkField.DataTextField = "num"
        oBoundField.DataField = "num"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        'oHyperLinkField.NavigateUrl = "~/hr/leave.aspx?smid=hr&smode=1&fmode=modify"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "姓名"
        oBoundField.DataField = "idname"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 50
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "假別"
        oBoundField.DataField = "altype"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "請假日期"
        oBoundField.DataField = "albdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 80
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "時"
        oBoundField.DataField = "albhour"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "分"
        oBoundField.DataField = "albmin"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "結束日期"
        oBoundField.DataField = "aledate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 80
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "時"
        oBoundField.DataField = "alehour"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "分"
        oBoundField.DataField = "alemin"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "填寫人"
        oBoundField.DataField = "sname"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 50
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "建立時間"
        oBoundField.DataField = "cdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 80
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "請假原因"
        oBoundField.DataField = "alreason"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)
    End Sub
    Function CellSet(text As String, title As Integer)
        Dim tCell As New TableCell
        'tCell.HorizontalAlign = HorizontalAlign.Center
        If (title = 1) Then
            tCell.BackColor = Drawing.Color.DeepSkyBlue
            tCell.Text = text
        Else
            tCell.Width = 300
        End If
        tCell.Font.Size = 10
        tCell.BorderWidth = 1
        tCell.Wrap = True
        Return tCell
    End Function
    Function CellSetWithTB(Txtxid As String)
        Dim tCell As New TableCell
        Dim Txtx As New TextBox
        Txtx.ID = Txtxid
        Txtx.Width = 300
        tCell.Controls.Add(Txtx)
        Return tCell
    End Function

    Function CellSetWithExtender(LBxid As String, Txtxid As String, ddeid As String)
        Dim tCell As New TableCell
        Dim dde As New DropDownExtender
        Dim Txtx As New TextBox
        Dim LBx As New ListBox

        LBx.ID = LBxid
        LBx.AutoPostBack = True
        LBx.Rows = 30
        AddHandler LBx.SelectedIndexChanged, AddressOf LB_SelectedIndexChanged
        tCell.Controls.Add(LBx)
        Txtx.ID = Txtxid
        Txtx.Width = 300
        tCell.Controls.Add(Txtx)
        dde.TargetControlID = Txtxid
        dde.ID = ddeid
        dde.DropDownControlID = LBxid
        tCell.Controls.Add(dde)
        Return tCell
    End Function
    Protected Sub LB_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Txtx As TextBox
        Dim str() As String
        str = Split(sender.ID, "_")
        Txtx = UpdT.FindControl("txt_" & str(1))
        Txtx.Text = sender.SelectedValue

        FilterT.Visible = False
        UpdT.Visible = True
        'MsgBox("LBS")
        gv1.Visible = False
    End Sub

    Protected Sub FTDDL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FTDDL.SelectedIndexChanged
        Dim dstr, rule As String
        If (FTDDL.SelectedIndex <> 3) Then
            'FTDDL.Visible = True
            'FilterT.Visible = True
            'UpdT.Visible = False
            'gv1.Visible = True
            If (FTDDL.SelectedIndex = 2) Then
                TxtSDate.Enabled = False
                TxtEDate.Enabled = False
            End If
            dstr = "1900/01/01"
            rule = " where albdate>='" & dstr & "' and aledate<='" & dstr & "'"
            If (FTDDL.SelectedIndex = 0) Then
                dstr = Format(Now(), "yyyy/MM/dd")
                rule = " where ((albdate<='" & dstr & "' and aledate>='" & dstr & "') or (" &
                            "aledate>='" & dstr & "' and aledate <='" & dstr & "') or (" &
                            "albdate>='" & dstr & "' and albdate <='" & dstr & "'))"
            ElseIf (FTDDL.SelectedIndex = 1) Then
                dstr = Format(Now.AddDays(1), "yyyy/MM/dd") '如是昨天則用Now.AddDays(-1)
                rule = " where ((albdate<='" & dstr & "' and aledate>='" & dstr & "') or (" &
                            "aledate>='" & dstr & "' and aledate <='" & dstr & "') or (" &
                            "albdate>='" & dstr & "' and albdate <='" & dstr & "'))"
            End If
            WriteSelectItemToGridView(rule)
        Else
            'FilterT.Visible = False
            'UpdT.Visible = True
            'gv1.Visible = False
            IniUpdField()
        End If
    End Sub

    Function CellSetWithCalenderExtender(Txtxid As String, ceid As String)
        Dim tCell As New TableCell
        Dim ce As New CalendarExtender
        Dim Txtx As New TextBox

        Txtx.ID = Txtxid
        Txtx.Width = 300
        tCell.Controls.Add(Txtx)
        ce.TargetControlID = Txtxid
        ce.ID = ceid
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        Return tCell
    End Function
    Sub CreateUpdItem()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Btnx As Button

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("序號", 1))
        tRow.Cells.Add(CellSet("", 0))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("人員", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_user", "txt_user", "dde_user"))
        UpdT.Controls.Add(tRow)
        'If (fmode = "modify") Then
        '    CType(UpdT.FindControl("lb_user"), ListBox).Enabled = False
        '    CType(UpdT.FindControl("txt_user"), TextBox).Enabled = False
        'ElseIf (fmode = "add") Then
        '    CType(UpdT.FindControl("lb_user"), ListBox).Enabled = True
        '    CType(UpdT.FindControl("txt_user"), TextBox).Enabled = True
        'End If
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("假別", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_altype", "txt_altype", "dde_altype"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("建立日期", 1))
        tRow.Cells.Add(CellSet("", 0))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("開始日期", 1))
        tRow.Cells.Add(CellSetWithCalenderExtender("txt_albdate", "ce_albdate"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("開始小時", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_albhour", "txt_albhour", "dde_albhour"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("開始分", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_albmin", "txt_albmin", "dde_albmin"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("結束日期", 1))
        tRow.Cells.Add(CellSetWithCalenderExtender("txt_aledate", "ce_aledate"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("結束小時", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_alehour", "txt_alehour", "dde_alehour"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("結束分", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_alemin", "txt_alemin", "dde_alemin"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("原因", 1))
        tRow.Cells.Add(CellSetWithTB("txt_alreason"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.BackColor = Drawing.Color.AntiqueWhite
        tRow.HorizontalAlign = HorizontalAlign.Center
        tRow.Font.Bold = True
        Dim Labelx As Label

        Dim Chkx As New CheckBox
        tCell = New TableCell
        tCell.ColumnSpan = 2
        tCell.BorderWidth = 1
        Chkx.ID = "chk_action"
        'If (fmode = "modify") Then
        '    Chkx.Text = "選取後執行刪除"
        '    'CommUtil.DisableObjectByPermission(Chkx, permssa100, "d")
        'ElseIf (fmode = "add") Then
        '    Chkx.Text = "新增確認"
        'End If
        AddHandler Chkx.CheckedChanged, AddressOf Chkx_CheckedChanged
        Chkx.AutoPostBack = True
        Labelx = New Label
        Labelx.ID = "label2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Btnx = New Button
        Btnx.ID = "btn_action"
        'If (fmode = "modify") Then
        '    Btnx.Text = "更新"
        'ElseIf (fmode = "add") Then
        '    Btnx.Text = "新增"
        '    Btnx.Enabled = False
        'End If
        AddHandler Btnx.Click, AddressOf BtnxAction_Click
        tCell.Controls.Add(Btnx)
        Labelx = New Label
        Labelx.ID = "label1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Btnx = New Button
        Btnx.ID = "btn_cancel"
        Btnx.Text = "取消"
        AddHandler Btnx.Click, AddressOf BtnxCancel_Click
        tCell.Controls.Add(Btnx)

        tCell.Controls.Add(Chkx)
        tRow.Controls.Add(tCell)
        UpdT.Controls.Add(tRow)
    End Sub
    Protected Sub Chkx_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Btnx As Button
        Btnx = UpdT.FindControl("btn_action")
        If (fmode = "modify") Then
            If (sender.Checked = False) Then
                Btnx.Text = "更新"
                Btnx.OnClientClick = ""
            Else
                Btnx.Text = "刪除"
                Btnx.OnClientClick = "return confirm('要刪除嗎')"
            End If
        ElseIf (fmode = "add") Then
            If (sender.Checked) Then
                Btnx.Enabled = True
                'Btnx.OnClientClick = "return confirm('要新增嗎')"
            Else
                Btnx.Enabled = False
                Btnx.OnClientClick = ""
            End If
        End If
    End Sub
    Protected Sub BtnxAction_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Chkx As CheckBox
        Dim actionok As Boolean
        If (FTDDL.SelectedIndex = 3) Then
            fmode = "add"
        Else
            fmode = "modify"
        End If
        If (RecordFieldCheck()) Then
            Chkx = UpdT.FindControl("chk_action")
            If (fmode = "add") Then
                If (Chkx.Checked) Then
                    actionok = InsertFCRecord()
                Else
                    CommUtil.ShowMsg(Me, "新增check沒打勾")
                    Exit Sub
                End If
            ElseIf (fmode = "modify") Then
                If (Chkx.Checked) Then 'delete
                    actionok = DeleteFCRecord()
                Else 'update
                    actionok = UpdateFCRecord()
                End If
            End If
            If (actionok) Then
                If (FTDDL.SelectedIndex = 2) Then
                    FilterALSearch()
                    FTDDL.Visible = True
                    AddT.Visible = False
                    FilterT.Visible = True
                    UpdT.Visible = False
                    gv1.Visible = True
                Else
                    Response.Redirect("~/hr/leave.aspx?smid=hr&smode=1&fmode=show")
                End If
            End If
        Else
            CommUtil.ShowMsg(Me, "有欄位空白")
        End If
    End Sub
    Function RecordFieldCheck()
        RecordFieldCheck = True
        If (CType(UpdT.FindControl("txt_user"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_altype"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_albdate"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_albhour"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_albmin"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_aledate"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_alehour"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_alemin"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_alreason"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
    End Function
    Function UpdateFCRecord()
        Dim num As Long
        Dim altype, albdate, albhour, albmin, aledate, alehour, alemin, alreason
        num = CLng(UpdT.Rows(0).Cells(1).Text)
        altype = TransALRealToSymbol(CType(UpdT.FindControl("txt_altype"), TextBox).Text)
        albdate = CType(UpdT.FindControl("txt_albdate"), TextBox).Text
        albhour = CType(UpdT.FindControl("txt_albhour"), TextBox).Text
        albmin = CInt(CType(UpdT.FindControl("txt_albmin"), TextBox).Text)
        aledate = CType(UpdT.FindControl("txt_aledate"), TextBox).Text
        alehour = CType(UpdT.FindControl("txt_alehour"), TextBox).Text
        alemin = CType(UpdT.FindControl("txt_alemin"), TextBox).Text
        alreason = CType(UpdT.FindControl("txt_alreason"), TextBox).Text
        SqlCmd = "update dbo.[@HALRT] set " &
        "altype='" & altype & "', albdate= '" & albdate & "' , albhour= '" & albhour & "', " &
        "albmin= '" & albmin & "' , aledate= '" & aledate & "', " &
        "alehour='" & alehour & "' , alemin= '" & alemin & "', alreason = '" & alreason & "' " &
        "where num = " & num
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
        UpdateFCRecord = True
    End Function
    Function InsertFCRecord()
        Dim num As Long
        Dim nullflag As Boolean
        SqlCmd = "SELECT max(num) FROM dbo.[@HALRT]"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        If (Not IsDBNull(dr(0))) Then
            num = dr(0) + 1
            nullflag = False
        Else
            nullflag = True
        End If
        dr.Close()
        connsap.Close()
        Dim idstr(), id, idname, altype, albdate, albhour, albmin, aledate, alehour, alemin, alreason, createdate, cid, cname
        cid = Session("s_id")
        cname = Session("s_name")
        createdate = UpdT.Rows(3).Cells(1).Text
        idstr = Split(CType(UpdT.FindControl("txt_user"), TextBox).Text, " ")
        id = idstr(0)
        idname = idstr(1)
        altype = TransALRealToSymbol(CType(UpdT.FindControl("txt_altype"), TextBox).Text)
        albdate = CType(UpdT.FindControl("txt_albdate"), TextBox).Text
        albhour = CType(UpdT.FindControl("txt_albhour"), TextBox).Text
        albmin = CInt(CType(UpdT.FindControl("txt_albmin"), TextBox).Text)
        aledate = CType(UpdT.FindControl("txt_aledate"), TextBox).Text
        alehour = CType(UpdT.FindControl("txt_alehour"), TextBox).Text
        alemin = CType(UpdT.FindControl("txt_alemin"), TextBox).Text
        alreason = CType(UpdT.FindControl("txt_alreason"), TextBox).Text
        SqlCmd = "insert into [dbo].[@HALRT] (id,idname,altype,albdate,albhour,albmin,aledate,alehour,alemin,cdate,alreason,sempid,sname) " &
        "values('" & id & "','" & idname & "','" & altype & "','" & albdate & "','" & albhour & "','" & albmin & "','" & aledate & "','" & alehour & "','" & alemin & "', " &
                "'" & createdate & "','" & alreason & "','" & cid & "','" & cname & "')"
        '讀回,check 是否已寫入
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()

        If (nullflag) Then
            SqlCmd = "SELECT max(num) FROM dbo.[@HALRT]"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (IsDBNull(dr(0))) Then
                CommUtil.ShowMsg(Me, "資料沒寫入")
                InsertFCRecord = False
            Else
                InsertFCRecord = True
            End If
            dr.Close()
            connsap.Close()
        Else
            SqlCmd = "SELECT altype FROM dbo.[@HALRT] where num=" & num
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr.HasRows) Then
                InsertFCRecord = True
            Else
                CommUtil.ShowMsg(Me, "資料沒寫入")
                InsertFCRecord = False
            End If
            dr.Close()
            connsap.Close()
        End If
    End Function
    Function DeleteFCRecord()
        Dim num As Long
        Dim connsap1 As New SqlConnection
        'MsgBox(DateDiff(DateInterval.Day, CDate(UpdT.Rows(6).Cells(1).Text), Now()))
        'Exit Sub
        DeleteFCRecord = True
        'If (DateDiff(DateInterval.Day, CDate(UpdT.Rows(6).Cells(1).Text), Now()) > 6) Then
        '    CommUtil.ShowMsg(Me, "此單已建立超過6天 ,不能刪除 , 請將狀態改為已取消即可")
        '    DeleteFCRecord = False
        '    Exit Function
        'Else
        '    ucode = UpdT.Rows(0).Cells(1).Text
        'End If
        num = CLng(UpdT.Rows(0).Cells(1).Text)
        SqlCmd = "SELECT count(T0.num) " &
        "FROM dbo.[@HALRT] T0 " &
        "where T0.num=" & num
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        If (dr(0) = 0) Then
            CommUtil.ShowMsg(Me, "單號已被他人刪除")
            DeleteFCRecord = False
        Else
            SqlCmd = "delete from dbo.[@HALRT] " &
                    "where num = " & num
            DeleteFCRecord = CommUtil.SqlSapExecute("del", SqlCmd, connsap1)
            connsap1.Close()
        End If
        dr.Close()
        connsap.Close()
    End Function
    Protected Sub BtnxCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (FTDDL.SelectedIndex = 2) Then
            FilterALSearch()
            FTDDL.Visible = True
            AddT.Visible = False
            FilterT.Visible = True
            UpdT.Visible = False
            gv1.Visible = True
        ElseIf (FTDDL.SelectedIndex = 3) Then
            Response.Redirect("~/hr/leave.aspx?smid=hr&smode=1&fmode=show")
        End If
    End Sub
    Function TransALSymbolToReal(altype As String)
        Dim altypestr As String
        If (altype = "T") Then
            altypestr = "事假"
        ElseIf (altype = "S") Then
            altypestr = "特休"
        ElseIf (altype = "K") Then
            altypestr = "病假"
        ElseIf (altype = "D") Then
            altypestr = "喪假"
        ElseIf (altype = "M") Then
            altypestr = "婚假"
        ElseIf (altype = "P") Then
            altypestr = "公假"
        ElseIf (altype = "O") Then
            altypestr = "其它"
        Else
            altypestr = ""
        End If
        Return altypestr
    End Function

    Function TransALRealToSymbol(altypestr As String)
        Dim altype As String
        If (altypestr = "事假") Then
            altype = "T"
        ElseIf (altypestr = "特休") Then
            altype = "S"
        ElseIf (altypestr = "病假") Then
            altype = "K"
        ElseIf (altypestr = "喪假") Then
            altype = "D"
        ElseIf (altypestr = "婚假") Then
            altype = "M"
        ElseIf (altypestr = "公假") Then
            altype = "P"
        ElseIf (altypestr = "其它") Then
            altype = "O"
        Else
            altype = ""
        End If
        Return altype
    End Function
    Sub ShowUpdItemData()
        Dim num As Long
        Dim id, idname, altype, altypestr, builtdate, startdate, bhour, bmin, enddate, ehour, emin, sname, reason As String
        Dim Chkx As CheckBox
        Chkx = UpdT.FindControl("chk_action")
        If (fmode = "modify") Then
            CType(UpdT.FindControl("lb_user"), ListBox).Enabled = False
            CType(UpdT.FindControl("txt_user"), TextBox).Enabled = False
            CType(UpdT.FindControl("chk_action"), CheckBox).Text = "選取後執行刪除"
            If (Not Chkx.Checked) Then
                CType(UpdT.FindControl("btn_action"), Button).Text = "更新"
            Else
                CType(UpdT.FindControl("btn_action"), Button).Text = "刪除"
            End If
        ElseIf (fmode = "add") Then
            CType(UpdT.FindControl("lb_user"), ListBox).Enabled = True
            CType(UpdT.FindControl("txt_user"), TextBox).Enabled = True
            CType(UpdT.FindControl("chk_action"), CheckBox).Text = "新增確認"
            CType(UpdT.FindControl("btn_action"), Button).Text = "新增"
            CType(UpdT.FindControl("btn_action"), Button).Enabled = False
        End If
        num = Request.QueryString("num")
        If (fmode = "modify") Then
            SqlCmd = "SELECT T0.idname,T0.altype,convert(char(12),T0.cdate,111) as cdate,convert(char(12),T0.albdate,111) as albdate, " &
                     "T0.albhour,T0.albmin,convert(char(12),T0.aledate,111) as aledate, " &
                     "T0.alehour,T0.alemin,T0.sname,T0.alreason,T0.id " &
                     "FROM dbo.[@HALRT] T0 where num=" & num
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            'If (dr.HasRows) Then
            dr.Read()
            id = dr(11)
            idname = dr(0)
            altype = dr(1)
            builtdate = dr(2)
            startdate = dr(3)
            bhour = dr(4)
            bmin = dr(5)
            enddate = dr(6)
            ehour = dr(7)
            emin = dr(8)
            sname = dr(9)
            reason = dr(10)
            altypestr = TransALSymbolToReal(altype)
            'End If
            dr.Close()
            connsap.Close()
            UpdT.Rows(0).Cells(1).Text = num
            CType(UpdT.FindControl("txt_user"), TextBox).Text = id & " " & idname
            CType(UpdT.FindControl("txt_altype"), TextBox).Text = altypestr
            'If (fmode = "modify") Then
            UpdT.Rows(3).Cells(1).Text = builtdate
            'ElseIf (fmode = "add") Then
            'UpdT.Rows(3).Cells(1).Text = Format(Now, "yyyy/MM/dd")
            'End If
            CType(UpdT.FindControl("txt_albdate"), TextBox).Text = startdate
            CType(UpdT.FindControl("txt_albhour"), TextBox).Text = bhour
            CType(UpdT.FindControl("txt_albmin"), TextBox).Text = bmin
            CType(UpdT.FindControl("txt_aledate"), TextBox).Text = enddate
            CType(UpdT.FindControl("txt_alehour"), TextBox).Text = ehour
            CType(UpdT.FindControl("txt_alemin"), TextBox).Text = emin
            CType(UpdT.FindControl("txt_alreason"), TextBox).Text = reason
        ElseIf (fmode = "add") Then
            'UpdT.Rows(3).Cells(1).Text = Format(Now, "yyyy/MM/dd")
        End If
    End Sub
    Sub IniUpdField()
        UpdT.Rows(0).Cells(1).Text = ""
        CType(UpdT.FindControl("txt_user"), TextBox).Text = ""
        CType(UpdT.FindControl("txt_altype"), TextBox).Text = ""
        CType(UpdT.FindControl("txt_albdate"), TextBox).Text = ""
        CType(UpdT.FindControl("txt_albhour"), TextBox).Text = "08"
        CType(UpdT.FindControl("txt_albmin"), TextBox).Text = "30"
        CType(UpdT.FindControl("txt_aledate"), TextBox).Text = ""
        CType(UpdT.FindControl("txt_alehour"), TextBox).Text = "17"
        CType(UpdT.FindControl("txt_alemin"), TextBox).Text = "30"
        CType(UpdT.FindControl("txt_alreason"), TextBox).Text = ""
        UpdT.Rows(3).Cells(1).Text = Format(Now, "yyyy/MM/dd")
    End Sub
End Class