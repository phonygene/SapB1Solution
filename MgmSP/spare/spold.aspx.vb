Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class spold
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap As SqlDataReader
    Public permsp200 As String
    Public ScriptManager1 As New ScriptManager
    Public fmode, action As String
    Public ds As New DataSet
    Public BtnSPM As Button
    Public TxtKW, TxtCNo, TxtBeginDate, TxtEndDate As TextBox
    Public BtnSearch, BtnCNoSearch As Button
    Public DDLWhs As DropDownList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim str() As String
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        permsp200 = CommUtil.GetAssignRight("sp200", Session("s_id"))
        CommUtil.DisableObjectByPermission(FTDDL, permsp200, "n")
        fmode = Request.QueryString("fmode")
        FTCreate()
        CreateAddItem()
        Dim controlobj As Control
        controlobj = CommUtil.GetPostBackControl(Page) ' or CommUtil.GetPostBackControl(sender)
        If (controlobj IsNot Nothing) Then
            str = Split(controlobj.ID, "_")
            If (controlobj.ID = "FTDDL") Then
                If (FTDDL.SelectedIndex = 0) Then
                    fmode = "show"
                ElseIf (FTDDL.SelectedIndex = 1) Then
                    fmode = "add"
                End If
            ElseIf (controlobj.ID = "btn_action" Or controlobj.ID = "btn_cancel" Or controlobj.ID = "btn_del") Then
                If (FTDDL.SelectedIndex = 1) Then
                    fmode = "add"
                Else

                End If
                action = fmode
                fmode = "show"
            ElseIf (controlobj.ID = "btn_spm" Or controlobj.ID = "btn_search") Then
                fmode = "show"
            ElseIf (str(0) = "lb") Then
                If (FTDDL.SelectedIndex = 1) Then
                    fmode = "add"
                Else
                    'fmode = "modify" 原本之fmode = Request.QueryString("fmode") 就是對的
                End If
                'ElseIf (controlobj.ID = "chk_action") Then

            Else
                If (FTDDL.SelectedIndex = 1) Then
                    fmode = "add"
                Else
                    'fmode = "modify" 原本之fmode = Request.QueryString("fmode") 就是對的
                End If
            End If
        End If
        WriteLBCombo()
        If (Not IsPostBack) Then
            TxtBeginDate.Text = "2023/01/01"
            TxtEndDate.Text = Format(Now, "yyyy/MM/dd")
            WriteSelectItemToGridView(1)
        Else

        End If
        If (fmode = "show" Or fmode = "showpost") Then
            FTDDL.Visible = True
            AddT.Visible = False
            FilterT.Visible = True
            gv1.Visible = True
        ElseIf (fmode = "add" Or fmode = "inout" Or fmode = "modify") Then
            FTDDL.Visible = False
            AddT.Visible = True
            FilterT.Visible = False
            gv1.Visible = False
        End If

        If (fmode = "add" Or fmode = "inout" Or fmode = "modify") Then
            ShowAddForm()
        End If
        'MsgBox(TxtBeginDate.Text & " " & TxtEndDate.Text)
        If (fmode = "showpost") Then
            MaterialInOutPosted()
        ElseIf (fmode = "show") Then
            'MsgBox(TxtBeginDate.Text)
            'WriteSelectItemToGridView(1)
        End If
    End Sub
    Sub WriteLBCombo()
        Dim LBx As ListBox
        LBx = AddT.FindControl("lb_whs")
        LBx.Items.Clear()
        LBx.Items.Add("C01")
        LBx.Items.Add("C02")
        LBx.SelectedIndex = -1

        LBx = AddT.FindControl("lb_dtype")
        LBx.Items.Clear()
        LBx.Items.Add("收入")
        LBx.Items.Add("發出")
        LBx.SelectedIndex = -1

        AddWhsItem()
    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        Dim LKHyper As LinkButton
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            If (fmode = "show") Then
                LKHyper = New LinkButton
                LKHyper.Text = "異動單建立"
                LKHyper.Font.Underline = False
                LKHyper.PostBackUrl = "~/spare/spold.aspx?smid=sp&smode=2&fmode=inout&itemcode=" & e.Row.Cells(1).Text & "&itemname=" & e.Row.Cells(2).Text &
                                  "&onhand=" & e.Row.Cells(3).Text
                CommUtil.DisableObjectByPermission(LKHyper, permsp200, "n")
                e.Row.Cells(0).Controls.Add(LKHyper)

                LKHyper = New LinkButton
                LKHyper.Text = e.Row.Cells(1).Text
                LKHyper.Font.Underline = False
                LKHyper.PostBackUrl = "~/spare/spold.aspx?smid=sp&smode=2&fmode=modify&itemcode=" & e.Row.Cells(1).Text
                'LKHyper.PostBackUrl = "~/spare/spold.aspx?smid=sp&smode=2&fmode=inout&num=" & e.Row.Cells(1).Text
                CommUtil.DisableObjectByPermission(LKHyper, permsp200, "m")
                e.Row.Cells(1).Controls.Add(LKHyper)

                LKHyper = New LinkButton
                LKHyper.Text = e.Row.Cells(7).Text
                LKHyper.Font.Underline = False
                LKHyper.PostBackUrl = "~/spare/spold.aspx?smid=sp&smode=2&fmode=showpost&itemcode=" & e.Row.Cells(1).Text &
                                  "&begindate=" & TxtBeginDate.Text & "&enddate=" & TxtEndDate.Text
                e.Row.Cells(7).Controls.Add(LKHyper)
                '
                SqlCmd = "SELECT itemname,onhand FROM OITM where itemcode='" & e.Row.Cells(1).Text & "'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    e.Row.Cells(2).Text = dr(0)
                    e.Row.Cells(4).Text = CInt(dr(1))
                Else
                    e.Row.Cells(4).Text = "NA"
                End If
                dr.Close()
                connsap.Close()
            ElseIf (fmode = "showpost") Then
                e.Row.Cells(0).Text = realindex + 1
                If (e.Row.Cells(3).Text = 0) Then
                    e.Row.Cells(3).Text = ""
                End If
                If (e.Row.Cells(4).Text = 0) Then
                    e.Row.Cells(4).Text = ""
                End If
            End If

        End If
    End Sub
    Sub WriteSelectItemToGridView(filtertype As Integer)
        Dim frule, resultrule As String
        FTDDL.SelectedIndex = 0
        ds.Reset()
        SetGridViewStyle()
        SetShowALGridViewFields()
        resultrule = GetSearchRuleString()
        If (filtertype = 1) Then
            If (resultrule <> "") Then
                frule = " where T0.onhand>0 and " & GetSearchRuleString() & " order by T0.itemcode"
            Else
                frule = " where T0.onhand>0 order by T0.itemcode"
            End If
        Else
            frule = " where T0.onhand>0 order by T0.itemcode"
        End If
        'MsgBox(frule)
        'SqlCmd = "SELECT inout='異動單',postrecord='過帳記錄',T0.itemcode,T0.itemname,T0.onhand , T0.location,convert(varchar,T0.crdate,111) As crdate, " &
        '    "T1.onhand As saponhand " &
        '    "FROM dbo.[@SPOMT] T0 INNER JOIN OITM T1 ON T0.itemcode=T1.itemcode " & frule
        SqlCmd = "SELECT inout='異動單',postrecord='過帳記錄',T0.itemcode,T0.itemname,T0.onhand , T0.location,convert(varchar,T0.crdate,111) As crdate, " &
            "saponhand=0 " &
            "FROM dbo.[@SPOMT] T0 " & frule
        'MsgBox(SqlCmd)
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()

        frule = ""
        If (filtertype = 1) Then
            If (resultrule <> "") Then
                frule = " where T0.onhand=0 and " & GetSearchRuleString() & " order by T0.itemcode"
            Else
                frule = " where T0.onhand=0 order by T0.itemcode"
            End If
        Else
            frule = " where T0.onhand=0 order by T0.itemcode"
        End If

        SqlCmd = "SELECT inout='異動單',postrecord='過帳記錄',T0.itemcode,T0.itemname,T0.onhand , T0.location,convert(varchar,T0.crdate,111) As crdate, " &
            "saponhand=0 " &
            "FROM dbo.[@SPOMT] T0 " & frule
        'MsgBox(SqlCmd)
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
    End Sub

    Function GetSearchRuleString()
        Dim whs As String
        Dim str(), frule As String
        Dim ruleflag As Boolean
        ruleflag = False
        frule = ""
        If (TxtKW.Text <> "") Then
            frule = frule & " T0.itemcode like '%" & TxtKW.Text & "%'"
            ruleflag = True
        End If
        If (DDLWhs.SelectedIndex <> 0) Then
            str = Split(DDLWhs.SelectedValue, "_")
            whs = str(0)
            If (ruleflag) Then
                frule = frule & " and T0.whsjudge='" & whs & "'"
            Else
                frule = frule & " T0.whsjudge='" & whs & "'"
            End If
            ruleflag = True
            'MsgBox("HH")
        Else
            'MsgBox(DDLWhs.SelectedIndex & " " & DDLWhs.SelectedValue)
        End If
        'If (ruleflag) Then
        '    frule = frule & " order by T0.itemcode"
        'Else
        '    frule = " order by T0.itemcode"
        'End If

        Return frule
    End Function

    Sub SetGridViewStyle()
        If (fmode <> "showpost") Then
            gv1.AllowPaging = True
            gv1.PageSize = 15
            gv1.PagerStyle.HorizontalAlign = HorizontalAlign.Center
        Else
            gv1.AllowPaging = False
        End If
        gv1.AutoGenerateColumns = False
        'gv1.ShowHeader = True
        'gv1.AllowSorting = True
        'gv1.Font.Size = FontSize.Smaller
        'gv1.ForeColor =
        gv1.GridLines = GridLines.Both

        gv1.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
        'gv1.HeaderStyle.BackColor =
        'gv1.RowStyle.BackColor
        'gv1.AlternatingRowStyle.BackColor
        gv1.HeaderStyle.ForeColor = Drawing.Color.White
    End Sub

    Sub SetShowALGridViewFields() ' 1:主檔建立  2:異動單建立
        Dim oBoundField As BoundField
        'Dim oHyperLinkField As HyperLinkField
        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "異動"
        'oHyperLinkField.DataTextField = "num"
        oBoundField.DataField = "inout"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        'oHyperLinkField.NavigateUrl = "~/hr/leave.aspx?smid=hr&smode=1&fmode=modify"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "料號"
        oBoundField.DataField = "itemcode"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 80
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "說明"
        oBoundField.DataField = "itemname"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 200
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "庫存"
        oBoundField.DataField = "onhand"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        oBoundField.ItemStyle.Width = 20
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "SAP庫存"
        oBoundField.DataField = "saponhand"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        oBoundField.ItemStyle.Width = 20
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "櫃位"
        oBoundField.DataField = "location"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 20
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "建立日期"
        oBoundField.DataField = "crdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "過帳記錄"
        oBoundField.DataField = "postrecord"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 40
        gv1.Columns.Add(oBoundField)
    End Sub
    Sub SetMaterialInOutGridViewFields()
        Dim oBoundField As BoundField
        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "項次"
        oBoundField.DataField = "icount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "料號"
        oBoundField.DataField = "itemcode"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.ItemStyle.Width = 30
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "過帳日期"
        oBoundField.DataField = "ddate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.DataFormatString = "{0:yyyy/MM/dd}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "文件編號"
        oBoundField.DataField = "num"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.Columns.Add(oBoundField)

        'oBoundField = New BoundField
        'oBoundField.HeaderText = "倉庫"
        'oBoundField.DataField = "whsjudge"
        'oBoundField.ShowHeader = True
        'oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        'oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        'gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "收貨量"
        oBoundField.DataField = "inamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "發貨量"
        oBoundField.DataField = "outamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "建立人"
        oBoundField.DataField = "cname"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "備註"
        oBoundField.DataField = "reason"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)
    End Sub
    Sub ShowAddForm()
        If (fmode = "add" Or fmode = "modify") Then
            AddT.Rows(0).Cells(0).Text = "主檔資料建立"
            AddT.Rows(3).Cells(0).Text = "建立日期"
            AddT.Rows(4).Visible = False
            AddT.Rows(5).Visible = False
            AddT.Rows(6).Visible = False
            AddT.Rows(7).Visible = False
            AddT.Rows(8).Visible = True
            AddT.Rows(9).Visible = True
            If (fmode = "add") Then
                CType(AddT.FindControl("txt_itemcode"), TextBox).Enabled = True
                CType(AddT.FindControl("txt_itemname"), TextBox).Enabled = True
                AddT.Rows(3).Cells(1).Text = Format(Now, "yyyy/MM/dd")
            Else
                IniAddField()
                SqlCmd = "SELECT itemname,convert(varchar,crdate,111),whsjudge,location FROM dbo.[@SPOMT] where itemcode='" & Request.QueryString("itemcode") & "'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                CType(AddT.FindControl("txt_itemcode"), TextBox).Enabled = False
                CType(AddT.FindControl("txt_itemname"), TextBox).Enabled = False
                CType(AddT.FindControl("txt_itemcode"), TextBox).Text = Request.QueryString("itemcode")
                CType(AddT.FindControl("txt_itemname"), TextBox).Text = dr(0)
                AddT.Rows(3).Cells(1).Text = dr(1)
                CType(AddT.FindControl("txt_whs"), TextBox).Text = dr(2)
                CType(AddT.FindControl("txt_location"), TextBox).Text = dr(3)
                dr.Close()
                connsap.Close()
            End If
        ElseIf (fmode = "inout") Then
            IniInOutField()
            AddT.Rows(0).Cells(0).Text = "異動單建立"
            AddT.Rows(3).Cells(0).Text = "異動日期"
            AddT.Rows(4).Visible = True
            AddT.Rows(5).Visible = True
            AddT.Rows(6).Visible = True
            AddT.Rows(7).Visible = True
            AddT.Rows(8).Visible = False
            AddT.Rows(9).Visible = False
            CType(AddT.FindControl("txt_itemcode"), TextBox).Enabled = False
            CType(AddT.FindControl("txt_itemname"), TextBox).Enabled = False
            CType(AddT.FindControl("txt_itemcode"), TextBox).Text = Request.QueryString("itemcode")
            CType(AddT.FindControl("txt_itemname"), TextBox).Text = Request.QueryString("itemname")
            AddT.Rows(3).Cells(1).Text = Format(Now, "yyyy/MM/dd")
            AddT.Rows(4).Cells(1).Text = Request.QueryString("onhand")
        End If
        If (fmode = "add") Then
            CType(AddT.FindControl("chk_action"), CheckBox).Text = "新增確認"
            CType(AddT.FindControl("btn_action"), Button).Text = "新增"
        ElseIf (fmode = "modify") Then
            CType(AddT.FindControl("chk_action"), CheckBox).Text = "修改確認"
            CType(AddT.FindControl("btn_action"), Button).Text = "修改"
        Else
            CType(AddT.FindControl("chk_action"), CheckBox).Text = "異動確認"
            CType(AddT.FindControl("btn_action"), Button).Text = "異動"
        End If
        CType(AddT.FindControl("btn_action"), Button).Enabled = False
    End Sub
    Sub CreateAddItem()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Btnx As Button
        Dim Chkx As CheckBox
        Dim Labelx As Label
        'row=0
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.HorizontalAlign = HorizontalAlign.Center
        tCell = New TableCell
        tCell.ColumnSpan = 2
        tCell.BorderWidth = 1
        tRow.Controls.Add(tCell)
        AddT.Controls.Add(tRow)
        '------------------------------------------------
        'row=1
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("料號", 1))
        tRow.Cells.Add(CellSetWithTB("txt_itemcode"))
        AddT.Controls.Add(tRow)
        '-------------------------------------------------
        'row=2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("說明", 1))
        tRow.Cells.Add(CellSetWithTB("txt_itemname"))
        AddT.Controls.Add(tRow)
        '------------------------------------------------ 
        'row=3
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("", 1)) '建立日期 or 異動日期
        tRow.Cells.Add(CellSet("", 0))
        AddT.Controls.Add(tRow)
        '------------------------------------------------
        'row=4
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("目前庫存", 1))
        tRow.Cells.Add(CellSet("", 0))
        AddT.Controls.Add(tRow)
        'If (ftype = "add" Or ftype = "modify") Then
        'tRow.Visible = False
        'End If
        '-----------------------------------------------
        'row=5
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("異動數量", 1))
        tRow.Cells.Add(CellSetWithTB("txt_qty"))
        AddT.Controls.Add(tRow)
        'If (ftype = "add" Or ftype = "modify") Then
        '    tRow.Visible = False
        'End If
        '------------------------------------------------
        'row=6
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("異動種類", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_dtype", "txt_dtype", "dde_dtype"))
        AddT.Controls.Add(tRow)
        'If (ftype = "add" Or ftype = "modify") Then
        '    tRow.Visible = False
        'End If
        '------------------------------------------------
        'row=7
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("異動原因", 1))
        tRow.Cells.Add(CellSetWithTB("txt_reason"))
        AddT.Controls.Add(tRow)
        'If (ftype = "add" Or ftype = "modify") Then
        '    tRow.Visible = False
        'End If
        '------------------------------------------------
        'row=8
        'If (ftype = "add") Then
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("歸類倉別", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_whs", "txt_whs", "dde_whs"))
        AddT.Controls.Add(tRow)
        'If (ftype = "inout") Then
        '    tRow.Visible = False
        'End If
        '------------------------------------------------
        'row=9
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("櫃位", 1))
        tRow.Cells.Add(CellSetWithTB("txt_location"))
        AddT.Controls.Add(tRow)
        'If (ftype = "inout") Then
        '    tRow.Visible = False
        'End If
        'End If
        '------------------------------------------------
        'row=10
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.BackColor = Drawing.Color.AntiqueWhite
        tRow.HorizontalAlign = HorizontalAlign.Center
        tRow.Font.Bold = True


        tCell = New TableCell
        tCell.ColumnSpan = 2
        tCell.BorderWidth = 1
        Chkx = New CheckBox
        Chkx.ID = "chk_action"
        'If (fmode = "add") Then
        '    Chkx.Text = "打勾後新增"
        'Else
        '    Chkx.Text = "打勾後修改"
        'End If
        AddHandler Chkx.CheckedChanged, AddressOf Chkx_CheckedChanged
        Chkx.AutoPostBack = True
        Labelx = New Label
        Labelx.ID = "label2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Btnx = New Button
        Btnx.ID = "btn_action"
        'If (fmode = "add") Then
        '    Btnx.Text = "新增"
        'Else
        '    Btnx.Text = "修改"
        'End If
        Btnx.Enabled = False
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
        AddT.Controls.Add(tRow)

        If (fmode = "modify") Then
            SqlCmd = "SELECT count(*) FROM dbo.[@SPOMPT] where itemcode='" & Request.QueryString("itemcode") & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) = 0) Then
                tRow = New TableRow()
                tRow.BorderWidth = 1
                tRow.Font.Bold = True
                tRow.HorizontalAlign = HorizontalAlign.Center
                tCell = New TableCell
                tCell.ColumnSpan = 2
                tCell.BorderWidth = 1

                Chkx = New CheckBox
                Chkx.ID = "chk_del"
                Chkx.Text = "尚未有過帳資料,打勾後可刪除"
                AddHandler Chkx.CheckedChanged, AddressOf ChkDelx_CheckedChanged
                Chkx.AutoPostBack = True
                Labelx = New Label
                Labelx.ID = "labeldel"
                Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"

                Btnx = New Button
                Btnx.ID = "btn_del"
                Btnx.Text = "刪除"
                Btnx.Enabled = False
                AddHandler Btnx.Click, AddressOf BtnxDel_Click
                tCell.Controls.Add(Btnx)
                tCell.Controls.Add(Labelx)
                tCell.Controls.Add(Chkx)
                tRow.Controls.Add(tCell)
                AddT.Controls.Add(tRow)
            End If
            dr.Close()
            connsap.Close()
        End If
    End Sub
    Protected Sub BtnxDel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SqlCmd = "delete from dbo.[@SPOMT] " &
                "where itemcode = '" & Request.QueryString("itemcode") & "'"
        CommUtil.SqlSapExecute("del", SqlCmd, connsap)
        connsap.Close()
        WriteSelectItemToGridView(1)
    End Sub
    Protected Sub BtnxCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        WriteSelectItemToGridView(1)
    End Sub
    Protected Sub Chkx_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Btnx As Button
        Btnx = AddT.FindControl("btn_action")
        If (sender.Checked) Then
            Btnx.Enabled = True
            'Btnx.OnClientClick = "return confirm('要新增嗎')"
        Else
            Btnx.Enabled = False
            Btnx.OnClientClick = ""
        End If
    End Sub
    Protected Sub ChkDelx_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Btnx As Button
        Btnx = AddT.FindControl("btn_del")
        If (sender.Checked) Then
            Btnx.Enabled = True
            'Btnx.OnClientClick = "return confirm('要新增嗎')"
        Else
            Btnx.Enabled = False
            Btnx.OnClientClick = ""
        End If
    End Sub
    Function RecordFieldCheck()
        RecordFieldCheck = True
        If (CType(AddT.FindControl("txt_itemcode"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(AddT.FindControl("txt_itemname"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (fmode = "inout") Then
            If (CType(AddT.FindControl("txt_qty"), TextBox).Text = "") Then
                RecordFieldCheck = False
            End If
            If (CType(AddT.FindControl("txt_dtype"), TextBox).Text = "") Then
                RecordFieldCheck = False
            End If
            If (CType(AddT.FindControl("txt_reason"), TextBox).Text = "") Then
                RecordFieldCheck = False
            End If
        ElseIf (fmode = "add" Or fmode = "modify") Then
            If (CType(AddT.FindControl("txt_whs"), TextBox).Text = "") Then
                RecordFieldCheck = False
            End If
        End If
    End Function
    Function InsertMasterDataRecord()
        Dim resultflag As Boolean
        resultflag = False
        Dim itemcode, itemname, whs, cid, cname, crdate, location As String
        cid = Session("s_id")
        cname = Session("s_name")
        crdate = Format(Now, "yyyy/MM/dd")
        itemcode = CType(AddT.FindControl("txt_itemcode"), TextBox).Text
        itemname = CType(AddT.FindControl("txt_itemname"), TextBox).Text

        ' If (FTDDL.SelectedIndex = 1) Then
        SqlCmd = "SELECT count(*) FROM dbo.[@SPOMT] where itemcode='" & itemcode & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        If (dr(0) <> 0) Then
            CommUtil.ShowMsg(Me, itemcode & "料號已存在")
            dr.Close()
            connsap.Close()
            resultflag = False
        Else
            dr.Close()
            connsap.Close()
            location = CType(AddT.FindControl("txt_location"), TextBox).Text
            whs = CType(AddT.FindControl("txt_whs"), TextBox).Text
            SqlCmd = "insert into [dbo].[@SPOMT] (cid,cname,itemcode,itemname,crdate,whsjudge,location) " &
                        "values('" & cid & "','" & cname & "','" & itemcode & "','" & itemname & "','" & crdate & "', " &
                    "'" & whs & "','" & location & "')"
            CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
            connsap.Close()
            '讀回,check 是否已寫入
            SqlCmd = "SELECT count(*) FROM dbo.[@SPOMT] where itemcode='" & itemcode & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) <> 0) Then
                resultflag = True
            Else
                CommUtil.ShowMsg(Me, "資料沒寫入")
                resultflag = False
            End If
            dr.Close()
            connsap.Close()
        End If
        '        End If
        Return resultflag
    End Function
    Function UpdateFCRecord()
        Dim whs, location, itemcode As String
        itemcode = CType(AddT.FindControl("txt_itemcode"), TextBox).Text
        whs = CType(AddT.FindControl("txt_whs"), TextBox).Text
        location = CType(AddT.FindControl("txt_location"), TextBox).Text
        SqlCmd = "update dbo.[@SPOMT] set " &
        "whsjudge='" & whs & "', location= '" & location & "' where itemcode = '" & itemcode & "'"
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
        UpdateFCRecord = True
    End Function
    Function InsertInOutRecord()
        Dim num As Long
        Dim nullflag, resultflag As Boolean
        resultflag = False
        SqlCmd = "SELECT max(num) FROM dbo.[@SPOMPT]"
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
        Dim itemcode, dtype, reason, cid, cname, ddate As String
        Dim qty As Integer
        cid = Session("s_id")
        cname = Session("s_name")
        ddate = Format(Now, "yyyy/MM/dd")
        itemcode = CType(AddT.FindControl("txt_itemcode"), TextBox).Text
        qty = CInt(CType(AddT.FindControl("txt_qty"), TextBox).Text)
        dtype = CType(AddT.FindControl("txt_dtype"), TextBox).Text
        reason = CType(AddT.FindControl("txt_reason"), TextBox).Text
        SqlCmd = "insert into [dbo].[@SPOMPT] (cid,cname,itemcode,dtype,reason,ddate,qty) " &
                 "values('" & cid & "','" & cname & "','" & itemcode & "','" & dtype & "','" & reason & "','" & ddate & "', " & qty & ")"
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()
        '讀回,check 是否已寫入
        If (nullflag) Then
            SqlCmd = "SELECT max(num) FROM dbo.[@SPOMPT]"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (IsDBNull(dr(0))) Then
                CommUtil.ShowMsg(Me, "資料沒寫入")
                resultflag = False
            Else
                resultflag = True
            End If
            dr.Close()
            connsap.Close()
        Else
            SqlCmd = "SELECT itemcode FROM dbo.[@SPOMPT] where num=" & num
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr.HasRows) Then
                resultflag = True
            Else
                CommUtil.ShowMsg(Me, "資料沒寫入")
                resultflag = False
            End If
            dr.Close()
            connsap.Close()
            If (resultflag) Then
                'update庫存數
                If (dtype = "收入") Then
                    SqlCmd = "update dbo.[@SPOMT] set " &
                         "onhand=onhand+'" & qty & "' where itemcode = '" & itemcode & "'"
                Else
                    SqlCmd = "update dbo.[@SPOMT] set " &
                         "onhand=onhand-'" & qty & "' where itemcode = '" & itemcode & "'"
                End If
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
            End If
        End If
        Return resultflag
    End Function
    Protected Sub BtnxAction_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Chkx As CheckBox
        Dim actionok As Boolean
        If (RecordFieldCheck()) Then
            Chkx = AddT.FindControl("chk_action")
            'If (fmode = "add") Then
            If (Chkx.Checked) Then
                If (action = "add") Then
                    actionok = InsertMasterDataRecord()
                    'MsgBox(fmode)
                ElseIf (action = "modify") Then
                    actionok = UpdateFCRecord()
                ElseIf (action = "inout") Then
                    actionok = InsertInOutRecord()
                End If
            Else
                CommUtil.ShowMsg(Me, "新增(修改或異動)check沒打勾")
                Exit Sub
            End If
            If (actionok) Then
                WriteSelectItemToGridView(1)
            End If
        Else
            CommUtil.ShowMsg(Me, "有欄位空白")
        End If
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
        Txtx.AutoPostBack = True
        If (Txtxid = "txt_itemcode") Then
            AddHandler Txtx.TextChanged, AddressOf Txtx_TextChanged
        End If
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
        Txtx = AddT.FindControl("txt_" & str(1))
        Txtx.Text = sender.SelectedValue

        '        FilterT.Visible = False
        '        AddT.Visible = True
        'MsgBox("LBS")
        '        gv1.Visible = False
    End Sub
    Sub FTCreate()
        FilterTCreate()
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
        '--------------------------------
        Labelx = New Label()
        Labelx.ID = "label_kw"
        Labelx.Text = "料號關鍵字:"
        Labelx.Font.Size = 10
        tCell.Controls.Add(Labelx)
        TxtKW = New TextBox()
        TxtKW.ID = "txt_kw"
        TxtKW.Width = 100
        tCell.Controls.Add(TxtKW)
        '-----------------------------------------
        Labelx = New Label()
        Labelx.ID = "label_adds3"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp或&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        'ron
        DDLWhs = New DropDownList()
        DDLWhs.ID = "ddl_whs"
        'DDLWhs.AutoPostBack = True
        tCell.Controls.Add(DDLWhs)
        '--------------------------------------
        Labelx = New Label()
        Labelx.ID = "label_adds1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnSearch = New Button()
        BtnSearch.ID = "btn_search"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnSearch.Text = "尋找"
        BtnSearch.Font.Size = 10
        AddHandler BtnSearch.Click, AddressOf BtnSearch_Click
        tCell.Controls.Add(BtnSearch)
        tRow.Cells.Add(tCell)

        Labelx = New Label()
        Labelx.ID = "label_adds2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp或&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnSPM = New Button()
        BtnSPM.ID = "btn_spm"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnSPM.Text = "顯示全部"
        BtnSPM.Font.Size = 10
        AddHandler BtnSPM.Click, AddressOf BtnSPM_Click
        tCell.Controls.Add(BtnSPM)
        tRow.Cells.Add(tCell)


        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Right
        Labelx = New Label()
        Labelx.ID = "label_begin"
        Labelx.Text = "&nbsp&nbsp&nbsp異動期間:"
        Labelx.Font.Size = 10
        tCell.Controls.Add(Labelx)
        TxtBeginDate = New TextBox()
        TxtBeginDate.ID = "txt_begin"
        TxtBeginDate.Width = 70
        TxtBeginDate.AutoPostBack = True
        AddHandler TxtBeginDate.TextChanged, AddressOf TxtBeginDate_TextChanged
        tCell.Controls.Add(TxtBeginDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtBeginDate.ID
        ce.ID = "ce_begindate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(TxtBeginDate)
        '-------------------------------
        Labelx = New Label()
        Labelx.ID = "label_end"
        Labelx.Text = "~"
        tCell.Controls.Add(Labelx)
        TxtEndDate = New TextBox()
        TxtEndDate.ID = "txt_end"
        TxtEndDate.Width = 70
        TxtEndDate.AutoPostBack = True
        AddHandler TxtEndDate.TextChanged, AddressOf TxtEndDate_TextChanged
        tCell.Controls.Add(TxtEndDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtEndDate.ID
        ce.ID = "ce_enddate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(TxtEndDate)
        tRow.Cells.Add(tCell)
        FilterT.Rows.Add(tRow)
    End Sub
    Protected Sub BtnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '        gv1.Visible = True
        '        If (TxtKW.Text <> "") Then
        fmode = "show"
            WriteSelectItemToGridView(1)
        'Else
        '    gv1.Visible = False
        '    CommUtil.ShowMsg(Me, "需輸入關鍵字")
        'End If
    End Sub
    Protected Sub BtnSPM_Click(sender As Object, e As EventArgs)
        'gv1.Visible = True
        fmode = "show"
        gv1.PageIndex = Request.QueryString("indexpage")
        WriteSelectItemToGridView(0)
    End Sub
    Protected Sub FTDDL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FTDDL.SelectedIndexChanged
        If (FTDDL.SelectedIndex = 0) Then
            FTDDL.Visible = True
            FilterT.Visible = True
            AddT.Visible = False
            gv1.Visible = True
            WriteSelectItemToGridView(1)
        ElseIf (FTDDL.SelectedIndex = 1) Then
            FTDDL.Visible = False
            FilterT.Visible = False
            AddT.Visible = True
            gv1.Visible = False
            IniAddField()
        End If
    End Sub
    Sub IniAddField()
        CType(AddT.FindControl("txt_itemcode"), TextBox).Text = ""
        CType(AddT.FindControl("txt_itemname"), TextBox).Text = ""
        CType(AddT.FindControl("txt_whs"), TextBox).Text = ""
        CType(AddT.FindControl("txt_location"), TextBox).Text = ""
        CType(AddT.FindControl("chk_action"), CheckBox).Checked = False
    End Sub
    Sub IniInOutField()
        CType(AddT.FindControl("txt_itemcode"), TextBox).Text = ""
        CType(AddT.FindControl("txt_itemname"), TextBox).Text = ""
        CType(AddT.FindControl("txt_qty"), TextBox).Text = ""
        CType(AddT.FindControl("txt_dtype"), TextBox).Text = ""
        CType(AddT.FindControl("txt_reason"), TextBox).Text = ""
        CType(AddT.FindControl("chk_action"), CheckBox).Checked = False
    End Sub
    Sub MaterialInOutPosted()

        Dim itemcode As String
        itemcode = Request.QueryString("itemcode")
        TxtBeginDate.Text = Request.QueryString("begindate")
        TxtEndDate.Text = Request.QueryString("enddate")
        ds.Reset()
        SetGridViewStyle()
        SetMaterialInOutGridViewFields()
        'MsgBox(TxtBeginDate.Text & " " & TxtEndDate.Text)
        SqlCmd = "SELECT icount=0,outamount=0,num,qty As inamount,ddate,cname,reason,itemcode FROM dbo.[@SPOMPT] " &
                "where itemcode='" & itemcode & "' and dtype='收入' and ddate>='" & TxtBeginDate.Text & "' " &
                "And ddate <='" & TxtEndDate.Text & "'"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        SqlCmd = "SELECT icount=0,inamount=0,num,qty As outamount,ddate,cname,reason FROM dbo.[@SPOMPT] " &
                "where itemcode='" & itemcode & "' and dtype='發出' and ddate>='" & TxtBeginDate.Text & "' " &
                "And ddate <='" & TxtEndDate.Text & "'"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()

        ds.Tables(0).DefaultView.Sort = "ddate"
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "無任何過帳資料")
        End If
    End Sub
    Protected Sub TxtBeginDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        WriteSelectItemToGridView(1) '日期變更後 , 在此值才會生效 , 且需執行此程式 , 才能讓過帳資料所帶日期參數更新
    End Sub
    Protected Sub TxtEndDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        WriteSelectItemToGridView(1)
    End Sub

    Sub Txtx_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim Txtx As TextBox = sender
        SqlCmd = "SELECT itemname FROM OITM where itemcode='" & CStr(CType(AddT.FindControl(Txtx.ID), TextBox).Text) & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            CType(AddT.FindControl("txt_itemname"), TextBox).Text = dr(0)
        Else
            CommUtil.ShowMsg(Me, "非 Sap 上有登錄之料件")
            CType(AddT.FindControl("txt_itemname"), TextBox).Text = ""
        End If
        dr.Close()
        connsap.Close()
    End Sub
    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        fmode = "show"
        WriteSelectItemToGridView(1)
        FTDDL.Visible = True
        AddT.Visible = False
        FilterT.Visible = True
        gv1.Visible = True
    End Sub
    Sub AddWhsItem()
        SqlCmd = "SELECT T0.[WhsCode], T0.[WhsName] " &
        "FROM OWHS T0 order by T0.WhsCode"
        DDLWhs.Items.Clear()
        DDLWhs.Items.Add("請選擇所屬倉別")
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            Do While (drsap.Read())
                DDLWhs.Items.Add(drsap(0) & "_" & drsap(1))
            Loop
        End If
        'DDLWhs.SelectedIndex = 0
        drsap.Close()
        connsap.Close()
    End Sub
End Class