Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class freport
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap As SqlDataReader
    Public TxtSDate, TxtEDate, TxtMaterialID As TextBox
    Public DDLMaterialType, DDLReportType As DropDownList
    Public BtnFilter As Button
    Public ds As New DataSet
    Public permshr100 As String
    Public ScriptManager1 As New ScriptManager
    Public TotalPrice As Double
    Public mode As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        mode = Request.QueryString("mode")
        FTCreate()
        If (Not IsPostBack) Then
            TxtMaterialID.Text = Request.QueryString("keycode")
            TxtSDate.Text = Request.QueryString("begindate")
            TxtEDate.Text = Request.QueryString("enddate")
            DDLMaterialType.SelectedIndex = Request.QueryString("materialtype")
            DDLReportType.SelectedIndex = Request.QueryString("reportindex")
            gv1.PageIndex = Request.QueryString("indexpage")
            ViewState("keycode") = TxtMaterialID.Text
            ViewState("begindate") = TxtSDate.Text
            ViewState("enddate") = TxtEDate.Text
            ViewState("materialtype") = DDLMaterialType.SelectedIndex
            ViewState("reportindex") = DDLReportType.SelectedIndex
            ViewState("indexpage") = gv1.PageIndex
        Else
            TxtMaterialID.Text = ViewState("keycode")
            TxtSDate.Text = ViewState("begindate")
            TxtEDate.Text = ViewState("enddate")
            DDLMaterialType.SelectedIndex = ViewState("materialtype")
            DDLReportType.SelectedIndex = ViewState("reportindex")
            gv1.PageIndex = ViewState("indexpage")
        End If

        If (mode = "detaillist") Then
            ShowDetailList()
        Else
            ShowList()
        End If
    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        Dim Hyper As HyperLink
        'Dim Hyper As LinkButton
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            e.Row.Cells(0).Text = realindex + 1
            If (mode <> "detaillist") Then
                SqlCmd = "Select Itemname FROM OITM " &
                    "WHERE itemcode = '" & e.Row.Cells(1).Text & "'"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                drL.Read()
                e.Row.Cells(2).Text = drL(0)
                drL.Close()
                connL.Close()
                Hyper = New HyperLink()
                'Hyper = New LinkButton()
                Hyper.Text = "過帳記錄"
                Hyper.NavigateUrl = "freport.aspx?smid=fd&smode=1&mode=detaillist&indexpage=" & gv1.PageIndex & "&itemcode=" & e.Row.Cells(1).Text &
                    "&begindate=" & TxtSDate.Text & "&enddate=" & TxtEDate.Text & "&keycode=" & TxtMaterialID.Text &
                    "&materialtype=" & DDLMaterialType.SelectedIndex & "&reportindex=" & DDLReportType.SelectedIndex
                'Hyper.PostBackUrl = "freport.aspx?smid=fd&smode=1&mode=detaillist&indexpage=" & gv1.PageIndex & "&itemcode=" & e.Row.Cells(1).Text
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_purchase_" & e.Row.Cells(0).Text
                e.Row.Cells(5).Controls.Add(Hyper)
                If (e.Row.Cells(4).Text < 0) Then
                    e.Row.Cells(4).BackColor = Drawing.Color.LightSalmon
                End If
                If (e.Row.Cells(3).Text < 0) Then
                    e.Row.Cells(3).BackColor = Drawing.Color.LightSalmon
                End If
            Else
                If (e.Row.Cells(8).Text < 0) Then
                    e.Row.Cells(8).BackColor = Drawing.Color.LightSalmon
                    If (e.Row.Cells(4).Text <> 0) Then
                        e.Row.Cells(4).Text = "A/P發票-" & e.Row.Cells(4).Text
                    Else
                        e.Row.Cells(4).Text = "無"
                    End If
                    If (e.Row.Cells(5).Text <> 0) Then
                        CommUtil.ShowMsg(Me, "A/P貸項通知有目標文件(" & e.Row.Cells(5).Text & "), 但程式無對應, 請洽程式設計師")
                    Else
                        e.Row.Cells(5).Text = "無"
                    End If
                Else
                    If (e.Row.Cells(4).Text <> 0) Then
                        e.Row.Cells(4).Text = "收貨採購-" & e.Row.Cells(4).Text
                    Else
                        e.Row.Cells(4).Text = "無"
                    End If
                    If (e.Row.Cells(5).Text <> 0) Then
                        e.Row.Cells(5).Text = "貸項通知-" & e.Row.Cells(5).Text
                    Else
                        e.Row.Cells(5).Text = "無"
                    End If
                End If
                If (e.Row.Cells(13).Text < 0) Then
                    e.Row.Cells(13).BackColor = Drawing.Color.LightSalmon
                End If
                If (e.Row.Cells(11).Text = 0) Then
                    e.Row.Cells(11).Text = ""
                End If
                If (e.Row.Cells(12).Text = 0) Then
                    e.Row.Cells(12).Text = ""
                End If
                e.Row.Cells(14).ToolTip = e.Row.Cells(14).Text
                If (e.Row.Cells(14).Text.Length > 30) Then
                    e.Row.Cells(14).Text = e.Row.Cells(14).Text.Substring(0, 30) + "..."
                End If
            End If
        End If
    End Sub
    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        If (mode = "detaillist") Then
            ShowDetailList()
        Else
            ShowList()
        End If
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim Hyper As HyperLink
        Dim ce As CalendarExtender
        tRow = New TableRow()
        tRow.BorderWidth = 1
        '-----------------------------------------
        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Left
        DDLMaterialType = New DropDownList()
        DDLMaterialType.ID = "ddl_mtype"
        DDLMaterialType.Width = 200
        tCell.Controls.Add(DDLMaterialType)
        DDLMaterialType.Items.Clear()
        DDLMaterialType.Items.Add("請選擇料件種類")
        DDLMaterialType.Items.Add("原料")
        DDLMaterialType.Items.Add("半成品")
        DDLMaterialType.Items.Add("製成品")

        '-----------------------------------------
        Labelx = New Label()
        Labelx.ID = "label_rtype"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        DDLReportType = New DropDownList()
        DDLReportType.ID = "ddl_rtype"
        DDLReportType.Width = 200
        tCell.Controls.Add(DDLReportType)
        DDLReportType.Items.Clear()
        DDLReportType.Items.Add("請選擇報表種類")
        DDLReportType.Items.Add("進料報表")
        DDLReportType.Items.Add("銷售報表")

        Labelx = New Label()
        Labelx.ID = "label_mid"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp料號: "
        tCell.Controls.Add(Labelx)
        TxtMaterialID = New TextBox()
        TxtMaterialID.ID = "txt_mid"
        TxtMaterialID.Width = 100
        tCell.Controls.Add(TxtMaterialID)

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

        Labelx = New Label()
        Labelx.ID = "label_hyper"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Hyper = New HyperLink
        Hyper.ID = "hyper_back0"
        Hyper.Text = "回前頁"
        Hyper.NavigateUrl = "freport.aspx?smid=fd&smode=1&mode=''&indexpage=" & Request.QueryString("indexpage") &
                    "&begindate=" & Request.QueryString("begindate") & "&enddate=" & Request.QueryString("enddate") & "&keycode=" & Request.QueryString("keycode") &
                    "&materialtype=" & Request.QueryString("materialtype") & "&reportindex=" & Request.QueryString("reportindex")
        Hyper.Font.Underline = False
        tCell.Controls.Add(Hyper)
        If (mode <> "detaillist") Then
            Hyper.Visible = False
        Else
            Hyper.Visible = True
        End If

        tRow.Cells.Add(tCell)
        FT.Rows.Add(tRow)

    End Sub
    Protected Sub BtnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (DDLMaterialType.SelectedIndex = 0 Or DDLReportType.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "請選擇料件或報表種類")
            Exit Sub
        End If
        If (DDLMaterialType.SelectedIndex = 2 Or DDLMaterialType.SelectedIndex = 3 Or DDLReportType.SelectedIndex = 2) Then
            CommUtil.ShowMsg(Me, "尚未發展")
            Exit Sub
        End If
        ShowList()
    End Sub
    Sub ShowList()
        Dim frule, idstring, str() As String
        Dim allitemcode As Boolean
        allitemcode = True
        frule = ""
        If (DDLMaterialType.SelectedIndex = 1) Then
            frule = " T0.docdate >='" & TxtSDate.Text & "' and T0.docdate <='" & TxtEDate.Text & "' and (Left(itemcode,1)='0' or Left(itemcode,1)='3') " &
                    "and Left(itemcode, 2) <> '3A' and  Left(itemcode, 2) <> '3J'"
        ElseIf (DDLMaterialType.SelectedIndex = 2) Then
            frule = " T0.docdate >='" & TxtSDate.Text & "' and T0.docdate <='" & TxtEDate.Text & "' and (Left(itemcode,1)='2')"
        ElseIf (DDLMaterialType.SelectedIndex = 3) Then
            frule = " T0.docdate >='" & TxtSDate.Text & "' and T0.docdate <='" & TxtEDate.Text & "' and " &
                    "Left(itemcode,1)= '1' and Left(itemcode,2) <> '1A' and Left(itemcode,2) <> '1J'"
        ElseIf (DDLMaterialType.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "需選擇料件種類")
            Exit Sub
        End If
        idstring = TxtMaterialID.Text
        If ((TxtSDate.Text <> "" And TxtEDate.Text <> "")) Then
            str = Split(idstring, "*")
            For i = 0 To UBound(str)
                If (str(i) <> "") Then
                    allitemcode = False
                End If
            Next
            If (allitemcode = False) Then
                For i = 0 To UBound(str)
                    If (str(i) <> "") Then
                        frule = frule & " and itemcode like '%" & str(i) & "%'"
                    End If
                Next
            End If
            If (DDLReportType.SelectedIndex = 1) Then
                PurchaseReport(frule)
            End If
        Else
            CommUtil.ShowMsg(Me, "需設定日期區間")
        End If
    End Sub
    Sub PurchaseReport(frule As String)
        SqlCmd = "SELECT IsNull(sum(T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0)),0) " &
                    "FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where" & frule
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        TotalPrice = drsap(0)
        drsap.Close()
        connsap.Close()
        SqlCmd = "SELECT IsNull(0-sum(T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0)),0) " &
                    "FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where" & frule
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        TotalPrice = TotalPrice + drsap(0)
        drsap.Close()
        connsap.Close()
        ds.Reset()
        SetGridViewStyle()
        SetMaterialGridViewFields()
        SqlCmd = "SELECT itemcode,IsNull(0-sum(T1.Quantity),0) As totalqty,IsNull(0-sum(T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0)),0) As totalprice,  " &
                    "itemname='',icount=0,status=0 " &
                    "FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where" & frule & " group by itemcode order by totalprice desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        SqlCmd = "SELECT itemcode,IsNull(sum(T1.Quantity),0) As totalqty,IsNull(sum(T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0)),0) As totalprice,  " &
                    "itemname='',icount=0,status=0 " &
                    "FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where" & frule & " group by itemcode order by totalprice desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()

        'ds.Tables(0).DefaultView.Sort = "itemcode"
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
    End Sub

    Sub ShowDetailList()
        Dim frule, itemcode As String
        frule = ""
        itemcode = Request.QueryString("itemcode")
        frule = " T0.docdate >='" & TxtSDate.Text & "' and T0.docdate <='" & TxtEDate.Text & "' and T1.itemcode='" & itemcode & "'"
        If ((TxtSDate.Text <> "" And TxtEDate.Text <> "")) Then
            If (DDLReportType.SelectedIndex = 1) Then
                APListOfItemcodeReport(frule)
            End If
        Else
            CommUtil.ShowMsg(Me, "需設定日期區間")
        End If
    End Sub
    Sub APListOfItemcodeReport(frule As String)
        SqlCmd = "SELECT IsNull(sum(T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0)),0) " &
                    "FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where" & frule
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        TotalPrice = drsap(0)
        drsap.Close()
        connsap.Close()
        SqlCmd = "SELECT IsNull(0-sum(T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0)),0) " &
                    "FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where" & frule
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        TotalPrice = TotalPrice + drsap(0)
        drsap.Close()
        connsap.Close()
        ds.Reset()
        SetGridViewStyle()
        SetMaterialInOutGridViewFields()
        SqlCmd = "SELECT T1.itemcode,T1.Quantity As inamount,T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0) As tprice,  " &
                    "T0.docdate,doctype='A/P發票',T0.docnum,T1.whscode,T1.AcctCode,t1.price as unitprice,T0.comments, " &
                    "T2.itemname,icount=0,T1.rate,T0.DiscPrcnt as discount,T0.doccur,T0.docdate,IsNull(T1.TrgetEntry,0) As targetentry,IsNull(T1.BaseEntry,0) as baseentry " &
                    "FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "INNER JOIN OITM T2 ON T1.Itemcode=T2.Itemcode " &
                    "where" & frule
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        SqlCmd = "SELECT T1.itemcode,0-T1.Quantity As inamount,0-T1.[TotalSumSy]*(1-T0.[DiscPrcnt]/100.0) As tprice,  " &
                    "T0.docdate,doctype='A/P貸項通知',T0.docnum,T1.whscode,T1.AcctCode,t1.price as unitprice,T0.comments, " &
                    "T2.itemname,icount=0,T1.rate,T0.DiscPrcnt as discount,T0.doccur,T0.docdate,IsNull(T1.TrgetEntry,0) As targetentry,IsNull(T1.BaseEntry,0) as baseentry " &
                    "FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "INNER JOIN OITM T2 ON T1.Itemcode=T2.Itemcode " &
                    "where" & frule
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        ds.Tables(0).DefaultView.Sort = "docdate"
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "無任何A/P發票過帳記錄")
        End If
    End Sub
    Sub SetGridViewStyle()
        gv1.AutoGenerateColumns = False
        gv1.ShowFooter = True
        gv1.AllowPaging = True
        gv1.PageSize = 20
        gv1.PagerStyle.HorizontalAlign = HorizontalAlign.Center

            'gv1.AllowSorting = True
            'gv1.Font.Size = FontSize.Smaller
            'gv1.ForeColor =
            gv1.GridLines = GridLines.Both


        gv1.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.FooterStyle.HorizontalAlign = HorizontalAlign.Center
        'gv1.HeaderStyle.BackColor =
        'gv1.RowStyle.BackColor
        'gv1.AlternatingRowStyle.BackColor
        'gv1.HeaderStyle.ForeColor
    End Sub

    Sub SetMaterialGridViewFields()
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
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "說明"
        oBoundField.DataField = "itemname"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        'oBoundField = New BoundField
        'oBoundField.HeaderText = "倉庫"
        'oBoundField.DataField = "whscode"
        'oBoundField.ShowHeader = True
        'oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        'oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        'gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "總進數量"
        oBoundField.DataField = "totalqty"
        oBoundField.ShowHeader = True
        oBoundField.FooterText = "進貨總價"
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "總價(NTD)"
        oBoundField.DataField = "totalprice"
        oBoundField.FooterText = Format(TotalPrice, "###,###")
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:N0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "單據資訊"
        oBoundField.DataField = "status"
        oBoundField.FooterText = "NTD"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
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
        oBoundField.HeaderText = "過帳日期"
        oBoundField.DataField = "docdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.DataFormatString = "{0:yyyy/MM/dd}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "文件種類"
        oBoundField.DataField = "doctype"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "文件編號"
        oBoundField.DataField = "docnum"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "來源文件"
        oBoundField.DataField = "baseentry"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "目標文件"
        oBoundField.DataField = "targetentry"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "倉庫"
        oBoundField.DataField = "whscode"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "會計科目"
        oBoundField.DataField = "AcctCode"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

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
        oBoundField.HeaderText = "單價"
        oBoundField.DataField = "unitprice"
        oBoundField.ShowHeader = True
        oBoundField.FooterText = "Summary"
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:N0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "幣別"
        oBoundField.DataField = "doccur"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "匯率"
        oBoundField.DataField = "rate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F3}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "折扣%"
        oBoundField.DataField = "discount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F3}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "總價(NTD)"
        oBoundField.DataField = "tprice"
        oBoundField.ShowHeader = True
        oBoundField.FooterText = Format(TotalPrice, "###,###")
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:N0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "備註"
        oBoundField.DataField = "comments"
        oBoundField.ShowHeader = True
        oBoundField.FooterText = "NTD"
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)
    End Sub
End Class