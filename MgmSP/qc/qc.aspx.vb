Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class qc
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap As SqlDataReader
    Public TxtKW, TxtPO, TxtNum, TxtBeginDate, TxtEndDate, TxtPOFilter As TextBox
    Public DDLID, DDLIQCID, DDLVender, DDLMtype, DDLResult, DDLAudit, DDLRecord As DropDownList
    Public BtnSearch, BtnUmitFilter, BtnIQCFilter As Button
    Public BtnPO As Button
    Public ds As New DataSet
    Public permsqc100 As String

    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        Dim Hyper As HyperLink
        Dim tinamount As Integer
        Dim vcode As String
        vcode = ""
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If (DDLFun.SelectedIndex = 1) Then '建立
                realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
                e.Row.Cells(0).Text = realindex + 1
                Hyper = New HyperLink()
                If (ds.Tables(0).Rows(realindex)("u_F7") = 0) Then
                    Hyper.Text = "建規格"
                    Hyper.NavigateUrl = "iqc.aspx?smid=qc&smode=1&mode=edit&iqctype=2&funindex=1&indexpage=" & gv1.PageIndex &
                        "&itemcode=" & e.Row.Cells(1).Text & "&kw=" & TxtKW.Text &
                        "&itemname=" & e.Row.Cells(2).Text
                    CommUtil.DisableObjectByPermission(Hyper, permsqc100, "n")
                Else
                    Hyper.Text = "查詢修改"
                    Hyper.NavigateUrl = "iqc.aspx?smid=qc&smode=1&mode=showvalue&iqctype=1&funindex=1&indexpage=" & gv1.PageIndex &
                    "&itemcode=" & e.Row.Cells(1).Text &
                    "&itemname=" & e.Row.Cells(2).Text
                End If
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_rawsta_" & e.Row.Cells(0).Text
                e.Row.Cells(4).Controls.Add(Hyper)
            ElseIf (DDLFun.SelectedIndex = 2) Then
                realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
                e.Row.Cells(0).Text = realindex + 1
                Hyper = New HyperLink()
                Hyper.Text = "查詢修改"
                Hyper.NavigateUrl = "iqc.aspx?smid=qc&smode=1&mode=showvalue&iqctype=1&funindex=2&indexpage=" & gv1.PageIndex &
                    "&itemcode=" & e.Row.Cells(1).Text &
                    "&itemname=" & e.Row.Cells(2).Text
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_rawsta_" & e.Row.Cells(0).Text
                'CommUtil.DisableObjectByPermission(Hyper, permsqc100, "m")
                e.Row.Cells(4).Controls.Add(Hyper)
            ElseIf (DDLFun.SelectedIndex = 4) Then
                realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
                e.Row.Cells(0).Text = realindex + 1
                SqlCmd = "SELECT T0.CardCode FROM OPOR T0 where T0.docnum=" & CInt(e.Row.Cells(4).Text)
                drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (drsap.HasRows) Then
                    drsap.Read()
                    vcode = drsap(0)
                    connsap.Close()
                End If
                If (vcode = "T021" Or vcode = "T021-1") Then
                    SqlCmd = "Select T0.quantity from dbo.cnc1 T0 where T0.sappo=" & CInt(e.Row.Cells(4).Text) & "and T0.itemcode='" & e.Row.Cells(2).Text & "'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    If (dr.HasRows) Then
                        dr.Read()
                        e.Row.Cells(6).Text = dr(0)
                    Else
                        CommUtil.ShowMsg(Me, "在cnc加工單找不到此PO:" & e.Row.Cells(4).Text & "及料號--" & e.Row.Cells(2).Text)
                    End If
                    dr.Close()
                    conn.Close()
                End If
                Hyper = New HyperLink()
                If (ds.Tables(0).Rows(realindex)("action") = 0) Then
                    Hyper.Text = "建IQC單"
                    Hyper.NavigateUrl = "iqc.aspx?smid=qc&smode=1&iqctype=3&mode=edit&funindex=4&indexpage=" & gv1.PageIndex
                    CommUtil.DisableObjectByPermission(Hyper, permsqc100, "n")
                Else
                    Hyper.Text = "查詢修改"
                    Hyper.NavigateUrl = "iqc.aspx?smid=qc&smode=1&mode=showvalue&iqctype=4&funindex=4&docnum=" & ds.Tables(0).Rows(realindex)("u_docnum") &
                        "&itemname=" & e.Row.Cells(3).Text & "&indexpage=" & gv1.PageIndex
                End If
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_rawsta_" & e.Row.Cells(1).Text
                'CommUtil.DisableObjectByPermission(Hyper, permsqc100, "m")
                e.Row.Cells(10).Controls.Add(Hyper)
                If (e.Row.Cells(9).Text = "檢驗中") Then
                    e.Row.Cells(9).BackColor = Drawing.Color.Yellow
                End If
            ElseIf (DDLFun.SelectedIndex = 3) Then
                realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
                e.Row.Cells(0).Text = realindex + 1
                If (e.Row.Cells(4).Text = "CNC部" Or e.Row.Cells(4).Text = "CNC部門") Then
                    '處理itemname
                    SqlCmd = "select T0.itemname " &
                    "from dbo.OITM T0 where T0.itemcode='" & e.Row.Cells(2).Text & "'"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    If (dr.HasRows) Then
                        dr.Read()
                        e.Row.Cells(3).Text = dr(0)
                    Else
                        CommUtil.ShowMsg(Me, "在SAP找不到此料號--" & e.Row.Cells(2).Text)
                        connsap.Close()
                        Exit Sub
                    End If
                    dr.Close()
                    connsap.Close()
                End If
                SqlCmd = "Select IsNull(sum(u_inamount),0) FROM dbo.[@UIQT] T0 " &
                "where T0.u_itemcode='" & ds.Tables(0).Rows(realindex)("itemcode") & "' and T0.u_po=" & ds.Tables(0).Rows(realindex)("docnum")
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    tinamount = dr(0)
                Else
                    tinamount = 0
                End If
                dr.Close()
                connsap.Close()
                e.Row.Cells(5).Text = CInt(e.Row.Cells(5).Text)
                e.Row.Cells(6).Text = CInt(tinamount)
                e.Row.Cells(7).Text = CInt(ds.Tables(0).Rows(realindex)("quantity") - tinamount)

                Hyper = New HyperLink()
                If (CInt(e.Row.Cells(7).Text) > 0) Then
                    SqlCmd = "SELECT count(*) FROM dbo.[@UMIT] T0 " &
                "where T0.u_itemcode='" & ds.Tables(0).Rows(realindex)("itemcode") & "'"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    dr.Read()
                    If (dr(0) <> 0) Then
                        Hyper.Text = "建IQC單"
                        Hyper.NavigateUrl = "iqc.aspx?smid=qc&smode=1&mode=edit&iqctype=3&funindex=3&po=" & ds.Tables(0).Rows(realindex)("docnum") &
                    "&rest_inamount=" & CInt(e.Row.Cells(7).Text) & "&itemcode=" & e.Row.Cells(2).Text & "&po_amount=" & CInt(e.Row.Cells(5).Text) &
                    "&itemname=" & e.Row.Cells(3).Text & "&indexpage=" & gv1.PageIndex
                    Else
                        Hyper.Text = "建規格"
                        Hyper.NavigateUrl = "iqc.aspx?smid=qc&smode=1&mode=edit&iqctype=2&funindex=3&itemcode=" & e.Row.Cells(2).Text &
                        "&itemname=" & e.Row.Cells(3).Text & "&indexpage=" & gv1.PageIndex & "&po=" & e.Row.Cells(1).Text
                    End If
                    dr.Close()
                    connsap.Close()
                    Hyper.Font.Underline = False
                    Hyper.ID = "hyper_rawsta_" & e.Row.Cells(0).Text
                    CommUtil.DisableObjectByPermission(Hyper, permsqc100, "n")
                    e.Row.Cells(8).Controls.Add(Hyper)
                Else
                    e.Row.Cells(8).Text = "結束"
                End If

            End If
        End If
    End Sub

    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        'MsgBox(gv1.PageIndex)
        If (DDLFun.SelectedIndex = 2) Then
            FilterUmitSearch()
        ElseIf (DDLFun.SelectedIndex = 1) Then
            'MsgBox(gv1.PageIndex)
            UMITKWSearch(TxtKW.Text)
        ElseIf (DDLFun.SelectedIndex = 4) Then
            FilterIQCSearch()
        ElseIf (DDLFun.SelectedIndex = 3) Then
            ShowPOItem()
        End If
    End Sub

    Public ScriptManager1 As New ScriptManager
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim triggerid As String
        'Dim control As Control
        'If (IsPostBack) Then
        '    triggerid = Page.Request.Params.Get("__EVENTTARGET")
        '    If (triggerid.Length <> 0) Then
        '        Control = Page.FindControl(triggerid)
        '        If (Not (Control Is Nothing)) Then
        '            If (Control.ID = "DDLFun") Then
        '                'ddlidflag = False
        '                FT.EnableViewState = False
        '            End If
        '        End If
        '    End If
        'End If
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim iqctype As Integer
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        permsqc100 = CommUtil.GetAssignRight("qc100", Session("s_id"))
        If (Not IsPostBack) Then
            DDLFun.SelectedIndex = Request.QueryString("funindex")
            'iqctype = Request.QueryString("funindex")
            'If (iqctype = 1) Then
            '    DDLFun.SelectedIndex = 2
            'ElseIf (iqctype = 2) Then
            '    DDLFun.SelectedIndex = 1
            'ElseIf (iqctype = 3) Then
            '    DDLFun.SelectedIndex = 3
            'ElseIf (iqctype = 4) Then
            '    DDLFun.SelectedIndex = 4
            'End If
        End If
        If (DDLFun.SelectedIndex <> 0) Then
            FTItemMasterCreate()
            FTIQCCreate()
        End If
        If (DDLFun.SelectedIndex = 1) Then
            FTIMC.Visible = True
            FTIMS.Visible = False
            FTIQCC.Visible = False
            FTIQCS.Visible = False
        ElseIf (DDLFun.SelectedIndex = 2) Then
            FTIMC.Visible = False
            FTIMS.Visible = True
            FTIQCC.Visible = False
            FTIQCS.Visible = False
        ElseIf (DDLFun.SelectedIndex = 3) Then
            FTIMC.Visible = False
            FTIMS.Visible = False
            FTIQCC.Visible = True
            FTIQCS.Visible = False
        ElseIf (DDLFun.SelectedIndex = 4) Then
            FTIMC.Visible = False
            FTIMS.Visible = False
            FTIQCC.Visible = False
            FTIQCS.Visible = True
        End If
        If (Not IsPostBack) Then
            gv1.PageIndex = Request.QueryString("indexpage")
            If (DDLFun.SelectedIndex = 4) Then
                InsertIQCItemcodeList()
                InitIQCShow()
            ElseIf (DDLFun.SelectedIndex = 2) Then
                gv1.Visible = True
                InsertUMITItemcodeList()
                FilterUmitSearch()
            ElseIf (DDLFun.SelectedIndex = 1) Then
                TxtKW.Text = Request.QueryString("kw")
                UMITKWSearch(TxtKW.Text)
            ElseIf (DDLFun.SelectedIndex = 3) Then
                gv1.Visible = True
                TxtPO.Text = Request.QueryString("po")
                ShowPOItem()
            End If
        End If
        'DDLFun.Visible = False
    End Sub


    Protected Sub DDLFun_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DDLFun.SelectedIndexChanged
        If (DDLFun.SelectedIndex = 0) Then

        ElseIf (DDLFun.SelectedIndex = 2) Then
            InsertUMITItemcodeList()
        ElseIf (DDLFun.SelectedIndex = 4) Then
            InsertIQCItemcodeList()
        End If
        gv1.Visible = False
        gv1.PageIndex = 0
    End Sub

    Sub InsertUMITItemcodeList()
        SqlCmd = "SELECT distinct T0.u_itemcode,T1.Itemname FROM dbo.[@UMIT] T0 INNER JOIN OITM T1 ON T0.u_Itemcode=T1.Itemcode order by T0.u_itemcode"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            DDLID.Items.Clear()
            DDLID.Items.Add("全部已建測試規格料號")
            Do While (drsap.Read())
                DDLID.Items.Add(drsap(0) & " " & drsap(1))
            Loop
        End If
        drsap.Close()
        connsap.Close()
    End Sub

    Sub InsertIQCItemcodeList()
        SqlCmd = "SELECT distinct T0.u_itemcode,T1.Itemname FROM dbo.[@UMIT] T0 INNER JOIN OITM T1 ON T0.u_Itemcode=T1.Itemcode order by T0.u_itemcode"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            DDLIQCID.Items.Clear()
            DDLIQCID.Items.Add("全部已建測試規格料號")
            Do While (drsap.Read())
                DDLIQCID.Items.Add(drsap(0) & " " & drsap(1))
            Loop
        End If
        drsap.Close()
        connsap.Close()
    End Sub
    Sub FTItemMasterCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        tRow = New TableRow()
            tRow.BorderWidth = 1
            tCell = New TableCell()
            tCell.HorizontalAlign = HorizontalAlign.Left
            '--------------------------------
            Labelx = New Label()
            Labelx.ID = "label_kw"
            Labelx.Text = "料號關鍵字:"
            tCell.Controls.Add(Labelx)
            TxtKW = New TextBox()
            TxtKW.ID = "txt_kw"
            TxtKW.Width = 100
            tCell.Controls.Add(TxtKW)
            '-----------------------------------------
            Labelx = New Label()
            Labelx.ID = "label_adds1"
            Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
            tCell.Controls.Add(Labelx)
            BtnSearch = New Button()
            BtnSearch.ID = "btn_search"
            'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
            BtnSearch.Text = "尋找"
        AddHandler BtnSearch.Click, AddressOf BtnSearch_Click
        tCell.Controls.Add(BtnSearch)
            tRow.Cells.Add(tCell)
        FTIMC.Rows.Add(tRow)
        '''''''以上是Function row
        '''''''以下是Filter Row
        tRow = New TableRow()
            tRow.BorderWidth = 1
            tCell = New TableCell()
            tCell.ColumnSpan = 3
            tCell.HorizontalAlign = HorizontalAlign.Left

            DDLID = New DropDownList()
            DDLID.ID = "ddl_id"
            DDLID.Width = 600
            'DDLID.EnableViewState = ddlidflag
            'DDLID.EnableViewState = False
            tCell.Controls.Add(DDLID)

            Labelx = New Label()
            Labelx.ID = "label_adds3"
            Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
            tCell.Controls.Add(Labelx)
            BtnUmitFilter = New Button()
            BtnUmitFilter.ID = "btn_filter"
            BtnUmitFilter.Text = "篩選"
            AddHandler BtnUmitFilter.Click, AddressOf BtnUmitFilter_Click
            tCell.Controls.Add(BtnUmitFilter)

            tRow.Cells.Add(tCell)
        FTIMS.Rows.Add(tRow)
    End Sub

    Sub FTIQCCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim ce As CalendarExtender

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left

        Labelx = New Label()
        Labelx.ID = "label_po"
        Labelx.Text = "採購號:"
        tCell.Controls.Add(Labelx)
        TxtPO = New TextBox()
        TxtPO.ID = "txt_po"
        TxtPO.Width = 40
        tCell.Controls.Add(TxtPO)
        '-----------------------------------------
        Labelx = New Label()
        Labelx.ID = "label_addiqcs2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnPO = New Button()
        BtnPO.ID = "btn_po"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnPO.Text = "顯示料件"
        AddHandler BtnPO.Click, AddressOf BtnPO_Click
        tCell.Controls.Add(BtnPO)
        tRow.Cells.Add(tCell)

        FTIQCC.Rows.Add(tRow)
        '''''''以上是Function row

        '''''''以下是Filter Row
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.ColumnSpan = 3
        tCell.HorizontalAlign = HorizontalAlign.Left

        Labelx = New Label()
        Labelx.ID = "label_iqcnum"
        Labelx.Text = "單號:"
        tCell.Controls.Add(Labelx)
        TxtNum = New TextBox()
        TxtNum.ID = "txt_num"
        TxtNum.Width = 40
        tCell.Controls.Add(TxtNum)
        '----------------------------------------
        DDLIQCID = New DropDownList()
        DDLIQCID.ID = "ddl_iqcid"
        DDLIQCID.Width = 180
        DDLIQCID.SelectedIndex = 0
        'DDLID.Items.Clear()
        'DDLID.Items.Add("選擇料號")
        'DDLIQCID.EnableViewState = ddlidflag
        tCell.Controls.Add(DDLIQCID)
        '--------------------------------
        DDLVender = New DropDownList()
        DDLVender.ID = "ddl_vender"
        DDLVender.Width = 100
        DDLVender.Items.Clear()
        DDLVender.Items.Add("廠商")
        DDLVender.SelectedIndex = 0
        tCell.Controls.Add(DDLVender)
        '-------------------------------
        Labelx = New Label()
        Labelx.ID = "label_begin"
        Labelx.Text = "開始日期:"
        tCell.Controls.Add(Labelx)
        TxtBeginDate = New TextBox()
        TxtBeginDate.ID = "txt_begin"
        TxtBeginDate.Width = 70
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
        Labelx.Text = "結束日期:"
        tCell.Controls.Add(Labelx)
        TxtEndDate = New TextBox()
        TxtEndDate.ID = "txt_end"
        TxtEndDate.Width = 70
        tCell.Controls.Add(TxtEndDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtEndDate.ID
        ce.ID = "ce_enddate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(TxtEndDate)
        '-------------------------------
        Labelx = New Label()
        Labelx.ID = "label_pof"
        Labelx.Text = "採購單:"
        tCell.Controls.Add(Labelx)
        TxtPOFilter = New TextBox()
        TxtPOFilter.ID = "txt_pof"
        TxtPOFilter.Width = 40
        tCell.Controls.Add(TxtPOFilter)
        '--------------------------------
        DDLMtype = New DropDownList()
        DDLMtype.ID = "ddl_mtype"
        DDLMtype.Width = 80
        DDLMtype.Items.Clear()
        DDLMtype.Items.Add("進料類別")
        DDLMtype.SelectedIndex = 0
        tCell.Controls.Add(DDLMtype)
        '--------------------------------
        DDLResult = New DropDownList()
        DDLResult.ID = "ddl_result"
        DDLResult.Width = 80
        DDLResult.Items.Clear()
        DDLResult.Items.Add("選擇結果")
        DDLResult.SelectedIndex = 0
        tCell.Controls.Add(DDLResult)
        '--------------------------------
        DDLAudit = New DropDownList()
        DDLAudit.ID = "ddl_audit"
        DDLAudit.Width = 80
        DDLAudit.Items.Clear()
        DDLAudit.Items.Add("未審單")
        DDLAudit.SelectedIndex = 0
        tCell.Controls.Add(DDLAudit)

        'tRow.Cells.Add(tCell)
        '----------------------------------
        'tCell = New TableCell()
        'tCell.BorderWidth = 1
        'tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label()
        Labelx.ID = "label_iqcadds3"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnIQCFilter = New Button()
        BtnIQCFilter.ID = "btn_iqcfilter"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnIQCFilter.Text = "篩選"
        AddHandler BtnIQCFilter.Click, AddressOf BtnIQCFilter_Click
        tCell.Controls.Add(BtnIQCFilter)

        tRow.Cells.Add(tCell)
        FTIQCS.Rows.Add(tRow)

    End Sub

    Protected Sub BtnUmitFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        gv1.Visible = True
        FilterUmitSearch()
    End Sub

    Protected Sub BtnIQCFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gv1.Visible = True
        FilterIQCSearch()
    End Sub

    Sub InitIQCShow()
        gv1.Visible = True
        FilterIQCSearch()
    End Sub

    Protected Sub BtnPO_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        gv1.Visible = True
        ShowPOItem()

        'FilterIQCSearch()
    End Sub

    Sub ShowPOItem()
        Dim po As Long
        Dim vcode As String
        po = CLng(TxtPO.Text)
        ds.Reset()
        SetGridViewStyle()
        SetPOGridViewFields()
        SqlCmd = "SELECT T0.CardCode FROM OPOR T0 where T0.docnum=" & po
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            drsap.Read()
            vcode = drsap(0)
            connsap.Close()
            drsap.Close()
            If (vcode = "T021" Or vcode = "T021-1") Then
                SqlCmd = "Select icount=0,famount=0,restamount=0,status=0,Dscription='',T0.itemcode,T0.quantity,T0.sappo as docnum, " &
                "cardname='CNC部' " &
                "from dbo.cnc1 T0 where T0.sappo=" & po
                ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
                conn.Close()
                gv1.DataSource = ds.Tables(0)
                gv1.DataBind()
            Else
                'SqlCmd = "SELECT icount=0,famount=0,restamount=0,status=0,T1.itemcode,T1.Dscription,T0.docnum,T0.cardname,T1.quantity " &
                '"FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry " &
                '"where T0.docnum=" & po & " order by T1.itemcode"
                SqlCmd = "SELECT icount=0,famount=0,restamount=0,status=0,T1.itemcode,(select itemname from oitm where itemcode=T1.itemcode) as Dscription, " &
                "docnum=" & po & ",(select cardname from opor where docnum=" & po & ") as cardname,Sum(T1.quantity) as quantity " &
                "FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T0.docnum=" & po & " group by T1.itemcode order by T1.itemcode"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
                connsap.Close()
                gv1.DataSource = ds.Tables(0)
                gv1.DataBind()
            End If
        Else
            CommUtil.ShowMsg(Me, "無此PO號")
        End If
    End Sub

    Sub FilterUmitSearch()
        Dim rule As String
        Dim itemcode As String
        Dim filterflag As Boolean
        Dim str() As String
        filterflag = False
        rule = " where "

        If (DDLID.SelectedIndex <> 0) Then
            str = Split(DDLID.SelectedValue, " ")
            itemcode = str(0)
            If (filterflag = True) Then
                rule = rule & " and u_itemcode='" & itemcode & "' "
            Else
                rule = rule & " u_itemcode='" & itemcode & "' "
            End If
            filterflag = True
        End If

        If (filterflag = False) Then
            rule = ""
        End If
        rule = rule & " order by u_itemcode"
        'MsgBox(rule)
        WriteSelectItemToGridView(rule)
    End Sub

    Sub FilterIQCSearch()
        Dim rule As String
        Dim docnum, po As Long
        Dim mtype, judge As Integer
        Dim itemcode, vender, begindate, enddate As String
        Dim filterflag As Boolean
        Dim status As String
        Dim str() As String
        filterflag = False
        rule = " where "
        If (TxtNum.Text <> "") Then
            docnum = CLng(TxtNum.Text)
            rule = rule & "u_docnum=" & docnum
            filterflag = True
        End If

        If (DDLIQCID.SelectedIndex <> 0) Then
            str = Split(DDLIQCID.SelectedValue, " ")
            itemcode = str(0)
            If (filterflag = True) Then
                rule = rule & " and u_itemcode='" & itemcode & "' "
            Else
                rule = rule & " u_itemcode='" & itemcode & "' "
            End If
            filterflag = True
        End If

        If (DDLVender.SelectedIndex <> 0) Then
            vender = DDLID.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_vender='" & vender & "' "
            Else
                rule = rule & " u_vender='" & vender & "' "
            End If
            filterflag = True
        End If
        If (TxtBeginDate.Text <> "" And TxtEndDate.Text <> "") Then
            begindate = TxtBeginDate.Text
            enddate = TxtEndDate.Text
            If (filterflag = True) Then
                rule = rule & " and (u_cdate>='" & begindate & "' and u_cdate<='" & enddate & "') "
            Else
                rule = rule & " (u_cdate>='" & begindate & "' and u_cdate<='" & enddate & "') "
            End If
            filterflag = True
        End If
        If (TxtPOFilter.Text <> "") Then
            po = CLng(TxtPOFilter.Text)
            If (filterflag = True) Then
                rule = rule & " and u_po=" & po
            Else
                rule = rule & "u_po=" & po
            End If
            filterflag = True
        End If

        If (DDLMtype.SelectedIndex <> 0) Then
            If (DDLMtype.SelectedValue = "銑件") Then
                mtype = 1
            ElseIf (DDLMtype.SelectedValue = "車件") Then
                mtype = 2
            ElseIf (DDLMtype.SelectedValue = "市購") Then
                mtype = 3
            ElseIf (DDLMtype.SelectedValue = "鈑金") Then
                mtype = 4
            ElseIf (DDLMtype.SelectedValue = "骨架") Then
                mtype = 5
            End If

            If (filterflag = True) Then
                rule = rule & " and u_mtype=" & mtype
            Else
                rule = rule & " u_mtype=" & mtype
            End If
            filterflag = True
        End If
        If (DDLResult.SelectedIndex <> 0) Then
            If (DDLResult.SelectedValue = "允收") Then
                judge = 1
            ElseIf (DDLResult.SelectedValue = "拒收") Then
                judge = 2
            ElseIf (DDLResult.SelectedValue = "特採") Then
                judge = 3
            ElseIf (DDLResult.SelectedValue = "其他") Then
                judge = 4
            End If

            If (filterflag = True) Then
                rule = rule & " and u_judge=" & judge
            Else
                rule = rule & " u_judge=" & judge
            End If
            filterflag = True
        End If

        If (DDLAudit.SelectedIndex <> 0) Then
            status = DDLAudit.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_status='" & status & "' "
            Else
                rule = rule & " u_status='" & status & "' "
            End If
            filterflag = True
        End If

        If (filterflag = False) Then
            rule = ""
        End If
        rule = rule & " order by u_docnum desc"
        'MsgBox(rule)
        WriteSelectItemToGridView(rule)
    End Sub

    Sub WriteSelectItemToGridView(rule As String)
        'MsgBox(rule)
        '        Dim po As Long
        '        po = CLng(TxtPOFilter.Text)
        ds.Reset()
        SetGridViewStyle()
        If (DDLFun.SelectedIndex = 2) Then
            SetUmitGridViewFields()
            SqlCmd = "SELECT distinct icount=0,status=1,T0.u_itemcode,T1.Itemname,T0.u_mapno " &
            "FROM dbo.[@UMIT] T0 INNER JOIN OITM T1 ON T0.u_Itemcode=T1.Itemcode " & rule
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()
            gv1.DataSource = ds.Tables(0)
            gv1.DataBind()
        ElseIf (DDLFun.SelectedIndex = 4) Then
            SetIQCGridViewFields()
            SqlCmd = "SELECT icount=0,action=1,T0.u_itemcode,T1.Itemname,T0.u_docnum, " &
            "T0.u_po,T0.u_amount,T0.u_inamount,T0.u_famount,T0.u_cdate,T0.u_status " &
            "FROM dbo.[@UIQT] T0 INNER JOIN OITM T1 ON T0.u_Itemcode=T1.Itemcode " & rule
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()
            gv1.DataSource = ds.Tables(0)
            gv1.DataBind()
        End If
    End Sub

    Sub SetGridViewStyle()
        'gv1.AutoGenerateColumns = False
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

    Sub SetUmitGridViewFields()
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
        oBoundField.DataField = "u_itemcode"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "說明"
        oBoundField.DataField = "itemname"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "對應編號"
        oBoundField.DataField = "u_mapno"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "動作"
        oBoundField.DataField = "status"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)
    End Sub

    Sub SetUmitKWSearchGridViewFields()
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

        oBoundField = New BoundField
        oBoundField.HeaderText = "對應編號"
        oBoundField.DataField = "u_F6"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "動作"
        oBoundField.DataField = "u_F7"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)
    End Sub

    Sub SetIQCGridViewFields()
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
        oBoundField.HeaderText = "單號"
        oBoundField.DataField = "u_docnum"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "料號"
        oBoundField.DataField = "u_itemcode"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "說明"
        oBoundField.DataField = "itemname"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "PO號"
        oBoundField.DataField = "u_po"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "建立日期"
        oBoundField.DataField = "u_cdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:yyyy/MM/dd}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "總量"
        oBoundField.DataField = "u_amount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "進量"
        oBoundField.DataField = "u_inamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "已檢"
        oBoundField.DataField = "u_famount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "狀態"
        oBoundField.DataField = "u_status"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "動作"
        oBoundField.DataField = "action"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)
    End Sub

    Sub SetPOGridViewFields()
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
        oBoundField.HeaderText = "PO號"
        oBoundField.DataField = "docnum"
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
        oBoundField.DataField = "Dscription"
        oBoundField.ShowHeader = True
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "Vender"
        oBoundField.DataField = "cardname"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "總量"
        oBoundField.DataField = "quantity"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "已建"
        oBoundField.DataField = "famount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "剩餘"
        oBoundField.DataField = "restamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "狀態"
        oBoundField.DataField = "status"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)
    End Sub

    Protected Sub BtnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gv1.Visible = True
        If (TxtKW.Text <> "") Then
            UMITKWSearch(TxtKW.Text)
        Else
            CommUtil.ShowMsg(Me, "需輸入關鍵字")
        End If
    End Sub

    Sub UMITKWSearch(kw As String)
        'MsgBox(kw)
        ds.Reset()
        SetGridViewStyle()
        SetUmitKWSearchGridViewFields()
        SqlCmd = "SELECT icount=0,T0.itemcode,T0.itemname,IsNull(T0.u_F6,0) As u_F6,IsNull(T0.u_F7,0) As u_F7 " &
                "FROM OITM T0 " &
                "where T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%' order by T0.itemcode"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
    End Sub
End Class