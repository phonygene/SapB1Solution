Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Imports System.IO
Imports System.Web.Caching
Imports System.Collections
Imports System.Collections.Generic
Partial Public Class molist
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap As New SqlConnection
    Public conn As New SqlConnection
    Public dr, drsap As SqlDataReader
    Public SqlCmd As String
    Public smid As String
    Public smode As Integer
    Public wotype As String
    Public permsmf000 As String
    Public permsmf202 As String
    Public permsmf201 As String
    Public TxtSapWo, TxtCus, TxtBeginDate, TxtEndDate As TextBox
    Public DDLModel, DDLWoType As DropDownList
    Public BtnFilter, BtnFilterReset As Button
    Public rule As String
    Public FileUL As FileUpload
    Public ChkDel As CheckBox
    Public indexpage As Integer
    Public ScriptManager1 As New ScriptManager

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim upditem As String
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        If (IsPostBack) Then
            rule = ViewState("rule")
        Else
            rule = ""
            gv1.PageIndex = Request.QueryString("indexpage")
        End If

        permsmf000 = CommUtil.GetAssignRight("mf000", Session("s_id"))
        permsmf202 = CommUtil.GetAssignRight("mf202", Session("s_id"))
        permsmf201 = CommUtil.GetAssignRight("mf201", Session("s_id"))
        upditem = Request.QueryString("upditem")
        smode = Request.QueryString("smode")
        smid = Request.QueryString("smid")
        If (smode = 1) Then
            wotype = "一般"
        ElseIf (smode = 2) Then
            wotype = "備庫"
        ElseIf (smode = 3) Then
            wotype = "半成品"
        End If
        gv1.HeaderStyle.Width = 10
        If (upditem = "pofrom") Then
            UpdatePOFrom()
        End If
        If (upditem = "updstatus") Then
            UpdatePOStatus()
        End If
        CreateFT()
        'If (Not Page.IsPostBack) Then
        If (ViewState("begindate") = "") Then
            GetNormalWo(wotype)
        End If
        'End If
    End Sub

    Sub CreateFT()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim ce As CalendarExtender

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.ColumnSpan = 3
        tCell.HorizontalAlign = HorizontalAlign.Left

        Labelx = New Label()
        Labelx.ID = "label_sapwo"
        Labelx.Text = "Sap工單號:"
        tCell.Controls.Add(Labelx)
        TxtSapWo = New TextBox()
        TxtSapWo.ID = "txt_sapwo"
        TxtSapWo.Width = 40
        tCell.Controls.Add(TxtSapWo)

        'Labelx = New Label()
        'Labelx.ID = "label_wsn"
        'Labelx.Text = "&nbsp&nbsp&nbsp&nbsp自訂單號:"
        'tCell.Controls.Add(Labelx)
        'TxtWsn = New TextBox()
        'TxtWsn.ID = "txt_wsn"
        'TxtWsn.Width = 100
        'tCell.Controls.Add(TxtWsn)

        Labelx = New Label()
        Labelx.ID = "label_cus"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp客戶:"
        tCell.Controls.Add(Labelx)
        TxtCus = New TextBox()
        TxtCus.ID = "txt_wsn"
        TxtCus.Width = 100
        tCell.Controls.Add(TxtCus)

        Labelx = New Label()
        Labelx.ID = "label_model"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        If (Session("usingwhs") = "C02") Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                     "FROM dbo.[@UMMD] T0 where T0.u_mtype='SPI' or T0.u_mtype='AOI' or " &
                     "T0.u_mtype='3DAOI' order by T0.u_model,T0.u_mcode"
        ElseIf (Session("usingwhs") = "C01") Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                     "FROM dbo.[@UMMD] T0 where T0.u_mtype='ICT' order by T0.u_model,T0.u_mcode"
            'Else
            'CommUtil.ShowMsg(Me, "倉別設定須為C01 or C02已決定是ICT or AOI")
        End If
        DDLModel = New DropDownList()
        DDLModel.ID = "ddl_model"
        DDLModel.Width = 150
        DDLModel.Items.Clear()
        DDLModel.Items.Add("選擇機型")
        If (Session("usingwhs") = "C02" Or Session("usingwhs") = "C01") Then
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    DDLModel.Items.Add(drsap(0))
                Loop
            End If
            drsap.Close()
            connsap.Close()
        End If
        'AddHandler DDLModel.SelectedIndexChanged, AddressOf DDLModel_SelectedIndexChanged
        'DDLModel.AutoPostBack = True
        'DDLModel.SelectedIndex = 0
        tCell.Controls.Add(DDLModel)

        Labelx = New Label()
        Labelx.ID = "label_shipdate"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp已出貨日期:"
        tCell.Controls.Add(Labelx)
        TxtBeginDate = New TextBox()
        TxtBeginDate.ID = "txt_begindate"
        TxtBeginDate.Width = 100
        tCell.Controls.Add(TxtBeginDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtBeginDate.ID
        ce.ID = "ce_begindate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)

        Labelx = New Label()
        Labelx.ID = "label_shipdate1"
        Labelx.Text = "-"
        tCell.Controls.Add(Labelx)
        TxtEndDate = New TextBox()
        TxtEndDate.ID = "txt_enddate"
        TxtEndDate.Width = 100
        tCell.Controls.Add(TxtEndDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtEndDate.ID
        ce.ID = "ce_enddate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)

        Labelx = New Label()
        Labelx.ID = "label_filter"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnFilter = New Button()
        BtnFilter.ID = "btn_po"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnFilter.Text = "篩選"
        AddHandler BtnFilter.Click, AddressOf BtnFilter_Click
        tCell.Controls.Add(BtnFilter)
        tRow.Cells.Add(tCell)

        Labelx = New Label()
        Labelx.ID = "label_reset"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnFilterReset = New Button()
        BtnFilterReset.ID = "btn_filterreset"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnFilterReset.Text = "重設"
        AddHandler BtnFilterReset.Click, AddressOf BtnFilterReset_Click
        tCell.Controls.Add(BtnFilterReset)
        tRow.Cells.Add(tCell)
        FilterT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.BackColor = Drawing.Color.AntiqueWhite
        tCell = New TableCell()
        tCell.ColumnSpan = 3
        tCell.HorizontalAlign = HorizontalAlign.Left

        Labelx = New Label()
        Labelx.ID = "label_fileul"
        Labelx.Text = "選擇上傳之聯絡單檔案"
        tCell.Controls.Add(Labelx)
        FileUL = New FileUpload()
        FileUL.ID = "fileul"
        tCell.Controls.Add(FileUL)

        ChkDel = New CheckBox
        ChkDel.ID = "chk_del"
        ChkDel.Text = "刪除聯絡單"
        ChkDel.AutoPostBack = True
        AddHandler ChkDel.CheckedChanged, AddressOf ChkDel_CheckedChanged
        tCell.Controls.Add(ChkDel)


        tRow.Cells.Add(tCell)

        'Labelx = New Label()
        'Labelx.ID = "label_wotype"
        'Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        'tCell.Controls.Add(Labelx)
        'DDLWoType = New DropDownList()
        'DDLWoType.ID = "ddl_wotype"
        'DDLWoType.Width = 150
        'DDLWoType.Items.Clear()
        'DDLWoType.Items.Add("在製工單")
        'DDLWoType.Items.Add("備貨工單")
        'DDLWoType.Items.Add("半成品模組工單")
        'AddHandler DDLWoType.SelectedIndexChanged, AddressOf DDLWoType_SelectedIndexChanged
        'DDLWoType.AutoPostBack = True
        'tCell.Controls.Add(DDLWoType)

        FilterT.Rows.Add(tRow)
    End Sub
    Protected Sub DDLWoType_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

    End Sub
    Protected Sub BtnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        rule = GetFilterRule()
        ViewState("rule") = rule
        If (TxtBeginDate.Text = "" And TxtEndDate.Text = "") Then
            GetNormalWo(wotype)
        Else
            GetShippedWo()
        End If
        ViewState("begindate") = TxtBeginDate.Text
    End Sub

    Protected Sub BtnFilterReset_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        rule = ""
        ViewState("rule") = rule
        TxtSapWo.Text = ""
        TxtCus.Text = ""
        TxtBeginDate.Text = ""
        TxtEndDate.Text = ""
        DDLModel.SelectedIndex = 0
        GetNormalWo(wotype)
        ViewState("begindate") = TxtBeginDate.Text
    End Sub
    Function GetFilterRule()
        Dim rule As String
        Dim filterflag As Boolean
        filterflag = False
        rule = " and "
        If (TxtSapWo.Text <> "") Then
            rule = rule & "docnum=" & CLng(TxtSapWo.Text)
            filterflag = True
        End If

        'If (TxtWsn.Text <> "") Then
        '    If (filterflag = True) Then
        '        rule = rule & " and wsn='" & TxtWsn.Text & "'"
        '    Else
        '        rule = rule & " wsn='" & TxtWsn.Text & "'"
        '    End If
        '    filterflag = True
        'End If

        If (TxtCus.Text <> "") Then
            If (filterflag = True) Then
                rule = rule & " and cus_name like '%" & TxtCus.Text & "%' "
            Else
                rule = rule & " cus_name like '%" & TxtCus.Text & "%' "
            End If
            filterflag = True
        End If

        If (DDLModel.SelectedIndex <> 0) Then
            If (filterflag = True) Then
                rule = rule & " and model='" & DDLModel.SelectedValue & "'"
            Else
                rule = rule & " model='" & DDLModel.SelectedValue & "'"
            End If
            filterflag = True
        End If
        'If (TxtBeginDate.Text <> "" And TxtEndDate.Text <> "") Then
        '    If (filterflag = True) Then
        '        rule = rule & " and (ship_date>='" & TxtBeginDate.Text & "' and ship_date<='" & TxtEndDate.Text & "') "
        '    Else
        '        rule = rule & " (ship_date>='" & TxtBeginDate.Text & "' and ship_date<='" & TxtEndDate.Text & "') "
        '    End If
        '    filterflag = True
        'End If
        If (TxtBeginDate.Text <> "" And TxtEndDate.Text <> "") Then
            If (filterflag = True) Then
                rule = rule & " and (status>='" & TxtBeginDate.Text & "' and status<='" & TxtEndDate.Text & "') "
            Else
                rule = rule & " (status>='" & TxtBeginDate.Text & "' and status<='" & TxtEndDate.Text & "') "
            End If
            filterflag = True
        End If

        'rule = rule & " order by model,docnum"
        If (filterflag = False) Then
            rule = ""
        End If
        GetFilterRule = rule
    End Function
    Sub UpdatePOFrom()
        Dim wsn As String
        Dim postatus As Integer
        postatus = Request.QueryString("postatus")
        wsn = Request.QueryString("wsn")
        If (postatus = 10) Then
            postatus = 20
        ElseIf (postatus = 20) Then
            postatus = 30
        ElseIf (postatus = 30) Then
            postatus = 40
        ElseIf (postatus = 40) Then
            postatus = 50
        ElseIf (postatus = 50) Then
            postatus = 60
        ElseIf (postatus = 60) Then
            postatus = 70
        ElseIf (postatus = 70) Then
            postatus = 10
        End If
        SqlCmd = "update dbo.[worksn] set getpo= " & postatus & " where wsn='" & wsn & "'"
        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub

    Sub UpdatePOStatus()
        'Dim count As Integer
        Dim wsn As String
        Dim postatus As Integer
        postatus = Request.QueryString("postatus")
        wsn = Request.QueryString("wsn")
        ' CommUtil.ShowMsg(Me,postatus & "-" & wsn)
        If (postatus = 10) Then
            postatus = 20
        ElseIf (postatus = 20) Then
            postatus = 30
        ElseIf (postatus = 30) Then
            postatus = 40
        ElseIf (postatus = 40) Then
            postatus = 50
        ElseIf (postatus = 50) Then
            postatus = 60
        ElseIf (postatus = 60) Then
            postatus = 70
        ElseIf (postatus = 70) Then
            postatus = 10
        End If
        SqlCmd = "update dbo.[worksn] set f_stat= " & postatus & " where wsn='" & wsn & "'"
        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub
    Sub GetShippedWo() 'just test 2023/04/28
        Dim ds As New DataSet
        Dim SelectCmd As String
        If (rule <> "") Then
            If (Session("usingwhs") = "C01") Then
                SelectCmd = "Select T1.status As ship_date, ship_set=1, " &
                         "T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                         "T0.resolution , T0.f_set , T0.camera_brand , T0.f_stat , " &
                         "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.CDate,T0.docnum " &
                         "From dbo.[work_records] T1 inner join dbo.[worksn] T0 On T1.wsn=T0.wsn where Left(T0.wsn,1)='I' " & rule &
                         " and T1.dpart=5 and T1.iseq=4 order by T0.wsn,T1.status"
            Else
                SelectCmd = "Select T1.status As ship_date, ship_set=1, " &
                         "T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                         "T0.resolution , T0.f_set , T0.camera_brand , T0.f_stat , " &
                         "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.CDate,T0.docnum " &
                         "From dbo.[work_records] T1 inner join dbo.[worksn] T0 On T1.wsn=T0.wsn where Left(T0.wsn,1)<>'I' " & rule &
                         " and T1.dpart=5 and T1.iseq=4 order by T0.wsn,T1.status"
            End If
            ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
            conn.Close()
        End If
        '--------------------------------------------------
        If (ds.Tables(0).Rows.Count <> 0) Then
            ds.Tables(0).Columns.Add("cno")
        Else
            CommUtil.ShowMsg(Me, "無任何資料")
        End If
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        'If (Session("s_id") = "ron" Or Session("s_id") = "ltx") Then
        ShowInfo()
        'End If
    End Sub
    Sub GetNormalWo(wotype As String)
        Dim ds, ds1 As New DataSet
        Dim SelectCmd As String
        If (rule = "") Then
            '先建立Wsn,但還未建立sap
            If (Session("usingwhs") = "C01") Then
                SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                             "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                             "T0.note ,T0.comm,T0.model_set,T0.mfmes,convert(char(12),T0.CDate,111) as CDate,T0.ship_date,T0.docnum " &
                             "From dbo.[worksn] T0 where Left(T0.wsn,1)='I' and T0.docnum= 0 " & rule & " order by T0.f_stat,T0.model,T0.cdate desc"
            Else
                SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                             "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                             "T0.note ,T0.comm,T0.model_set,T0.mfmes,convert(char(12),T0.CDate,111) as CDate,T0.ship_date,T0.docnum " &
                             "From dbo.[worksn] T0 where Left(T0.wsn,1)<>'I' and T0.docnum= 0 " & rule & " order by T0.f_stat,T0.model,T0.cdate desc"
            End If
            ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
            conn.Close()

            '已建立Sap 未結案且是母工單
            SelectCmd = "Select T0.[DocNum] " &
                "FROM dbo.OWOR T0 " &
                "WHERE T0.Warehouse='" & Session("usingwhs") & " ' and T0.[DocNum] = T0.[U_F16] and T0.[Status] <> 'L' " &
                "And T0.[Status] <> 'C' " &
                "ORDER BY T0.docnum,T0.itemcode"
            drsap = CommUtil.SelectSapSqlUsingDr(SelectCmd, connsap)
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    'check sap工單是否已建立wsn
                    SelectCmd = "Select T0.wsn " &
                                "From dbo.[worksn] T0 where T0.docnum=" & drsap(0)
                    dr = CommUtil.SelectLocalSqlUsingDr(SelectCmd, conn)
                    If (Not dr.HasRows) Then
                        conn.Close()
                        '已建立Sap 但未建立wsn
                        SelectCmd = "Select T0.[DocNum],T0.[PlannedQty] As model_set,T0.[DueDate] As ship_date,convert(char(12),T0.[PostDate],111) As cdate " &
                        "FROM dbo.OWOR T0 " &
                        "WHERE T0.[DocNum] = " & drsap(0)
                        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SelectCmd, conn)
                    End If
                    dr.Close()
                    conn.Close()
                Loop
            End If
            drsap.Close()
            connsap.Close()
        End If

        '---------------出貨數不等於製作數且sap 已建之工單
        If (Session("usingwhs") = "C01") Then
            SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                         "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                         "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.cdate,T0.ship_date,T0.docnum " &
                         "From dbo.[worksn] T0 where Left(T0.wsn,1)='I' and T0.docnum<> 0 and T0.model_set<>T0.ship_set " & rule & " order by T0.model,T0.cdate desc"
            ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
            conn.Close()
        Else
            If (Request.QueryString("sort") = "getpo") Then
                SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                         "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                         "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.cdate,T0.ship_date,T0.docnum " &
                         "From dbo.[worksn] T0 where (getpo=20 or getpo=70) and Left(T0.wsn,1)<>'I' and T0.docnum<> 0 and T0.model_set<>T0.ship_set " & rule & " order by T0.model,T0.cdate desc"
                ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
                conn.Close()
                SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
             "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
             "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.cdate,T0.ship_date,T0.docnum " &
             "From dbo.[worksn] T0 where getpo<>20 and getpo<>70 and Left(T0.wsn,1)<>'I' and T0.docnum<> 0 and T0.model_set<>T0.ship_set " & rule & " order by T0.model,T0.cdate desc"
                ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
                conn.Close()
            Else
                SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                         "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                         "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.cdate,T0.ship_date,T0.docnum " &
                         "From dbo.[worksn] T0 where Left(T0.wsn,1)<>'I' and T0.docnum<> 0 and T0.model_set<>T0.ship_set " & rule & " order by T0.model,T0.cdate desc"
                ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
                conn.Close()
            End If
        End If


        '----------------------------------------------------
        '--------------- 已結案且sap 已建之工單
        If (rule <> "") Then
            If (Session("usingwhs") = "C01") Then
                SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                         "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                         "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.cdate,T0.ship_date,T0.docnum " &
                         "From dbo.[worksn] T0 where Left(T0.wsn,1)='I' and T0.docnum<> 0 and T0.model_set=T0.ship_set " & rule & " order by T0.model,T0.cdate desc"
            Else
                SelectCmd = "Select T0.wsn , T0.getpo , T0.cus_name , T0.company ,T0.model , " &
                         "T0.resolution , T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                         "T0.note ,T0.comm,T0.model_set,T0.mfmes,T0.cdate,T0.ship_date,T0.docnum " &
                         "From dbo.[worksn] T0 where Left(T0.wsn,1)<>'I' and T0.docnum<> 0 and T0.model_set=T0.ship_set " & rule & " order by T0.model,T0.cdate desc"
            End If
            ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SelectCmd, conn)
            conn.Close()
        End If
        '--------------------------------------------------
        If (ds.Tables(0).Rows.Count <> 0) Then
            ds.Tables(0).Columns.Add("cno")
        Else
            CommUtil.ShowMsg(Me, "無任何資料")
        End If
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        'If (Session("s_id") = "ron") Then
        ShowInfo()
        'End If
    End Sub

    Sub ShowInfo()
        If (InStr(permsmf000, "s")) Then
            Label1.Visible = True
            Dim SqlCmd As String
            Dim nowdate, str(), BeginDate, EndDate As String
            Dim fromyear, frommonth, toyear, tomonth, lastmonth As String
            Dim totalaoiship, totalspiship, thismonthship, lastmonthship As Integer
            If (TxtBeginDate.Text = "" And TxtEndDate.Text = "") Then
                nowdate = Format(Now(), "yyyy/MM/dd")
                str = Split(nowdate, "/")
                '上月出貨
                If ((CInt(str(1)) - 1) = 0) Then
                    fromyear = CStr(CInt(str(0)) - 1)
                    frommonth = "12"
                    toyear = str(0)
                    tomonth = "01"
                Else
                    fromyear = str(0)
                    frommonth = CStr(CInt(str(1)) - 1)
                    If (Len(frommonth) = 1) Then
                        frommonth = "0" & frommonth
                    End If
                    toyear = str(0)
                    tomonth = str(1)
                End If
                lastmonth = frommonth
                BeginDate = fromyear & "/" & frommonth & "/01"
                EndDate = toyear & "/" & tomonth & "/01"
                'MsgBox(BeginDate & "~" & EndDate)
                If (Session("usingwhs") = "C01") Then
                    SqlCmd = "Select count(*) " &
                "From dbo.[work_records] T0 where status >='" & BeginDate & "' and status <'" & EndDate & "' and " &
                "dpart=5 and iseq=4 and Left(T0.wsn,1)='I'"
                Else
                    SqlCmd = "Select count(*) " &
                "From dbo.[work_records] T0 where status >='" & BeginDate & "' and status <'" & EndDate & "' and " &
                "dpart=5 and iseq=4 and Left(T0.wsn,1)<>'I'"
                End If
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                lastmonthship = dr(0)
                dr.Close()
                conn.Close()
                '本月出貨
                fromyear = str(0)
                frommonth = str(1)
                If ((CInt(str(1)) + 1) > 12) Then
                    toyear = CStr(CInt(str(0)) + 1)
                    tomonth = "01"
                Else
                    toyear = str(0)
                    tomonth = CStr(CInt(str(1)) + 1)
                    If (Len(tomonth) = 1) Then
                        tomonth = "0" & tomonth
                    End If
                End If
                BeginDate = fromyear & "/" & frommonth & "/01"
                EndDate = toyear & "/" & tomonth & "/01"
                'MsgBox(BeginDate & "~" & EndDate)
                If (Session("usingwhs") = "C01") Then
                    SqlCmd = "Select count(*) " &
                "From dbo.[work_records] T0 where status >='" & BeginDate & "' and status <'" & EndDate & "' and " &
                "dpart=5 and iseq=4 and Left(T0.wsn,1)='I'"
                Else
                    SqlCmd = "Select count(*) " &
                "From dbo.[work_records] T0 where status >='" & BeginDate & "' and status <'" & EndDate & "' and " &
                "dpart=5 and iseq=4 and Left(T0.wsn,1)<>'I'"
                End If
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                thismonthship = dr(0)
                dr.Close()
                conn.Close()

                If (Session("usingwhs") = "C02") Then
                    BeginDate = str(0) & "/01/01"
                    SqlCmd = "Select count(*) " &
                    "From dbo.[work_records] T0 where status >='" & BeginDate & "' and status <='" & nowdate & "' and " &
                    "dpart=5 and iseq=4 and Left(T0.wsn,1)='O'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    totalaoiship = dr(0)
                    dr.Close()
                    conn.Close()
                    SqlCmd = "Select count(*) " &
                    "From dbo.[work_records] T0 where status >='" & BeginDate & "' and status <='" & nowdate & "' and " &
                    "dpart=5 and iseq=4 and Left(T0.wsn,1)='S'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    totalspiship = dr(0)
                    dr.Close()
                    conn.Close()

                    Label1.Text = "上月(" & lastmonth & "月)出貨: " & CStr(lastmonthship) & "台  本月(" & str(1) & "月)出貨: " & CStr(thismonthship) & "台" &
                              "   今年(" & str(0) & "年)到現在共出貨:" & CStr(totalaoiship + totalspiship) & "台(AOI: " & CStr(totalaoiship) & "  SPI:" &
                              " " & CStr(totalspiship) & ")"
                Else
                    Label1.Text = "上月(" & lastmonth & "月)出貨: " & CStr(lastmonthship) & "台ICT  本月(" & str(1) & "月)出貨: " & CStr(thismonthship) & "台ICT" &
                              "   今年(" & str(0) & "年)到現在共出貨:" & CStr(totalaoiship + totalspiship) & "台ICT"
                End If
            Else
                If (Session("usingwhs") = "C01") Then
                    SqlCmd = "Select count(*) " &
                    "From dbo.[work_records] T0 where status >='" & TxtBeginDate.Text & "' and status <'" & TxtEndDate.Text & "' and " &
                    "dpart=5 and iseq=4 and Left(T0.wsn,1)='I'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    Label1.Text = TxtBeginDate.Text & "~" & TxtEndDate.Text & "共出貨:" & CStr(dr(0)) & "台ICT"
                Else
                    SqlCmd = "Select count(*) " &
                    "From dbo.[work_records] T0 where status >='" & TxtBeginDate.Text & "' and status <'" & TxtEndDate.Text & "' and " &
                    "dpart=5 and iseq=4 and Left(T0.wsn,1)<>'I'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    Label1.Text = TxtBeginDate.Text & "~" & TxtEndDate.Text & "共出貨:" & CStr(dr(0)) & "台AOI"
                End If
                dr.Close()
                conn.Close()
            End If
        Else
            Label1.Visible = False
        End If
    End Sub
    Protected Sub gv1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim itemcode As String
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='lightgreen'")
            '設定光棒顏色，當滑鼠 onMouseOver 時驅動
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            '當 onMouseOut 也就是滑鼠移開時，要恢復原本的顏色
            Dim SqlCmd As String
            Dim btnx As Button

            Dim Hyper, Hyper1, Hyper2, Hyper3, HyperFile As New HyperLink
            If (e.Row.Cells(0).Text <> 0) Then
                Hyper1.Text = e.Row.Cells(0).Text
                Hyper1.NavigateUrl = "wolist.aspx?modocnum=" & e.Row.Cells(0).Text & "&wsn=" & e.Row.Cells(1).Text &
                                     "&smid=molist&smode=0&indexpage=" & gv1.PageIndex
            End If
            Hyper1.Font.Underline = False
            CommUtil.DisableObjectByPermission(Hyper1, permsmf202, "e")
            e.Row.Cells(0).Controls.Add(Hyper1)

            If (e.Row.Cells(1).Text <> "NA") Then '此欄位設定null顯示NA
                Hyper.Text = e.Row.Cells(1).Text
                Hyper.NavigateUrl = "moadd_sys.aspx?wsn=" & e.Row.Cells(1).Text & "&mode=modify" &
                                    "&smid=molist&smode=0&indexpage=" & gv1.PageIndex
                CommUtil.DisableObjectByPermission(Hyper, permsmf201, "e")
            Else
                Hyper.Text = "建立"
                Hyper.NavigateUrl = "moadd_sys.aspx?smid=molist&smode=4&mode=create&docnum=" & e.Row.Cells(0).Text & '這裡的 smid 及 mode 數值要配合Mysite1.Master之設定
                                    "&smid=molist&smode=0&indexpage=" & gv1.PageIndex
                CommUtil.DisableObjectByPermission(Hyper, permsmf201, "m")
                SqlCmd = "Select T0.[itemcode] " &
                        "FROM dbo.OWOR T0 " &
                        "WHERE T0.[DocNum] = " & e.Row.Cells(0).Text
                drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                itemcode = ""
                If (drsap.HasRows) Then
                    drsap.Read()
                    itemcode = drsap(0)
                    drsap.Close()
                    connsap.Close()
                End If
                SqlCmd = "SELECT T0.u_model " &
                             "FROM dbo.[@UMMD] T0 where T0.u_mcode='" & Left(itemcode, 4) & "'"
                drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (drsap.HasRows) Then
                    drsap.Read()
                    e.Row.Cells(5).Text = drsap(0)
                End If
                drsap.Close()
                connsap.Close()
            End If
            Hyper.Font.Underline = False

            e.Row.Cells(1).Controls.Add(Hyper)

            If (e.Row.Cells(1).Text <> "NA") Then
                If (e.Row.Cells(4).Text = "1") Then
                    e.Row.Cells(4).Text = "捷智"
                ElseIf (e.Row.Cells(4).Text = "2") Then
                    e.Row.Cells(4).Text = "捷智通"
                ElseIf (e.Row.Cells(4).Text = "3") Then
                    e.Row.Cells(4).Text = "捷豐"
                End If

                If (e.Row.Cells(2).Text = "10") Then
                    Hyper2.Text = "未簽單"
                    e.Row.Cells(2).BackColor = Drawing.Color.White
                ElseIf (e.Row.Cells(2).Text = "20") Then
                    Hyper2.Text = "聯絡單"
                    e.Row.Cells(2).BackColor = Drawing.Color.Yellow
                ElseIf (e.Row.Cells(2).Text = "30") Then
                    Hyper2.Text = "備庫"
                    e.Row.Cells(2).BackColor = Drawing.Color.LightSalmon
                ElseIf (e.Row.Cells(2).Text = "40") Then
                    Hyper2.Text = "Demo"
                    e.Row.Cells(2).BackColor = Drawing.Color.LightBlue
                ElseIf (e.Row.Cells(2).Text = "50") Then
                    Hyper2.Text = "研發用"
                    e.Row.Cells(2).BackColor = Drawing.Color.LightCyan
                ElseIf (e.Row.Cells(2).Text = "60") Then
                    Hyper2.Text = "訂單暫停"
                    e.Row.Cells(2).BackColor = Drawing.Color.Red
                ElseIf (e.Row.Cells(2).Text = "70") Then
                    Hyper2.Text = "暫停領料"
                    e.Row.Cells(2).BackColor = Drawing.Color.LightPink
                End If
                Hyper2.NavigateUrl = "molist.aspx?upditem=pofrom&wsn=" & e.Row.Cells(1).Text & "&postatus=" &
                                     e.Row.Cells(2).Text & "&smid=molist&smode=" & smode & "&indexpage=" & gv1.PageIndex &
                                     "&sort=" & Request.QueryString("sort")
                Hyper2.Font.Underline = False
                CommUtil.DisableObjectByPermission(Hyper2, permsmf000, "m")
                e.Row.Cells(2).Controls.Add(Hyper2)

                'If (e.Row.Cells(13).Text = "10") Then
                '    Hyper3.Text = "未開工"
                '    e.Row.Cells(13).BackColor = Drawing.Color.White
                'ElseIf (e.Row.Cells(13).Text = "20") Then
                '    Hyper3.Text = "組裝中"
                '    e.Row.Cells(13).BackColor = Drawing.Color.Yellow
                'ElseIf (e.Row.Cells(13).Text = "30") Then
                '    Hyper3.Text = "佈線中"
                '    e.Row.Cells(13).BackColor = Drawing.Color.LightSalmon
                'ElseIf (e.Row.Cells(13).Text = "40") Then
                '    Hyper3.Text = "測試中"
                '    e.Row.Cells(13).BackColor = Drawing.Color.LightPink
                'ElseIf (e.Row.Cells(13).Text = "50") Then
                '    Hyper3.Text = "檢驗中"
                '    e.Row.Cells(13).BackColor = Drawing.Color.LightPink
                'ElseIf (e.Row.Cells(13).Text = "60") Then
                '    Hyper3.Text = "已完成"
                '    e.Row.Cells(13).BackColor = Drawing.Color.LimeGreen
                'ElseIf (e.Row.Cells(13).Text = "70") Then
                '    Hyper3.Text = "已出貨"
                '    e.Row.Cells(13).BackColor = Drawing.Color.LimeGreen
                'End If
                'Hyper3.Font.Underline = False
                'CommUtil.DisableObjectByPermission(Hyper3, permsmf000, "e")
                'Hyper3.NavigateUrl = "molist.aspx?upditem=updstatus&wsn=" & e.Row.Cells(1).Text & "&postatus=" &
                '              e.Row.Cells(13).Text & "&indexpage=" & gv1.PageIndex & "&smid=molist&smode=" & smode
                'e.Row.Cells(13).Controls.Add(Hyper3)
                If (CInt(e.Row.Cells(7).Text) = 0) Then
                    e.Row.Cells(13).Text = "已轉出"
                ElseIf (CInt(e.Row.Cells(9).Text) = CInt(e.Row.Cells(7).Text)) Then
                    e.Row.Cells(13).Text = "已出貨"
                    e.Row.Cells(13).BackColor = Drawing.Color.LightBlue

                ElseIf (CInt(e.Row.Cells(8).Text) = CInt(e.Row.Cells(7).Text)) Then
                    e.Row.Cells(13).Text = "已完工"
                    e.Row.Cells(13).BackColor = Drawing.Color.LightGreen
                Else
                    SqlCmd = "Select count(*) " &
                    "From dbo.[work_records] T0 where wsn='" & e.Row.Cells(1).Text & "' and (status='進行中' or " &
                    "status='已完工' or status='已包裝' or status='已出貨' or status='已完成') and dpart<>1"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    If (dr(0) <> 0) Then
                        e.Row.Cells(13).Text = "開工中"
                        e.Row.Cells(13).BackColor = Drawing.Color.Yellow
                        Hyper3.Text = "看進度"
                        Hyper3.NavigateUrl = "wostamodify.aspx?wsn=" & e.Row.Cells(1).Text & "&mode=show&source=frommolist&indexpage=" & gv1.PageIndex
                        Hyper3.Font.Underline = True
                        e.Row.Cells(13).Controls.Add(Hyper3)
                    Else
                        e.Row.Cells(13).Text = "未開工"
                        Hyper3.Text = "未開工"
                        Hyper3.NavigateUrl = "wostamodify.aspx?wsn=" & e.Row.Cells(1).Text & "&mode=show&source=frommolist&indexpage=" & gv1.PageIndex
                        Hyper3.Font.Underline = True
                        e.Row.Cells(13).Controls.Add(Hyper3)
                    End If
                    conn.Close()
                End If
                If (e.Row.Cells(6).Text = "1") Then
                    e.Row.Cells(6).Text = "20um"
                ElseIf (e.Row.Cells(6).Text = "2") Then
                    e.Row.Cells(6).Text = "15um"
                ElseIf (e.Row.Cells(6).Text = "3") Then
                    e.Row.Cells(6).Text = "12um"
                ElseIf (e.Row.Cells(6).Text = "4") Then
                    e.Row.Cells(6).Text = "10um"
                ElseIf (e.Row.Cells(6).Text = "5") Then
                    e.Row.Cells(6).Text = "8um"
                ElseIf (e.Row.Cells(6).Text = "6") Then
                    e.Row.Cells(6).Text = "7um"
                ElseIf (e.Row.Cells(6).Text = "7") Then
                    e.Row.Cells(6).Text = "6um"
                ElseIf (e.Row.Cells(6).Text = "8") Then
                    e.Row.Cells(6).Text = "5.5um"
                ElseIf (e.Row.Cells(6).Text = "9") Then
                    e.Row.Cells(6).Text = "3um"
                End If
            End If
            e.Row.Cells(12).ToolTip = e.Row.Cells(12).Text
            If (e.Row.Cells(12).Text.Length > 6) Then
                e.Row.Cells(12).Text = e.Row.Cells(12).Text.Substring(0, 5) + "..."
            End If

            e.Row.Cells(14).ToolTip = e.Row.Cells(14).Text
            If (e.Row.Cells(14).Text.Length > 6) Then
                e.Row.Cells(14).Text = e.Row.Cells(14).Text.Substring(0, 5) + "..."
            End If

            e.Row.Cells(15).ToolTip = e.Row.Cells(15).Text
            If (e.Row.Cells(15).Text.Length > 6) Then
                e.Row.Cells(15).Text = e.Row.Cells(15).Text.Substring(0, 5) + "..."
            End If

            If (CInt(e.Row.Cells(8).Text) > 0 And CInt(e.Row.Cells(8).Text) < CInt(e.Row.Cells(7).Text)) Then
                e.Row.Cells(8).BackColor = Drawing.Color.Yellow
            ElseIf (CInt(e.Row.Cells(8).Text) = CInt(e.Row.Cells(7).Text)) Then
                e.Row.Cells(8).BackColor = Drawing.Color.LightGreen
            End If
            SqlCmd = "Select status,count(*) " &
                    "From dbo.[work_records] T0 where wsn='" & e.Row.Cells(1).Text & "' and status<>'' and " &
                    "dpart=5 and iseq=4 group by status"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            Do While (dr.Read())
                e.Row.Cells(9).ToolTip = e.Row.Cells(9).ToolTip + dr(0) + " *" + CStr(dr(1)) + vbCrLf
            Loop
            dr.Close()
            conn.Close()
            If (CInt(e.Row.Cells(9).Text) > 0 And CInt(e.Row.Cells(9).Text) < CInt(e.Row.Cells(7).Text)) Then
                e.Row.Cells(9).BackColor = Drawing.Color.Yellow
            ElseIf (CInt(e.Row.Cells(9).Text) = CInt(e.Row.Cells(7).Text)) Then
                e.Row.Cells(9).BackColor = Drawing.Color.LightBlue
            End If

            If (e.Row.Cells(1).Text <> "NA") Then
                Dim targetDir As String
                '---Dim appPath As String
                Dim targetPath As String
                Dim filename As String
                'filename = e.Row.Cells(1).Text & ".pdf"
                filename = GetCLAttachedFileName("current", e.Row.Cells(1).Text) & ".pdf"
                'targetDir = Application("localdir") & "CLFormFile\" '"C:\SapErp\Uploads\"
                targetDir = HttpContext.Current.Server.MapPath("~/") & "AttachFile\CLFormFile\"
                '----appPath = Request.PhysicalApplicationPath '應用程式目錄
                targetPath = targetDir & filename

                btnx = New Button
                btnx.ID = e.Row.Cells(1).Text
                btnx.Font.Size = 10
                If (System.IO.File.Exists(targetPath)) Then
                    If (ChkDel.Checked) Then
                        btnx.Text = "刪除"
                        btnx.BackColor = Drawing.Color.Red
                    Else
                        btnx.Text = "下載"
                        btnx.BackColor = Drawing.Color.LightGreen
                    End If
                Else
                    btnx.Text = "上傳"
                End If
                btnx.Width = 40
                btnx.Height = 20
                AddHandler btnx.Click, AddressOf btnx_Click
                e.Row.Cells(16).Controls.Add(btnx)
            End If
        ElseIf (e.Row.RowType = DataControlRowType.Header) Then
            Dim HyperHead As HyperLink
            HyperHead = New HyperLink
            HyperHead.Text = e.Row.Cells(2).Text
            HyperHead.NavigateUrl = "molist.aspx?sort=getpo&smid=molist&smode=" & smode & "&indexpage=" & gv1.PageIndex
            HyperHead.Font.Underline = True
            HyperHead.ForeColor = Drawing.Color.White
            e.Row.Cells(2).Controls.Add(HyperHead)

            HyperHead = New HyperLink
            HyperHead.Text = e.Row.Cells(5).Text
            HyperHead.NavigateUrl = "molist.aspx?sort=model&smid=molist&smode=" & smode & "&indexpage=" & gv1.PageIndex
            HyperHead.Font.Underline = True
            HyperHead.ForeColor = Drawing.Color.White
            e.Row.Cells(5).Controls.Add(HyperHead)
        End If
    End Sub
    Protected Sub dDDLPOPurPose_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim dDDL As DropDownList
        Dim sqlresult As Boolean
        Dim str, str1 As String
        str = "A21100701"
        str1 = "研發用"
        'InitLocalSQLConnection()
        dDDL = Me.FindControl("dDDLP_1")
        'CommUtil.ShowMsg(Me,dDDL.SelectedValue)
        SqlCmd = "update dbo.[worksn] set getpo= '訂單暫停' where wsn='" & str & "'"
        'myCommand = New SqlCommand(SqlCmd, conn)
        'count = myCommand.ExecuteNonQuery()
        sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        If (sqlresult) Then
            Response.Redirect("molist.aspx?smid=molist&smode=1")
        Else
            CommUtil.ShowMsg(Me, "更新失敗")
        End If
        conn.Close()
    End Sub

    Protected Sub gv1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        If (TxtBeginDate.Text = "" And TxtEndDate.Text = "") Then
            GetNormalWo(wotype)
        Else
            GetShippedWo()
        End If
        'GetNormalWo(wotype)
    End Sub
    Function GetCLAttachedFileName(type As String, wsn As String)
        Dim filename As String
        Dim conn As New SqlConnection
        Dim dr As SqlDataReader
        filename = ""
        SqlCmd = "Select attachfileno " &
        "from [dbo].[worksn] " &
        "where wsn='" & wsn & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            If (type = "current") Then
                If (dr(0) = 0) Then
                    filename = wsn
                Else
                    filename = wsn & "(" & CStr(dr(0)) & ")"
                End If
            Else 'next
                filename = wsn & "(" & CStr(dr(0) + 1) & ")"
            End If
        End If
        dr.Close()
        conn.Close()
        Return filename
    End Function
    Protected Sub btnx_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim targetDir, CLHttpFile As String
        'Dim appPath As String
        Dim targetPath As String
        Dim filename, nameext As String
        Dim str() As String
        Dim p As New Process()
        Dim url As String
        url = Application("http")
        targetDir = Application("localdir") & "CLFormFile\"
        'filename = sender.ID
        'appPath = Request.PhysicalApplicationPath '應用程式目錄
        If (sender.Text = "上傳") Then
            If (FileUL.HasFile) Then
                str = Split(FileUL.FileName, ".")
                nameext = str(1)
                If (nameext <> "pdf") Then
                    CommUtil.ShowMsg(Me, "要上傳檔案需為pdf")
                    Exit Sub
                End If
                filename = GetCLAttachedFileName("next", sender.ID)
                targetPath = targetDir & filename & ".pdf"
                CLHttpFile = HttpContext.Current.Server.MapPath("~/") & "AttachFile\CLFormFile\" & filename & ".pdf"
                FileUL.SaveAs(targetPath)
                FileUL.SaveAs(CLHttpFile)
                CommUtil.ShowMsg(Me, "檔案上傳成功")
                sender.Text = "下載"
                sender.BackColor = Drawing.Color.LightGreen
                SqlCmd = "update dbo.[worksn] set attachfileno= attachfileno+1 where wsn='" & sender.ID & "'"
                CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
            Else
                CommUtil.ShowMsg(Me, "未指定上傳檔案")
            End If
        ElseIf (sender.Text = "下載") Then
            'Dim cacheEnum As IDictionaryEnumerator '= Cache.GetEnumerator()
            'cacheEnum = Cache.GetEnumerator()
            'Do While (cacheEnum.MoveNext())
            '    Cache.Remove(cacheEnum.Key.ToString())
            'Loop
            'p.StartInfo.FileName = Application("localdir") & "copyfile.bat"
            'p.StartInfo.Arguments = " " & Application("localdir") & "MachineCL\" & filename & ".pdf " & HttpContext.Current.Server.MapPath("~/") & "TempFile\" & filename & ".pdf"
            'p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized 'WindowStyle可以設定開啟視窗的大小
            'p.Start()
            'p.WaitForExit(3000)
            'p.Close()
            'p.Dispose()
            filename = GetCLAttachedFileName("current", sender.ID)
            Dim tpath As String = url & "AttachFile/CLFormFile/" & filename & ".pdf"
            ScriptManager.RegisterStartupScript(Me, Page.GetType, "Script", "showDisplay1('" & tpath & "');", True)

        ElseIf (sender.Text = "刪除") Then
            filename = GetCLAttachedFileName("current", sender.ID)
            targetPath = targetDir & filename & ".pdf"
            CLHttpFile = HttpContext.Current.Server.MapPath("~/") & "AttachFile\CLFormFile\" & filename & ".pdf"
            IO.File.Delete(targetPath)
            IO.File.Delete(CLHttpFile)
            sender.Text = "上傳"
            ChkDel.Checked = False
            sender.BackColor = Nothing
            CommUtil.ShowMsg(Me, "刪除成功")
        End If
        ViewState("indexpage") = indexpage
    End Sub

    Protected Sub ChkDel_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        GetNormalWo(wotype) '為了讓button文字變成刪除
        ViewState("indexpage") = indexpage
    End Sub
End Class