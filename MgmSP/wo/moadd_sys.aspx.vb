Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit

Partial Public Class moadd_sys
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public connsap, connsap1 As New SqlConnection
    Public SqlCmd As String
    Public oCompany As New SAPbobsCOM.Company
    Public dr As SqlDataReader
    Public permsmf201 As String
    Public ret As Long
    Public TxtWsn As TextBox
    Public ChkDel As CheckBox
    Public indexpage As Integer
    Public ScriptManager1 As New ScriptManager
    'Public ScriptManager1 As New ScriptManager

    'Public Sub InitSAPSQLConnection1(ByVal DestIP As String, ByVal HostName As String)
    '    connsap.ConnectionString = "Data Source= " & DestIP & ";uid=sa;pwd=sap19690123;database=" & HostName
    '    connsap.Open()
    '    If (connsap.State <> 1) Then
    '        CommUtil.ShowMsg(Me,"連線失敗")
    '    End If
    'End Sub
    Public Function InitSAPConnection(ByVal DestIP As String, ByVal HostName As String) As Long
        oCompany.Server = DestIP
        oCompany.CompanyDB = HostName
        oCompany.UserName = Session("sapid")
        oCompany.Password = Session("sappwd")
        oCompany.UseTrusted = False
        oCompany.DbUserName = "sa"
        oCompany.DbPassword = "sap19690123"
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
        InitSAPConnection = oCompany.Connect
    End Function

    Public Sub CloseSAPConnection()
        oCompany.Disconnect()
    End Sub

    Public Sub CreateTable(ByVal tTable As Table, ByVal row As Integer, ByVal col As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tRow.Cells.Add(tCell)
            Next
            Me.Table1.Rows.Add(tRow)
        Next
        tTable.BorderStyle = BorderStyle.Solid
        tTable.GridLines = GridLines.Both
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        If (IsPostBack) Then
            indexpage = ViewState("indexpage")
        Else
            indexpage = Request.QueryString("indexpage")
            ViewState("indexpage") = indexpage
        End If

        permsmf201 = CommUtil.GetAssignRight("mf201", Session("s_id"))
        ShowWorkOrderOfCreate()
    End Sub

    Sub SetRowBackColor(ByVal tTable As Table, ByVal row As Integer)
        If (row Mod 2) Then
            tTable.Rows(row).BackColor = Drawing.Color.LightBlue
        Else
            tTable.Rows(row).BackColor = Drawing.Color.Azure
        End If
    End Sub

    Function GetSysWsn()
        Dim wsn, wsnstr As String
        Dim i As Integer
        Dim dDDL As DropDownList
        Dim model_str() As String
        wsnstr = ""
        wsn = ""
        dDDL = Table1.FindControl(ViewState("model_mo"))
        model_str = Split(dDDL.SelectedValue, "-")
        'InitLocalSQLConnection()
        For i = 1 To 30
            wsnstr = Format(Now(), "yyMMdd")
            wsn = CInt(wsnstr) * 100 + i
            If (Left(model_str(0), 1) = "6") Then
                wsn = "S" & wsn
            ElseIf (Left(model_str(0), 1) = "7" Or Left(model_str(0), 1) = "8") Then
                wsn = "O" & wsn
            ElseIf (Left(model_str(0), 1) = "S") Then
                wsn = "F" & wsn
            Else
                wsn = "I" & wsn
            End If
            SqlCmd = "Select count(*) From dbo.[worksn] T0 where T0.wsn='" & wsn & "'"
            'myCommand = New SqlCommand(SqlCmd, conn)
            'dr = myCommand.ExecuteReader()
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            dr.Read()
            If (dr(0) = 0) Then
                Exit For
            End If
            dr.Close()
            conn.Close()
        Next
        Return wsn
    End Function

    Protected Sub dDDL_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        'Table1.Rows(2).Cells(1).Text = GetSysWsn()
        'ViewState("now_wsn") = Table1.Rows(2).Cells(1).Text
        TxtWsn.Text = GetSysWsn()
        ViewState("now_wsn") = TxtWsn.Text
        ViewState("indexpage") = indexpage
        'CommUtil.ShowMsg(Me,ViewState("now_wsn"))
    End Sub

    Sub ShowWorkOrderOfCreate()
        Dim ce As CalendarExtender
        'CreateTable(Table1, 2, 8)
        Dim rRBL As RadioButtonList
        Dim tTxt As TextBox
        Dim dDDL As DropDownList
        Dim lLbl As Label
        Dim cChk As CheckBox
        Dim tBtn As Button
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim wsn, mode, now_wsn As String
        Dim docnum As Long
        Dim drwsn, drsap, drsap1 As SqlDataReader
        Dim itemcode_flag, qty_flag, whs_flag, docnum_flag As Boolean
        Dim createdate As String
        Dim ddlindex, ti As Integer
        createdate = Format(Now(), "yyyy/MM/dd")
        itemcode_flag = False
        qty_flag = False
        whs_flag = False
        docnum_flag = False
        wsn = ""
        now_wsn = ""
        If (Not IsPostBack) Then
            'CommUtil.ShowMsg(Me,"Not")
            mode = Request.QueryString("mode")
            If (mode = "modify" Or mode = "modify1") Then
                wsn = Request.QueryString("wsn")
            ElseIf (mode = "create") Then
                docnum = Request.QueryString("docnum")
            Else
                docnum = 0
            End If
        Else
            'CommUtil.ShowMsg(Me,"Yes")
            mode = ViewState("mode")
            wsn = ViewState("wsn")
            docnum = ViewState("docnum")
            now_wsn = ViewState("now_wsn")
            'CommUtil.ShowMsg(Me,mode & "-" & wsn)
        End If

        If (mode = "modify" Or mode = "modify1") Then
            'InitLocalSQLConnection()
            SqlCmd = "Select  T0.docnum , T0.getpo , T0.cus_name , T0.company , T0.model , T0.resolution , " &
                            "T0.f_set , T0.ship_set , T0.camera_brand , T0.f_stat , " &
                            "T0.note , T0.comm , T0.creater, T0.sales,T0.model_set,T0.ship_label, " &
                            "T0.lens,T0.belt,T0.antibelt,T0.light,T0.mfmes,T0.cdate,T0.ship_date " &
                            "From dbo.[worksn] T0 where T0.wsn='" & wsn & "'"
            'myCommand = New SqlCommand(SqlCmd, conn)
            'drwsn = myCommand.ExecuteReader()
            drwsn = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            drwsn.Read()
            If (drwsn(0) <> 0) Then
                'InitSAPSQLConnection()
                SqlCmd = "SELECT T0.[DocNum], T0.[ItemCode], T0.[Comments], T0.[PlannedQty] , " &
                "T0.[CmpltQty], T0.[DueDate], T0.[Warehouse], T0.[Status] ,T0.[PostDate],T0.OriginNum " &
                "FROM dbo.OWOR T0 WHERE T0.[DocNum] ='" & drwsn(0) & "'"
                'myCommand = New SqlCommand(SqlCmd, connsap)
                'drsap = myCommand.ExecuteReader()
                drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                drsap.Read()
            End If
        ElseIf (mode = "create") Then ' create
            'InitSAPSQLConnection()
            SqlCmd = "SELECT T0.[DocNum], T0.[ItemCode], T0.[Comments], T0.[PlannedQty] , " &
            "T0.[CmpltQty], T0.[DueDate], T0.[Warehouse], T0.[Status] ,T0.[PostDate],T0.OriginNum " &
            "FROM dbo.OWOR T0 WHERE T0.[DocNum] ='" & docnum & "'"
            'myCommand = New SqlCommand(SqlCmd, connsap)
            'drsap = myCommand.ExecuteReader()
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            drsap.Read()
        End If
        Table1.BorderStyle = BorderStyle.Solid
        Table1.GridLines = GridLines.Both

        'Title
        i = 0
        tRow = New TableRow()
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = 6
        Table1.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).BackColor = Drawing.Color.Beige
        Table1.Rows(i).Font.Bold = True
        If (mode = "modify") Then
            'Table1.Rows(i).Cells(0).Text = "捷智科技工單--修改中..."
            lLbl = New Label
            lLbl.ID = "lLbl_" & i & i
            lLbl.Text = "捷智科技工單--修改中...&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
            Table1.Rows(i).Cells(0).Controls.Add(lLbl)
            ChkDel = New CheckBox
            ChkDel.ID = "Chk_del"
            ChkDel.Text = "刪除此單"
            ChkDel.ForeColor = Drawing.Color.Red
            CommUtil.DisableObjectByPermission(ChkDel, permsmf201, "d")
            Table1.Rows(i).Cells(0).Controls.Add(ChkDel)
        Else
            Table1.Rows(i).Cells(0).Text = "捷智科技工單--建立中..."
        End If

        'SAP 工號
        i = i + 1
        tRow = New TableRow()
        For j = 0 To 5
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(3).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(5).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        Table1.Rows(i).Cells(2).Font.Bold = True
        Table1.Rows(i).Cells(4).Font.Bold = True
        Table1.Rows(i).BackColor = Drawing.Color.Beige
        Table1.Rows(i).Cells(0).Text = "SAP工號"
        Table1.Rows(i).Cells(2).Text = "MO 料號"
        Table1.Rows(i).Cells(4).Text = "倉別"
        If (mode = "create1") Then
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & i & i
            ViewState("itemcode_mo") = tTxt.ID
            CommUtil.DisableObjectByPermission(tTxt, permsmf201, "n")
            Table1.Rows(i).Cells(3).Controls.Add(tTxt)
            itemcode_flag = True
        ElseIf (mode = "create") Then
            Table1.Rows(i).Cells(3).Text = drsap(1)
        Else
            'Table1.Rows(i).Cells(1).Text = drwsn(0)
            If (drwsn(0) <> 0) Then
                Table1.Rows(i).Cells(3).Text = drsap(1)
            Else
                tTxt = New TextBox()
                tTxt.ID = "tTxt_" & i & i
                ViewState("itemcode_mo") = tTxt.ID
                CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
                Table1.Rows(i).Cells(3).Controls.Add(tTxt)
                itemcode_flag = True
            End If
        End If
        ViewState("itemcode_flag") = itemcode_flag

        If (mode = "modify") Then
            If (drwsn(0) <> 0) Then
                Table1.Rows(i).Cells(5).Text = drsap(6)
            Else
                tTxt = New TextBox()
                tTxt.ID = "tTxt_" & i + 1 & i
                ViewState("whs_mo") = tTxt.ID
                CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
                Table1.Rows(i).Cells(5).Controls.Add(tTxt)
                whs_flag = True
            End If
        ElseIf (mode = "create") Then
            Table1.Rows(i).Cells(5).Text = drsap(6)
        Else 'create1
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & i + 1 & i
            tTxt.AutoPostBack = True
            ViewState("whs_mo") = tTxt.ID
            CommUtil.DisableObjectByPermission(tTxt, permsmf201, "n")
            Table1.Rows(i).Cells(5).Controls.Add(tTxt)
            whs_flag = True
            If (Session("usingwhs") = "C01") Then
                tTxt.Text = "C01"
            Else
                tTxt.Text = "C02"
            End If
        End If
        ViewState("whs_flag") = whs_flag
        'sap 單號 Text
        If (mode = "modify") Then
            tTxt = New TextBox() 'sap 單號 Text
            tTxt.ID = "tTxt_" & i + 2 & i
            If (drwsn(0) <> 0) Then
                tTxt.Text = drwsn(0)
                Table1.Rows(i).Cells(1).Text = drwsn(0)
            Else
                tTxt.Text = 0
                Table1.Rows(i).Cells(1).Text = 0
            End If
            docnum_flag = True
            ViewState("docnum_mo") = tTxt.ID
            CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
            Table1.Rows(i).Cells(1).Controls.Add(tTxt)
        ElseIf (mode = "create") Then
            'tTxt.Text = docnum
            Table1.Rows(i).Cells(1).Text = docnum
            ViewState("docnum_mo") = docnum
            docnum_flag = False
        ElseIf (mode = "create1") Then
            'tTxt.Text = 0
            Table1.Rows(i).Cells(1).Text = 0
            ViewState("docnum_mo") = 0
            docnum_flag = False
        End If
        ViewState("docnum_flag") = docnum_flag
        '以下開始各列資料
        i = i + 1
        tRow = New TableRow()
        For j = 0 To 5
            tCell = New TableCell()
            tCell.Width = New Unit("16.6%")
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(3).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(5).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        Table1.Rows(i).Cells(2).Font.Bold = True
        Table1.Rows(i).Cells(4).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "系統工號"
        Table1.Rows(i).Cells(2).Text = "開單日期"
        Table1.Rows(i).Cells(4).Text = "填寫人"

        TxtWsn = New TextBox()
        TxtWsn.ID = "txt_wsn"
        TxtWsn.Width = 100
        CommUtil.DisableObjectByPermission(TxtWsn, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(TxtWsn)
        If (mode = "modify") Then
            If (now_wsn = "") Then
                'Table1.Rows(i).Cells(1).Text = wsn
                TxtWsn.Text = wsn
            Else
                'Table1.Rows(i).Cells(1).Text = now_wsn
                TxtWsn.Text = now_wsn
            End If
            'If (drwsn(0) <> 0) Then
            Table1.Rows(i).Cells(3).Text = drwsn(21) 'drsap(8) 'postdate
            'Else

            'End If
            Table1.Rows(i).Cells(5).Text = drwsn(12)
        ElseIf (mode = "create") Then
            Table1.Rows(i).Cells(3).Text = createdate 'drsap(8)
            Table1.Rows(i).Cells(5).Text = Session("s_name")
        Else 'create1
            Table1.Rows(i).Cells(3).Text = createdate 'Now()
            Table1.Rows(i).Cells(5).Text = Session("s_name")
        End If
        ViewState("now_wsn") = TxtWsn.Text 'Table1.Rows(i).Cells(1).Text
        ViewState("createdate") = Table1.Rows(i).Cells(3).Text
        i = i + 1
        tRow = New TableRow()
        For j = 0 To 3
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).ColumnSpan = 3
        Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(3).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        Table1.Rows(i).Cells(2).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "訂單來源"
        Table1.Rows(i).Cells(2).Text = "負責業務"
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_" & i & i
        rRBL.Items.Add("台北捷智")
        rRBL.Items.Add("深圳捷智通")
        rRBL.Items.Add("昆山捷豐")
        rRBL.Items(0).Value = 1
        rRBL.Items(1).Value = 2
        rRBL.Items(2).Value = 3
        rRBL.RepeatDirection = RepeatDirection.Vertical
        If (mode = "modify") Then
            rRBL.SelectedValue = drwsn(3)
        End If
        ViewState("company_mo") = rRBL.ID
        CommUtil.DisableObjectByPermission(rRBL, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(rRBL)

        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i & i
        If (mode = "modify") Then
            tTxt.Text = drwsn(13)
        End If
        ViewState("sales_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
        Table1.Rows(i).Cells(3).Controls.Add(tTxt)

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 4
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).ColumnSpan = 2
        Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(3).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(4).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        Table1.Rows(i).Cells(2).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "捷智機型"
        Table1.Rows(i).Cells(2).Text = "客戶名稱"
        'InitLocalSQLConnection()
        'InitSAPSQLConnection1()
        If (Session("usingwhs") = "C02") Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                     "FROM dbo.[@UMMD] T0 where T0.u_mtype='SPI' or T0.u_mtype='AOI' or " &
                     "T0.u_mtype='3DAOI' order by T0.u_model,T0.u_mcode"
        ElseIf (Session("usingwhs") = "C01") Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                     "FROM dbo.[@UMMD] T0 where T0.u_mtype='ICT' order by T0.u_model,T0.u_mcode"
        Else
            CommUtil.ShowMsg(Me, "倉別設定須為C01 or C02已決定是ICT or AOI")
        End If
        dDDL = New DropDownList()
        dDDL.Items.Add("請選擇")
        If (Session("usingwhs") = "C02" Or Session("usingwhs") = "C01") Then
            drsap1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)

            ti = 1
            If (drsap1.HasRows) Then
                Do While (drsap1.Read())
                    dDDL.Items.Add(drsap1(0) & "-" & drsap1(1))
                    If (mode = "modify") Then
                        If (drsap1(0) = drwsn(4)) Then
                            dDDL.Items(ti).Selected = True
                            ddlindex = ti
                        End If
                    ElseIf (mode = "create") Then
                        'nothing
                    End If
                    ti = ti + 1
                Loop
            End If
            drsap1.Close()
            connsap1.Close()
        End If
        dDDL.ID = "ddl_" & i & i
        ViewState("model_mo") = dDDL.ID
        AddHandler dDDL.SelectedIndexChanged, AddressOf dDDL_SelectedIndexChanged
        dDDL.AutoPostBack = True
        CommUtil.DisableObjectByPermission(dDDL, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(dDDL)

        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i & i
        If (mode = "modify") Then
            tTxt.Text = drwsn(2)
        End If
        ViewState("cus_name_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
        Table1.Rows(i).Cells(3).Controls.Add(tTxt)

        lLbl = New Label
        lLbl.ID = "lLbl_" & i & i
        lLbl.Text = "數量:"
        If (mode = "modify") Then
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & i + 1 & i
            tTxt.Width = 40
            ViewState("plannedqty_mo") = tTxt.ID
            'If (drwsn(0) <> 0) Then
            'ViewState("plannedqty_mo") = CInt(drsap(3))
            'lLbl.Text = lLbl.Text & CInt(drsap(3))
            'Table1.Rows(i).Cells(4).Controls.Add(lLbl)
            'Else
            Table1.Rows(i).Cells(4).Controls.Add(lLbl)
            tTxt.Text = drwsn(14)
            CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
            Table1.Rows(i).Cells(4).Controls.Add(tTxt)
            qty_flag = True
            'End If
        ElseIf (mode = "create") Then
            lLbl.Text = lLbl.Text & CInt(drsap(3))
            Table1.Rows(i).Cells(4).Controls.Add(lLbl)
            ViewState("plannedqty_mo") = CInt(drsap(3))
        Else 'create1
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & i + 1 & i
            tTxt.Width = 40
            ViewState("plannedqty_mo") = tTxt.ID
            Table1.Rows(i).Cells(4).Controls.Add(lLbl)
            CommUtil.DisableObjectByPermission(tTxt, permsmf201, "n")
            Table1.Rows(i).Cells(4).Controls.Add(tTxt)
            qty_flag = True
        End If
        ViewState("qty_flag") = qty_flag
        'CommUtil.ShowMsg(Me,qty_flag & "-1")
        i = i + 1
        tRow = New TableRow()
        For j = 0 To 3
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).ColumnSpan = 2
        Table1.Rows(i).Cells(3).ColumnSpan = 2
        Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(3).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        Table1.Rows(i).Cells(2).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "機型標籤"
        Table1.Rows(i).Cells(2).Text = "出貨日期"
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i & i
        If (mode = "modify") Then
            tTxt.Text = drwsn(15)
        End If
        ViewState("ship_label_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "n")
        Table1.Rows(i).Cells(1).Controls.Add(tTxt)

        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i + 1 & i
        If (mode = "modify") Then
            'If (drwsn(0) <> 0) Then
            tTxt.Text = drwsn(22) 'drsap(5)
            'End If
        ElseIf (mode = "create") Then
            tTxt.Text = drsap(5)
        End If
        ViewState("duedate_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
        Table1.Rows(i).Cells(3).Controls.Add(tTxt)
        ce = New CalendarExtender
        ce.TargetControlID = tTxt.ID
        ce.ID = "ce_shipdate"
        ce.Format = "yyyy/MM/dd"
        Table1.Rows(i).Cells(3).Controls.Add(ce)

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 1
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).ColumnSpan = 7
        'Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "解析度"
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_" & i & i
        rRBL.Items.Add("20um")
        rRBL.Items.Add("15um")
        rRBL.Items.Add("12um")
        rRBL.Items.Add("10um")
        rRBL.Items.Add("8um")
        rRBL.Items.Add("7um")
        rRBL.Items.Add("6um")
        rRBL.Items.Add("5.5um")
        rRBL.Items.Add("3um")
        rRBL.Items.Add("其它")
        rRBL.Items.Add("無需求")
        rRBL.Items(0).Value = 1
        rRBL.Items(1).Value = 2
        rRBL.Items(2).Value = 3
        rRBL.Items(3).Value = 4
        rRBL.Items(4).Value = 5
        rRBL.Items(5).Value = 6
        rRBL.Items(6).Value = 7
        rRBL.Items(7).Value = 8
        rRBL.Items(8).Value = 9
        rRBL.Items(9).Value = 10
        rRBL.Items(10).Value = 11
        rRBL.RepeatDirection = RepeatDirection.Vertical
        If (mode = "modify") Then
            rRBL.SelectedValue = drwsn(5)
        End If
        'rRBL.Width = 400
        ViewState("resolution_mo") = rRBL.ID
        CommUtil.DisableObjectByPermission(rRBL, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(rRBL)
        If (Session("usingwhs") = "C02") Then
            rRBL.Enabled = True
        Else
            rRBL.Enabled = False
            rRBL.SelectedValue = 11
        End If
        'tTxt = New TextBox()
        'tTxt.ID = "tTxt_" & i
        'tTxt.Width = 40
        'Table1.Rows(i).Cells(2).Controls.Add(tTxt)

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = 8
        Table1.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).BackColor = Drawing.Color.Beige
        Table1.Rows(i).Font.Bold = True
        Table1.Rows(i).Cells(0).Text = "系統規格"

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 3
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).ColumnSpan = 2
        Table1.Rows(i).Cells(3).ColumnSpan = 2
        Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(3).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        Table1.Rows(i).Cells(2).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "相機"
        Table1.Rows(i).Cells(2).Text = "鏡頭"
        'InitSAPSQLConnection1()

        SqlCmd = "SELECT T0.itemname,T0.Itemcode " &
                     "FROM dbo.[OITM] T0 where T0.qrygroup50='Y' order by T0.itemcode"

        'myCommand = New SqlCommand(SqlCmd, connsap1)
        'drsap1 = myCommand.ExecuteReader()
        drsap1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap1)
        dDDL = New DropDownList()
            dDDL.Items.Add("請選擇")
            dDDL.Items.Add("無需求")
            ti = 2
            If (drsap1.HasRows) Then
                Do While (drsap1.Read())
                    dDDL.Items.Add(drsap1(0))
                    If (mode = "modify") Then
                        If (drsap1(0) = drwsn(8)) Then
                            dDDL.Items(ti).Selected = True
                            ddlindex = ti
                        End If
                    End If
                    ti = ti + 1
                Loop
            End If
            drsap1.Close()
            connsap1.Close()
            dDDL.ID = "ddl_" & i & i
            ViewState("camera_brand_mo") = dDDL.ID
            'AddHandler dDDL.SelectedIndexChanged, AddressOf dDDL_SelectedIndexChanged
            'dDDL.AutoPostBack = True
            CommUtil.DisableObjectByPermission(dDDL, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(dDDL)
        If (Session("usingwhs") = "C02") Then
            dDDL.Enabled = True
        Else
            dDDL.Enabled = False
            dDDL.SelectedIndex = 1
        End If

        'tTxt = New TextBox()
        'tTxt.ID = "tTxt_" & i & i
        'If (mode = "modify") Then
        '    tTxt.Text = drwsn(8)
        'End If
        ''If (mode = "modify") Then
        'ViewState("camera_brand_mo") = tTxt.ID
        ''End If
        'DisableObjectByPermission(tTxt, permsmf201, "nm")
        'Table1.Rows(i).Cells(1).Controls.Add(tTxt)

        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i + 1 & i
        If (mode = "modify") Then
            tTxt.Text = drwsn(16)
        ElseIf (mode = "create1") Then
            tTxt.Text = "無需求"
        End If
        ViewState("lens_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
        Table1.Rows(i).Cells(3).Controls.Add(tTxt)
        If (Session("usingwhs") = "C02") Then
            tTxt.Enabled = True
        Else
            tTxt.Enabled = False
        End If

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 3
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).ColumnSpan = 2
        Table1.Rows(i).Cells(3).ColumnSpan = 2
        'Table1.Rows(i).Cells(1).HorizontalAlign = HorizontalAlign.Center
        'Table1.Rows(i).Cells(3).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).Cells(0).Font.Bold = True
        Table1.Rows(i).Cells(2).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "軌道皮帶"
        Table1.Rows(i).Cells(2).Text = "燈盤選擇"
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_" & i & i
        rRBL.Items.Add("平皮帶")
        rRBL.Items.Add("時規皮帶")
        rRBL.Items.Add("無需求")
        rRBL.Items(0).Value = 1
        rRBL.Items(1).Value = 2
        rRBL.Items(2).Value = 3
        If (mode = "modify") Then
            rRBL.SelectedValue = drwsn(17)
        End If
        rRBL.RepeatDirection = RepeatDirection.Vertical
        ViewState("belt_mo") = rRBL.ID
        CommUtil.DisableObjectByPermission(rRBL, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(rRBL)
        If (Session("usingwhs") = "C02") Then
            rRBL.Enabled = True
        Else
            rRBL.Enabled = False
            rRBL.SelectedValue = 3
        End If

        cChk = New CheckBox()
        cChk.ID = "cChk_" & i & i
        cChk.Text = "防靜電需求"
        If (mode = "modify") Then
            cChk.Checked = drwsn(18)
        End If
        ViewState("antibelt_mo") = cChk.ID
        CommUtil.DisableObjectByPermission(cChk, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(cChk)
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_" & i + 1 & i
        rRBL.Items.Add("OPT")
        rRBL.Items.Add("V5")
        rRBL.Items.Add("V7")
        rRBL.Items.Add("無需求")
        rRBL.Items(0).Value = 1
        rRBL.Items(1).Value = 2
        rRBL.Items(2).Value = 3
        rRBL.Items(3).Value = 4
        If (mode = "modify") Then
            rRBL.SelectedValue = drwsn(19)
        End If
        rRBL.RepeatDirection = RepeatDirection.Vertical
        ViewState("light_mo") = rRBL.ID
        CommUtil.DisableObjectByPermission(rRBL, permsmf201, "m")
        Table1.Rows(i).Cells(3).Controls.Add(rRBL)
        If (Session("usingwhs") = "C02") Then
            rRBL.Enabled = True
        Else
            rRBL.Enabled = False
            rRBL.SelectedValue = 4
        End If
        i = i + 1
        tRow = New TableRow()
        For j = 0 To 1
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(1).ColumnSpan = 7
        Table1.Rows(i).Cells(0).Font.Bold = True
        SetRowBackColor(Table1, i)
        Table1.Rows(i).Cells(0).Text = "備註"
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i & i
        If (mode = "modify") Then
            tTxt.Text = drwsn(10)
        End If
        tTxt.Width = 700
        ViewState("note_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
        Table1.Rows(i).Cells(1).Controls.Add(tTxt)

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = 8
        Table1.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).BackColor = Drawing.Color.Beige
        Table1.Rows(i).Font.Bold = True
        Table1.Rows(i).Cells(0).Text = "聯絡訊息"

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = 8
        Table1.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i & i
        tTxt.Width = 1000
        tTxt.Height = 100
        If (mode = "modify") Then
            tTxt.Text = drwsn(11)
        End If
        tTxt.TextMode = TextBoxMode.MultiLine
        ViewState("comm_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
        Table1.Rows(i).Cells(0).Controls.Add(tTxt)

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = 8
        Table1.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).BackColor = Drawing.Color.Beige
        Table1.Rows(i).Font.Bold = True
        Table1.Rows(i).Cells(0).Text = "製造紀錄"

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = 8
        Table1.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & i & i
        tTxt.Width = 1000
        tTxt.Height = 50
        If (mode = "modify") Then
            tTxt.Text = drwsn(20)
        End If
        tTxt.TextMode = TextBoxMode.MultiLine
        ViewState("mfmes_mo") = tTxt.ID
        CommUtil.DisableObjectByPermission(tTxt, permsmf201, "m")
        Table1.Rows(i).Cells(0).Controls.Add(tTxt)

        i = i + 1
        tRow = New TableRow()
        For j = 0 To 0
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Table1.Rows.Add(tRow)
        Table1.Rows(i).Cells(0).ColumnSpan = 8
        Table1.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        Table1.Rows(i).BackColor = Drawing.Color.Goldenrod
        Table1.Rows(i).Font.Bold = True
        ViewState("gmode") = mode
        tBtn = New Button()
        tBtn.ID = "tBtn_" & i & i
        tBtn.Text = "儲存"
        CommUtil.DisableObjectByPermission(tBtn, permsmf201, "m")
        Table1.Rows(i).Cells(0).Controls.Add(tBtn)
        AddHandler tBtn.Click, AddressOf tBtn_Click
        cChk = New CheckBox()
        cChk.ID = "cChk_" & i & i
        cChk.Text = "產生SAP系統母工單"
        ViewState("genchk_mo") = cChk.ID
        CommUtil.DisableObjectByPermission(cChk, permsmf201, "n")
        Table1.Rows(i).Cells(0).Controls.Add(cChk)
        If (Table1.Rows(1).Cells(1).Text <> 0) Then
            cChk.Enabled = False
        End If

        'If (mode = "modify") Then
        ViewState("mode") = mode '"modify"
        ViewState("wsn") = wsn
        ViewState("docnum") = docnum
        'End If
        If (mode = "modify" Or mode = "modify1") Then
            If (drwsn(0) <> 0 Or mode = "create") Then
                drsap.Close()
                connsap.Close()
            End If
            drwsn.Close()
            conn.Close()
        ElseIf (mode = "create") Then
            drsap.Close()
            connsap.Close()
        End If
    End Sub

    Protected Sub tBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If (ViewState("mode") = "modify" Or ViewState("mode") = "modify1") Then
            If (ChkDel.Checked) Then
                SqlCmd = "delete from dbo.[worksn] where wsn='" & ViewState("wsn") & "'"
                CommUtil.SqlLocalExecute("del", SqlCmd, conn)
                conn.Close()
                SqlCmd = "delete from dbo.[work_records] where wsn='" & ViewState("wsn") & "'"
                CommUtil.SqlLocalExecute("del", SqlCmd, conn)
                conn.Close()
                Response.Redirect("molist.aspx?smid=molist&smode=1&indexpage=" & indexpage)
            End If
        End If
        Dim cChk As CheckBox
        Dim genchk_mo As Integer
        'cChk = Table1.FindControl(ViewState("genchk_mo"))
        'genchk_mo = cChk.Checked
        'If () Then
        'sender.OnClientClick = "return confirm('要複製嗎')"        '
        Dim rRBL As RadioButtonList
        Dim tTxt As TextBox
        Dim dDDL As DropDownList

        Dim str1() As String
        Dim docnum_mo As Long
        Dim company_mo, plannedqty_mo, resolution_mo, belt_mo, antibelt_mo As Integer
        Dim itemcode_mo, whs_mo, sales_mo, model_mo, cus_name_mo As String
        Dim camera_brand_mo, lens_mo, light_mo, note_mo, comm_mo, mfmes_mo As String
        Dim ship_label_mo, duedate_mo As String
        Dim mode, wsn, now_wsn, creater As String
        'Dim count As Integer
        'Dim stastr As String
        Dim createdate As String
        Dim sqlresult As Boolean
        createdate = ViewState("createdate") 'Format(Now(), "yyyy/MM/dd")
        ViewState("mode") = ViewState("mode")
        If (ViewState("docnum_flag")) Then
            tTxt = Table1.FindControl(ViewState("docnum_mo"))
            docnum_mo = CLng(tTxt.Text)
        Else
            docnum_mo = ViewState("docnum_mo")
        End If

        If (ViewState("itemcode_flag")) Then
            tTxt = Table1.FindControl(ViewState("itemcode_mo"))
            itemcode_mo = tTxt.Text
            If (tTxt.Text = "") Then
                CommUtil.ShowMsg(Me, "料號需填入")
                Exit Sub
            End If
        End If
        If (ViewState("whs_flag")) Then
            tTxt = Table1.FindControl(ViewState("whs_mo"))
            whs_mo = tTxt.Text
            If (tTxt.Text = "") Then
                CommUtil.ShowMsg(Me, "倉別需填入")
                Exit Sub
            End If
        End If
        rRBL = Table1.FindControl(ViewState("company_mo"))
        If (rRBL.SelectedValue = "") Then
            CommUtil.ShowMsg(Me, "訂單來源需選擇")
            Exit Sub
        End If
        company_mo = rRBL.SelectedValue
        tTxt = Table1.FindControl(ViewState("sales_mo"))
        sales_mo = tTxt.Text
        dDDL = Table1.FindControl(ViewState("model_mo"))
        If (dDDL.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "需選擇機型")
            Exit Sub
        End If
        str1 = Split(dDDL.SelectedValue, "-")
        model_mo = str1(0)

        tTxt = Table1.FindControl(ViewState("cus_name_mo"))
        cus_name_mo = tTxt.Text
        If (tTxt.Text = "") Then
            CommUtil.ShowMsg(Me, "客戶名須指定")
            Exit Sub
        End If

        If (ViewState("qty_flag")) Then
            tTxt = Table1.FindControl(ViewState("plannedqty_mo"))
            If (tTxt.Text = "") Then
                CommUtil.ShowMsg(Me, "數量需填入")
                Exit Sub
            End If
            plannedqty_mo = CInt(tTxt.Text)
        Else
            plannedqty_mo = ViewState("plannedqty_mo")
        End If

        tTxt = Table1.FindControl(ViewState("duedate_mo"))
        If (tTxt.Text = "") Then
            CommUtil.ShowMsg(Me, "出貨日期須指定")
            Exit Sub
        End If
        duedate_mo = tTxt.Text

        'CommUtil.ShowMsg(Me,plannedqty_mo)
        rRBL = Table1.FindControl(ViewState("resolution_mo"))
        If (rRBL.SelectedValue = "") Then
            CommUtil.ShowMsg(Me, "解析度需選擇")
            Exit Sub
        End If
        resolution_mo = rRBL.SelectedValue

        dDDL = Table1.FindControl(ViewState("camera_brand_mo"))
        camera_brand_mo = dDDL.Text
        tTxt = Table1.FindControl(ViewState("lens_mo"))
        lens_mo = tTxt.Text
        rRBL = Table1.FindControl(ViewState("belt_mo"))
        If (rRBL.SelectedValue = "") Then
            CommUtil.ShowMsg(Me, "皮帶需選擇")
            Exit Sub
        End If
        belt_mo = rRBL.SelectedValue
        cChk = Table1.FindControl(ViewState("antibelt_mo"))
        antibelt_mo = cChk.Checked
        rRBL = Table1.FindControl(ViewState("light_mo"))
        If (rRBL.SelectedValue = "") Then
            CommUtil.ShowMsg(Me, "燈盤需選擇")
            Exit Sub
        End If
        light_mo = rRBL.SelectedValue

        tTxt = Table1.FindControl(ViewState("note_mo"))
        note_mo = tTxt.Text
        tTxt = Table1.FindControl(ViewState("comm_mo"))
        comm_mo = tTxt.Text
        tTxt = Table1.FindControl(ViewState("mfmes_mo"))
        mfmes_mo = tTxt.Text
        cChk = Table1.FindControl(ViewState("genchk_mo"))
        genchk_mo = cChk.Checked
        tTxt = Table1.FindControl(ViewState("ship_label_mo"))
        ship_label_mo = tTxt.Text
        'tTxt = Table1.FindControl(ViewState("duedate_mo"))
        'duedate_mo = tTxt.Text

        mode = ViewState("mode")
        wsn = ViewState("wsn")
        'now_wsn = ViewState("now_wsn") '當不用textbox要enable
        now_wsn = TxtWsn.Text
        Dim ok As Boolean
        'Dim dpart, i, j, k As Integer
        ok = True
        Dim model_set_org, finishcount, ship_set As Integer
        finishcount = 0
        If (mode = "modify" Or mode = "modify1") Then
            SqlCmd = "select T0.model_set,T0.ship_set from dbo.worksn T0 where wsn='" & wsn & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                model_set_org = dr(0)
                ship_set = dr(1)
            End If
            dr.Close()
            connsap.Close()
            If (plannedqty_mo < ship_set) Then
                CommUtil.ShowMsg(Me, "已出貨" & ship_set & "台, 故數量修改不能小於出貨數")
                Exit Sub
            End If
            SqlCmd = "update dbo.[worksn] set company= " & company_mo & ",model_set=" & plannedqty_mo & ",resolution=" & resolution_mo & "," &
                     "belt=" & belt_mo & ",antibelt=" & antibelt_mo & ",sales='" & sales_mo & "'," &
                     "model='" & model_mo & "',cus_name='" & cus_name_mo & "',camera_brand='" & camera_brand_mo & "'," &
                     "lens='" & lens_mo & "',light=" & light_mo & ",note='" & note_mo & "',comm='" & comm_mo & "',mfmes='" & mfmes_mo & "'," &
                     "ship_label='" & ship_label_mo & "',docnum=" & docnum_mo & ",wsn='" & now_wsn & "',ship_date='" & duedate_mo & "' " &
                     "where wsn='" & wsn & "'"
            sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
            conn.Close()
            If (sqlresult = False) Then
                ok = False
                CommUtil.ShowMsg(Me, "更新失敗")
            Else
                If (wsn <> now_wsn) Then '更新work_records內wsn之records為now_wsn
                    SqlCmd = "update dbo.[work_records] set wsn='" & now_wsn & "' where wsn='" & wsn & "'"
                    CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                    conn.Close()
                    SqlCmd = "update dbo.[omri] set wsn='" & now_wsn & "' where wsn='" & wsn & "'"
                    CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                    conn.Close()
                End If
                If (model_set_org > plannedqty_mo) Then '數量減少
                    If (plannedqty_mo = 0) Then '如現在數量改為0 , 則關於此單work_records 都刪除
                        SqlCmd = "delete from dbo.[work_records] where wsn='" & now_wsn & "'"
                        CommUtil.SqlLocalExecute("del", SqlCmd, conn)
                        conn.Close()
                        SqlCmd = "update dbo.[worksn] set f_set=0 where wsn='" & now_wsn & "'"
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                    Else '若不為0 , 則刪除減少部分 , 並將原note 之wseq改掉
                        SqlCmd = "delete from dbo.[work_records] " &
                        "where wsn='" & now_wsn & "' and wseq <=" & model_set_org & " and wseq >" & plannedqty_mo
                        CommUtil.SqlLocalExecute("del", SqlCmd, conn)
                        conn.Close()
                        SqlCmd = "update dbo.[work_records] set wseq=" & plannedqty_mo + 1 &
                        " where wsn='" & now_wsn & "' and wseq=" & model_set_org + 1
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                        SqlCmd = "select count(*) from dbo.[work_records] where wsn='" & now_wsn & "' and dpart=5 and " &
                                "iseq=3 and (status='已完工' or status='已包裝')"
                        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                        dr.Read()
                        finishcount = dr(0)
                        dr.Close()
                        conn.Close()
                        SqlCmd = "update dbo.[worksn] set f_set=" & finishcount & " where wsn='" & now_wsn & "'"
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                    End If
                ElseIf (model_set_org < plannedqty_mo) Then
                    If (model_set_org = 0) Then '如原數量為0 , 則關於此單work_records 都要重建(含note)
                        GenWorkRecords("full", plannedqty_mo, now_wsn, 1)
                    Else '若不為0 , 則新建增加部分 , 並將原note 之wseq改掉
                        SqlCmd = "update dbo.[work_records] set wseq=" & plannedqty_mo + 1 &
                        " where wsn='" & now_wsn & "' and wseq=" & model_set_org + 1
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                        GenWorkRecords("partial", plannedqty_mo, now_wsn, model_set_org + 1)
                    End If
                End If
            End If
            'conn.Close()
        Else ' create mode
            If (now_wsn = "") Then 'kkkkk
                CommUtil.ShowMsg(Me, "需選擇機型產生自建碼")
                Exit Sub
            End If
            '以下要Insert資料到jtdb資料庫
            creater = Table1.Rows(2).Cells(5).Text
            SqlCmd = "Insert into dbo.[worksn] (company,model_set,resolution,belt,antibelt,sales,model,cus_name,camera_brand," &
                     "lens,light,note,comm,mfmes,ship_label,docnum,wsn,cdate,ship_date,creater)" &
                     "values(" & company_mo & "," & plannedqty_mo & "," & resolution_mo & "," & belt_mo & "," &
                     antibelt_mo & ",'" & sales_mo & "','" & model_mo & "','" & cus_name_mo & "','" & camera_brand_mo &
                     "','" & lens_mo & "','" & light_mo & "','" & note_mo & "','" & comm_mo & "','" & mfmes_mo &
                     "','" & ship_label_mo & "'," & docnum_mo & ",'" & now_wsn & "','" & createdate & "','" & duedate_mo & "','" & creater & "')"
            sqlresult = CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
            conn.Close()
            If (sqlresult) Then
                '產生 work_records
                ok = GenWorkRecords("full", plannedqty_mo, now_wsn, 1)
                'For dpart = 1 To 5
                '    If (dpart = 1 Or dpart = 2 Or dpart = 3 Or dpart = 4) Then
                '        j = 5
                '    ElseIf (dpart = 5) Then
                '        j = 4
                '    End If
                '    For i = 1 To j
                '        For k = 1 To plannedqty_mo + 1
                '            If (k = (plannedqty_mo + 1)) Then
                '                SqlCmd = "Insert into dbo.[work_records] (wsn,dpart,iseq,wseq,id) " &
                '                "values('" & now_wsn & "'," & dpart & "," & i & "," & k & ",'" & Session("s_id") & "')"
                '            Else
                '                If (dpart = 5) Then
                '                    If (i = 1 Or i = 2) Then
                '                        stastr = "未進行"
                '                    ElseIf (i = 3) Then
                '                        stastr = "未出貨"
                '                    Else
                '                        stastr = ""
                '                    End If
                '                Else
                '                    stastr = "未進行"
                '                End If
                '                SqlCmd = "Insert into dbo.[work_records] (wsn,dpart,iseq,wseq,id,status) " &
                '                "values('" & now_wsn & "'," & dpart & "," & i & "," & k & ",'" & Session("s_id") & "','" & stastr & "')"
                '            End If
                '            sqlresult = CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
                '            conn.Close()
                '            If (sqlresult = False) Then
                '                ok = False
                '                CommUtil.ShowMsg(Me, "新增系統工單料件狀態表失敗")
                '            End If
                '        Next
                '    Next
                'Next
            Else
                ok = False
                CommUtil.ShowMsg(Me, "新增系統工單失敗")
            End If
        End If
        If (genchk_mo) Then '產生sap 工單
            If (Table1.Rows(1).Cells(1).Text = 0) Then
                If (itemcode_mo <> "" And whs_mo <> "" And plannedqty_mo <> 0 And duedate_mo <> "") Then
                    SapWoGen(itemcode_mo, whs_mo, plannedqty_mo, duedate_mo, model_mo, cus_name_mo, camera_brand_mo, lens_mo, now_wsn) 'duedate:到期日期
                Else
                    CommUtil.ShowMsg(Me, "料號,倉別,到期(出貨)日期不能空白及數量不能為0")
                End If
            Else
                CommUtil.ShowMsg(Me, "SAP工單已存在, 不再產生")
            End If
        End If
        If (ok) Then
            Response.Redirect("molist.aspx?smid=molist&smode=1&indexpage=" & indexpage)
        End If
        'Session.Remove("postdate_mo")'刪除單一session
        'Session("postdate_mo") = Nothing''刪除單一session
        'Session.Abandon()''刪除所有session
    End Sub

    Function GenWorkRecords(mode As String, plannedqty As Integer, now_wsn As String, fromwseq As Integer)
        Dim i, j, k, dpart As Integer
        Dim stastr As String
        Dim sqlresult As Boolean
        Dim GenWoFlag As Boolean
        GenWoFlag = False
        For dpart = 1 To 5
            If (dpart = 1 Or dpart = 2 Or dpart = 3 Or dpart = 4) Then
                j = 5
            ElseIf (dpart = 5) Then
                j = 4
            End If
            For i = 1 To j
                For k = fromwseq To plannedqty + 1
                    If (k = (plannedqty + 1)) Then
                        If (mode <> "full") Then
                            Exit For
                        End If
                        SqlCmd = "Insert into dbo.[work_records] (wsn,dpart,iseq,wseq,id) " &
                        "values('" & now_wsn & "'," & dpart & "," & i & "," & k & ",'" & Session("s_id") & "')"
                    Else
                        If (dpart = 5) Then
                            If (i = 1 Or i = 2) Then
                                stastr = "未進行"
                            ElseIf (i = 3) Then
                                stastr = "未出貨"
                            Else
                                stastr = ""
                            End If
                        Else
                            stastr = "未進行"
                        End If
                        SqlCmd = "Insert into dbo.[work_records] (wsn,dpart,iseq,wseq,id,status) " &
                        "values('" & now_wsn & "'," & dpart & "," & i & "," & k & ",'" & Session("s_id") & "','" & stastr & "')"
                    End If
                    sqlresult = CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
                    conn.Close()
                    If (sqlresult = False) Then
                        GenWoFlag = False
                        CommUtil.ShowMsg(Me, "新增系統工單料件狀態表失敗")
                    Else
                        GenWoFlag = True
                    End If
                Next
            Next
        Next
        Return GenWoFlag
    End Function
    Sub SapWoGen(ByVal itemcode As String, ByVal whs As String, ByVal plannedqty As Integer, ByVal DueDate As String, ByVal model As String, ByVal cus_name As String, ByVal camera As String, ByVal lens As String, ByVal wsn As String)
        Dim vWo As SAPbobsCOM.ProductionOrders
        Dim docnum As Long
        Dim sqlresult As Boolean
        ret = InitSAPConnection(Session("usingserver"), Session("usingdb"))
        If (ret <> 0) Then
            CommUtil.ShowMsg(Me, "連線失敗")
            Exit Sub
        Else
            vWo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
            vWo.ItemNo = itemcode
            vWo.PlannedQuantity = plannedqty
            vWo.DueDate = DueDate
            vWo.Warehouse = whs
            vWo.UserFields.Fields.Item("U_F16").Value = "9999999"
            vWo.Remarks = model & " " & cus_name & "*" & plannedqty & " " & camera & "-" & lens
            If (0 <> vWo.Add()) Then
                CommUtil.ShowMsg(Me, "Failed to add WorkOrder item(可check看交貨日期是否小於今日日期): " & vWo.ItemNo) 'If failed, show a message
                vWo = Nothing
                CloseSAPConnection()
            Else
                vWo = Nothing
                SqlCmd = "select T0.docnum from dbo.OWOR T0 where T0.U_F16= '9999999' and T0.Status<>'C' and T0.Status<>'L'"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    docnum = dr(0)
                    dr.Close()
                    connsap.Close()

                    SqlCmd = "update dbo.OWOR set U_F16= '" & docnum & "' where docnum='" & docnum & "'"

                    sqlresult = CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                    connsap.Close()
                    If (sqlresult) Then
                        CommUtil.ShowMsg(Me, "SAP母工單產生成功")

                        SqlCmd = "update dbo.[worksn] set docnum=" & docnum & " " &
                        "where wsn='" & wsn & "'"

                        sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        If (sqlresult = False) Then
                            CommUtil.ShowMsg(Me, "更新worksn 的SAP號碼失敗")
                        End If
                        conn.Close()
                        CloseSAPConnection()
                        Response.Redirect("molist.aspx?smid=molist&smode=1")
                    Else
                        CommUtil.ShowMsg(Me, "更新失敗")
                        CloseSAPConnection()
                    End If
                Else
                    CommUtil.ShowMsg(Me, "查詢不到已產生之母工單 , 請至SAP 體check 看看")
                    CloseSAPConnection()
                End If
            End If
        End If
        'CloseSAPConnection()
    End Sub
End Class