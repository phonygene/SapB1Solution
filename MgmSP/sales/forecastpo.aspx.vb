Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class forecastpo
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public connsap As New SqlConnection
    Public SqlCmd As String
    Public ds As New DataSet
    Public dr As SqlDataReader
    Public permssa100 As String
    Public smode As Integer
    Public DDLModel, DDLCus, DDLArea, DDLSales, DDLPixel, DDLResolution, DDLStatus, DDLShipLoc As DropDownList
    Public TxtShipDate As TextBox
    Public ScriptManager1 As New ScriptManager
    Public machinetype As String
    Public BtnFilter, BtnNoFilter As Button
    Public RadioListSPP As RadioButtonList
    Public rule As String
    Public machineradioindex As Integer
    Public LtnAdd As LinkButton
    Public fmode As String
    'Public LBModel, LBCus, LBArea, LBSales, LBPixel, LBResolution, LBStatus As New ListBox
    Public gmodel, gcusname, gsales_area, gamount, gshipdate, gsales_person, gcamera, gresolution, gstatus, gmemo, gshiploc As String
    'Public TxtModel, TxtCspec, TxtCus, TxtArea, TxtSales, TxtPixel, TxtResolution, TxtStatus, TxtShipDateUpd As New TextBox
    Sub FTCreate()
        Dim ce As CalendarExtender
        Dim tRow As New TableRow
        Dim tCell As TableCell
        Dim Labelx As Label

        tCell = New TableCell
        RadioListSPP = New RadioButtonList
        RadioListSPP.ID = "radiolist_spp"
        RadioListSPP.Items.Add("近期預估")
        RadioListSPP.Items.Add("遠期預估")
        RadioListSPP.Items(0).Value = 1
        RadioListSPP.Items(1).Value = 2
        RadioListSPP.Width = 200
        RadioListSPP.RepeatDirection = RepeatDirection.Vertical
        AddHandler RadioListSPP.SelectedIndexChanged, AddressOf RadioListSPP_SelectedIndexChanged
        RadioListSPP.AutoPostBack = True
        RadioListSPP.SelectedIndex = 0
        tCell.Controls.Add(RadioListSPP)
        tRow.Cells.Add(tCell)

        tCell = New TableCell
        DDLModel = New DropDownList()
        DDLModel.ID = "ddl_model"
        DDLModel.Width = 150
        tCell.Controls.Add(DDLModel)

        DDLCus = New DropDownList()
        DDLCus.ID = "ddl_cus"
        DDLCus.Width = 100
        tCell.Controls.Add(DDLCus)

        DDLArea = New DropDownList()
        DDLArea.ID = "ddl_area"
        DDLArea.Width = 100
        tCell.Controls.Add(DDLArea)

        DDLShipLoc = New DropDownList()
        DDLShipLoc.ID = "ddl_shiploc"
        DDLShipLoc.Width = 80
        tCell.Controls.Add(DDLShipLoc)

        TxtShipDate = New TextBox
        TxtShipDate.Width = 100
        TxtShipDate.ID = "txt_shipdate"
        tCell.Controls.Add(TxtShipDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtShipDate.ID
        ce.ID = "ce_shipdate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)

        DDLSales = New DropDownList()
        DDLSales.ID = "ddl_sales"
        DDLSales.Width = 100
        tCell.Controls.Add(DDLSales)

        DDLPixel = New DropDownList()
        DDLPixel.ID = "ddl_pixel"
        DDLPixel.Width = 100
        tCell.Controls.Add(DDLPixel)

        DDLResolution = New DropDownList()
        DDLResolution.ID = "ddl_reso"
        DDLResolution.Width = 100
        tCell.Controls.Add(DDLResolution)

        DDLStatus = New DropDownList()
        DDLStatus.ID = "ddl_status"
        DDLStatus.Width = 100
        tCell.Controls.Add(DDLStatus)

        Labelx = New Label()
        Labelx.ID = "label_filter"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnFilter = New Button()
        BtnFilter.ID = "btn_filter"
        BtnFilter.Text = "篩選"
        AddHandler BtnFilter.Click, AddressOf BtnFilter_Click
        tCell.Controls.Add(BtnFilter)

        Labelx = New Label()
        Labelx.ID = "label_nofilter"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnNoFilter = New Button()
        BtnNoFilter.ID = "btn_nofilter"
        BtnNoFilter.Text = "不篩選"
        AddHandler BtnNoFilter.Click, AddressOf BtnNoFilter_Click
        tCell.Controls.Add(BtnNoFilter)
        tRow.Cells.Add(tCell)

        tCell = New TableCell
        LtnAdd = New LinkButton()
        LtnAdd.ID = "ltn_add"
        LtnAdd.Width = 150
        LtnAdd.Text = "新增預估"
        LtnAdd.PostBackUrl = "~/sales/forecastpo.aspx?smode=1&fmode=add"
        CommUtil.DisableObjectByPermission(LtnAdd, permssa100, "n")
        tCell.Controls.Add(LtnAdd)
        tRow.Cells.Add(tCell)

        FT.Rows.Add(tRow)
    End Sub

    Protected Sub BtnFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Response.Redirect("~/sales/forecastpo.aspx?smode=1&machineradioindex=" & MachineOption.SelectedIndex & "&sspradioindex=" & RadioListSPP.SelectedIndex & "&fmode=" & fmode)
        ListT.Dispose()
        ShowForeCastList()
    End Sub
    Protected Sub BtnNoFilter_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ListT.Dispose()
        ResetFilterField()
        ShowForeCastList()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        smode = Request.QueryString("smode")
        permssa100 = CommUtil.GetAssignRight("sa100", Session("s_id"))
        Page.Form.Controls.Add(ScriptManager1)
        fmode = Request.QueryString("fmode")
        FTCreate()
        CreateUpdItem()
        If (Not IsPostBack) Then
            machineradioindex = Request.QueryString("machineradioindex")
            MachineOption.SelectedIndex = machineradioindex
            RadioListSPP.SelectedIndex = Request.QueryString("sspradioindex")
            'fmode = Request.QueryString("fmode")
        Else
            'UpdT.FindControl("")
            'BtnDelete.OnClientClick = "return confirm('要刪除嗎')"
        End If

        If (fmode = "show") Then
            ListT.Visible = True
            UpdT.Visible = False
            FT.Visible = True
            MachineOption.Enabled = True
        Else
            ListT.Visible = False
            UpdT.Visible = True
            FT.Visible = False
            MachineOption.Enabled = False
        End If

        WriteFilterCombo()
        CreateHyperMenu()
        If (Not IsPostBack) Then
            ListT.Dispose()
            ShowForeCastList()
        End If
        If (fmode <> "show") Then
            ShowUpdItemData()
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
        tRow.Cells.Add(CellSet("單號", 1))
        tRow.Cells.Add(CellSet("", 0))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("機型", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_model", "txt_model", "dde_model"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("說明", 1))
        tRow.Cells.Add(CellSet("", 0))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("廠商", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_cus", "txt_cus", "dde_cus"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("銷售區域", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_area", "txt_area", "dde_area"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("何處出貨", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_shiploc", "txt_shiploc", "dde_shiploc"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("數量", 1))
        tRow.Cells.Add(CellSetWithTB("txt_amount"))
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
        tRow.Cells.Add(CellSet("預交日期", 1))
        tRow.Cells.Add(CellSetWithCalenderExtender("txt_shipdateupd", "ce_shipdateupd"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("銷售員", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_sales", "txt_sales", "dde_sales"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("像素", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_pixel", "txt_pixel", "dde_pixel"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("解析度", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_resolution", "txt_resolution", "dde_resolution"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("狀態", 1))
        tRow.Cells.Add(CellSetWithExtender("lb_status", "txt_status", "dde_status"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("備註", 1))
        tRow.Cells.Add(CellSetWithTB("txt_comment"))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("更新備註", 1))
        tRow.Cells.Add(CellSet("", 0))
        UpdT.Controls.Add(tRow)
        '------------------------------------------------
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tRow.Cells.Add(CellSet("更新次數", 1))
        tRow.Cells.Add(CellSet("", 0))
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
        If (fmode = "modify") Then
            Chkx.Text = "選取後執行刪除"
            CommUtil.DisableObjectByPermission(Chkx, permssa100, "d")
        ElseIf (fmode = "add") Then
            Chkx.Text = "新增確認"
        End If
        AddHandler Chkx.CheckedChanged, AddressOf Chkx_CheckedChanged
        Chkx.AutoPostBack = True
        Labelx = New Label
        Labelx.ID = "label2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Btnx = New Button
        Btnx.ID = "btn_action"
        If (fmode = "modify") Then
            Btnx.Text = "更新"
        ElseIf (fmode = "add") Then
            Btnx.Text = "新增"
            Btnx.Enabled = False
        End If
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
    Protected Sub BtnxCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Response.Redirect("~/sales/forecastpo.aspx?smode=1&machineradioindex=" & MachineOption.SelectedIndex & "&sspradioindex=" & RadioListSPP.SelectedIndex & "&fmode=show")
    End Sub
    Protected Sub BtnxAction_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Chkx As CheckBox
        Dim ucode As Long
        Dim actionok As Boolean
        If (RecordFieldCheck()) Then
            Chkx = UpdT.FindControl("chk_action")
            If (fmode = "add") Then
                If (Chkx.Checked) Then
                    SqlCmd = "SELECT IsNull(Max(cast(T0.Code as int)),0) from [dbo].[@UPSP] T0"
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    dr.Read()
                    ucode = dr(0) + 1
                    dr.Close()
                    connsap.Close()
                    actionok = InsertFCRecord(ucode)
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
                Response.Redirect("~/sales/forecastpo.aspx?smode=1&machineradioindex=" & MachineOption.SelectedIndex & "&sspradioindex=" & RadioListSPP.SelectedIndex & "&fmode=show")
            End If
        Else
            CommUtil.ShowMsg(Me, "有欄位空白")
        End If
    End Sub
    Function RecordFieldCheck()
        RecordFieldCheck = True
        If (CType(UpdT.FindControl("txt_model"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_model"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_cus"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_area"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_amount"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_shipdateupd"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_sales"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_pixel"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_resolution"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_status"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
        If (CType(UpdT.FindControl("txt_shiploc"), TextBox).Text = "") Then
            RecordFieldCheck = False
        End If
    End Function
    Function InsertFCRecord(ucode As Integer)
        Dim model, cspec, cusname, sales_area, createdate, shipdate, comment, status, shiploc As String
        Dim sales_person, camera_pixel, resolution As String
        Dim amount As Integer
        Dim updcount As Integer
        Dim ptype As Integer
        If (RadioListSPP.SelectedIndex = 0) Then
            ptype = 1
        Else
            ptype = 2
        End If
        updcount = 0
        model = CType(UpdT.FindControl("txt_model"), TextBox).Text
        cspec = UpdT.Rows(2).Cells(1).Text
        cusname = CType(UpdT.FindControl("txt_cus"), TextBox).Text
        sales_area = CType(UpdT.FindControl("txt_area"), TextBox).Text
        shiploc = CType(UpdT.FindControl("txt_shiploc"), TextBox).Text
        amount = CInt(CType(UpdT.FindControl("txt_amount"), TextBox).Text)
        createdate = UpdT.Rows(7).Cells(1).Text
        shipdate = CType(UpdT.FindControl("txt_shipdateupd"), TextBox).Text
        sales_person = CType(UpdT.FindControl("txt_sales"), TextBox).Text
        camera_pixel = CType(UpdT.FindControl("txt_pixel"), TextBox).Text
        resolution = CType(UpdT.FindControl("txt_resolution"), TextBox).Text
        status = CType(UpdT.FindControl("txt_status"), TextBox).Text
        comment = CType(UpdT.FindControl("txt_comment"), TextBox).Text
        SqlCmd = "insert into [dbo].[@UPSP] (code,name,u_updcount,u_amount,u_status,u_model,u_cspec,u_cusname,u_sales_area,u_shiploc,u_createdate,u_shipdate,u_sales_person,u_camera_pixel,u_resolution,u_comment,u_mtype,u_ptype) " &
        "values(" & ucode & "," & ucode & "," & updcount & "," & amount & ",'" & status & "','" & model & "','" & cspec & "','" & cusname & "','" & sales_area & "','" & shiploc & "', " &
                "'" & createdate & "','" & shipdate & "','" & sales_person & "','" & camera_pixel & "','" & resolution & "','" & comment & "','" & machinetype & "'," & ptype & ")"
        '讀回,check 是否已寫入
        CommUtil.SqlSapExecute("ins", SqlCmd, connsap)
        connsap.Close()

        SqlCmd = "SELECT T0.U_model " &
        "FROM dbo.[@UPSP] T0 " &
        "where T0.code='" & ucode & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            InsertFCRecord = True
        Else
            CommUtil.ShowMsg(Me, "資料沒寫入")
            InsertFCRecord = False
        End If
        dr.Close()
        connsap.Close()
    End Function
    Function UpdateFCRecord()
        Dim ucode As String
        Dim updcount, newupdc As Integer
        Dim lastupdmemo As String
        Dim model, cusname, sales_area, shipdate, comment, status, shiploc As String
        Dim sales_person, camera_pixel, resolution As String
        Dim cspec, createdate As String
        Dim amount As Integer
        Dim ptype As Integer
        model = CType(UpdT.FindControl("txt_model"), TextBox).Text
        cusname = CType(UpdT.FindControl("txt_cus"), TextBox).Text
        sales_area = CType(UpdT.FindControl("txt_area"), TextBox).Text
        shiploc = CType(UpdT.FindControl("txt_shiploc"), TextBox).Text
        amount = CInt(CType(UpdT.FindControl("txt_amount"), TextBox).Text)
        shipdate = CType(UpdT.FindControl("txt_shipdateupd"), TextBox).Text
        sales_person = CType(UpdT.FindControl("txt_sales"), TextBox).Text
        camera_pixel = CType(UpdT.FindControl("txt_pixel"), TextBox).Text
        resolution = CType(UpdT.FindControl("txt_resolution"), TextBox).Text
        status = CType(UpdT.FindControl("txt_status"), TextBox).Text
        comment = CType(UpdT.FindControl("txt_comment"), TextBox).Text
        cspec = UpdT.Rows(2).Cells(1).Text
        createdate = UpdT.Rows(7).Cells(1).Text

        If (RadioListSPP.SelectedIndex = 0) Then
            ptype = 1
        Else
            ptype = 2
        End If
        ucode = UpdT.Rows(0).Cells(1).Text
        updcount = CInt(UpdT.Rows(15).Cells(1).Text)
        'If(Sheets(sh).Cells(i, 14) gmodel , gcusname, gsales_area, gamount, gshipdate, gsales_person, gcamera, gresulution, gstatus, gmemo
        If (gmodel <> model Or gcusname <> cusname Or gsales_area <> sales_area Or gamount <> amount _
       Or gshipdate <> shipdate Or gsales_person <> sales_person Or gcamera <> camera_pixel Or gresolution <> resolution _
       Or gmemo <> comment Or gstatus <> status Or gshiploc <> shiploc) Then
            'lastupdmemo = Format(Now(), "yyyy/MM/dd") & "-更改 "
            lastupdmemo = ""
            If (gmodel <> model) Then
                lastupdmemo = lastupdmemo & gmodel & "-->" & model & ","
            End If
            If (gcusname <> cusname) Then
                lastupdmemo = lastupdmemo & gcusname & "-->" & cusname & ","
            End If
            If (gsales_area <> sales_area) Then
                lastupdmemo = lastupdmemo & gsales_area & "-->" & sales_area & ","
            End If
            If (gamount <> gamount) Then
                lastupdmemo = lastupdmemo & gamount & "-->" & gamount & ","
            End If
            If (gshipdate <> shipdate) Then
                lastupdmemo = lastupdmemo & gshipdate & "-->" & shipdate & ","
            End If
            If (gsales_person <> sales_person) Then
                lastupdmemo = lastupdmemo & gsales_person & "-->" & sales_person & ","
            End If
            If (gcamera <> camera_pixel) Then
                lastupdmemo = lastupdmemo & gcamera & "-->" & camera_pixel & ","
            End If
            If (gresolution <> resolution) Then
                'MsgBox gresolution & "---" & Sheets(sh).Cells(i, 11)
                lastupdmemo = lastupdmemo & gresolution & "-->" & resolution & ","
            End If
            If (gmemo <> comment) Then
                If (gmemo <> "") Then
                    If (Trim(comment) <> "") Then
                        lastupdmemo = lastupdmemo & "備註:" & gmemo & "-->" & Trim(comment) & ","
                    Else
                        lastupdmemo = lastupdmemo & "備註:" & gmemo & "--> 刪除備註,"
                    End If
                Else
                    If (Trim(comment) <> "") Then
                        lastupdmemo = lastupdmemo & "備註:" & "原空白改為-->" & Trim(comment) & ","
                    End If
                End If
            End If
            If (gstatus <> status) Then
                lastupdmemo = lastupdmemo & gstatus & "-->" & status & ","
            End If
            If (gshiploc <> shiploc) Then
                lastupdmemo = lastupdmemo & gshiploc & "-->" & shiploc & ","
            End If
        Else
            CommUtil.ShowMsg(Me, "沒任何更動")
            Exit Function
            'lastupdmemo = UpdT.Rows(13).Cells(1).Text
        End If
        lastupdmemo = Format(Now(), "yyyy/MM/dd") & "-更改 " & lastupdmemo
        SqlCmd = "SELECT T0.u_updcount " &
        "FROM dbo.[@UPSP] T0 " &
        "where T0.code='" & ucode & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        'MsgBox temprs.Fields(0).Value & "---" & updcount
        If (dr(0) <> updcount) Then
            CommUtil.ShowMsg(Me, "單號:" & ucode & "已先行被他人更新, 請從新檢視是否需再更新")
            UpdateFCRecord = False
            Exit Function
        End If
        dr.Close()
        connsap.Close()
        'createdate不更新
        newupdc = updcount + 1
        SqlCmd = "update [dbo].[@UPSP] set " &
        "u_amount=" & amount & ", u_model= '" & model & "' , u_cspec= '" & cspec & "', " &
        "u_cusname= '" & cusname & "' , u_sales_area= '" & sales_area & "', u_shiploc='" & shiploc & "', " &
        "u_shipdate='" & shipdate & "' , u_sales_person= '" & sales_person & "', u_camera_pixel = '" & camera_pixel & "', " &
        "u_resolution='" & resolution & "' , u_status= '" & status & "', u_comment = '" & comment & "', " &
        "u_mtype='" & machinetype & "' , u_lastupdmemo='" & lastupdmemo & "' , u_updcount= " & newupdc & ",u_ptype=" & ptype & " " &
        "where code = '" & ucode & "'"
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
        UpdateFCRecord = True
    End Function
    Function DeleteFCRecord()
        Dim ucode As String
        Dim connsap1 As New SqlConnection
        'MsgBox(DateDiff(DateInterval.Day, CDate(UpdT.Rows(6).Cells(1).Text), Now()))
        'Exit Sub
        DeleteFCRecord = True
        If (DateDiff(DateInterval.Day, CDate(UpdT.Rows(7).Cells(1).Text), Now()) > 6) Then
            CommUtil.ShowMsg(Me, "此單已建立超過6天 ,不能刪除 , 請將狀態改為已取消即可")
            DeleteFCRecord = False
            Exit Function
        Else
            ucode = UpdT.Rows(0).Cells(1).Text
        End If
        SqlCmd = "SELECT count(T0.u_model) " &
        "FROM dbo.[@UPSP] T0 " &
        "where T0.code='" & ucode & "'"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        If (dr(0) = 0) Then
            CommUtil.ShowMsg(Me, "單號已被他人刪除")
            DeleteFCRecord = False
        Else
            SqlCmd = "delete from [dbo].[@UPSP] " &
                    "where code = '" & ucode & "'"
            DeleteFCRecord = CommUtil.SqlSapExecute("del", SqlCmd, connsap1)
            connsap1.Close()
        End If
        dr.Close()
        connsap.Close()
    End Function
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
            Else
                Btnx.Enabled = False
            End If
        End If
    End Sub

    Sub ShowUpdItemData()
        If (Session("grp") = "JF" Or Session("grp") = "JT") Then
            CType(UpdT.FindControl("txt_area"), TextBox).Enabled = False
            CType(UpdT.FindControl("lb_area"), ListBox).Enabled = False
        End If
        Dim nowmodel As String
        Dim ucode As Long
        Dim amount, updcount As Integer
        Dim cspec, cusname, sales_area, createdate, shipdate, sales_person, camera_pixel, resolution, status, comment, lastupdmemo, shiploc As String
        If (fmode = "modify") Then
            ucode = Request.QueryString("num")
            SqlCmd = "select code, u_model , u_cspec, u_cusname, u_sales_area,u_amount, u_createdate, " &
                "u_shipdate, u_sales_person, u_camera_pixel, u_resolution, u_status, u_comment,u_updcount, " &
                "IsNull(u_lastupdmemo,''),IsNull(u_shiploc,'') from [dbo].[@UPSP] where code=" & ucode
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            'If (dr.HasRows) Then
            dr.Read()
            nowmodel = dr(1)
            cspec = dr(2)
            cusname = dr(3)
            sales_area = dr(4)
            amount = dr(5)
            createdate = dr(6)
            shipdate = dr(7)
            sales_person = dr(8)
            camera_pixel = dr(9)
            resolution = dr(10)
            status = dr(11)
            comment = dr(12)
            lastupdmemo = dr(14)
            updcount = dr(13)
            shiploc = dr(15)

            gmodel = dr(1)
            gcusname = dr(3)
            gsales_area = dr(4)
            gamount = dr(5)
            gshipdate = dr(7)
            gsales_person = dr(8)
            gcamera = dr(9)
            gresolution = dr(10)
            gstatus = dr(11)
            gmemo = dr(12)
            gshiploc = dr(15)
            'End If
            dr.Close()
            connsap.Close()
            UpdT.Rows(0).Cells(1).Text = ucode
            CType(UpdT.FindControl("txt_model"), TextBox).Text = nowmodel
            UpdT.Rows(2).Cells(1).Text = cspec
            CType(UpdT.FindControl("txt_cus"), TextBox).Text = cusname
            CType(UpdT.FindControl("txt_area"), TextBox).Text = sales_area
            CType(UpdT.FindControl("txt_shiploc"), TextBox).Text = shiploc
            CType(UpdT.FindControl("txt_amount"), TextBox).Text = amount
            UpdT.Rows(7).Cells(1).Text = createdate
            CType(UpdT.FindControl("txt_shipdateupd"), TextBox).Text = shipdate
            CType(UpdT.FindControl("txt_sales"), TextBox).Text = sales_person
            CType(UpdT.FindControl("txt_pixel"), TextBox).Text = camera_pixel
            CType(UpdT.FindControl("txt_resolution"), TextBox).Text = resolution
            CType(UpdT.FindControl("txt_status"), TextBox).Text = status
            'If (comment <> "0" And comment <> "") Then
            CType(UpdT.FindControl("txt_comment"), TextBox).Text = comment
            'End If
            UpdT.Rows(15).Cells(1).Text = updcount 'temp save
            UpdT.Rows(14).Cells(1).Text = lastupdmemo
        ElseIf (fmode = "add") Then
            UpdT.Rows(7).Cells(1).Text = Format(Now, "yyyy/MM/dd")
            CType(UpdT.FindControl("txt_shiploc"), TextBox).Text = ""
            If (Session("grp") = "JF") Then
                CType(UpdT.FindControl("lb_area"), ListBox).SelectedIndex = 2
                CType(UpdT.FindControl("txt_area"), TextBox).Text = CType(UpdT.FindControl("lb_area"), ListBox).SelectedValue
            ElseIf (Session("grp") = "JT") Then
                CType(UpdT.FindControl("lb_area"), ListBox).SelectedIndex = 3
                CType(UpdT.FindControl("txt_area"), TextBox).Text = CType(UpdT.FindControl("lb_area"), ListBox).SelectedValue
            End If
        End If
    End Sub
    Sub ShowForeCastList()
        If (Session("grp") = "JF") Then
            DDLArea.SelectedIndex = 2
            DDLArea.Enabled = False
        ElseIf (Session("grp") = "JT") Then
            DDLArea.SelectedIndex = 3
            DDLArea.Enabled = False
        End If
        rule = FilterSearch()
        CreateHead()
        CreateItemList(rule)
        PutDataToItemList(rule)
    End Sub
    Sub CreateItemList(rule As String)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim j, k As Integer
        Dim premodel As String
        Dim nowmodel As String
        Dim first As Boolean
        first = True
        'i = 2
        k = 0
        SqlCmd = "update [dbo].[@UPSP] set " &
                "u_status= '預估延遲' " &
                "where u_status= '預估中' and u_shipdate < GETDATE()"
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
        SqlCmd = "select u_model " &
                "from [dbo].[@UPSP] " & rule
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            premodel = ""
            Do While (dr.Read())
                nowmodel = dr(0)
                If (premodel <> nowmodel And first = False) Then
                    tRow = New TableRow()
                    tRow.BackColor = Drawing.Color.White
                    tRow.HorizontalAlign = HorizontalAlign.Center
                    For j = 0 To 14
                        tCell = New TableCell()
                        tCell.Font.Size = 10
                        tCell.BorderWidth = 1
                        tRow.Cells.Add(tCell)
                    Next
                    ListT.Rows.Add(tRow)
                    premodel = nowmodel
                    k = k + 1
                    tRow = New TableRow()
                    tRow.HorizontalAlign = HorizontalAlign.Center
                    If (k Mod 2) Then
                        tRow.BackColor = Drawing.Color.PapayaWhip
                    Else
                        tRow.BackColor = Drawing.Color.Lavender
                    End If
                    For j = 0 To 14
                        tCell = New TableCell()
                        tCell.BorderWidth = 1
                        tCell.Font.Size = 10
                        If (j = 2 Or j = 13) Then
                            tCell.HorizontalAlign = HorizontalAlign.Left
                        End If
                        tRow.Cells.Add(tCell)
                    Next
                    ListT.Rows.Add(tRow)
                Else
                    k = k + 1
                    tRow = New TableRow()
                    tRow.HorizontalAlign = HorizontalAlign.Center
                    If (k Mod 2) Then
                        tRow.BackColor = Drawing.Color.PapayaWhip
                    Else
                        tRow.BackColor = Drawing.Color.Lavender
                    End If
                    For j = 0 To 14
                        tCell = New TableCell()
                        tCell.BorderWidth = 1
                        tCell.Font.Size = 10
                        If (j = 2 Or j = 13) Then
                            tCell.HorizontalAlign = HorizontalAlign.Left
                        End If
                        tRow.Cells.Add(tCell)
                    Next
                    ListT.Rows.Add(tRow)
                    premodel = nowmodel
                    first = False
                End If
            Loop
        End If
        dr.Close()
        connsap.Close()
        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.White
        tRow.HorizontalAlign = HorizontalAlign.Center
        For j = 0 To 14
            tCell = New TableCell()
            tCell.Font.Size = 10
            tCell.BorderWidth = 1
            tRow.Cells.Add(tCell)
        Next
        ListT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.White
        tRow.HorizontalAlign = HorizontalAlign.Center
        For j = 0 To 14
            tCell = New TableCell()
            tCell.Font.Size = 10
            tCell.BorderWidth = 1
            tRow.Cells.Add(tCell)
        Next
        ListT.Rows.Add(tRow)
    End Sub

    Sub PutDataToItemList(rule As String)
        Dim i, j As Integer
        Dim count As Integer
        Dim modelc As Integer
        Dim premodel As String
        Dim nowmodel As String
        Dim amount, updcount, ucode1 As Integer
        Dim cspec, cusname, sales_area, createdate, shipdate, sales_person, camera_pixel, resolution, status, comment, lastupdmemo, shiploc As String
        Dim first As Boolean
        Dim Ltn As LinkButton
        first = True
        count = 0
        i = 2
        SqlCmd = "select code, u_model , u_cspec, u_cusname, u_sales_area,u_amount,IsNull(u_shiploc,'台北發貨'), u_createdate, " &
                "u_shipdate, u_sales_person, u_camera_pixel, u_resolution, u_status, u_comment,u_updcount, " &
                "IsNull(u_lastupdmemo,'') from [dbo].[@UPSP] " & rule
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            premodel = ""
            Do While (dr.Read())
                ucode1 = dr(0)
                nowmodel = dr(1)
                cspec = dr(2)
                cusname = dr(3)
                sales_area = dr(4)
                amount = dr(5)
                shiploc = dr(6)
                createdate = dr(7)
                shipdate = dr(8)
                sales_person = dr(9)
                camera_pixel = dr(10)
                resolution = dr(11)
                status = dr(12)
                comment = dr(13)
                updcount = dr(14)
                lastupdmemo = dr(15)

                'ListT.Rows(i).Cells(j).Text
                If (premodel <> nowmodel And first = False) Then
                    For j = 0 To 14
                        If (j = 4) Then
                            ListT.Rows(i).Cells(j).Text = premodel & "台數"
                            ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightGreen
                        End If
                        If (j = 5) Then
                            ListT.Rows(i).Cells(j).Text = modelc
                            ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightGreen
                        End If
                    Next
                    modelc = dr(5)
                    premodel = nowmodel
                    i = i + 1
                    For j = 0 To 14
                        If (j = 0) Then
                            Ltn = New LinkButton
                            Ltn.ID = "ltn_modify_" & dr(j)
                            'Ltn.Width = 150
                            Ltn.Text = dr(0) '"修改"
                            Ltn.PostBackUrl = "~/sales/forecastpo.aspx?smode=1&fmode=modify&num=" & dr(0)
                            CommUtil.DisableObjectByPermission(Ltn, permssa100, "m")
                            ListT.Rows(i).Cells(j).Controls.Add(Ltn)
                        Else
                            ListT.Rows(i).Cells(j).Text = dr(j)
                            If (j = 7 And DateDiff("d", CDate(dr(7)), Now()) <= 6) Then
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.Yellow
                            ElseIf (j = 7 And DateDiff("d", CDate(dr(7)), Now()) > 6 And DateDiff("d", CDate(dr(7)), Now()) <= 10) Then
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightBlue
                            End If
                            If (j = 8 And DateDiff("d", Now(), CDate(dr(8))) < 30) Then
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.Pink
                            End If
                            If (j = 14) Then
                                ListT.Rows(i).Cells(j).ToolTip = lastupdmemo
                            End If
                            If (j = 6 And ListT.Rows(i).Cells(j).Text <> "台北發貨") Then
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightSalmon
                            End If
                        End If
                    Next
                Else
                    For j = 0 To 14
                        If (j = 0) Then
                            Ltn = New LinkButton
                            Ltn.ID = "ltn_modify_" & dr(j)
                            'Ltn.Width = 150
                            Ltn.Text = dr(0) '"修改"
                            Ltn.PostBackUrl = "~/sales/forecastpo.aspx?smode=1&fmode=modify&num=" & dr(0)
                            CommUtil.DisableObjectByPermission(Ltn, permssa100, "m")
                            ListT.Rows(i).Cells(j).Controls.Add(Ltn)
                        Else
                            ListT.Rows(i).Cells(j).Text = dr(j)
                            If (j = 7 And DateDiff("d", CDate(dr(7)), Now()) <= 6) Then '
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.Yellow
                            ElseIf (j = 7 And DateDiff("d", CDate(dr(7)), Now()) > 6 And DateDiff("d", CDate(dr(7)), Now()) <= 10) Then
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightBlue
                            End If
                            If (j = 8 And DateDiff("d", Now(), CDate(dr(8))) < 30) Then
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.Pink
                            End If
                            If (j = 14) Then
                                ListT.Rows(i).Cells(j).ToolTip = lastupdmemo
                            End If
                            If (j = 6 And ListT.Rows(i).Cells(j).Text <> "台北發貨") Then
                                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightSalmon
                            End If
                        End If
                    Next
                    modelc = modelc + amount
                    premodel = nowmodel
                    first = False
                End If
                count = count + amount
                If (ListT.Rows(i).Cells(12).Text = "預估延遲") Then
                    ListT.Rows(i).Cells(12).BackColor = Drawing.Color.Red
                End If
                i = i + 1
            Loop
        End If
        dr.Close()
        connsap.Close()
        For j = 0 To 14
            If (j = 4) Then
                ListT.Rows(i).Cells(j).Text = premodel & "台數"
                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightGreen
            End If
            If (j = 5) Then
                ListT.Rows(i).Cells(j).Text = modelc
                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightGreen
            End If
        Next
        i = i + 1
        For j = 0 To 14
            If (j = 4) Then
                ListT.Rows(i).Cells(j).Text = "All台數"
                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightGreen
            End If
            If (j = 5) Then
                ListT.Rows(i).Cells(j).Text = count
                ListT.Rows(i).Cells(j).BackColor = Drawing.Color.LightGreen
            End If
        Next
    End Sub
    Sub CreateHyperMenu()
        Dim Hyper As HyperLink
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        tRow = New TableRow()
        For i = 0 To 1
            tCell = New TableCell()
            tRow.Cells.Add(tCell)
        Next
        Me.HyperMenuT.Rows.Add(tRow)
        j = 0
        Hyper = New HyperLink()
        Hyper.ID = "index"
        Hyper.Text = "首頁"
        Hyper.NavigateUrl = "../index.aspx?smid=index"
        Hyper.BackColor = Drawing.Color.Aqua
        Hyper.Font.Underline = False
        Hyper.Width = 130
        Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
        Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
        Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
        j = j + 1
        If (smode <> j) Then
            Hyper = New HyperLink()
            Hyper.ID = "hyper_forcast"
            Hyper.Text = "預估訂單"
            Hyper.NavigateUrl = "~/sales/forecastpo.aspx?smode=1"
            Hyper.BackColor = Drawing.Color.Aqua
            Hyper.Font.Underline = False
            Hyper.Width = 130
            CommUtil.DisableObjectByPermission(Hyper, permssa100, "e")
            Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
            Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
            Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
        Else
            Me.HyperMenuT.Rows(0).Cells(j).Text = "預估訂單"
            Me.HyperMenuT.Rows(0).Cells(j).Width = 130
            Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
            Me.HyperMenuT.Rows(0).Cells(j).BackColor = Drawing.Color.Gainsboro
        End If
    End Sub

    Public Sub CreateHead()
        Dim tCell As TableCell
        Dim tRow As TableRow
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Wrap = False
        If (MachineOption.SelectedIndex = 0) Then
            If (RadioListSPP.SelectedValue = 1) Then
                tCell.Text = "AOI近期預估訂單List(篩選條件:如上面下拉式條件)"
            Else
                tCell.Text = "AOI遠期預估訂單List(篩選條件:如上面下拉式條件)"
            End If
        Else
            If (RadioListSPP.SelectedValue = 2) Then
                tCell.Text = "ICT近期預估訂單List(篩選條件:如上面下拉式條件)"
            Else
                tCell.Text = "ICT遠期預估訂單List(篩選條件:如上面下拉式條件)"
            End If
        End If
        tCell.ColumnSpan = 14
        tCell.Font.Bold = True
        tRow.Cells.Add(tCell)
        ListT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tRow.Font.Bold = True
        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "單號"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "機型"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "說明"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "客戶"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "銷售區域"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "數量"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "何處出貨"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "建立日期"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "預交日期"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "銷售員"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "相機"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "解析度"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "狀態"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "備註"
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.BackColor = Drawing.Color.DeepSkyBlue
        tCell.Font.Size = 10
        tCell.Text = "更新次"
        tRow.Cells.Add(tCell)

        ListT.Rows.Add(tRow)
    End Sub

    Sub ResetFilterField()
        DDLModel.SelectedIndex = 0
        DDLArea.SelectedIndex = 0
        DDLCus.SelectedIndex = 0
        TxtShipDate.Text = ""
        DDLSales.SelectedIndex = 0
        DDLPixel.SelectedIndex = 0
        DDLResolution.SelectedIndex = 0
        DDLStatus.SelectedIndex = 0
    End Sub
    Sub WriteFilterCombo()
        Dim i As Integer
        Dim k As Integer
        Dim model, mdesc, mtype As String
        Dim LBx As ListBox
        If (MachineOption.SelectedIndex = 0) Then
            machinetype = "AOI"
        ElseIf (MachineOption.SelectedIndex = 1) Then
            machinetype = "ICT"
        End If

        If (MachineOption.SelectedIndex = 0) Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                    "FROM dbo.[@UMMD] T0 where T0.u_mtype='SPI' or T0.u_mtype='AOI' or T0.u_mtype='3DAOI' " &
                    "order by T0.u_model,T0.u_mcode"
        ElseIf (MachineOption.SelectedIndex = 1) Then
            SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                    "FROM dbo.[@UMMD] T0 where T0.u_mtype='ICT' order by T0.u_model,T0.u_mcode"
        End If
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        DDLModel.Items.Clear()
        DDLModel.Items.Add("選機型")
        LBx = UpdT.FindControl("lb_model")
        LBx.Items.Clear()
        LBx.Items.Add("選機型")
        k = 1
        If (dr.HasRows) Then
            Do While (dr.Read())
                model = dr(0)
                mdesc = dr(1)
                mtype = dr(2)
                DDLModel.Items.Add(mtype & "-" & model & "-" & mdesc)
                LBx.Items.Add(mtype & "-" & model & "-" & mdesc)
                k = k + 1
            Loop
        End If
        DDLModel.SelectedIndex = 0
        LBx.SelectedIndex = 0
        dr.Close()
        connsap.Close()

        DDLArea.Items.Clear()
        DDLArea.Items.Add("選區域")
        DDLArea.Items.Add("台北捷智")
        DDLArea.Items.Add("華東捷豐")
        DDLArea.Items.Add("華南捷智通")
        DDLArea.SelectedIndex = 0
        LBx = UpdT.FindControl("lb_area")
        LBx.Items.Clear()
        LBx.Items.Add("選區域")
        LBx.Items.Add("台北捷智")
        LBx.Items.Add("華東捷豐")
        LBx.Items.Add("華南捷智通")
        LBx.SelectedIndex = 0

        DDLShipLoc.Items.Clear()
        DDLShipLoc.Items.Add("選發貨區")
        DDLShipLoc.Items.Add("台北發貨")
        DDLShipLoc.Items.Add("華南發貨")
        DDLShipLoc.Items.Add("Demo機出貨")
        DDLShipLoc.SelectedIndex = 0
        LBx = UpdT.FindControl("lb_shiploc")
        LBx.Items.Clear()
        LBx.Items.Add("選發貨區")
        LBx.Items.Add("台北發貨")
        LBx.Items.Add("華南發貨")
        LBx.Items.Add("Demo機出貨")
        LBx.SelectedIndex = 0

        DDLStatus.Items.Clear()
        DDLStatus.Items.Add("未結案")
        DDLStatus.Items.Add("預估中")
        DDLStatus.Items.Add("預估延遲")
        DDLStatus.Items.Add("已發單(無訂單)")
        DDLStatus.Items.Add("已發單(有訂單)")
        DDLStatus.Items.Add("已出貨")
        DDLStatus.Items.Add("已取消")
        DDLStatus.Items.Add("待驗收")
        DDLStatus.SelectedIndex = 0
        LBx = UpdT.FindControl("lb_status")
        LBx.Items.Clear()
        LBx.Items.Add("未結案")
        LBx.Items.Add("預估中")
        LBx.Items.Add("預估延遲")
        LBx.Items.Add("已發單(無訂單)")
        LBx.Items.Add("已發單(有訂單)")
        LBx.Items.Add("已出貨")
        LBx.Items.Add("已取消")
        LBx.Items.Add("待驗收")
        LBx.SelectedIndex = 0

        DDLPixel.Items.Clear()
        DDLPixel.Items.Add("請選像素")
        LBx = UpdT.FindControl("lb_pixel")
        LBx.Items.Clear()
        LBx.Items.Add("請選像素")
        If (MachineOption.SelectedIndex = 0) Then
            DDLPixel.Items.Add("不確定")
            LBx.Items.Add("不確定")
        Else
            DDLPixel.Items.Add("不需要")
            LBx.Items.Add("不需要")
        End If
        DDLPixel.Items.Add("6.5M")
        DDLPixel.Items.Add("12M")
        DDLPixel.Items.Add("25M")
        LBx.Items.Add("6.5M")
        LBx.Items.Add("12M")
        LBx.Items.Add("25M")
        DDLPixel.SelectedIndex = 0
        LBx.SelectedIndex = 0
        'If (MachineOption.SelectedIndex = 0) Then
        '    DDLPixel.SelectedIndex = 0
        '    LBx.SelectedIndex = 0
        'Else
        '    DDLPixel.SelectedIndex = 1
        '    LBx.SelectedIndex = 1
        'End If

        DDLResolution.Items.Clear()
        DDLResolution.Items.Add("選擇解析度")
        LBx = UpdT.FindControl("lb_resolution")
        LBx.Items.Clear()
        LBx.Items.Add("選擇解析度")
        If (MachineOption.SelectedIndex = 0) Then
            DDLResolution.Items.Add("不確定")
            LBx.Items.Add("不確定")
        Else
            DDLResolution.Items.Add("不需要")
            LBx.Items.Add("不需要")
        End If
        DDLResolution.Items.Add("2.5u")
        DDLResolution.Items.Add("3u")
        DDLResolution.Items.Add("5u")
        DDLResolution.Items.Add("5.5u")
        DDLResolution.Items.Add("6u")
        DDLResolution.Items.Add("6.7u")
        DDLResolution.Items.Add("7u")
        DDLResolution.Items.Add("8u")
        DDLResolution.Items.Add("10u")
        DDLResolution.Items.Add("12u")
        DDLResolution.Items.Add("15u")
        DDLResolution.Items.Add("20u")
        DDLResolution.Items.Add("60u")
        DDLResolution.Items.Add("100u")
        LBx.Items.Add("2.5u")
        LBx.Items.Add("3u")
        LBx.Items.Add("5u")
        LBx.Items.Add("5.5u")
        LBx.Items.Add("6u")
        LBx.Items.Add("6.7u")
        LBx.Items.Add("7u")
        LBx.Items.Add("8u")
        LBx.Items.Add("10u")
        LBx.Items.Add("12u")
        LBx.Items.Add("15u")
        LBx.Items.Add("20u")
        LBx.Items.Add("60u")
        LBx.Items.Add("100u")
        DDLResolution.SelectedIndex = 0
        LBx.SelectedIndex = 0
        'If (MachineOption.SelectedIndex = 0) Then
        '    DDLResolution.SelectedIndex = 0
        '    LBx.SelectedIndex = 0
        'Else
        '    DDLResolution.SelectedIndex = 1
        '    LBx.SelectedIndex = 1
        'End If

        'TxtShipDate.Text = Format(Now, "yyyy/MM/dd")

        SqlCmd = "SELECT distinct T0.U_cusname " &
            "FROM dbo.[@UPSP] T0"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            DDLCus.Items.Clear()
            DDLCus.Items.Add("客戶名")
            LBx = UpdT.FindControl("lb_cus")
            LBx.Items.Clear()
            LBx.Items.Add("客戶名")
            i = 1
            Do While (dr.Read())
                DDLCus.Items.Add(dr(0))
                LBx.Items.Add(dr(0))
                i = i + 1
            Loop
        End If
        dr.Close()
        connsap.Close()
        DDLCus.SelectedIndex = 0
        LBx.SelectedIndex = 0

        SqlCmd = "SELECT distinct T0.U_sales_person " &
            "FROM dbo.[@UPSP] T0"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            DDLSales.Items.Clear()
            DDLSales.Items.Add("銷售員")
            LBx = UpdT.FindControl("lb_sales")
            LBx.Items.Clear()
            LBx.Items.Add("銷售員")
            i = 1
            Do While (dr.Read())
                DDLSales.Items.Add(dr(0))
                LBx.Items.Add(dr(0))
                i = i + 1
            Loop
        End If
        dr.Close()
        connsap.Close()
        DDLSales.SelectedIndex = 0
        LBx.SelectedIndex = 0
    End Sub

    Function FilterSearch()
        Dim rule As String
        Dim str() As String
        Dim model, cusname, sales_area, shipdate, sales_person, camera_pixel, resolution, status, shiploc As String
        Dim filterflag As Boolean

        filterflag = False
        rule = " where u_mtype='" & machinetype & "' and "
        If (DDLModel.SelectedIndex <> 0) Then
            str = Split(DDLModel.SelectedValue, "-")
            model = str(1)
            rule = rule & "u_model='" & model & "' "
            filterflag = True
        End If
        If (DDLCus.SelectedIndex <> 0) Then
            cusname = DDLCus.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_cusname='" & cusname & "' "
            Else
                rule = rule & " u_cusname='" & cusname & "' "
            End If
            filterflag = True
        End If
        If (DDLArea.SelectedIndex <> 0) Then
            sales_area = DDLArea.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_sales_area='" & sales_area & "' "
            Else
                rule = rule & " u_sales_area='" & sales_area & "' "
            End If
            filterflag = True
        End If
        If (DDLShipLoc.SelectedIndex <> 0) Then
            shiploc = DDLShipLoc.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_shiploc='" & shiploc & "' "
            Else
                rule = rule & " u_shiploc='" & shiploc & "' "
            End If
            filterflag = True
        End If
        If (TxtShipDate.Text <> "") Then
            shipdate = TxtShipDate.Text
            If (filterflag = True) Then
                rule = rule & " and u_shipdate='" & shipdate & "' "
            Else
                rule = rule & " u_shipdate <='" & shipdate & "' "
            End If
            filterflag = True
        End If
        If (DDLSales.SelectedIndex <> 0) Then
            sales_person = DDLSales.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_sales_person='" & sales_person & "' "
            Else
                rule = rule & " u_sales_person='" & sales_person & "' "
            End If
            filterflag = True
        End If
        If (DDLPixel.SelectedIndex <> 0) Then
            camera_pixel = DDLPixel.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_camera_pixel='" & camera_pixel & "' "
            Else
                rule = rule & " u_camera_pixel='" & camera_pixel & "' "
            End If
            filterflag = True
        End If
        If (DDLResolution.SelectedIndex <> 0) Then
            resolution = DDLResolution.SelectedValue
            If (filterflag = True) Then
                rule = rule & " and u_resolution='" & resolution & "' "
            Else
                rule = rule & " u_resolution='" & resolution & "' "
            End If
            filterflag = True
        End If
        If (DDLStatus.SelectedValue <> "") Then
            status = DDLStatus.SelectedValue
            If (status <> "未結案") Then
                If (filterflag = True) Then
                    rule = rule & " and u_status='" & status & "' "
                Else
                    rule = rule & " u_status='" & status & "' "
                End If
            Else
                If (filterflag = True) Then
                    rule = rule & " and (u_status='預估中' or u_status='預估延遲' or u_status='已發單(無訂單)' or u_status='已發單(有訂單)' or u_status='待驗收') "
                Else
                    rule = rule & " (u_status='預估中' or u_status='預估延遲' or u_status='已發單(無訂單)' or u_status='已發單(有訂單)' or u_status='待驗收') "
                End If
            End If
            filterflag = True
        End If
        If (filterflag = False) Then
            rule = rule & " (u_status='預估中' or u_status='預估延遲' or u_status='已發單(無訂單)' or u_status='已發單(有訂單)' or u_status='待驗收') "
        End If
        If (RadioListSPP.SelectedValue = 1) Then
            rule = rule & " and u_ptype=1 "
        ElseIf (RadioListSPP.SelectedValue = 2) Then
            rule = rule & " and u_ptype=2 "

        End If
        rule = rule & "order by u_model,u_status desc,u_shipdate,u_sales_area,u_sales_person"
        Return rule
        'ShowPrePOList rule
    End Function
    '
    Protected Sub RadioListSPP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        '用response 或底下mark 都可以 , 差別在於用response會init filter 條件(如要保留,可用下述或在網址上把這些參數加上=>較麻煩)
        Response.Redirect("~/sales/forecastpo.aspx?smode=1&machineradioindex=" & MachineOption.SelectedIndex & "&sspradioindex=" & RadioListSPP.SelectedIndex & "&fmode=" & fmode)
        'ListT.Dispose()
        'ShowForeCastList()
    End Sub

    Protected Sub MachineOption_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MachineOption.SelectedIndexChanged
        '用response 或底下mark 都可以 , 差別在於用response會init filter 條件
        Response.Redirect("~/sales/forecastpo.aspx?smode=1&machineradioindex=" & MachineOption.SelectedIndex & "&sspradioindex=" & RadioListSPP.SelectedIndex & "&fmode=" & fmode)
        'ListT.Dispose()
        'ShowForeCastList()
    End Sub

    Protected Sub LB_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Txtx As TextBox
        Dim str(), str1() As String
        str = Split(sender.ID, "_")
        If (str(1) <> "model") Then
            Txtx = UpdT.FindControl("txt_" & str(1))
            Txtx.Text = sender.SelectedValue
        Else
            Txtx = UpdT.FindControl("txt_" & str(1))
            str1 = Split(sender.SelectedValue, "-")
            Txtx.Text = str1(1)
            UpdT.Rows(2).Cells(1).Text = str1(2)
        End If
    End Sub
End Class