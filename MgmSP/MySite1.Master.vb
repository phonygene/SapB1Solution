Imports System.Data
Imports System.Data.SqlClient

Partial Public Class MySite1
    Inherits System.Web.UI.MasterPage
    Dim CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public connsap As New SqlConnection
    Public SqlCmd As String
    Public dr As SqlDataReader
    Public DDLWhs As DropDownList
    Public DDLDBS As DropDownList

    Sub HyperMainMenuGen(row As Integer, objid As String, text As String, url As String, width As Integer, perms As String, keyp As String)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Hyper As HyperLink
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        If (row Mod 2) Then
            tCell.BackColor = Drawing.Color.Aqua
        Else
            tCell.BackColor = Drawing.Color.LightPink
        End If
        Hyper = New HyperLink()
        Hyper.ID = objid
        Hyper.Text = text
        If (width <> 0) Then
            Hyper.Width = width
        End If
        Hyper.NavigateUrl = url
        Hyper.Font.Underline = False
        tCell.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='Gainsboro'")
        tCell.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        tCell.Controls.Add(Hyper)
        If (perms <> "nouse") Then
            CommUtil.DisableObjectByPermission(Hyper, perms, keyp)
        End If
        tRow.Cells.Add(tCell)
        tRow.HorizontalAlign = HorizontalAlign.Center
        TMainMenu.Rows.Add(tRow)
        'TMainMenu.Rows(i).Cells(0).HorizontalAlign = HorizontalAlign.Center
        'TMainMenu.Rows(i).Cells(0).Controls.Add(Hyper)
    End Sub

    'Sub HyperMainMenuGenOtherMethod(row As Integer, objid As String, text As String, url As String, width As Integer, perms As String, keyp As String)
    '    Dim Hyper As HyperLink
    '    Hyper = New HyperLink()
    '    Hyper.ID = objid
    '    Hyper.Text = text
    '    If (row Mod 2) Then
    '        Hyper.BackColor = Drawing.Color.Aqua
    '    Else
    '        Hyper.BackColor = Drawing.Color.LightPink
    '    End If
    '    If (width <> 0) Then
    '        Hyper.Width = width
    '    End If
    '    Hyper.NavigateUrl = url
    '    Hyper.Font.Underline = False
    '    Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='Gainsboro'")
    '    Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
    '    If (perms <> "nouse") Then
    '        CommUtil.DisableObjectByPermission(Hyper, perms, keyp)
    '    End If
    '    Me.TMainMenu.Rows(row).Cells(0).HorizontalAlign = HorizontalAlign.Center
    '    Me.TMainMenu.Rows(row).Cells(0).Controls.Add(Hyper)
    'End Sub

    Sub HyperSubMenuGen(tRow As TableRow, col As Integer, objid As String, text As String, url As String, width As Integer, perms As String, keyp As String)
        Dim tCell As TableCell
        Dim Hyper As HyperLink
        tCell = New TableCell()
        tCell.BackColor = Drawing.Color.LightBlue
        If (Request.QueryString("smode") <> col) Then
            Hyper = New HyperLink()
            Hyper.ID = objid
            Hyper.Text = text
            If (width <> 0) Then
                Hyper.Width = width
            End If
            Hyper.NavigateUrl = url
            Hyper.Font.Underline = False
            tCell.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='Gainsboro'")
            tCell.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            tCell.Controls.Add(Hyper)
            If (perms <> "nouse") Then
                CommUtil.DisableObjectByPermission(Hyper, perms, keyp)
            End If
        Else
            tCell.Text = text
            If (width <> 0) Then
                tCell.Width = width
            End If
            tCell.BackColor = Drawing.Color.Gainsboro
        End If
        tRow.Cells.Add(tCell)
    End Sub

    Sub DataBaseDropList(tRow As TableRow)
        Dim tCell As TableCell
        tCell = New TableCell()
        DDLDBS = New DropDownList()
        'CommUtil.InitSAPSQLConnection(connsap)
        SqlCmd = "SELECT name, database_id, create_date FROM sys.databases"
        'myCommand = New SqlCommand(SqlCmd, connsap)
        'dr = myCommand.ExecuteReader()
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        Do While (dr.Read())
            If (dr(0) <> "master" And dr(0) <> "tempdb" And dr(0) <> "model" And dr(0) <> "msdb" And dr(0) <> "SBO-COMMON") Then
                DDLDBS.Items.Add(dr(0))
            End If
        Loop
        dr.Close()
        DDLDBS.SelectedValue = Session("usingdb")
        connsap.Close()
        DDLDBS.ID = "ddl_dbs"
        DDLDBS.Width = 100
        AddHandler DDLDBS.SelectedIndexChanged, AddressOf DDLDBS_SelectedIndexChanged
        DDLDBS.AutoPostBack = True
        DDLDBS.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='Gainsboro'")
        DDLDBS.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        'If (Session("s_id") = "ron" Or Session("s_id") = "su") Then
        'DDLDBS.Enabled = True
        'Else
        DDLDBS.Enabled = False
        'End If
        tCell.Controls.Add(DDLDBS)
        tRow.Cells.Add(tCell)
    End Sub

    Sub WhsDropList(tRow As TableRow)
        Dim tCell As TableCell
        tCell = New TableCell()
        DDLWhs = New DropDownList()
        DDLWhs.Items.Clear()
        DDLWhs.Items.Add("C01 ICT")
        DDLWhs.Items.Add("C02 AOI")
        DDLWhs.SelectedValue = Session("usingwhsfull")

        'SqlCmd = "SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0 order by T0.WhsCode"
        'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        'count = 0
        'DDLWhs.Items.Add("請選擇倉別")
        'Do While (dr.Read())
        '    DDLWhs.Items.Add(dr(0) & " " & dr(1))
        '    count = count + 1
        'Loop
        'dr.Close()
        'DDLWhs.SelectedValue = Session("usingwhsfull")
        'connsap.Close()
        DDLWhs.ID = "ddl_whs"
        DDLWhs.Width = 100
        AddHandler DDLWhs.SelectedIndexChanged, AddressOf DDLWhs_SelectedIndexChanged
        DDLWhs.AutoPostBack = True
        DDLWhs.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='Gainsboro'")
        DDLWhs.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        tCell.Controls.Add(DDLWhs)
        tRow.Cells.Add(tCell)
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim tRow As TableRow
        Dim i As Integer
        Dim perm As String
        Dim mmenusize, smenusize As Integer

        mmenusize = 130
        smenusize = 100
        'Dim tCell As TableCell
        If (Session("s_id") <> "" And Not IsPostBack()) Then
            'DDLServer.Visible = True
            'DDLServer.Items.Clear()
            'InitSAPSQLConnection(Session("usingserver"), "")
            'SqlCmd = "SELECT name, database_id, create_date FROM sys.databases"
            'myCommand = New SqlCommand(SqlCmd, connsap)
            'dr = myCommand.ExecuteReader()
            'DDLServer.Items.Add("請選擇SAP資料庫")
            'Do While (dr.Read())
            '    If (dr(0) <> "master" And dr(0) <> "tempdb" And dr(0) <> "model" And dr(0) <> "msdb" And dr(0) <> "SBO-COMMON") Then
            '        DDLServer.Items.Add(dr(0))
            '    End If
            'Loop
            'dr.Close()
            'CloseSAPSQLConnection()
            'DDLServer.SelectedValue = Session("usingdb")

            'DDLWhs.Visible = True
            'InitSAPSQLConnection()
            'SqlCmd = "SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0 order by T0.WhsCode"
            'myCommand = New SqlCommand(SqlCmd, connsap)
            'dr = myCommand.ExecuteReader()
            'DDLWhs.Items.Clear()
            'DDLWhs.Items.Add("請選擇倉別")
            'Do While (dr.Read())
            '    DDLWhs.Items.Add(dr(0) & " " & dr(1))
            'Loop
            'dr.Close()
            'DDLWhs.SelectedValue = Session("usingwhsfull")
            'CloseSAPSQLConnection()
        End If
        '        If (IsPostBack) Then
        'For i = 0 To 15
        '    tRow = New TableRow()
        '    tRow.BorderWidth = 1
        '    For j = 0 To 0
        '        tCell = New TableCell()
        '        tRow.Cells.Add(tCell)
        '    Next
        '    Me.TMainMenu.Rows.Add(tRow)
        'Next
        Dim indexpage As Integer
        If (Session("s_id") <> "") Then
            indexpage = Request.QueryString("indexpage")
            tRow = New TableRow()
            tRow.BorderWidth = 1
            tRow.HorizontalAlign = HorizontalAlign.Center
            i = 0
            If (Session("s_id") = "ron") Then
                HyperMainMenuGen(i, "", "上線人數: " & Application("user_sessions"), "", mmenusize, "nouse", "")
                i = i + 1
            End If
            HyperMainMenuGen(i, "logout", "登出" & "(" & Session("s_id") & ")", "~/usermgm/logout.aspx", mmenusize, "nouse", "")
            '-------------------------------------------
            'smid是用來控制顯示目前所選主menu上之sub menu (不然以下所寫之sub menu 會互相干涉)
            'smode是用來控制目前所在之sub menu , 讓其無法呈現超聯結 , 而顯示反白(表示目前處於此sub menu中)

            If (Session("actmode") <> "signoff" And Session("actmode") <> "single_signoff" And Session("actmode") <> "todoitem" And Session("actmode") <> "informtraceperson") Then
                i = i + 1
                HyperMainMenuGen(i, "index", "首頁", "index.aspx?smid=index&smode=0", mmenusize, "nouse", "")
                If (Request.QueryString("smid") = "index") Then
                    HyperSubMenuGen(tRow, 1, "pwdchange", "修改密碼", "~/usermgm/pwdchange.aspx?smid=index&smode=1", smenusize, "nouse", "")
                    HyperSubMenuGen(tRow, 2, "addsapid", "設定SAP帳密", "~/usermgm/addsapid.aspx?smid=index&smode=2", smenusize, "nouse", "")
                    DataBaseDropList(tRow)
                    WhsDropList(tRow)
                    TSubMenu.Rows.Add(tRow)
                End If
                '-------------------------------------------
                perm = CommUtil.GetAssignRight("ac000", Session("s_id"))
                If (InStr(perm, "n")) Then
                    i = i + 1
                    HyperMainMenuGen(i, "userlist", "帳號管理", "~/usermgm/userlist.aspx?smid=userlist", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "userlist") Then
                        HyperSubMenuGen(tRow, 1, "useradd", "新增帳號", "~/usermgm/useradd.aspx?smid=userlist&smode=1", smenusize, perm, "n")
                        TSubMenu.Rows.Add(tRow)
                    End If
                End If
                '-------------------------------------------
                perm = CommUtil.GetAssignRight("mf000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, "molist", "製造管理", "~/wo/molist.aspx?smid=molist&smode=1", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "molist") Then 'smid:橫向次menu smode:次menu順序
                        HyperSubMenuGen(tRow, 1, "wo_" & i, "機台總表", "~/wo/molist.aspx?smid=molist&smode=1&indexpage=" & indexpage, smenusize, "nouse", "")
                        'HyperSubMenuGen(tRow, 2, "wo_" & i + 1, "備庫工單", "~/wo/molist.aspx?smid=molist&smode=2", smenusize, "nouse", "")
                        'HyperSubMenuGen(tRow, 3, "wo_" & i + 2, "半成品工單", "~/wo/molist.aspx?smid=molist&smode=3", smenusize, "nouse", "")
                        perm = CommUtil.GetAssignRight("mf201", Session("s_id"))
                        HyperSubMenuGen(tRow, 4, "wo_" & i + 3, "工單開立", "~/wo/moadd_sys.aspx?mode=create1&smid=molist&smode=4", smenusize, perm, "e")
                        perm = CommUtil.GetAssignRight("mf203", Session("s_id")) 'p203 是發料通知
                        HyperSubMenuGen(tRow, 5, "wo_" & i + 4, "領料需求", "~/wo/issuedwoinfo.aspx?smid=molist&smode=5", smenusize, perm, "e")
                        perm = CommUtil.GetAssignRight("mf205", Session("s_id"))
                        HyperSubMenuGen(tRow, 6, "wo_" & i + 5, "機台進度", "~/wo/wostatus.aspx?smid=molist&smode=6", smenusize, perm, "e")
                        perm = CommUtil.GetAssignRight("mf204", Session("s_id"))
                        HyperSubMenuGen(tRow, 7, "wo_" & i + 6, "加工總表", "~/cnc/cncmain.aspx?act=showlist&smid=molist&smode=7", smenusize, perm, "e")
                        perm = "nouse"
                        HyperSubMenuGen(tRow, 8, "wo_" & i + 7, "料況查詢", "~/wo/pcmaterial.aspx?act=showlist&smid=molist&smode=8", smenusize, perm, "")
                        TSubMenu.Rows.Add(tRow)
                        'HyperSubMenuGen(tRow, 8, "wo_" & i + 7, "機台規格", "~/cnc/cncmain.aspx?act=showlist&smid=molist&smode=8", smenusize, perm, "e")
                        'TSubMenu.Rows.Add(tRow)
                    End If
                End If
                '-------------------------------------------
                perm = CommUtil.GetAssignRight("qc000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "品質管理", "~/qc/qc.aspx?smid=qc&smode=1&funindex=4", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "qc") Then 'smid:橫向次menu smode:次menu順序
                        perm = CommUtil.GetAssignRight("qc100", Session("s_id"))
                        'HyperSubMenuGen(tRow, 1, "qc_" & i, "進料檢驗", "~/qc/iqc.aspx?smid=qc&smode=1&mode=showempty&iqctype=0", smenusize, perm, "e")
                        HyperSubMenuGen(tRow, 1, "qc_" & i, "進料檢驗", "~/qc/qc.aspx?smid=qc&smode=1&funindex=4", smenusize, perm, "e")
                        perm = CommUtil.GetAssignRight("qc200", Session("s_id"))
                        HyperSubMenuGen(tRow, 2, "qc_" & i + 1, "在線檢驗", "~/qc/pqc.aspx?smid=qc&smode=2", smenusize, perm, "e")
                        perm = CommUtil.GetAssignRight("qc300", Session("s_id"))
                        HyperSubMenuGen(tRow, 3, "qc_" & i + 2, "出貨檢驗", "~/qc/oqc.aspx?smid=qc&smode=3", smenusize, perm, "e")
                        TSubMenu.Rows.Add(tRow)
                    End If
                End If
                '-------------------------------------------
                perm = CommUtil.GetAssignRight("sp000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "備品管理", "~/spare/spmaterial.aspx?smid=sp&smode=1&mode=init&allwhs=single", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "sp") Then 'smid:橫向次menu smode:次menu順序
                        perm = CommUtil.GetAssignRight("sp100", Session("s_id"))
                        HyperSubMenuGen(tRow, 1, "sp_" & i, "備品料況", "~/spare/spmaterial.aspx?smid=sp&smode=1&mode=init&allwhs=single", smenusize, perm, "e")
                        perm = CommUtil.GetAssignRight("sp200", Session("s_id"))
                        HyperSubMenuGen(tRow, 2, "sp_" & i + 1, "未列帳舊料管控", "~/spare/spold.aspx?smid=sp&smode=2&fmode=show", 120, perm, "e")
                        TSubMenu.Rows.Add(tRow)
                    End If
                End If
                '-------------------------------------------
                perm = CommUtil.GetAssignRight("sa000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "業務管理", "~/sales/forecastpo.aspx?smode=1&machineradioindex=0&sspradioindex=0&fmode=show", mmenusize, perm, "e")
                    'If (Request.QueryString("smid") = "sa") Then
                    'perm = CommUtil.GetAssignRight("sa100", Session("s_id"))
                    'HyperSubMenuGen(tRow, 1, "sa_" & i, "預估訂單", "~/sales/forecastpo.aspx?smid=sa&smode=1", smenusize, perm, "e")
                    'TSubMenu.Rows.Add(tRow)
                    'End If
                End If

                '-------------------------------------------
                perm = CommUtil.GetAssignRight("sg000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "簽核管理", "~/signoff/signoff.aspx?smid=sg&smode=1&signflowmode=0", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "sg") Then
                        perm = CommUtil.GetAssignRight("sg100", Session("s_id"))
                        HyperSubMenuGen(tRow, 1, "sg_" & i, "簽核總表", "~/signoff/signoff.aspx?smid=sg&smode=1&signflowmode=0", smenusize, perm, "e")
                        TSubMenu.Rows.Add(tRow)
                        perm = CommUtil.GetAssignRight("sg200", Session("s_id"))
                        If (InStr(perm, "e")) Then
                            HyperSubMenuGen(tRow, 2, "sg_" & i + 1, "簽核內容", "", smenusize, perm, "e")
                            TSubMenu.Rows.Add(tRow)
                        End If
                        perm = CommUtil.GetAssignRight("sg300", Session("s_id"))
                        If (InStr(perm, "e")) Then
                            HyperSubMenuGen(tRow, 3, "sg_" & i + 2, "簽核人設定", "~/signoff/signoffsetup.aspx?smid=sg&smode=3&mode=init", smenusize, perm, "e")
                            TSubMenu.Rows.Add(tRow)
                        End If

                        perm = "nouse"
                        'perm = CommUtil.GetAssignRight("sg400", Session("s_id"))
                        'If (InStr(perm, "e")) Then '設定自己,不須設限
                        HyperSubMenuGen(tRow, 4, "sg_" & i + 3, "代理簽核設定", "~/signoff/agnsetup.aspx?smid=sg&smode=4", smenusize, perm, "") '最後一個空白,表示不設限
                        TSubMenu.Rows.Add(tRow)
                        'End If
                        perm = CommUtil.GetAssignRight("sg500", Session("s_id"))
                        If (InStr(perm, "e")) Then
                            HyperSubMenuGen(tRow, 5, "sg_" & i + 4, "管理簽核", "~/signoff/signoffvip.aspx?smid=sg&smode=5", smenusize, perm, "e")
                            TSubMenu.Rows.Add(tRow)
                        End If
                        perm = "nouse"
                        'perm = CommUtil.GetAssignRight("sg500", Session("s_id"))
                        'If (InStr(perm, "e")) Then
                        HyperSubMenuGen(tRow, 6, "sg_" & i + 5, "單據追蹤", "~/signoff/signofftodo.aspx?smid=sg&smode=6&formtypeindex=0&inchargeindex=9999&uid=" & Session("s_id") & "&inchargeid=" & Session("s_id"), smenusize, perm, "")
                        TSubMenu.Rows.Add(tRow)
                        'End If
                        perm = CommUtil.GetAssignRight("sg700", Session("s_id"))
                        If (InStr(perm, "e")) Then
                            HyperSubMenuGen(tRow, 7, "sg_" & i + 6, "設定工具", "~/signoff/signofftool.aspx?smid=sg&smode=7", smenusize, perm, "e")
                            TSubMenu.Rows.Add(tRow)
                        End If
                    End If
                End If
                perm = CommUtil.GetAssignRight("hr000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "人事管理", "~/hr/leave.aspx?smid=hr&smode=1&fmode=show", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "hr") Then
                        perm = CommUtil.GetAssignRight("hr100", Session("s_id"))
                        HyperSubMenuGen(tRow, 1, "hr_" & i, "請假狀況", "~/hr/leave.aspx?smid=hr&smode=1&fmode=show", smenusize, perm, "e")
                        TSubMenu.Rows.Add(tRow)
                        perm = CommUtil.GetAssignRight("hr200", Session("s_id"))
                        If (InStr(perm, "e")) Then
                            HyperSubMenuGen(tRow, 2, "hr_" & i + 1, "外出狀況", "~/hr/outside.aspx?smid=hr&smode=2&fmode=show", smenusize, perm, "e")
                            TSubMenu.Rows.Add(tRow)
                        End If
                        perm = CommUtil.GetAssignRight("hr300", Session("s_id"))
                        If (InStr(perm, "e")) Then
                            HyperSubMenuGen(tRow, 3, "hr_" & i + 2, "獎懲事蹟", "~/hr/workevent.aspx?smid=hr&smode=3&fmode=show", smenusize, perm, "e")
                            TSubMenu.Rows.Add(tRow)
                        End If
                    End If
                End If
                perm = CommUtil.GetAssignRight("pu000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "採購管理", "~/pu/qv.aspx?smid=pu&smode=1&fmode=show", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "pu") Then
                        perm = CommUtil.GetAssignRight("pu100", Session("s_id"))
                        HyperSubMenuGen(tRow, 1, "pu_" & i, "合格廠商", "~/pu/qv.aspx?smid=pu&smode=1&fmode=show", smenusize, perm, "e")
                        TSubMenu.Rows.Add(tRow)
                    End If
                End If
                perm = CommUtil.GetAssignRight("rd000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "研發管理", "~/rd/mcoderule.aspx?smid=rd&smode=1&fmode=show", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "rd") Then
                        perm = CommUtil.GetAssignRight("rd100", Session("s_id"))
                        HyperSubMenuGen(tRow, 1, "rd_" & i, "料號管理", "~/rd/mcoderule.aspx?smid=rd&smode=1&fmode=show", smenusize, perm, "e")
                        TSubMenu.Rows.Add(tRow)
                    End If
                End If
                Dim nowdate, str(), begindate As String
                nowdate = Format(Now, "yyyy/MM/dd")
                str = Split(nowdate, "/")
                begindate = str(0) & "/01/01"
                'begindate = "2020/10/01"
                perm = CommUtil.GetAssignRight("fd000", Session("s_id"))
                If (InStr(perm, "e")) Then
                    i = i + 1
                    HyperMainMenuGen(i, i, "財務管理", "~/fd/freport.aspx?smid=fd&begindate=" & begindate & "&enddate=" & nowdate & "&materialtype=1&reportindex=1", mmenusize, perm, "e")
                    If (Request.QueryString("smid") = "fd") Then
                        perm = CommUtil.GetAssignRight("fd100", Session("s_id"))
                        HyperSubMenuGen(tRow, 1, "fd_" & i, "報表查詢", "~/fd/freport.aspx?smid=fd&begindate=" & begindate & "&enddate=" & nowdate & "&materialtype=1&reportindex=1", smenusize, perm, "e")
                        TSubMenu.Rows.Add(tRow)
                        'perm = CommUtil.GetAssignRight("fd200", Session("s_id"))
                        'HyperSubMenuGen(tRow, 1, "fd_" & i, "固資列表", "~/fd/falist.aspx?smid=as&smode=1&fmode=show", smenusize, perm, "e")
                        'TSubMenu.Rows.Add(tRow)
                        'perm = CommUtil.GetAssignRight("fd300", Session("s_id"))
                        'HyperSubMenuGen(tRow, 1, "fd_" & i + 1, "固資異動", "~/fd/fachange.aspx?smid=as&smode=1&fmode=show", smenusize, perm, "e")
                        'TSubMenu.Rows.Add(tRow)
                    End If
                End If
            Else
                If (Session("actmode") = "todoitem" Or Session("actmode") = "informtraceperson") Then
                    HyperSubMenuGen(tRow, 6, "sg_6", "單據追蹤", "", smenusize, perm, "e")
                    TSubMenu.Rows.Add(tRow)
                ElseIf (Session("actmode") = "signoff" Or Session("actmode") = "single_signoff") Then
                    HyperSubMenuGen(tRow, 2, "sg_2", "簽核內容", "", smenusize, perm, "e")
                    TSubMenu.Rows.Add(tRow)
                End If
            End If
            ''-------------------------------------------
            'i = i + 1
            ''perm = CommUtil.GetAssignRight("p100", Session("s_id"))
            'HyperMainMenuGen(i, i, "研發管理", "", mmenusize, perm, "e")
            ''-------------------------------------------
            'i = i + 1
            ''perm = CommUtil.GetAssignRight("p100", Session("s_id"))
            'HyperMainMenuGen(i, i, "財務管理", "", mmenusize, perm, "e")
            ''-------------------------------------------
            'i = i + 1
            ''perm = CommUtil.GetAssignRight("p100", Session("s_id"))
            'HyperMainMenuGen(i, i, "客服管理", "", mmenusize, perm, "e")
            ''-------------------------------------------

            ''-------------------------------------------
        Else
            i = 0
            'If (Session("s_id") = "ron") Then
            'HyperMainMenuGen(i, "", "Timer數: " & Application("timer_count"), "", mmenusize, "nouse", "")
            'i = i + 1
            'End If
            HyperMainMenuGen(i, "login", "登入", "~/usermgm/login.aspx", 140, "nouse", "")
        End If
    End Sub

    Protected Sub DDLDBS_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        If (DDLDBS.SelectedIndex = 0) Then
            DDLDBS.SelectedValue = Session("usingdb")
            CommUtil.ShowMsg(Me,"需選擇資料庫")
        Else
            'CommUtil.InitSAPSQLConnection(connsap)
            SqlCmd = "SELECT T0.[WhsCode], T0.[WhsName] FROM OWHS T0 order by T0.WhsCode"
            'myCommand = New SqlCommand(SqlCmd, connsap)
            'dr = myCommand.ExecuteReader()
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            DDLWhs.Items.Clear()
            DDLWhs.Items.Add("請選擇倉別")
            Do While (dr.Read())
                DDLWhs.Items.Add(dr(0) & " " & dr(1))
            Loop
            dr.Close()
            connsap.Close()
            Session("usingdb") = DDLDBS.SelectedValue
        End If
    End Sub

    Protected Sub DDLWhs_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim str() As String
        'If (DDLWhs.SelectedIndex = 0) Then
        'DDLWhs.SelectedValue = Session("usingwhsfull")
        'CommUtil.ShowMsg(Me,"需選擇倉別")
        'Else
        str = Split(DDLWhs.SelectedValue, " ")
            Session("usingwhs") = str(0)
            Session("usingwhsfull") = DDLWhs.SelectedValue
        'End If
    End Sub
End Class