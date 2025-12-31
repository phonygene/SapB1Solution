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
        tCell = New TableCell()
        Hyper = New HyperLink()
        Hyper.ID = objid
        Hyper.Text = text
        If (width <> 0) Then
            Hyper.Width = width
        End If
        Hyper.NavigateUrl = url
        Hyper.Font.Underline = False
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
        If (Request.QueryString("smode") <> col) Then
            Hyper = New HyperLink()
            Hyper.ID = objid
            Hyper.Text = text
            If (width <> 0) Then
                Hyper.Width = width
            End If
            Hyper.NavigateUrl = url
            Hyper.Font.Underline = False
            tCell.Controls.Add(Hyper)
            If (perms <> "nouse") Then
                CommUtil.DisableObjectByPermission(Hyper, perms, keyp)
            End If
        Else
            tCell.Text = text
            If (width <> 0) Then
                tCell.Width = width
            End If
            tCell.CssClass = "submenu-active"
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
        'If (Session("s_id") = "ron" Or Session("s_id") = "su") Then
        'DDLDBS.Enabled = True
        'Else
        DDLDBS.Enabled = False
        'End If
        tCell.Controls.Add(DDLDBS)
        tRow.Cells.Add(tCell)
    End Sub

    ' [修改] 此版本不使用倉庫選單，保留函數但不再調用
    Sub WhsDropList(tRow As TableRow)
        ' 此版本不使用生產製造功能，不再顯示倉庫選單
        ' Dim tCell As TableCell
        ' tCell = New TableCell()
        ' DDLWhs = New DropDownList()
        ' DDLWhs.Items.Clear()
        ' DDLWhs.Items.Add("C01 ICT")
        ' DDLWhs.Items.Add("C02 AOI")
        ' DDLWhs.SelectedValue = Session("usingwhsfull")
        ' DDLWhs.ID = "ddl_whs"
        ' DDLWhs.Width = 100
        ' AddHandler DDLWhs.SelectedIndexChanged, AddressOf DDLWhs_SelectedIndexChanged
        ' DDLWhs.AutoPostBack = True
        ' DDLWhs.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='Gainsboro'")
        ' DDLWhs.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        ' tCell.Controls.Add(DDLWhs)
        ' tRow.Cells.Add(tCell)
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
                HyperMainMenuGen(i, "index", "首頁", "Home.aspx?smid=index&smode=0", mmenusize, "nouse", "")
                If (Request.QueryString("smid") = "index") Then
                    HyperSubMenuGen(tRow, 1, "pwdchange", "修改密碼", "~/usermgm/pwdchange.aspx?smid=index&smode=1", smenusize, "nouse", "")
                    ' [修改] SAP帳密設定只給管理員 (ron) 使用
                    If (Session("s_id") = "ron") Then
                        HyperSubMenuGen(tRow, 2, "addsapid", "設定SAP帳密", "~/usermgm/addsapid.aspx?smid=index&smode=2", smenusize, "nouse", "")
                        DataBaseDropList(tRow)
                    End If
                    ' [修改] 移除倉庫選單，此版本不使用
                    ' WhsDropList(tRow)
                    TSubMenu.Rows.Add(tRow)
                End If

                ' [修改] 費用申請 - 所有登入使用者皆可使用（移到前面，不需權限檢查）
                i = i + 1
                HyperMainMenuGen(i, "expense", "費用申請", "~/ExpenseClaimForm.aspx?smid=ec&smode=1", mmenusize, "nouse", "")

                ' [修改] 單據查詢 - 所有登入使用者皆可使用（移到前面）
                i = i + 1
                HyperMainMenuGen(i, "docsearch", "單據查詢", "~/DocumentSearch.aspx", mmenusize, "nouse", "")
                ' ============================================
                ' [2024-12-31] 暫時隱藏以下所有功能表項目
                ' 目前只顯示：登出、首頁、修改密碼、費用申請、單據查詢
                ' ============================================

                ' [隱藏] 帳號管理
                'perm = CommUtil.GetAssignRight("ac000", Session("s_id"))
                'If (InStr(perm, "n")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, "userlist", "帳號管理", "~/usermgm/userlist.aspx?smid=userlist", mmenusize, perm, "e")
                '    If (Request.QueryString("smid") = "userlist") Then
                '        HyperSubMenuGen(tRow, 1, "useradd", "新增帳號", "~/usermgm/useradd.aspx?smid=userlist&smode=1", smenusize, perm, "n")
                '        TSubMenu.Rows.Add(tRow)
                '    End If
                'End If

                ' [隱藏] 製造管理
                'perm = CommUtil.GetAssignRight("mf000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, "molist", "製造管理", "~/wo/molist.aspx?smid=molist&smode=1", mmenusize, perm, "e")
                'End If

                ' [隱藏] 品質管理
                'perm = CommUtil.GetAssignRight("qc000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "品質管理", "~/qc/qc.aspx?smid=qc&smode=1&funindex=4", mmenusize, perm, "e")
                'End If

                ' [隱藏] 備品管理
                'perm = CommUtil.GetAssignRight("sp000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "備品管理", "~/spare/spmaterial.aspx?smid=sp&smode=1&mode=init&allwhs=single", mmenusize, perm, "e")
                'End If

                ' [隱藏] 業務管理
                'perm = CommUtil.GetAssignRight("sa000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "業務管理", "~/sales/forecastpo.aspx?smode=1&machineradioindex=0&sspradioindex=0&fmode=show", mmenusize, perm, "e")
                'End If

                ' [隱藏] 簽核管理
                'perm = CommUtil.GetAssignRight("sg000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "簽核管理", "~/signoff/signoff.aspx?smid=sg&smode=1&signflowmode=0", mmenusize, perm, "e")
                'End If

                ' [隱藏] 人事管理
                'perm = CommUtil.GetAssignRight("hr000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "人事管理", "~/hr/leave.aspx?smid=hr&smode=1&fmode=show", mmenusize, perm, "e")
                'End If

                ' [隱藏] 採購管理
                'perm = CommUtil.GetAssignRight("pu000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "採購管理", "~/pu/qv.aspx?smid=pu&smode=1&fmode=show", mmenusize, perm, "e")
                'End If

                ' [隱藏] 研發管理
                'perm = CommUtil.GetAssignRight("rd000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "研發管理", "~/rd/mcoderule.aspx?smid=rd&smode=1&fmode=show", mmenusize, perm, "e")
                'End If

                ' [隱藏] 財務管理
                'perm = CommUtil.GetAssignRight("fd000", Session("s_id"))
                'If (InStr(perm, "e")) Then
                '    i = i + 1
                '    HyperMainMenuGen(i, i, "財務管理", "~/fd/freport.aspx?smid=fd&...", mmenusize, perm, "e")
                'End If

                ' ============================================
                ' 以上功能表項目暫時隱藏
                ' ============================================
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