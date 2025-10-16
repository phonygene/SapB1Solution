Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class cncmain
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public conn2, connsap As New SqlConnection
    Public myCommand, myCommand1 As SqlCommand
    Public SqlCmd As String
    Public oCompany As New SAPbobsCOM.Company
    Public ret As Long
    Public ds As New DataSet
    Public act As String
    Public permsmf204 As String
    Public dr, drsap As SqlDataReader
    Public modifynum As Long
    Public ScriptManager1 As New ScriptManager
    Public Sub InitLocalSQLConnection()
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        conn.Open()
    End Sub
    Public Sub InitLocalSQLConnection2()
        conn2.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        conn2.Open()
    End Sub

    Public Sub InitSAPSQLConnection()
        connsap.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
        connsap.Open()
    End Sub

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

    Public Sub CloseSAPSQLConnection()
        connsap.Close()
    End Sub
    Public Sub CloseLocalSQLConnection()
        conn.Close()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Dim act1 As String
        Page.Form.Controls.Add(ScriptManager1)
        permsmf204 = CommUtil.GetAssignRight("mf204", Session("s_id"))
        If (Not IsPostBack) Then 'Hyper不是 postback , 但button 是 , 故此處可以去除這些action
            gv1.PageIndex = Request.QueryString("indexpage")
            act = Request.QueryString("act")
            'MsgBox(act)
            act1 = Request.QueryString("act1")
            If (act = "updcncsta" Or act = "updpsta" Or act = "updrawsta") Then
                UpdateStatus()
            ElseIf (act = "del") Then
                DeleteWorking()
                Response.Redirect("cncmain.aspx?act=showlist&smid=molist&smode=7")
            ElseIf (act = "createwitem") Then
                'CreateCncWItem()
                'Response.Redirect("cncmain.aspx?act=showlist&smid=molist&smode=7")
            End If
            If (act1 = "add") Then
                CommUtil.ShowMsg(Me, "新增成功")
            ElseIf (act1 = "modify") Then
                CommUtil.ShowMsg(Me, "修改成功")
            End If
        End If
        modifynum = Request.QueryString("modifynum")
        ShowCnc()
        If (Not IsPostBack And (act = "createwitem" Or act = "modifywitem")) Then
            CreateCncWItem(act)
        End If
    End Sub

    Sub UpdateStatus()
        Dim num As Long
        Dim postatus As Integer
        Dim indexpage As Integer
        num = Request.QueryString("num")
        postatus = Request.QueryString("postatus")
        indexpage = Request.QueryString("indexpage")
        If (postatus = 0) Then
            postatus = 10
        ElseIf (postatus = 10) Then
            postatus = 20
        ElseIf (postatus = 20) Then
            postatus = 90
        ElseIf (postatus = 90) Then
            postatus = 0
        End If

        If (act = "updcncsta") Then
            SqlCmd = "update dbo.[ocnc] set cncsta= " & postatus & " where num=" & num
        ElseIf (act = "updpsta") Then
            SqlCmd = "update dbo.[ocnc] set psta= " & postatus & " where num=" & num
        ElseIf (act = "updrawsta") Then
            SqlCmd = "update dbo.[ocnc] set rawsta= " & postatus & " where num=" & num
        End If
        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub

    Sub DeleteWorking()
        Dim num, sappo As Long
        Dim indexpage As Integer
        num = Request.QueryString("num")
        sappo = Request.QueryString("sappo")
        indexpage = Request.QueryString("indexpage")
        SqlCmd = "delete from dbo.[ocnc]  where num=" & num
        CommUtil.SqlLocalExecute("del", SqlCmd, conn)
        conn.Close()
        'delete cnc1
        SqlCmd = "delete from dbo.[cnc1]  where sappo=" & sappo
        CommUtil.SqlLocalExecute("del", SqlCmd, conn)
        conn.Close()
        CommUtil.ShowMsg(Me, "已刪除")
    End Sub
    Sub ShowCnc()
        ShowCncAddAndFilter()
        ShowGridView()
    End Sub

    Sub ShowCncAddAndFilter()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim Txtx As TextBox
        Dim DDLx As DropDownList
        Dim Btnx As Button
        Dim ce As CalendarExtender
        Dim BtnTestx As Button

        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Wrap = False
        tCell.Font.Bold = True

        Labelx = New Label()
        Labelx.ID = "label_po"
        Labelx.Text = "採購單:"
        tCell.Controls.Add(Labelx)
        Txtx = New TextBox()
        Txtx.ID = "txt_po"
        ViewState("po") = Txtx.ID
        Txtx.Width = 50
        tCell.Controls.Add(Txtx)
        '----------------------------
        Labelx = New Label()
        Labelx.ID = "label_cdate"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp加工日期:"
        tCell.Controls.Add(Labelx)
        Txtx = New TextBox()
        Txtx.ID = "txt_cdate"
        ViewState("cdate") = Txtx.ID
        Txtx.Width = 80
        ce = New CalendarExtender
        ce.TargetControlID = Txtx.ID
        ce.ID = "ce_create"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(Txtx)
        ' ----------------------------
        Labelx = New Label()
        Labelx.ID = "label_shipdate"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp預交日期:"
        tCell.Controls.Add(Labelx)
        Txtx = New TextBox()
        Txtx.ID = "txt_shipdate"
        ViewState("shipdate") = Txtx.ID
        Txtx.Width = 80
        ce = New CalendarExtender
        ce.TargetControlID = Txtx.ID
        ce.ID = "ce_ship"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(Txtx)
        '----------------------------
        Labelx = New Label()
        Labelx.ID = "label_vender"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp廠商:"
        tCell.Controls.Add(Labelx)
        DDLx = New DropDownList()
        DDLx.ID = "ddl_vender"
        ViewState("vender") = DDLx.ID
        DDLx.Width = 80
        'SqlCmd = "SELECT IsNull(T0.[AliasName],''), T0.[CardCode] FROM OCRD T0 where T0.QryGroup10='Y' order by T0.CardCode"
        'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        DDLx.Items.Clear()
        DDLx.Items.Add("CNC部")
        'If (dr.HasRows) Then
        '    Do While (dr.Read())
        '        'DDLx.Items.Add(dr(1) & " " & dr(0))
        '        DDLx.Items.Add(dr(0))
        '    Loop
        'End If
        'DDLx.SelectedIndex = 0
        'dr.Close()
        'connsap.Close()

        'DDLx.Items.Clear()
        'DDLx.Items.Add("選擇廠商")
        'DDLx.Items.Add("CNC部")
        'DDLx.Items.Add("奇典")
        'DDLx.Items.Add("國興")
        'DDLx.SelectedIndex = 1
        tCell.Controls.Add(DDLx)
        '----------------------------
        Labelx = New Label()
        Labelx.ID = "label_comm"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp備註:"
        tCell.Controls.Add(Labelx)
        Txtx = New TextBox()
        Txtx.ID = "txt_comm"
        ViewState("comm") = Txtx.ID
        Txtx.Width = 400
        tCell.Controls.Add(Txtx)
        '----------------------------
        Labelx = New Label()
        Labelx.ID = "label_add"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Btnx = New Button()
        Btnx.ID = "btn_add_" & modifynum
        CommUtil.DisableObjectByPermission(Btnx, permsmf204, "m")
        Btnx.Text = "新增"
        AddHandler Btnx.Click, AddressOf Btnx_Click
        tCell.Controls.Add(Btnx)

        BtnTestx = New Button()
        BtnTestx.ID = "btn_test"
        CommUtil.DisableObjectByPermission(Btnx, permsmf204, "m")
        BtnTestx.Text = "Test"
        BtnTestx.Visible = False 'disable
        AddHandler BtnTestx.Click, AddressOf BtnTestx_Click
        tCell.Controls.Add(BtnTestx)

        tRow.Cells.Add(tCell)
        CncAddT.Rows.Add(tRow)

        '''''''''''''''''''''''''''''''''''''''''''''''Filter Table
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Wrap = False
        tCell.Text = "Filter"
        tCell.Font.Bold = True
        tRow.Cells.Add(tCell)
        CncFilterT.Rows.Add(tRow)
    End Sub

    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        'Dim btn As Button
        Dim tTxt As TextBox
        Dim cChk As CheckBox
        Dim Hyper As HyperLink
        Dim total, notf As Integer
        Dim pstatus As Integer
        Dim doingflag As Boolean
        doingflag = True
        'Dim indexpage As Integer
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            'If (ds.Tables(0).Rows(e.Row.RowIndex)("num") <> 0) Then 'CNC自製
            If (e.Row.Cells(2).Text <> 0) Then 'CNC自製
                e.Row.Cells(9).ToolTip = e.Row.Cells(9).Text
                If (e.Row.Cells(9).Text.Length > 25) Then
                    e.Row.Cells(9).Text = e.Row.Cells(9).Text.Substring(0, 24) + "..."
                End If
                Hyper = New HyperLink()
                Hyper.Text = "修改_" & e.Row.Cells(2).Text
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_modify_" & e.Row.RowIndex
                CommUtil.DisableObjectByPermission(Hyper, permsmf204, "n")
                Hyper.NavigateUrl = "cncmain.aspx?sappo=" & e.Row.Cells(1).Text &
                                "&indexpage=" & gv1.PageIndex &
                                "&cdate=" & e.Row.Cells(3).Text &
                                "&ship_date=" & e.Row.Cells(8).Text &
                                "&vender=" & e.Row.Cells(5).Text &
                                "&comm=" & e.Row.Cells(9).ToolTip &
                                "&modifynum=" & e.Row.Cells(2).Text &
                                "&act=modifywitem&smid=molist&smode=7"
                e.Row.Cells(2).Controls.Add(Hyper)

                '未完/總數
                SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & e.Row.Cells(1).Text
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                total = dr(0)
                dr.Close()
                conn.Close()
                SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & e.Row.Cells(1).Text & " and T0.stat<>90"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                notf = dr(0)
                dr.Close()
                conn.Close()
                e.Row.Cells(4).Text = notf & "/" & total

                Hyper = New HyperLink()
                'If (ds.Tables(0).Rows(e.Row.RowIndex)("cncsta") <> 20) Then
                If (e.Row.Cells(6).Text <> 20) Then
                    SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & e.Row.Cells(1).Text & " and (T0.stat=10 or T0.stat=90)"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    If (notf = 0) Then
                        Hyper.Text = "已完工"
                        e.Row.Cells(6).BackColor = Drawing.Color.LightGreen
                        pstatus = 90
                    ElseIf (dr(0) <> 0) Then
                        Hyper.Text = "加工中"
                        e.Row.Cells(6).BackColor = Drawing.Color.Yellow
                        pstatus = 10
                    Else
                        Hyper.Text = "未開工"
                        e.Row.Cells(6).BackColor = Drawing.Color.White
                        pstatus = 0
                        doingflag = False
                    End If
                    dr.Close()
                    conn.Close()
                    If (ds.Tables(0).Rows(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("cncsta") <> pstatus) Then
                        SqlCmd = "update dbo.[ocnc] set cncsta= " & pstatus & " where num=" & e.Row.Cells(2).Text 'ds.Tables(0).Rows(e.Row.RowIndex)("num")
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                    End If
                    'If (ds.Tables(0).Rows(e.Row.RowIndex)("cncsta") = 0) Then
                    '    Hyper.Text = "未開工"
                    '    e.Row.Cells(6).BackColor = Drawing.Color.White
                    'ElseIf (ds.Tables(0).Rows(e.Row.RowIndex)("cncsta") = 10) Then
                    '    Hyper.Text = "加工中"
                    '    e.Row.Cells(6).BackColor = Drawing.Color.Yellow
                    'ElseIf (ds.Tables(0).Rows(e.Row.RowIndex)("cncsta") = 20) Then
                    '    Hyper.Text = "暫停"
                    '    e.Row.Cells(6).BackColor = Drawing.Color.MediumSeaGreen
                    'ElseIf (ds.Tables(0).Rows(e.Row.RowIndex)("cncsta") = 90) Then
                    '    Hyper.Text = "已完工"
                    '    e.Row.Cells(6).BackColor = Drawing.Color.LightGreen
                    'End If
                Else
                    Hyper.Text = "暫停"
                    e.Row.Cells(6).BackColor = Drawing.Color.MediumVioletRed
                End If
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_cncsta_" & e.Row.Cells(2).Text
                'Hyper.NavigateUrl = "cncmain.aspx?act=updcncsta&num=" & e.Row.Cells(2).Text &
                '                "&postatus=" & ds.Tables(0).Rows(e.Row.RowIndex)("cncsta") & "&indexpage=" & gv1.PageIndex &
                '                "&smid=molist&smode=7"
                Hyper.NavigateUrl = "cncmain.aspx?act=updcncsta&num=" & e.Row.Cells(2).Text &
                                "&postatus=" & e.Row.Cells(6).Text & "&indexpage=" & gv1.PageIndex &
                                "&smid=molist&smode=7"
                CommUtil.DisableObjectByPermission(Hyper, permsmf204, "m")
                Hyper.Enabled = False
                e.Row.Cells(6).Controls.Add(Hyper)

                Hyper = New HyperLink()
                'If (ds.Tables(0).Rows(e.Row.RowIndex)("psta") <> 10) Then
                SqlCmd = "select Sum(T1.quantity),Sum(T1.opencreqty),count(*) " &
                "from dbo.OPOR T0 inner join POR1 T1 on T0.docentry=T1.docentry where T0.docnum=" &
                ds.Tables(0).Rows(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("sappo")
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    'MsgBox(ds.Tables(0).Rows(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("sappo") & " " & dr(0) & " " & dr(1) & " " & dr(2))
                    If (dr(1) = 0) Then
                        Hyper.Text = "已入庫"
                        e.Row.Cells(7).BackColor = Drawing.Color.LightGreen
                        pstatus = 90
                        Hyper.Enabled = False
                    ElseIf (total <> dr(2)) Then
                        If (doingflag = True) Then
                            Hyper.Text = "部分入庫"
                            e.Row.Cells(7).BackColor = Drawing.Color.MediumSeaGreen
                            pstatus = 20
                            Hyper.Enabled = False
                        Else
                            Hyper.Text = "細項差異"
                            e.Row.Cells(7).BackColor = Drawing.Color.Red
                            pstatus = 20
                            Hyper.Enabled = False
                        End If
                    ElseIf (total = dr(2)) Then
                        Hyper.Text = "未處理"
                        e.Row.Cells(7).BackColor = Drawing.Color.White
                        pstatus = 0
                    Else
                        Hyper.Text = "送表處"
                        e.Row.Cells(7).BackColor = Drawing.Color.Yellow
                        pstatus = 10
                    End If
                    If (dr(1) <> 0) Then
                        tTxt = New TextBox
                        tTxt.ID = "txt_pseq_" & e.Row.Cells(2).Text ' ds.Tables(0).Rows(e.Row.RowIndex)("num")
                        tTxt.Width = 30
                        tTxt.Text = e.Row.Cells(0).Text
                        tTxt.AutoPostBack = True
                        CommUtil.DisableObjectByPermission(tTxt, permsmf204, "n")
                        AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
                        e.Row.Cells(0).Controls.Add(tTxt)
                    End If
                    dr.Close()
                    connsap.Close()
                    Dim newpseq As Long
                    If (ds.Tables(0).Rows(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("psta") <> pstatus) Then
                        SqlCmd = "update dbo.[ocnc] set psta= " & pstatus & " where num=" & e.Row.Cells(2).Text 'ds.Tables(0).Rows(e.Row.RowIndex)("num")
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                        If (pstatus = 90) Then
                            SqlCmd = "select Max(pseq) from dbo.[ocnc]"
                            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
                            dr.Read()
                            If (dr(0) >= 1001) Then
                                newpseq = dr(0) + 1
                            Else
                                newpseq = 1001
                            End If
                            dr.Close()
                            conn.Close()
                            SqlCmd = "update dbo.[ocnc] set pseq=" & newpseq & " where num=" & e.Row.Cells(2).Text 'ds.Tables(0).Rows(e.Row.RowIndex)("num")
                            CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                            conn.Close()
                        End If
                    End If
                Else
                    CommUtil.ShowMsg(Me, "無此PO" & ds.Tables(0).Rows(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("sappo"))
                End If

                Hyper.Font.Underline = False
                Hyper.ID = "hyper_psta_" & e.Row.Cells(2).Text
                'Hyper.NavigateUrl = "cncmain.aspx?act=updpsta&num=" & e.Row.Cells(2).Text &
                '                "&postatus=" & ds.Tables(0).Rows(e.Row.RowIndex)("psta") & "&indexpage=" & gv1.PageIndex &
                '                "&smid=molist&smode=7"
                Hyper.NavigateUrl = "cncmain.aspx?act=updpsta&num=" & e.Row.Cells(2).Text &
                                "&postatus=" & e.Row.Cells(7).Text & "&indexpage=" & gv1.PageIndex &
                                "&smid=molist&smode=7"
                CommUtil.DisableObjectByPermission(Hyper, permsmf204, "m")
                Hyper.Enabled = False
                e.Row.Cells(7).Controls.Add(Hyper)

                'Hyper = New HyperLink()
                'If (e.Row.Cells(10).Text = 0) Then
                '    Hyper.Text = "未採購"
                '    e.Row.Cells(10).BackColor = Drawing.Color.White
                'ElseIf (e.Row.Cells(10).Text = 10) Then
                '    Hyper.Text = "已採購"
                '    e.Row.Cells(10).BackColor = Drawing.Color.Yellow
                'ElseIf (e.Row.Cells(10).Text = 20) Then
                '    Hyper.Text = "已來料"
                '    e.Row.Cells(10).BackColor = Drawing.Color.MediumSeaGreen
                'ElseIf (e.Row.Cells(10).Text = 90) Then
                '    Hyper.Text = "料確認"
                '    e.Row.Cells(10).BackColor = Drawing.Color.LightGreen
                'End If
                'Hyper.Font.Underline = False
                'Hyper.ID = "hyper_rawsta_" & e.Row.Cells(2).Text
                'Hyper.NavigateUrl = "cncmain.aspx?act=updrawsta&num=" & e.Row.Cells(2).Text &
                '                "&postatus=" & e.Row.Cells(10).Text & "&indexpage=" & gv1.PageIndex &
                '                "&smid=molist&smode=7"
                'CommUtil.DisableObjectByPermission(Hyper, permsmf204, "m")
                'e.Row.Cells(10).Controls.Add(Hyper)

                Dim allitemcount, comparecount, comparecount20, comparecount90 As Integer
                SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & e.Row.Cells(1).Text
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr.HasRows) Then
                    dr.Read()
                    allitemcount = dr(0)
                End If
                dr.Close()
                conn.Close()
                SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & e.Row.Cells(1).Text & " and rawstatus<>0"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr.HasRows) Then
                    dr.Read()
                    comparecount = dr(0)
                End If
                dr.Close()
                conn.Close()
                If (comparecount = 0) Then
                    e.Row.Cells(10).Text = "未採購"
                    e.Row.Cells(10).BackColor = Drawing.Color.White
                ElseIf (comparecount < allitemcount) Then
                    e.Row.Cells(10).Text = "部份採購"
                    e.Row.Cells(10).BackColor = Drawing.Color.Yellow
                Else
                    SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & e.Row.Cells(1).Text & " and rawstatus=20"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    If (dr.HasRows) Then
                        dr.Read()
                        comparecount20 = dr(0) '
                    End If
                    dr.Close()
                    conn.Close()
                    SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & e.Row.Cells(1).Text & " and rawstatus=90"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    If (dr.HasRows) Then
                        dr.Read()
                        comparecount90 = dr(0) '
                    End If
                    dr.Close()
                    conn.Close()
                    If (comparecount90 = allitemcount) Then
                        e.Row.Cells(10).Text = "料確認"
                        e.Row.Cells(10).BackColor = Drawing.Color.LightGreen
                    ElseIf ((comparecount20 + comparecount90) < allitemcount And comparecount20 > 0) Then
                        e.Row.Cells(10).Text = "部份來料"
                        e.Row.Cells(10).BackColor = Drawing.Color.Yellow
                    ElseIf (comparecount20 = allitemcount) Then
                        e.Row.Cells(10).Text = "已來料"
                        e.Row.Cells(10).BackColor = Drawing.Color.MediumSeaGreen
                    ElseIf (comparecount90 > 0) Then
                        e.Row.Cells(10).Text = "部份確認"
                        e.Row.Cells(10).BackColor = Drawing.Color.Yellow
                    Else
                        e.Row.Cells(10).Text = "已採購"
                        e.Row.Cells(10).BackColor = Drawing.Color.Yellow
                    End If
                End If
                cChk = New CheckBox
                cChk.ID = "chk_" & e.Row.RowIndex & "_" & e.Row.Cells(2).Text & "_" & gv1.PageIndex
                cChk.AutoPostBack = True
                AddHandler cChk.CheckedChanged, AddressOf cChk_CheckedChanged
                CommUtil.DisableObjectByPermission(Hyper, permsmf204, "d")
                e.Row.Cells(12).Controls.Add(cChk)
            Else '外發或自製未建單
                e.Row.Cells(9).ToolTip = e.Row.Cells(9).Text
                If (e.Row.Cells(9).Text.Length > 25) Then
                    e.Row.Cells(9).Text = e.Row.Cells(9).Text.Substring(0, 24) + "..."
                End If
                e.Row.Cells(6).Text = "NA"
                e.Row.Cells(10).Text = "NA"
                SqlCmd = "Select count(*) " &
                    "from dbo.OPOR T0 inner join POR1 T1 On T0.docentry=T1.docentry where T0.docnum=" & e.Row.Cells(1).Text
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                total = dr(0)
                dr.Close()
                connsap.Close()
                SqlCmd = "Select count(*) " &
                    "from dbo.OPOR T0 inner join POR1 T1 On T0.docentry=T1.docentry where T0.docnum=" &
                    e.Row.Cells(1).Text & "And T1.opencreqty<>0"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                dr.Read()
                notf = dr(0)
                dr.Close()
                connsap.Close()
                e.Row.Cells(4).Text = notf & "/" & total
                '以上共用
                If (e.Row.Cells(0).Text <> 0) Then '外發
                    SqlCmd = "Select Sum(T1.quantity),Sum(T1.opencreqty) " &
                    "from dbo.OPOR T0 inner join POR1 T1 On T0.docentry=T1.docentry where T0.docnum=" &
                    e.Row.Cells(1).Text
                    dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                    dr.Read()
                    If (dr(1) = 0) Then
                        e.Row.Cells(7).Text = "已入庫"
                        e.Row.Cells(7).BackColor = Drawing.Color.LightGreen
                    ElseIf (dr(0) <> dr(1)) Then
                        e.Row.Cells(7).Text = "部分入庫"
                        e.Row.Cells(7).BackColor = Drawing.Color.MediumSeaGreen
                    ElseIf (dr(0) = dr(1)) Then
                        e.Row.Cells(7).Text = "待入庫"
                        e.Row.Cells(7).BackColor = Drawing.Color.White
                    End If
                    dr.Close()
                    connsap.Close()
                Else '自製未建單
                    e.Row.Cells(7).Text = "NA"
                    Hyper = New HyperLink()
                    Hyper.Text = "建單"
                    Hyper.Font.Underline = False
                    Hyper.ID = "hyper_create_" & e.Row.RowIndex
                    Hyper.NavigateUrl = "cncmain.aspx?sappo=" & e.Row.Cells(1).Text &
                                "&indexpage=" & gv1.PageIndex &
                                "&CDate=" & e.Row.Cells(3).Text &
                                "&ship_date=" & e.Row.Cells(8).Text &
                                "&vender=" & e.Row.Cells(5).Text &
                                "&comm=" & e.Row.Cells(9).ToolTip &
                                "&modifynum=" & e.Row.Cells(2).Text &
                                "&act=createwitem&smid=molist&smode=7"
                    CommUtil.DisableObjectByPermission(Hyper, permsmf204, "n")
                    e.Row.Cells(2).Controls.Add(Hyper)
                End If
            End If
            If (e.Row.Cells(0).Text <> 0) Then
                Hyper = New HyperLink()
                Hyper.Text = "細項"
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_action_" & e.Row.RowIndex
                Hyper.NavigateUrl = "cncitems.aspx?sappo=" & e.Row.Cells(1).Text &
                                "&indexpage=" & gv1.PageIndex & "&num=" & e.Row.Cells(2).Text
                e.Row.Cells(11).Controls.Add(Hyper)
            End If
        End If
    End Sub

    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        ShowGridView()
    End Sub

    Sub ShowGridView()
        'Dim indexpage As Integer
        'indexpage = Request.QueryString("indexpage")
        'gv1.PageIndex = 2
        '在SAP中廠商是未結CNC的採購單且還沒在系統建單
        ds.Reset()
        SqlCmd = "Select T0.docnum " &
        "from dbo.OPOR T0 where T0.[DocStatus]='O' and T0.cardcode='T021'"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            Do While (drsap.Read())
                SqlCmd = "Select count(*) From dbo.[ocnc] T0 where T0.sappo=" & drsap(0)
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                If (dr(0) = 0) Then
                    conn.Close()
                    SqlCmd = "select action='',del='',notfinish=0,num=0,pseq=0,cncsta=0,rawsta=0,T0.docnum As sappo,T0.docdate As cdate," &
                    "T0.docduedate As ship_date,T0.comments As comm,T1.aliasname As vender,psta=0 " &
                    "from dbo.OPOR T0 inner join OCRD T1 on T0.cardcode=T1.cardcode where T1.QryGroup10='Y' " & '若無顯示 , 則 check OCRD 之QryGroup10 是否被清除
                    "and T0.[docnum]=" & drsap(0)
                    ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, conn)
                End If
                dr.Close()
                conn.Close()
            Loop
        End If
        drsap.Close()
        connsap.Close()

        SqlCmd = "Select action='',del='',notfinish=0,T0.num,T0.sappo ,T0.pseq, T0.cdate ,T0.vender , T0.cncsta,T0.psta,T0.rawsta, " &
                 "T0.ship_date,T0.comm from dbo.[ocnc] T0 " &
                 "where psta<>90 order by T0.pseq"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()

        'O and C(cancel,close)
        SqlCmd = "select action='',del='',notfinish=0,num=0,pseq=501,cncsta=0,rawsta=0,T0.docnum As sappo,T0.docdate As cdate," &
        "T0.docduedate As ship_date,T0.comments As comm,T1.aliasname As vender,psta=0 " &
        "from dbo.OPOR T0 inner join OCRD T1 on T0.cardcode=T1.cardcode where T1.QryGroup10='Y' " &
        "and T0.[DocStatus]='O' and T0.cardcode<>'T021'"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        '
        SqlCmd = "Select action='',del='',notfinish=0,T0.num,T0.sappo ,T0.pseq, T0.cdate ,T0.vender , T0.cncsta,T0.psta,T0.rawsta, " &
                 "T0.ship_date,T0.comm from dbo.[ocnc] T0 " &
                 "where psta=90 order by T0.pseq"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()

        'ds.Tables(0).Columns.Add("action")
        'ds.Tables(0).Columns.Add("del")
        'ds.Tables(0).Columns.Add("notfinish")
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
    End Sub
    Protected Sub Btnx_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim po As Long
        Dim createdate As String
        Dim ship_date As String
        Dim comm As String
        Dim vender As String
        Dim Txtx As TextBox
        Dim DDLx As DropDownList
        Dim pseq As Integer
        Dim mnum As Long
        Dim str() As String
        Dim go As Boolean
        go = True
        str = Split(sender.ID, "_")
        mnum = str(2)

        Txtx = CncAddT.FindControl(ViewState("po"))
        If (Txtx.Text = "" Or Txtx.Text = "0") Then
            CommUtil.ShowMsg(Me, "沒PO號或PO不能為0")
            go = False
        End If
        po = Txtx.Text
        Txtx = CncAddT.FindControl(ViewState("cdate"))
        If (Txtx.Text = "") Then
            CommUtil.ShowMsg(Me,"沒建立日期")
            go = False
        End If
        createdate = Txtx.Text
        Txtx = CncAddT.FindControl(ViewState("shipdate"))
        If (Txtx.Text = "") Then
            CommUtil.ShowMsg(Me,"沒出貨日期")
            go = False
        End If
        ship_date = Txtx.Text

        DDLx = CncAddT.FindControl(ViewState("vender"))
        If (DDLx.SelectedValue = "") Then
            CommUtil.ShowMsg(Me, "沒廠商")
            go = False
        End If
        vender = DDLx.SelectedValue

        Txtx = CncAddT.FindControl(ViewState("comm"))
        If (Txtx.Text = "") Then
            CommUtil.ShowMsg(Me,"沒備註")
            go = False
        End If
        comm = Txtx.Text
        If (go) Then
            ReSeqWorkingCnc()

            If (sender.Text = "新增") Then
                SqlCmd = "Select IsNull(max(pseq),0) from dbo.[ocnc] where psta<>90"
                dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                dr.Read()
                pseq = dr(0) + 1
                dr.Close()
                conn.Close()
                SqlCmd = "Insert into dbo.[ocnc] (sappo,pseq,cdate,ship_date,vender,comm) " &
                            "values(" & po & "," & pseq & ",'" & createdate & "','" &
                            ship_date & "','" & vender & "','" & comm & "')"
                CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
                conn.Close()
                'Insert cnc1 加工細項 from po
                SqlCmd = "select T1.itemcode,T1.dscription,T1.quantity,T1.whscode " &
                "from dbo.OPOR T0 inner join POR1 T1 on T0.docentry=T1.docentry where T0.docnum=" & po
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    Do While (dr.Read())
                        SqlCmd = "Insert into dbo.[cnc1] (sappo,itemcode,quantity) " &
                                            "values(" & po & ",'" & dr(0) & "'," & dr(2) & ")"
                        CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
                        conn.Close()
                    Loop
                End If
                dr.Close()
                connsap.Close()
                Response.Redirect("cncmain.aspx?act=showlist&smid=molist&smode=7&act1=add")
                CommUtil.ShowMsg(Me, "新增成功")
            ElseIf (sender.Text = "修改") Then
                SqlCmd = "update dbo.[ocnc] set ship_date='" & ship_date & "',comm='" & comm & "' " &
                "where num=" & mnum
                CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
                Response.Redirect("cncmain.aspx?act=showlist&smid=molist&smode=7&act1=modify")
                CommUtil.ShowMsg(Me, "修改成功")
            End If

            'clear
            Txtx = CncAddT.FindControl(ViewState("po"))
            Txtx.Text = ""

            Txtx = CncAddT.FindControl(ViewState("cdate"))
            Txtx.Text = ""

            Txtx = CncAddT.FindControl(ViewState("shipdate"))
            Txtx.Text = ""

            DDLx = CncAddT.FindControl(ViewState("vender"))
            DDLx.SelectedIndex = 0

            Txtx = CncAddT.FindControl(ViewState("comm"))
            Txtx.Text = ""

            CType(CncAddT.FindControl("txt_po"), TextBox).Enabled = True
            CType(CncAddT.FindControl("txt_cdate"), TextBox).Enabled = True
            CType(CncAddT.FindControl("ddl_vender"), DropDownList).Enabled = True
            CType(CncAddT.FindControl("btn_add_" & modifynum), Button).Text = "新增"
            CType(CncAddT.FindControl("btn_add_" & modifynum), Button).ID = "btn_add_0"
        End If
    End Sub

    Protected Sub BtnTestx_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        SqlCmd = "Select T0.sappo ,T0.rawsta from dbo.[ocnc] T0"
        drsap = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            Do While (drsap.Read())
                SqlCmd = "update dbo.[cnc1] set rawstatus= " & drsap(1) & " where sappo=" & drsap(0)
                CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
            Loop
        End If
        drsap.Close()
        connsap.Close()
    End Sub

    Protected Sub cChk_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim cb As CheckBox = sender
        Dim index, page As Integer
        Dim num As Long
        If (cb.Checked) Then
            index = CInt(Split(cb.ID, "_")(1))
            num = CLng(Split(cb.ID, "_")(2))
            page = CInt(Split(cb.ID, "_")(3))
            CType(gv1.Rows(index).FindControl("hyper_action_" & index), HyperLink).Text = "刪除"
            gv1.Rows(index).Cells(11).BackColor = Drawing.Color.Red
            CType(gv1.Rows(index).FindControl("hyper_action_" & index), HyperLink).NavigateUrl = "cncmain.aspx?act=del&num=" &
                num & "&indexpage=" & page & "&sappo=" & gv1.Rows(index).Cells(1).Text
        End If
    End Sub

    Sub CreateCncWItem(atype As String)
        Dim po As Long
        Dim createdate, vender As String
        Dim ship_date As String
        Dim comm As String
        Dim indexpage As Integer

        indexpage = Request.QueryString("indexpage")
        createdate = Request.QueryString("cdate")
        ship_date = Request.QueryString("ship_date")
        vender = Request.QueryString("vender")
        comm = Request.QueryString("comm")
        po = Request.QueryString("sappo")
        CType(CncAddT.FindControl("txt_po"), TextBox).Text = po
        CType(CncAddT.FindControl("txt_cdate"), TextBox).Text = createdate
        CType(CncAddT.FindControl("txt_shipdate"), TextBox).Text = ship_date
        CType(CncAddT.FindControl("txt_comm"), TextBox).Text = comm
        CType(CncAddT.FindControl("ddl_vender"), DropDownList).SelectedValue = vender

        If (atype = "modifywitem") Then
            CType(CncAddT.FindControl("txt_po"), TextBox).Enabled = False
            CType(CncAddT.FindControl("txt_cdate"), TextBox).Enabled = False
            CType(CncAddT.FindControl("ddl_vender"), DropDownList).Enabled = False
            CType(CncAddT.FindControl("btn_add_" & modifynum), Button).Text = "修改"
        Else
            CType(CncAddT.FindControl("txt_po"), TextBox).Enabled = True
            CType(CncAddT.FindControl("txt_cdate"), TextBox).Enabled = True
            CType(CncAddT.FindControl("ddl_vender"), DropDownList).Enabled = True
            CType(CncAddT.FindControl("btn_add_" & modifynum), Button).Text = "新增"
        End If
        '改為以上copy 至新增列 , 由新增建立
        'Dim po As Long
        'Dim createdate, vender As String
        'Dim ship_date As String
        'Dim comm As String
        'Dim pseq, indexpage As Integer
        'indexpage = Request.QueryString("indexpage")
        'createdate = Request.QueryString("cdate")
        'ship_date = Request.QueryString("ship_date")
        'vender = Request.QueryString("vender")
        'comm = Request.QueryString("comm")
        'po = Request.QueryString("sappo")

        'ReSeqWorkingCnc()

        'SqlCmd = "Select IsNull(max(pseq),0) from dbo.[ocnc] where psta<>90"
        'dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        'dr.Read()
        'pseq = dr(0) + 1
        'dr.Close()
        'conn.Close()

        'SqlCmd = "Insert into dbo.[ocnc] (sappo,pseq,cdate,ship_date,vender,comm) " &
        '                        "values(" & po & "," & pseq & ",'" & createdate & "','" &
        '                        ship_date & "','" & vender & "','" & comm & "')"
        'CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
        'conn.Close()
        ''Insert cnc1 加工細項 from po
        'SqlCmd = "select T1.itemcode,T1.dscription,T1.quantity,T1.whscode " &
        '"from dbo.OPOR T0 inner join POR1 T1 on T0.docentry=T1.docentry where T0.docnum=" & po
        'dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        'If (dr.HasRows) Then
        '    Do While (dr.Read())
        '        SqlCmd = "Insert into dbo.[cnc1] (sappo,itemcode,quantity) " &
        '                            "values(" & po & ",'" & dr(0) & "'," & dr(2) & ")"
        '        CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
        '        conn.Close()
        '    Loop
        'End If
        'dr.Close()
        'connsap.Close()

    End Sub

    Sub tTxt_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim pseq, pseqorg As Integer
        Dim targetnum As Long
        Dim str() As String

        str = Split(sender.ID, "_")
        targetnum = str(2)
        pseq = CInt(sender.Text)

        If (pseq = 0) Then
            CommUtil.ShowMsg(Me, "不能為0")
            Exit Sub
        ElseIf (pseq > 500) Then
            CommUtil.ShowMsg(Me, "不能大於500")
            Exit Sub
        End If
        SqlCmd = "Select pseq from dbo.[ocnc] where num=" & targetnum
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        pseqorg = dr(0)
        dr.Close()
        conn.Close()

        SqlCmd = "update dbo.[ocnc] set pseq= " & pseq & " where num=" & targetnum
        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        conn.Close()
        If (pseq > pseqorg) Then
            SqlCmd = "update dbo.[ocnc] set pseq= pseq-1 " &
            "where psta<>90 and pseq<=" & pseq & " and pseq>" & pseqorg & " and num<>" & targetnum
            CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
            conn.Close()
        Else
            SqlCmd = "update dbo.[ocnc] set pseq= pseq+1 " &
            "where psta<>90 and pseq>=" & pseq & " and pseq<" & pseqorg & " and num<>" & targetnum
            CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
            conn.Close()
        End If
        Response.Redirect("cncmain.aspx?act=showlist&smid=molist&smode=7")

    End Sub

    Sub ReSeqWorkingCnc()
        Dim pseq As Integer
        pseq = 1
        SqlCmd = "Select T0.num " &
         "from dbo.[ocnc] T0 " &
         "where T0.psta<>90 order by T0.pseq"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            Do While (dr.Read())
                SqlCmd = "update dbo.[ocnc] set pseq=" & pseq &
                "where num=" & dr(0)
                CommUtil.SqlLocalExecute("upd", SqlCmd, connsap)
                connsap.Close()
                pseq = pseq + 1
            Loop
            dr.Close()
        End If
        conn.Close()
    End Sub
End Class