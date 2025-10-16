Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Imports System.IO
Public Class signofftodo
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public CommSignOff As New CommSignOff
    Public connsap, conn, connsap1 As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap, dr1 As SqlDataReader
    Public ds As New DataSet
    Public ScriptManager1 As New ScriptManager
    Public formtypeindex, inchargeindex, traceindex, sfid As Integer
    Public act, actstr, sid_create, url, uid, inchargeid As String
    Public DDLAttaFile, DDLAnsFile As DropDownList
    Public docnum, xstdtnum As Long
    Public BtnSave As Button
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        gv1.PageIndex = Request.QueryString("indexpage")
        url = Application("http")
        actstr = Request.QueryString("actstr")
        Page.Form.Controls.Add(ScriptManager1)
        act = Request.QueryString("act")
        uid = Request.QueryString("uid")
        inchargeid = Request.QueryString("inchargeid")
        If (act = "showcontent") Then
            docnum = Request.QueryString("docnum")
            xstdtnum = Request.QueryString("num")
            sid_create = ""
            SqlCmd = "Select sfid,sid from [dbo].[@XASCH] " &
                        "where docnum=" & docnum
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                sfid = dr(0)
                sid_create = dr(1)
            End If
            dr.Close()
            connsap.Close()
        End If
        If (Not IsPostBack) Then
            formtypeindex = Request.QueryString("formtypeindex")
            inchargeindex = Request.QueryString("inchargeindex")
            traceindex = Request.QueryString("traceindex")
            SqlCmd = "Select T0.sfname,T0.sfid from dbo.[@XSFTT] T0 where todoflag=1 order by T0.sfid"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                DDLFormType.Items.Clear()
                DDLFormType.Items.Add("所有追蹤單據 0")
                Do While (dr.Read())
                    DDLFormType.Items.Add(dr(0) & " " & dr(1))
                Loop
            End If
            dr.Close()
            connsap.Close()
            'Dim inchargeid As String = Request.QueryString("uid")
            Dim ti As Integer
            Dim inchargeinclude, inchargematch As Boolean
            inchargeinclude = False
            inchargematch = False
            DDLInCharge.Items.Clear()
            DDLInCharge.Items.Add("所有負責人")
            SqlCmd = "SELECT distinct incharge FROM dbo.[@XSTDT] where incharge <> ''"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            ti = 1
            If (drsap.HasRows) Then
                Do While (drsap.Read())
                    If (Session("s_id") = drsap(0)) Then
                        inchargeinclude = True
                    End If
                    SqlCmd = "Select T0.name From dbo.[User] T0 where id = '" & drsap(0) & "'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    DDLInCharge.Items.Add(drsap(0) & " " & dr(0) & " 負責人")
                    If (drsap(0) = inchargeid And inchargeindex = 9999) Then
                        inchargeindex = ti
                        inchargematch = True
                    End If
                    dr.Close()
                    conn.Close()
                    ti = ti + 1
                Loop
            End If
            drsap.Close()
            connsap.Close()
            If (inchargeinclude = False) Then
                DDLInCharge.Items.Add(Session("s_id") & " " & Session("s_name") & " 負責人")
                If (inchargematch = False And inchargeindex = 9999) Then
                    inchargeindex = ti
                End If
            End If


            DDLTrace.Items.Clear()
            DDLTrace.Items.Add("所有追蹤人")
            SqlCmd = "SELECT distinct traceperson FROM dbo.[@XSTDT] where traceperson <> ''"
            drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (drsap.HasRows) Then
                'ti = 1
                Do While (drsap.Read())
                    SqlCmd = "Select T0.name From dbo.[User] T0 where id = '" & drsap(0) & "'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    DDLTrace.Items.Add(drsap(0) & " " & dr(0) & " 追蹤人")
                    dr.Close()
                    conn.Close()
                    'ti = ti + 1
                Loop
            End If
            drsap.Close()
            connsap.Close()

            ViewState("formtypeindex") = formtypeindex
            ViewState("inchargeindex") = inchargeindex
            ViewState("traceindex") = traceindex
            DDLFormType.SelectedIndex = formtypeindex
            DDLInCharge.SelectedIndex = inchargeindex
            DDLTrace.SelectedIndex = traceindex
            If (Request.QueryString("actstr") = "informsetincharge") Then
                CommUtil.ShowMsg(Me, "請先設定此單之追蹤單據負責人後再回簽核流程")
            End If
        Else
            formtypeindex = ViewState("formtypeindex")
            inchargeindex = ViewState("inchargeindex")
            traceindex = ViewState("traceindex")
        End If
        FTCreate()
        FT.Visible = False
        iframeContent.Visible = False
        If (act = "updsta") Then
            'UpdateStatus() 改由 textbox with listbox處理
        ElseIf (act = "showcontent") Then
            FT.Visible = True
            iframeContent.Visible = True
            DDLFormType.Visible = False
            DDLInCharge.Visible = False
            DDLTrace.Visible = False
            gv1.Visible = False
            ContentTCreate()
            ShowSighOffOrAnsPdf("ans")
        End If
        FormListDisplay(actstr)
        If (actstr = "informsetincharge" Or actstr = "todoitem" Or actstr = "informtraceperson") Then
            DDLFormType.Visible = False
            DDLInCharge.Visible = False
            DDLTrace.Visible = False
        End If
    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        Dim Hyper As HyperLink
        Dim tTxt As TextBox
        Dim ce As CalendarExtender
        Dim dDDL As DropDownList
        Dim LBx As ListBox
        Dim Labelx As Label
        Dim dde As New DropDownExtender
        Dim rtnfinish As Boolean
        Dim docstatus As String
        Dim sfid As Integer
        sfid = 0
        docstatus = ""
        rtnfinish = False
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim mtype As Integer 'mtype=1 表示是需要歸還之料件
            mtype = 0
            SqlCmd = "Select mtype FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & e.Row.Cells(1).Text & " and head=1"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                mtype = dr(0)
            End If
            dr.Close()
            connsap.Close()

            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            Hyper = New HyperLink()
            Hyper.Text = e.Row.Cells(0).Text
            Hyper.Font.Underline = False
            Hyper.ID = "hyper_docnum_" & ds.Tables(0).Rows(realindex)("num")
            If (mtype = 0) Then
                Hyper.NavigateUrl = "signofftodo.aspx?smid=sg&smode=6&act=showcontent&formtypeindex=" & formtypeindex &
                "&docnum=" & ds.Tables(0).Rows(realindex)("docentry") & "&inchargeindex=" & inchargeindex & "&actstr=" & actstr &
                "&uid=" & Request.QueryString("uid") & "&inchargeid=" & inchargeid & "&num=" & ds.Tables(0).Rows(realindex)("num") &
                "&traceindex=" & traceindex & "&indexpage=" & gv1.PageIndex
            Else
                SqlCmd = "Select status,sfid from [dbo].[@XASCH] where docnum=" & ds.Tables(0).Rows(realindex)("docentry")
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    docstatus = dr(0)
                    sfid = dr(1)
                End If
                dr.Close()
                connsap.Close()
                Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&status=" & docstatus & "&sfid=" & sfid &
                    "&formtypeindex=" & DDLFormType.SelectedIndex & "&inchargeindex=" & DDLInCharge.SelectedIndex &
                    "&traceindex=" & DDLTrace.SelectedIndex & "&maindocnum=" & ds.Tables(0).Rows(realindex)("docentry") &
                    "&docnum=" & ds.Tables(0).Rows(realindex)("docentry") & "&indexpage=" & gv1.PageIndex & "&inchargeid=" & inchargeid & "&fromasp=signofftodo"
            End If
            e.Row.Cells(0).Controls.Add(Hyper)

            SqlCmd = "Select T0.sfname from dbo.[@XSFTT] T0 where sfid=" & e.Row.Cells(2).Text
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                e.Row.Cells(2).Text = dr(0)
            End If
            dr.Close()
            connsap.Close()

            If (ds.Tables(0).Rows(realindex)("upddate") <> "1900/01/01" And ds.Tables(0).Rows(realindex)("upddate") <> "1900/1/1") Then
                e.Row.Cells(4).ToolTip = "最近更新日期-" & ds.Tables(0).Rows(realindex)("upddate")
            End If

            LBx = New ListBox
            LBx.ID = "lbsta_" & ds.Tables(0).Rows(realindex)("num") & "_" & e.Row.RowIndex
            LBx.AutoPostBack = True
            LBx.Rows = 21
            AddHandler LBx.SelectedIndexChanged, AddressOf LB_SelectedIndexChanged
            e.Row.Cells(4).Controls.Add(LBx)
            tTxt = New TextBox
            tTxt.ID = "txtsta_" & ds.Tables(0).Rows(realindex)("num")
            tTxt.Width = 40
            tTxt.Text = e.Row.Cells(4).Text
            e.Row.Cells(4).Controls.Add(tTxt)
            dde.TargetControlID = tTxt.ID
            dde.ID = "ddesta_" & ds.Tables(0).Rows(realindex)("num")
            dde.DropDownControlID = LBx.ID
            e.Row.Cells(4).Controls.Add(dde)
            If (ds.Tables(0).Rows(realindex)("incharge") = Session("s_id")) Then
                LBx.Enabled = True
            Else
                LBx.Enabled = False
            End If
            Labelx = New Label
            Labelx.ID = "labelsta_" & ds.Tables(0).Rows(realindex)("num")
            Labelx.Text = "%"
            e.Row.Cells(4).Controls.Add(Labelx)
            LBx.Items.Clear()
            For i = 0 To 100 Step 5
                LBx.Items.Add(i)
            Next
            If (tTxt.Text = 0) Then
                tTxt.BackColor = Drawing.Color.LightGray
            ElseIf (tTxt.Text <= 50) Then
                tTxt.BackColor = Drawing.Color.Yellow
            ElseIf (tTxt.Text <= 95) Then
                tTxt.BackColor = Drawing.Color.LightBlue
            ElseIf (tTxt.Text = 100) Then
                tTxt.BackColor = Drawing.Color.LightGreen
            End If

            tTxt = New TextBox
            tTxt.ID = "txttdate_" & ds.Tables(0).Rows(realindex)("num")
            tTxt.Width = 100
            If (e.Row.Cells(6).Text <> "1900/01/01") Then
                tTxt.Text = e.Row.Cells(6).Text
            Else
                tTxt.Text = ""
            End If
            tTxt.AutoPostBack = True
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            ce = New CalendarExtender
            ce.TargetControlID = tTxt.ID
            ce.ID = "cetdate_" & ds.Tables(0).Rows(realindex)("num")
            ce.Format = "yyyy/MM/dd"
            e.Row.Cells(6).Controls.Add(ce)
            e.Row.Cells(6).Controls.Add(tTxt)
            If (ds.Tables(0).Rows(realindex)("incharge") = Session("s_id")) Then
                tTxt.Enabled = True
            Else
                tTxt.Enabled = False
            End If
            Hyper = New HyperLink()
            If (mtype = 0) Then
                If (ds.Tables(0).Rows(realindex)("incharge") = Session("s_id")) Then
                    Hyper.Text = "請編輯" '-" & e.Row.Cells(7).Text.Substring(0, 5) + "..."
                Else
                    Hyper.Text = "請檢視"
                End If
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_num_" & ds.Tables(0).Rows(realindex)("num")
                Hyper.NavigateUrl = "signofftodo.aspx?smid=sg&smode=6&act=showcontent&formtypeindex=" & formtypeindex &
                "&docnum=" & ds.Tables(0).Rows(realindex)("docentry") & "&inchargeindex=" & inchargeindex & "&actstr=" & actstr &
                "&uid=" & Request.QueryString("uid") & "&inchargeid=" & inchargeid & "&num=" & ds.Tables(0).Rows(realindex)("num") &
                "&traceindex=" & traceindex & "&indexpage=" & gv1.PageIndex
            ElseIf (mtype = 1 Or mtype = 2) Then
                'Check 此單是否還有料件要返還
                SqlCmd = "Select sum(quantity),sum(rtnqty) FROM [dbo].[@XSMLS] T0 WHERE T0.[docentry] =" & ds.Tables(0).Rows(realindex)("docentry") & " And head=0"
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
                dr.Read()
                If (dr(0) = dr(1)) Then
                    rtnfinish = True
                End If
                dr.Close()
                conn.Close()
                If (Not rtnfinish) Then
                    If (ds.Tables(0).Rows(realindex)("incharge") = Session("s_id")) Then
                        SqlCmd = "Select docnum FROM [dbo].[@XASCH] T0 WHERE T0.[attadoc] ='" & e.Row.Cells(1).Text & "' and sfid=101 and status<>'F' and status<>'T' and status<>'C'"
                        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, conn)
                        If (dr.HasRows) Then
                            dr.Read()
                            'e.Row.Cells(7).Text = "還回簽核中(" & dr(0) & ")"
                            Hyper.Text = "返還簽核中(" & dr(0) & ")"
                            Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&status=" & docstatus & "&sfid=" & sfid &
                                            "&formtypeindex=" & DDLFormType.SelectedIndex & "&inchargeindex=" & DDLInCharge.SelectedIndex &
                                            "&traceindex=" & DDLTrace.SelectedIndex & "&maindocnum=" & ds.Tables(0).Rows(realindex)("docentry") &
                                            "&docnum=" & ds.Tables(0).Rows(realindex)("docentry") & "&indexpage=" & gv1.PageIndex & "&inchargeid=" & inchargeid &
                                            "&fromasp=signofftodo"
                        Else
                            Hyper.Text = "新增料件返還" '-" & e.Row.Cells(7).Text.Substring(0, 5) + "..."
                            Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&act=add101&status=A&sfid=101" &
                        "&formtypeindex=" & DDLFormType.SelectedIndex & "&inchargeindex=" & DDLInCharge.SelectedIndex &
                        "&traceindex=" & DDLTrace.SelectedIndex & "&maindocnum=" & ds.Tables(0).Rows(realindex)("docentry") & "&indexpage=" & gv1.PageIndex &
                        "&inchargeid=" & inchargeid & "&fromasp=signofftodo"
                        End If
                        dr.Close()
                        conn.Close()
                    Else
                        Hyper.Text = "請檢視"
                        Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&status=" & docstatus & "&sfid=" & sfid &
                                            "&formtypeindex=" & DDLFormType.SelectedIndex & "&inchargeindex=" & DDLInCharge.SelectedIndex &
                                            "&traceindex=" & DDLTrace.SelectedIndex & "&maindocnum=" & ds.Tables(0).Rows(realindex)("docentry") &
                                            "&docnum=" & ds.Tables(0).Rows(realindex)("docentry") & "&indexpage=" & gv1.PageIndex &
                                            "&inchargeid=" & inchargeid & "&fromasp=signofftodo"
                    End If
                Else
                    Hyper.Text = "已全部返還"
                    Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&actmode=single&status=" & docstatus & "&sfid=" & sfid &
                                            "&formtypeindex=" & DDLFormType.SelectedIndex & "&inchargeindex=" & DDLInCharge.SelectedIndex &
                                            "&traceindex=" & DDLTrace.SelectedIndex & "&maindocnum=" & ds.Tables(0).Rows(realindex)("docentry") &
                                            "&docnum=" & ds.Tables(0).Rows(realindex)("docentry") & "&indexpage=" & gv1.PageIndex &
                                            "&inchargeid=" & inchargeid & "&fromasp=signofftodo"
                End If
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_num_" & ds.Tables(0).Rows(realindex)("num")
            End If
            e.Row.Cells(7).Controls.Add(Hyper)

            If (ds.Tables(0).Rows(realindex)("status") <> 100) Then
                If (e.Row.Cells(8).Text = 0) Then
                    e.Row.Cells(8).Text = "本周尚未更新"
                    e.Row.Cells(8).BackColor = Drawing.Color.IndianRed
                Else
                    e.Row.Cells(8).Text = "本周已更新 " & e.Row.Cells(8).Text & " 次"
                    e.Row.Cells(8).BackColor = Drawing.Color.LightGreen
                End If
            Else
                e.Row.Cells(8).Text = "已結案"
            End If
            Dim inchargename As String
            inchargename = ""
            dDDL = New DropDownList
            dDDL.ID = "ddlincharge_" & ds.Tables(0).Rows(realindex)("num")
            dDDL.Width = 100
            SqlCmd = "Select T0.id,T0.name From dbo.[User] T0 where position <> 'NA' order by branch,grp"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dDDL.Items.Clear()
                dDDL.Items.Add("請選負責人")
                Do While (dr.Read())
                    dDDL.Items.Add(dr(0) & " " & dr(1))
                    If (dr(0) = e.Row.Cells(9).Text) Then
                        inchargename = dr(1)
                    End If
                Loop
            End If
            dr.Close()
            conn.Close()
            If (e.Row.Cells(9).Text = "") Then
                dDDL.SelectedIndex = 0
            Else
                dDDL.SelectedValue = e.Row.Cells(9).Text & " " & inchargename
            End If
            dDDL.AutoPostBack = True
            AddHandler dDDL.SelectedIndexChanged, AddressOf dDDL_SelectedIndexChanged
            e.Row.Cells(9).Controls.Add(dDDL)
            If (ds.Tables(0).Rows(realindex)("traceperson") = Session("s_id")) Then
                dDDL.Enabled = True
            Else
                dDDL.Enabled = False
            End If

            SqlCmd = "Select T0.name From dbo.[User] T0 where id='" & ds.Tables(0).Rows(realindex)("traceperson") & "'"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                e.Row.Cells(10).Text = dr(0)
            End If
            dr.Close()
            conn.Close()

            If (Request.QueryString("actstr") = "informsetincharge") Then
                Hyper = New HyperLink()
                Hyper.Text = "回簽核"
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_back_" & ds.Tables(0).Rows(realindex)("num")
                Hyper.NavigateUrl = "~/signoff/cLsignoff.aspx?smid=sg&smode=2&status=F" &
                                "&docnum=" & ds.Tables(0).Rows(realindex)("docentry") &
                                "&actmode=" & Request.QueryString("actmode") &
                                "&sfid=" & ds.Tables(0).Rows(realindex)("sfid")
                e.Row.Cells(11).Controls.Add(Hyper)
            End If
        End If
    End Sub
    Sub dDDL_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim DDLx As DropDownList = sender
        Dim num As Long
        Dim kw As String
        Dim str() As String
        str = Split(DDLx.ID, "_")
        num = str(1)
        kw = str(0)
        'MsgBox(txtkw)
        If (kw = "ddlincharge") Then
            If (DDLx.SelectedIndex = 0) Then
                SqlCmd = "update dbo.[@XSTDT] set incharge= '' where num=" & num
            Else
                SqlCmd = "update dbo.[@XSTDT] set incharge= '" & Split(DDLx.SelectedValue, " ")(0) & "' where num=" & num
            End If
            CommUtil.SqlSapExecute("upd", SqlCmd, conn)
            conn.Close()
        End If
    End Sub
    Sub tTxt_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim Txtx As TextBox = sender
        Dim inum As Long
        Dim txtkw As String
        Dim str() As String
        str = Split(Txtx.ID, "_")
        inum = str(1)
        txtkw = str(0)
        'MsgBox(txtkw)
        If (txtkw = "txttdate") Then
            SqlCmd = "update dbo.[@XSTDT] set tdate= '" & Txtx.Text & "' where num=" & inum
            CommUtil.SqlSapExecute("upd", SqlCmd, conn)
            conn.Close()
            'ElseIf (txtkw = "txtnote") Then
            '    SqlCmd = "update dbo.[@XSTDT] set note='" & Txtx.Text & "' where num=" & inum
            '    CommUtil.SqlSapExecute("upd", SqlCmd, conn)
            '    conn.Close()
        End If
    End Sub
    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        '以下或可用 Response.Redirect 帶參數方式執行
        'If (act = "mysearch") Then
        '    TxtDocnum.Text = Request.QueryString("txtdocnum")
        '    TxtKW.Text = Request.QueryString("txtkw")
        '    SearchDoc()
        'Else
        '    FormListDisplay(1)
        'End If
        FormListDisplay(Request.QueryString("actstr"))
    End Sub
    Sub FormListDisplay(actstr As String)
        Dim sfid As Integer
        Dim str() As String
        Dim inchargeid, traceid As String
        'Dim match As Boolean
        sfid = 0
        ds.Reset()
        SetGridViewStyle()
        SetFormListGridViewFields()
        str = Split(DDLFormType.SelectedValue, " ")
        sfid = str(1)
        If (DDLTrace.SelectedIndex <> 0) Then
            str = Split(DDLTrace.SelectedValue, " ")
            traceid = str(0)
        Else
            traceid = ""
        End If
        Dim sfid_rule As String
        inchargeid = Request.QueryString("inchargeid")

        If (actstr = "informsetincharge") Then
            sfid_rule = " and docentry=" & Request.QueryString("docentry")
        ElseIf (actstr = "todoitem") Then
            sfid_rule = " and incharge='" & inchargeid & "'"
        ElseIf (actstr = "informtraceperson") Then
            sfid_rule = " and incharge='" & inchargeid & "' and num=" & Request.QueryString("num")
        Else 'include actstr = "ddltypefilter" , actstr = "inchargefilter" actstr="tracefilter", acrstr="" 同,所以應不需此 else
            If (inchargeid <> "") Then
                If (sfid = 0) Then
                    sfid_rule = " and incharge='" & inchargeid & "'"
                Else
                    sfid_rule = " and sfid=" & sfid & " and incharge='" & inchargeid & "'"
                End If
            Else
                If (sfid = 0) Then
                    sfid_rule = ""
                Else
                    sfid_rule = " and sfid=" & sfid
                End If
            End If
            If (traceid <> "") Then
                sfid_rule = sfid_rule & " and traceperson='" & traceid & "'"
            End If
        End If
        SqlCmd = "SELECT docentry,sfid,convert(varchar(12), cdate, 111) as cdate,convert(varchar(12), tdate, 111) as tdate,status,subject,processrpt," &
                 "incharge,traceperson,convert(varchar(12), upddate,111) as upddate,num,updcount " &
                 "FROM dbo.[@XSTDT] T0 " &
                 "where status <> 100 " & sfid_rule & " order by sfid,docentry desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
        connsap1.Close()
        'MsgBox(SqlCmd)
        If (actstr <> "todoitem") Then
            SqlCmd = "SELECT docentry,sfid,convert(varchar(12), cdate, 111) as cdate,convert(varchar(12), tdate, 111) as tdate,status,subject,processrpt," &
                     "incharge,traceperson,convert(varchar(12), upddate,111) as upddate,num,updcount " &
                 "FROM dbo.[@XSTDT] T0 " &
                 "where status = 100 " & sfid_rule & " order by sfid,docentry desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        End If
        If (ds.Tables(0).Rows.Count = 0 And actstr = "ddltypefilter") Then
            CommUtil.ShowMsg(Me, "無任何資料")
        End If
        ds.Tables(0).Columns.Add("action")
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()

    End Sub

    Protected Sub DDLFormType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDLFormType.SelectedIndexChanged
        'Dim inchargeid As String
        'If (DDLInCharge.SelectedIndex = 0) Then
        '    inchargeid = ""
        'Else
        '    inchargeid = Split(DDLInCharge.SelectedValue, " ")(0)
        'End If
        Response.Redirect("~/signoff/signofftodo.aspx?smid=sg&smode=6&actstr=ddltypefilter&formtypeindex=" & DDLFormType.SelectedIndex &
                          "&inchargeindex=" & DDLInCharge.SelectedIndex & "&traceindex=" & DDLTrace.SelectedIndex & "&inchargeid=" & inchargeid &
                          "&uid=" & uid)
    End Sub
    Protected Sub DDLInCharge_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDLInCharge.SelectedIndexChanged
        Dim inchargeid As String
        If (DDLInCharge.SelectedIndex = 0) Then
            inchargeid = ""
        Else
            inchargeid = Split(DDLInCharge.SelectedValue, " ")(0)
        End If
        Response.Redirect("~/signoff/signofftodo.aspx?smid=sg&smode=6&actstr=inchargefilter&formtypeindex=" & DDLFormType.SelectedIndex &
                          "&inchargeindex=" & DDLInCharge.SelectedIndex & "&traceindex=" & DDLTrace.SelectedIndex & "&inchargeid=" & inchargeid &
                          "&uid=" & uid)
    End Sub

    Protected Sub DDLTrace_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDLTrace.SelectedIndexChanged
        'Dim inchargeid As String
        'If (DDLTrace.SelectedIndex = 0) Then
        '    inchargeid = ""
        'Else
        '    inchargeid = Split(DDLInCharge.SelectedValue, " ")(0)
        'End If
        Response.Redirect("~/signoff/signofftodo.aspx?smid=sg&smode=6&actstr=tracefilter&formtypeindex=" & DDLFormType.SelectedIndex &
                          "&inchargeindex=" & DDLInCharge.SelectedIndex & "&traceindex=" & DDLTrace.SelectedIndex & "&inchargeid=" & inchargeid &
                          "&uid=" & uid)
    End Sub
    Sub SetGridViewStyle()
        gv1.AutoGenerateColumns = False
        'gv1.ShowHeader = True
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

    Sub SetFormListGridViewFields()
        Dim oBoundField As BoundField
        gv1.Columns.Clear()
        oBoundField = New BoundField
        oBoundField.HeaderText = "序號"
        oBoundField.DataField = "num"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "對應簽核"
        oBoundField.DataField = "docentry"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "表單種類"
        oBoundField.DataField = "sfid"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Left
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "主旨"
        oBoundField.DataField = "subject"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.Width = 300
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "進度"
        oBoundField.DataField = "status"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "建立日期"
        oBoundField.DataField = "cdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "預定完成日期"
        oBoundField.DataField = "tdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "處理過程及附檔"
        oBoundField.DataField = "processrpt"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "更新狀態"
        oBoundField.DataField = "updcount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "負責人"
        oBoundField.DataField = "incharge"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "追蹤人"
        oBoundField.DataField = "traceperson"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        If (Request.QueryString("actstr") = "informsetincharge") Then
            oBoundField = New BoundField
            oBoundField.HeaderText = "動作"
            oBoundField.DataField = "action"
            oBoundField.ShowHeader = True
            oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
            gv1.Columns.Add(oBoundField)
        End If
    End Sub
    Sub UpdateStatus()
        Dim num As Long
        Dim status As Integer
        Dim nowday As String
        nowday = Format(Now(), "yyyy/MM/dd")
        num = Request.QueryString("num")
        status = Request.QueryString("status")
        If (status = 0) Then
            status = 10
        ElseIf (status = 10) Then
            status = 30
        ElseIf (status = 30) Then
            status = 50
        ElseIf (status = 50) Then
            status = 70
        ElseIf (status = 70) Then
            status = 90
        ElseIf (status = 90) Then
            status = 100
        ElseIf (status = 100) Then
            status = 0
        End If
        SqlCmd = "update dbo.[@XSTDT] set status= " & status & ",upddate='" & nowday & "' where num=" & num
        CommUtil.SqlSapExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Hyper As HyperLink
        Dim Labelx As Label
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        Hyper = New HyperLink
        Hyper.ID = "hyper_back"
        Hyper.Text = "回前頁"
        Hyper.NavigateUrl = "~/signoff/signofftodo.aspx?smid=sg&smode=6&formtypeindex=" & formtypeindex & "&inchargeindex=" & inchargeindex &
                            "&actstr=" & Request.QueryString("actstr") & "&uid=" & Request.QueryString("uid") & "&inchargeid=" & inchargeid &
                            "&num=" & Request.QueryString("num") & "&traceindex=" & traceindex & "&indexpage=" & Request.QueryString("indexpage")
        Hyper.Font.Underline = False
        tCell.Controls.Add(Hyper)

        Labelx = New Label()
        Labelx.ID = "label_fileul"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Dim FileUL As New FileUpload()
        FileUL.ID = "fileul"
        tCell.Controls.Add(FileUL)

        Dim ChkDel As New CheckBox
        ChkDel.ID = "chk_del"
        ChkDel.Text = "刪檔"
        ChkDel.AutoPostBack = True
        AddHandler ChkDel.CheckedChanged, AddressOf ChkDel_CheckedChanged
        tCell.Controls.Add(ChkDel)

        Labelx = New Label()
        Labelx.ID = "label_upfile"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Dim BtnFileAct As New Button
        BtnFileAct.ID = "btn_fileact"
        BtnFileAct.Text = "上傳"
        AddHandler BtnFileAct.Click, AddressOf BtnFileAct_Click
        tCell.Controls.Add(BtnFileAct)
        Labelx = New Label()
        Labelx.ID = "label_ddlansfile"
        Labelx.Text = "&nbsp&nbsp&nbsp問題回應文件&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        DDLAnsFile = New DropDownList
        DDLAnsFile.ID = "ddl_ansfile"
        DDLAnsFile.Width = 240
        AddHandler DDLAnsFile.SelectedIndexChanged, AddressOf DDLAnsFile_SelectedIndexChanged
        DDLAnsFile.AutoPostBack = True
        tCell.Controls.Add(DDLAnsFile)
        ShowAttachOrAndFileList("ans")
        'tRow.Cells.Add(tCell)
        If (inchargeid <> Session("s_id")) Then
            ChkDel.Enabled = False
            BtnFileAct.Enabled = False
            FileUL.Enabled = False
        End If

        Labelx = New Label()
        Labelx.ID = "label_ddlattafile"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp問題反應單文件&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        DDLAttaFile = New DropDownList
        DDLAttaFile.ID = "ddl_attafile"
        DDLAttaFile.Width = 240
        AddHandler DDLAttaFile.SelectedIndexChanged, AddressOf DDLAttaFile_SelectedIndexChanged
        DDLAttaFile.AutoPostBack = True
        tCell.Controls.Add(DDLAttaFile)
        ShowAttachOrAndFileList("atta")

        tRow.Cells.Add(tCell)
        FT.Rows.Add(tRow)
    End Sub
    Sub ShowAttachOrAndFileList(ftype As String)
        Dim di As DirectoryInfo
        Dim attachsel As String
        If (ftype = "ans") Then
            DDLAnsFile.Items.Clear()
            DDLAnsFile.Items.Add("請選擇欲顯示之回應檔案")
            attachsel = ""
            If (System.IO.Directory.Exists(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")) Then
                di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")
                Dim fi As FileInfo() = di.GetFiles(docnum & "_回應*")
                For i = 0 To fi.Length - 1
                    DDLAnsFile.Items.Add(fi(i).Name)
                    If (InStr(fi(i).Name, "(1)") <> 0) Then
                        attachsel = fi(i).Name
                    End If
                Next
                DDLAnsFile.SelectedValue = attachsel
            End If
        ElseIf (ftype = "atta") Then
            DDLAttaFile.Items.Clear()
            DDLAttaFile.Items.Add("請選擇欲顯示之簽核檔案")
            attachsel = ""
            If (System.IO.Directory.Exists(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")) Then
                di = New DirectoryInfo(HttpContext.Current.Server.MapPath("~/") & "AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\")
                Dim fi As FileInfo() = di.GetFiles(docnum & "*")
                For i = 0 To fi.Length - 1
                    If (InStr(fi(i).Name, "簽核") <> 0) Then
                        DDLAttaFile.Items.Add(fi(i).Name)
                    End If
                Next
                DDLAttaFile.SelectedValue = attachsel
            End If
        End If
    End Sub
    Protected Sub DDLAttaFile_SelectedIndexChanged(sender As Object, e As EventArgs)
        DDLAnsFile.SelectedIndex = 0
        If (DDLAttaFile.SelectedIndex <> 0) Then
            ShowSighOffOrAnsPdf("signoff")
        Else
            iframeContent.Attributes.Remove("src")
        End If
    End Sub
    Protected Sub DDLAnsFile_SelectedIndexChanged(sender As Object, e As EventArgs)
        DDLAttaFile.SelectedIndex = 0
        If (DDLAnsFile.SelectedIndex <> 0) Then
            ShowSighOffOrAnsPdf("ans")
        Else
            iframeContent.Attributes.Remove("src")
        End If
    End Sub
    Sub ShowSighOffOrAnsPdf(ftype As String)
        Dim httpfile As String
        Dim attachfile As String
        If (ftype = "signoff") Then
            If (DDLAttaFile.SelectedIndex <> 0) Then
                attachfile = DDLAttaFile.SelectedValue
            Else
                attachfile = ""
            End If
        Else
            If (DDLAnsFile.SelectedIndex <> 0) Then
                attachfile = DDLAnsFile.SelectedValue
            Else
                attachfile = ""
            End If
        End If
        If (attachfile <> "") Then
            httpfile = url & "AttachFile/SignOffsFormFiles/" & sid_create & "/" & sfid & "/" & attachfile
            iframeContent.Attributes.Remove("src")
            iframeContent.Attributes.Add("src", httpfile)
        Else
            iframeContent.Attributes.Remove("src")
        End If
    End Sub
    Protected Sub BtnFileAct_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim targetDir As String
        Dim attachfile As String
        Dim p As New Process()
        Dim FileUL As New FileUpload
        Dim localansdir As String
        Dim fname As String
        localansdir = Application("localdir") & "SignOffsFormFiles\"
        targetDir = HttpContext.Current.Server.MapPath("~/") & "\AttachFile\SignOffsFormFiles\" & sid_create & "\" & sfid & "\"
        'appPath = Request.PhysicalApplicationPath '應用程式目錄
        'If (Not System.IO.Directory.Exists(targetDir)) Then
        '    Directory.CreateDirectory(targetDir)
        'End If
        'If (Not System.IO.Directory.Exists(localansdir)) Then
        '    Directory.CreateDirectory(localansdir)
        'End If
        FileUL = CType(FT.FindControl("fileul"), FileUpload)
        If (sender.Text = "上傳") Then 'sssss
            If (FileUL.HasFile) Then
                fname = docnum & "_回應" & GetNextAttachedFileName() & "_" & FileUL.FileName
                FileUL.SaveAs(targetDir & fname)
                FileUL.SaveAs(localansdir & fname)
                SqlCmd = "update [dbo].[@XASCH] set attachfileno=attachfileno+1 " &
                             " where docnum=" & docnum
                CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
                connsap.Close()
                ShowAttachOrAndFileList("ans")
                DDLAnsFile.SelectedValue = fname
                ShowSighOffOrAnsPdf("ans")
            Else
                CommUtil.ShowMsg(Me, "無上傳檔案")
            End If
        ElseIf (sender.Text = "刪除") Then
            attachfile = targetDir & DDLAnsFile.SelectedValue
            IO.File.Delete(attachfile)
            IO.File.Delete(localansdir & DDLAnsFile.SelectedValue)
            sender.Text = "上傳"
            CType(FT.FindControl("chk_del"), CheckBox).Checked = False
            sender.BackColor = Nothing
            ShowAttachOrAndFileList("ans")
            DDLAnsFile.SelectedIndex = 0
            ShowSighOffOrAnsPdf("ans")
            DDLAttaFile.Enabled = True
        End If
    End Sub
    Protected Sub ChkDel_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim ChkDel As New CheckBox
        Dim BtnFileAct As New Button
        If (DDLAnsFile.SelectedIndex <> 0) Then
            ChkDel = CType(FT.FindControl("chk_del"), CheckBox)
            BtnFileAct = CType(FT.FindControl("btn_fileact"), Button)
            If (ChkDel.Checked) Then
                BtnFileAct.Text = "刪除"
                DDLAttaFile.Enabled = False
            Else
                BtnFileAct.Text = "上傳"
                DDLAttaFile.Enabled = True
            End If
        Else
            DDLAttaFile.Enabled = False
            CommUtil.ShowMsg(Me, "請先選擇欲被刪除之回應檔案")
        End If
    End Sub
    Function GetNextAttachedFileName()
        Dim nextfilename As String
        Dim connsap As New SqlConnection
        Dim dr As SqlDataReader
        nextfilename = ""
        SqlCmd = "Select attachfileno " &
        "from [dbo].[@XASCH] " &
        "where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            nextfilename = "(" & CStr(dr(0) + 1) & ")"
        End If
        dr.Close()
        connsap.Close()
        Return nextfilename
    End Function
    Sub ContentTCreate()
        InitTable()
        'WriteListBoxItemForXSTDT()
        ShowXSTDT()
    End Sub
    Sub InitTable()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor, NeedInputColor, WhiteColor As Drawing.Color
        NeedInputColor = Drawing.Color.AntiqueWhite
        WhiteColor = Drawing.Color.White
        tRow = New TableRow()
        'For j = 1 To 16
        '    tCell = New TableCell
        '    tCell.BorderWidth = 0
        '    tCell.Width = 200
        '    tCell.HorizontalAlign = HorizontalAlign.Center
        '    tRow.Controls.Add(tCell)
        'Next
        'contentT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        contentT.Font.Name = "標楷體"
        contentT.Font.Size = 14
        tRow = New TableRow()
        'CellSet(Text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer, align As String)
        tRow.Controls.Add(CommUtil.CellSet("主旨", 1, 1, False, 0, 0, "center", BColor))
        tRow.Controls.Add(CommUtil.CellSet("", 1, 1, False, 0, 0, "left", WhiteColor))
        contentT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.Controls.Add(CommUtil.CellSet("處理過程", 1, 1, False, 0, 0, "center", BColor))
        tRow.Controls.Add(CommUtil.CellSetWithTextBox(1, 1, "txt_processrpt", 5, 0, 800, NeedInputColor, "left"))
        contentT.Rows.Add(tRow)

        tRow = New TableRow()
        tCell = New TableCell()
        tCell.ColumnSpan = 2
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        BtnSave = New Button
        BtnSave.ID = "btn_save"
        BtnSave.Width = 60
        BtnSave.Text = "儲存"
        AddHandler BtnSave.Click, AddressOf BtnSave_Click
        tCell.Controls.Add(BtnSave)
        tRow.Cells.Add(tCell)
        contentT.Rows.Add(tRow) 'row=7
    End Sub

    Protected Sub LB_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tTxt As TextBox
        Dim str() As String
        Dim id As String
        Dim nowdate As String
        Dim gv1row As Integer
        Dim ds As New DataSet
        Dim connL As New SqlConnection
        Dim urlpara As String
        nowdate = Format(Now(), "yyyy/MM/dd")
        str = Split(sender.ID, "_")
        id = str(1)
        gv1row = str(2)
        tTxt = gv1.Rows(gv1row).FindControl("txtsta_" & id)
        tTxt.Text = sender.SelectedValue
        SqlCmd = "update dbo.[@XSTDT] set updcount= updcount+1,upddate='" & nowdate & "' where num=" & id & " and upddate <>'" & nowdate & "'"
        CommUtil.SqlSapExecute("upd", SqlCmd, conn)
        conn.Close()

        SqlCmd = "update dbo.[@XSTDT] set status= " & CInt(tTxt.Text) & " where num=" & id
        CommUtil.SqlSapExecute("upd", SqlCmd, conn)
        conn.Close()
        If (tTxt.Text = 0) Then
            tTxt.BackColor = Drawing.Color.LightGray
        ElseIf (tTxt.Text <= 50) Then
            tTxt.BackColor = Drawing.Color.Yellow
        ElseIf (tTxt.Text <= 95) Then
            tTxt.BackColor = Drawing.Color.LightBlue
        ElseIf (tTxt.Text = 100) Then
            tTxt.BackColor = Drawing.Color.LightGreen
            SqlCmd = "Select T0.traceperson As mailtoid,T0.subject,T0.sfid,T0.docentry,T0.num,T0.incharge,T1.sfname " &
                     "FROM [dbo].[@XSTDT] T0 Inner Join [dbo].[@XSFTT] T1 On T0.sfid=T1.sfid " &
                    "WHERE T0.[num] =" & id
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connL)
            urlpara = "?actmode=informtraceperson&uid=" & ds.Tables(0).Rows(0)("mailtoid") & "&inchargeid=" & ds.Tables(0).Rows(0)("incharge") &
                                        "&docentry=" & ds.Tables(0).Rows(0)("docentry") & "&sfid=" & ds.Tables(0).Rows(0)("sfid") &
                                        "&num=" & ds.Tables(0).Rows(0)("num")
            CommSignOff.InformByMail(Application("http"), "之問題反應完成通知", "進度完成", ds, urlpara)
            ds.Reset()
            connL.Close()
        End If
    End Sub
    Sub ShowXSTDT()
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        SqlCmd = "Select T0.subject,T0.processrpt " &
                     "FROM [dbo].[@XSTDT] T0 WHERE T0.[num] =" & xstdtnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

        If (drL.HasRows) Then
            drL.Read()
            contentT.Rows(0).Cells(1).Text = drL(0)
            CType(contentT.FindControl("txt_processrpt"), TextBox).Text = drL(1)
            If (inchargeid <> Session("s_id")) Then
                CType(contentT.FindControl("txt_processrpt"), TextBox).Enabled = False
                BtnSave.Enabled = False
            End If
            If ((System.Text.RegularExpressions.Regex.Matches(drL(1), "\r\n").Count + 1) <= 5) Then
                CType(contentT.FindControl("txt_processrpt"), TextBox).Rows = 5
            Else
                CType(contentT.FindControl("txt_processrpt"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(1), "\r\n").Count + 1
            End If
        End If
        drL.Close()
        connL.Close()
    End Sub
    Protected Sub BtnSave_Click(sender As Object, e As EventArgs)
        Dim nowdate As String
        nowdate = Format(Now(), "yyyy/MM/dd")
        SqlCmd = "update [dbo].[@XSTDT] set processrpt='" & CType(contentT.FindControl("txt_processrpt"), TextBox).Text & "' " &
                " where num=" & xstdtnum
        CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
        connsap.Close()
        SqlCmd = "update dbo.[@XSTDT] set updcount= updcount+1,upddate='" & nowdate & "' where num=" & xstdtnum & " and upddate <>'" & nowdate & "'"
        CommUtil.SqlSapExecute("upd", SqlCmd, conn)
        conn.Close()
        ShowXSTDT()
        CommUtil.ShowMsg(Me, "更新進度說明OK")
    End Sub
End Class