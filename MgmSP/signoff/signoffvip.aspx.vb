Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class signoffvip
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn, connsap1 As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap, dr1 As SqlDataReader
    Public ds As New DataSet
    Public ScriptManager1 As New ScriptManager
    Public DDLFormType, DDLFormStatus As DropDownList
    Public BtnFilter As Button
    Public BtnEmail As Button
    Public sid As String
    Public TxtDocnum, TxtKW As TextBox
    Public act As String
    Public permssg500 As String

    Protected Sub DDLFormType_SelectedIndexChanged(sender As Object, e As EventArgs)
        Response.Redirect("~/signoff/signoffvip.aspx?smid=sg&smode=5&formtypeindex=" & DDLFormType.SelectedIndex &
                          "&formstatusindex=" & DDLFormStatus.SelectedIndex)
    End Sub
    Protected Sub DDLFormStatus_SelectedIndexChanged(sender As Object, e As EventArgs)
        Response.Redirect("~/signoff/signoffvip.aspx?smid=sg&smode=5&formtypeindex=" & DDLFormType.SelectedIndex &
                          "&formstatusindex=" & DDLFormStatus.SelectedIndex)
    End Sub
    Protected Sub BtnFilter_Click(sender As Object, e As EventArgs)
        If (TxtDocnum.Text <> "") Then
            If (Not IsNumeric(TxtDocnum.Text)) Then
                CommUtil.ShowMsg(Me, "表單號欄位不是數字,請更正")
                Exit Sub
            End If
        End If
        If (TxtDocnum.Text <> "" Or TxtKW.Text <> "") Then
            Response.Redirect("~/signoff/signoffvip.aspx?smid=sg&smode=5&act=allsearch&txtdocnum=" & TxtDocnum.Text &
                          "&txtkw=" & TxtKW.Text)
        Else
            CommUtil.ShowMsg(Me, "表單號或主旨關鍵字須輸入")
        End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        sid = Session("s_id")
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        permssg500 = CommUtil.GetAssignRight("sg500", Session("s_id"))
        Page.Form.Controls.Add(ScriptManager1)
        FTCreate()
        If (Not IsPostBack) Then
            act = Request.QueryString("act")
            'If (act <> "") Then
            DDLFormType.SelectedIndex = Request.QueryString("formtypeindex")
            DDLFormStatus.SelectedIndex = Request.QueryString("formstatusindex")
            'End If
            If (act = "allsearch") Then
                TxtDocnum.Text = Request.QueryString("txtdocnum")
                TxtKW.Text = Request.QueryString("txtkw")
                SearchDoc()
            Else
                FormListDisplay(1)
            End If
        End If
        '以下為Test
        'Dim fileNameSign, targetPath As String
        'Dim pdfFiles(1) As String
        'Dim fileNameApproved As String = "6_Approved.pdf"
        'targetPath = HttpContext.Current.Server.MapPath("~/") & "SignOffsFormFiles\71\tedy\"
        'fileNameSign = "6_sign.pdf"
        'pdfFiles(0) = "6.pdf"
        'pdfFiles(1) = fileNameSign '會簽流程的pdf
        'CommUtil.mergePDF(pdfFiles, fileNameApproved, targetPath)
    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        Dim Hyper As HyperLink
        Dim sfid As Integer
        Dim seq As Integer
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            If (e.Row.Cells(3).Text = "0") Then
                e.Row.Cells(3).Text = "NA"
                'MsgBox(e.Row.Cells(2).Text)
            End If
            If (e.Row.Cells(4).Text = "") Then
                e.Row.Cells(4).Text = "NA"
            End If
            sfid = e.Row.Cells(1).Text
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            SqlCmd = "Select T0.sfname from dbo.[@XSFTT] T0 where T0.sfid=" & e.Row.Cells(1).Text
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            e.Row.Cells(1).Text = dr(0)
            dr.Close()
            connsap.Close()

            If (e.Row.Cells(6).Text = "E" Or e.Row.Cells(6).Text = "D" Or e.Row.Cells(6).Text = "B" Or e.Row.Cells(6).Text = "R") Then
                '照原sql 不修改
            ElseIf (e.Row.Cells(6).Text = "O" Or e.Row.Cells(6).Text = "F") Then
                SqlCmd = "SELECT uname,seq FROM dbo.[@XSPWT] where status=1 and docentry=" & e.Row.Cells(0).Text
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    e.Row.Cells(7).Text = dr(0)
                    seq = dr(1)
                End If
                dr.Close()
                connsap.Close()
            Else
                e.Row.Cells(7).Text = "NA"
            End If
            Hyper = New HyperLink
            Hyper.ID = "hyper_action"
            If (e.Row.Cells(6).Text = "O") Then
                '如果是最後一關 , 不能執行跳簽
                SqlCmd = "SELECT max(seq) FROM dbo.[@XSPWT] where signprop=0 and docentry=" & e.Row.Cells(0).Text
                dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
                If (dr.HasRows) Then
                    dr.Read()
                    If (dr(0) <> seq) Then
                        Hyper.Text = "執行跳簽"
                        act = "skip"
                    Else
                        Hyper.Text = "顯示"
                        act = "frommanage"
                    End If
                Else
                    Hyper.Text = "查無此單"
                End If
                dr.Close()
                connsap.Close()
            Else
                Hyper.Text = "顯示"
                act = "frommanage"
            End If
            Hyper.NavigateUrl = "cLsignoff.aspx?smid=sg&smode=2&act=" & act & "&status=" & e.Row.Cells(6).Text &
                                "&indexpage=" & gv1.PageIndex & "&docnum=" & e.Row.Cells(0).Text &
                                "&formstatusindex=" & DDLFormStatus.SelectedIndex &
                                "&formtypeindex=" & DDLFormType.SelectedIndex & "&sfid=" & sfid &
                                "&skipid=" & ds.Tables(0).Rows(realindex)("uid")

            Hyper.Font.Underline = False
            CommUtil.DisableObjectByPermission(Hyper, permssg500, "m")
            e.Row.Cells(9).Controls.Add(Hyper)

            If (e.Row.Cells(6).Text = "E") Then
                e.Row.Cells(6).Text = "編輯中"
            ElseIf (e.Row.Cells(6).Text = "D") Then
                e.Row.Cells(6).Text = "底稿"
            ElseIf (e.Row.Cells(6).Text = "O") Then
                e.Row.Cells(6).Text = "簽核中"
            ElseIf (e.Row.Cells(6).Text = "F") Then
                e.Row.Cells(6).Text = "簽核完成"
            ElseIf (e.Row.Cells(6).Text = "C") Then
                e.Row.Cells(6).Text = "作廢"
            ElseIf (e.Row.Cells(6).Text = "R") Then
                e.Row.Cells(6).Text = "抽回"
            ElseIf (e.Row.Cells(6).Text = "B") Then
                e.Row.Cells(6).Text = "駁回"
            ElseIf (e.Row.Cells(6).Text = "T") Then
                e.Row.Cells(6).Text = "已歸檔"
            End If
        End If
    End Sub
    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        '以下或可用 Response.Redirect 帶參數方式執行
        If (act = "allsearch") Then
            TxtDocnum.Text = Request.QueryString("txtdocnum")
            TxtKW.Text = Request.QueryString("txtkw")
            SearchDoc()
        Else
            FormListDisplay(1)
        End If
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        tRow = New TableRow()
        tCell = New TableCell()
        'tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        DDLFormType = New DropDownList
        DDLFormType.ID = "ddl_formtype"
        DDLFormType.Width = 250
        DDLFormType.AutoPostBack = True
        SqlCmd = "Select T0.sfname,T0.sfid,T0.sftypenote from dbo.[@XSFTT] T0 order by T0.sfid"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            DDLFormType.Items.Clear()
            DDLFormType.Items.Add("所有簽單 0")
            Do While (dr.Read())
                DDLFormType.Items.Add(dr(0) & " " & dr(1) & " " & dr(2))
            Loop
        End If
        dr.Close()
        connsap.Close()
        AddHandler DDLFormType.SelectedIndexChanged, AddressOf DDLFormType_SelectedIndexChanged
        tCell.Controls.Add(DDLFormType)

        Labelx = New Label
        Labelx.ID = "label_1"
        Labelx.Text = "&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        DDLFormStatus = New DropDownList
        DDLFormStatus.ID = "ddl_formstatus"
        DDLFormStatus.Width = 150
        DDLFormStatus.Items.Clear()
        DDLFormStatus.Items.Add("全部待簽單")
        DDLFormStatus.Items.Add("全部待送審單")
        DDLFormStatus.Items.Add("全部待歸檔")
        DDLFormStatus.Items.Add("全部已送審單")
        DDLFormStatus.Items.Add("全部已歸檔")
        DDLFormStatus.Items.Add("全部作廢單")
        DDLFormStatus.Items.Add("全部簽核單")
        DDLFormStatus.AutoPostBack = True
        AddHandler DDLFormStatus.SelectedIndexChanged, AddressOf DDLFormStatus_SelectedIndexChanged
        tCell.Controls.Add(DDLFormStatus)

        Labelx = New Label
        Labelx.ID = "label_2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp表單號:"
        tCell.Controls.Add(Labelx)
        TxtDocnum = New TextBox
        TxtDocnum.Width = 60
        tCell.Controls.Add(TxtDocnum)

        Labelx = New Label
        Labelx.ID = "label_3"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp主旨關鍵字:"
        tCell.Controls.Add(Labelx)
        TxtKW = New TextBox
        TxtKW.Width = 120
        tCell.Controls.Add(TxtKW)

        Labelx = New Label
        Labelx.ID = "label_4"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnFilter = New Button
        BtnFilter.ID = "btn_filter"
        BtnFilter.Text = "查詢"
        AddHandler BtnFilter.Click, AddressOf BtnFilter_Click
        tCell.Controls.Add(BtnFilter)

        tRow.Cells.Add(tCell)
        FT.Rows.Add(tRow)
    End Sub
    Sub FormListDisplay(displaymode As Integer)
        Dim sfid As Integer
        Dim str() As String
        sfid = 0
        ds.Reset()
        SetGridViewStyle()
        SetFormListGridViewFields()
        If (DDLFormType.SelectedIndex <> 0) Then
            str = Split(DDLFormType.SelectedValue, " ")
            sfid = str(1)
        End If
        Dim sfid_rule As String
        If (sfid = 0) Then
            sfid_rule = ""
        Else
            sfid_rule = " and T1.sfid=" & sfid
        End If

        If (DDLFormStatus.SelectedIndex = 0) Then ' 待簽核(關卡簽核)
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq,T0.uname,T0.uid " &
                    "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                    "where T0.signprop=0 and T0.status=1 and T1.status<>'B' and T1.status<>'R'" & sfid_rule &
                    " order by T0.signprop,T1.sfid,T0.uid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '關卡簽核
            connsap1.Close()
        ElseIf (DDLFormStatus.SelectedIndex = 1) Then ' 全部待送審單
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
             "FROM dbo.[@XASCH] T1 " &
             "where (T1.status='E' or T1.status='D' or T1.status='B' or T1.status='R')" & sfid_rule &
             " order by T1.sfid,T1.sid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (DDLFormStatus.SelectedIndex = 2) Then ' 全部待歸檔
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq,T0.uname,T0.uid " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=1 And T0.status=1" & sfid_rule &
                 " order by T1.sfid,T0.uid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (DDLFormStatus.SelectedIndex = 3) Then ' 全部已送審單
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.status=100" & sfid_rule &
                " order by T1.sfid,T0.uid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (DDLFormStatus.SelectedIndex = 4) Then ' 全部已歸檔
            SqlCmd = "Select T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,T0.seq,T0.uname,T0.uid " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.signprop=1 And T1.status='T'" & sfid_rule &
                 " order by T1.sfid,T0.uid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (DDLFormStatus.SelectedIndex = 5) Then ' 全部作廢單
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
             "FROM dbo.[@XASCH] T1 " &
             "where T1.status='C' " & sfid_rule &
             " order by T1.sfid,T1.sid,T1.docnum desc"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        ElseIf (DDLFormStatus.SelectedIndex = 6) Then ' 全部簽核單
            If (sfid_rule <> "") Then
                SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
                    "FROM dbo.[@XASCH] T1 " &
                    "where T1.sfid=" & sfid &
                    " order by T1.sfid,T1.sid,T1.docnum desc"
            Else
                SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
                    "FROM dbo.[@XASCH] T1 " &
                    " order by T1.sfid,T1.sid,T1.docnum desc"
            End If
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
            connsap1.Close()
        End If

        ds.Tables(0).Columns.Add("action")
        If (ds.Tables(0).Rows.Count = 0 And displaymode <> 1) Then
            CommUtil.ShowMsg(Me, "無任何資料")
        End If
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()

    End Sub
    Sub SearchDoc()
        ds.Reset()
        SetGridViewStyle()
        SetFormListGridViewFields()
        If (TxtDocnum.Text <> "") Then
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
             "FROM dbo.[@XASCH] T1 " &
             "where docnum=" & TxtDocnum.Text
        ElseIf (TxtKW.Text <> "") Then
            SqlCmd = "SELECT T1.price,T1.priceunit,T1.docnum,T1.subject,T1.sname As issuedperson,T1.sfid,T1.status,T1.docdate,seq=1,T1.sname As uname,T1.sid As uid " &
             "FROM dbo.[@XASCH] T1 " &
             "where subject like '%" & TxtKW.Text & "%' order by T1.sfid,T1.docnum desc"
        Else
            CommUtil.ShowMsg(Me, "表單號或主旨關鍵字須輸入")
            Exit Sub
        End If
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
        connsap1.Close()

        ds.Tables(0).Columns.Add("action")
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "查無條件設定之表單")
        End If
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
    End Sub
    Sub SetGridViewStyle()
        gv1.AutoGenerateColumns = False
        'gv1.ShowHeader = True
        gv1.AllowPaging = True
        gv1.PageSize = 10
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
        oBoundField.HeaderText = "單號"
        oBoundField.DataField = "docnum"
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
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "金額"
        oBoundField.DataField = "price"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:N0}"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "幣別"
        oBoundField.DataField = "priceunit"
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ShowHeader = True
        oBoundField.HtmlEncode = False
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "製單人"
        oBoundField.DataField = "issuedperson"
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

        oBoundField = New BoundField
        oBoundField.HeaderText = "覆核關卡"
        oBoundField.DataField = "uname"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "日期"
        oBoundField.DataField = "docdate"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Left
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

End Class