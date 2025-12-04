Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class spmaterial
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap As New SqlConnection
    Public SqlCmd As String
    Public dr, drsap As SqlDataReader
    Public ds As New DataSet
    Public BtnSPM As Button
    Public TxtKW, TxtCNo, TxtBeginDate, TxtEndDate As TextBox
    Public BtnSearch, BtnCNoSearch As Button
    Public permssp100 As String
    Public ScriptManager1 As New ScriptManager
    Public nowwhs As String
    Public mode As String
    Public TotalPrice As Long
    Public ChkWithZero, ChkAllWhs As CheckBox
    Public allwhs As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim str() As String
        Dim nowdate, year As String
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        permssp100 = CommUtil.GetAssignRight("sp100", Session("s_id"))
        Page.Form.Controls.Add(ScriptManager1)
        str = Split(DDLWhs.SelectedValue, "-")
        nowwhs = str(0)
        FTCreate()

        If (Not IsPostBack) Then
            mode = Request.QueryString("mode")
            AddWhsItem()
            str = Split(DDLWhs.SelectedValue, "_")
            nowwhs = str(0)
            allwhs = Request.QueryString("allwhs")
            If (mode = "inout") Then
                DDLWhs.SelectedIndex = Request.QueryString("ddlwhsindex")
                TxtBeginDate.Text = Request.QueryString("begindate")
                TxtEndDate.Text = Request.QueryString("enddate")

                If (allwhs = "all") Then
                    ChkAllWhs.Checked = True
                Else
                    ChkAllWhs.Checked = False
                End If
                str = Split(DDLWhs.SelectedValue, "_")
                nowwhs = str(0)
                MaterialInOut(allwhs)
                TxtKW.Text = Request.QueryString("kw")
                ViewState("allwhs") = allwhs
            ElseIf (mode = "init") Then
                nowdate = Format(Now, "yyyy/MM/dd")
                'MsgBox(nowdate)
                str = Split(nowdate, "/")
                year = str(0)
                TxtBeginDate.Text = year & "/01/01"
                TxtEndDate.Text = nowdate
                mode = "whssearch"
                ViewState("mode") = mode
                ViewState("allwhs") = allwhs
                DisplayWhsSearch()
            End If
        Else
            mode = ViewState("mode")
            allwhs = ViewState("allwhs")
            str = Split(DDLWhs.SelectedValue, "_")
            nowwhs = str(0)
            'MsgBox(allwhs)
        End If
        'ShowSapreMaterial()
    End Sub

    Sub AddWhsItem()
        SqlCmd = "SELECT T0.[WhsCode], T0.[WhsName] " &
        "FROM OWHS T0 order by T0.WhsCode"
        DDLWhs.Items.Clear()
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            Do While (drsap.Read())
                DDLWhs.Items.Add(drsap(0) & "_" & drsap(1))
            Loop
        End If
        If (Session("grp") = "JF") Then
            DDLWhs.SelectedValue = "S04_捷豐備品倉"
            DDLWhs.Enabled = False
        ElseIf (Session("grp") = "JT") Then
            DDLWhs.SelectedValue = "S05_捷智通備品倉"
            DDLWhs.Enabled = False
        Else
            DDLWhs.SelectedValue = "S04_捷豐備品倉"
            'DDLWhs.SelectedValue = "C02_AOI 倉"
            DDLWhs.Enabled = True
        End If
        drsap.Close()
        connsap.Close()
    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim realindex As Integer
        Dim Hyper As HyperLink
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            e.Row.Cells(0).Text = realindex + 1
            'MsgBox(mode)
            If (mode = "kwsearch" Or mode = "whssearch") Then
                Hyper = New HyperLink()
                Hyper.Text = "過帳記錄"
                Hyper.NavigateUrl = "spmaterial.aspx?smid=sp&smode=1&mode=inout&indexpage=" & gv1.PageIndex & "&itemcode=" & e.Row.Cells(1).Text &
                    "&ddlwhsindex=" & DDLWhs.SelectedIndex & "&kw=" & TxtKW.Text & "&begindate=" & TxtBeginDate.Text & "&enddate=" & TxtEndDate.Text &
                    "&allwhs=" & allwhs
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_inout_" & e.Row.Cells(0).Text
                CommUtil.DisableObjectByPermission(Hyper, permssp100, "m")
                If (InStr(permssp100, "p")) Then
                    e.Row.Cells(7).Controls.Add(Hyper)
                Else
                    e.Row.Cells(5).Controls.Add(Hyper)
                End If
            ElseIf (mode = "inout") Then
                If (e.Row.Cells(6).Text = 0) Then
                    e.Row.Cells(6).Text = ""
                End If
                If (e.Row.Cells(7).Text = 0) Then
                    e.Row.Cells(7).Text = ""
                End If
                If (e.Row.Cells(8).Text.Length > 60) Then
                    e.Row.Cells(8).ToolTip = e.Row.Cells(8).Text
                    e.Row.Cells(8).Text = e.Row.Cells(8).Text.Substring(0, 60) + "..."
                End If
            ElseIf (mode = "cnosearch") Then
                If (e.Row.Cells(8).Text = 0) Then
                    e.Row.Cells(8).Text = ""
                End If
                If (e.Row.Cells(7).Text = 0) Then
                    e.Row.Cells(7).Text = ""
                End If
            End If
        End If
    End Sub
    Protected Sub gv1_PageIndexChanging(sender As Object, e As GridViewPageEventArgs) Handles gv1.PageIndexChanging
        gv1.PageIndex = e.NewPageIndex
        If (mode = "kwsearch") Then
            DisplayKWSearch(TxtKW.Text)
        ElseIf (mode = "whssearch") Then
            DisplayWhsSearch()
        ElseIf (mode = "inout") Then

        End If
    End Sub

    Sub MaterialInOut(allwhs As String)
        Dim itemcode As String
        itemcode = Request.QueryString("itemcode")
        ds.Reset()
        SetGridViewStyle()
        SetMaterialInOutGridViewFields()
        If (allwhs = "all") Then
            SqlCmd = "SELECT icount=0,inamount=0,case when T1.Basetype=202 then '生產發貨' else '一般發貨' end As doctype, " &
                "T0.docdate,T0.docnum,T1.whscode,T1.AcctCode,T1.Quantity As outamount,T0.comments " &
                "FROM OIGE T0 INNER JOIN IGE1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "And T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "SELECT icount=0,outamount=0,case T1.Basetype when 202 then '生產收貨' else '一般收貨' end As doctype, " &
                "T0.docdate,T0.docnum,T1.whscode,T1.AcctCode,T1.Quantity As inamount,T0.comments " &
                "FROM OIGN T0 INNER JOIN IGN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,doctype='調撥單(出)',T0.docdate,T0.docnum, " &
                "T0.Filler As whscode,IsNull(T1.AcctCode,''),T1.Quantity As outamount,T0.comments " &
                "FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,outamount=0,doctype='調撥單(入)',T0.docdate,T0.docnum,T1.whscode, " &
                "IsNull(T1.AcctCode,''),T1.Quantity As inamount,T0.comments " &
                "FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,outamount=0,T1.whscode,IsNull(T1.AcctCode,''),doctype='收貨採購', " &
                "T0.docdate,T0.docnum,T1.Quantity As inamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,AcctCode='',doctype='退貨單', " &
                    "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                    "FROM ORPD T0 INNER JOIN RPD1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                    "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,AcctCode='',doctype='交貨單', " &
                    "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                    "FROM ODLN T0 INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                    "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,IsNull(T1.AcctCode,''),doctype='A/P貸項通知單', " &
                    "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                    "FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                    "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,outamount=0,T1.whscode,IsNull(T1.AcctCode,''),doctype='A/R貸項通知單', " &
                "T0.docdate,T0.docnum,T1.Quantity As inamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,IsNull(T1.AcctCode,''),doctype='AR發票', " &
                    "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                    "FROM OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry " &
                    "where T1.BaseType=-1 and T1.TargetType=-1 and " &
                    "T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                    "and T0.docdate <='" & TxtEndDate.Text & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()
        Else
            SqlCmd = "SELECT icount=0,inamount=0,case when T1.Basetype=202 then '生產發貨' else '一般發貨' end As doctype, " &
                "T0.docdate,T0.docnum,T1.whscode,T1.AcctCode,T1.Quantity As outamount,T0.comments " &
                "FROM OIGE T0 INNER JOIN IGE1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "And T0.docdate <='" & TxtEndDate.Text & "' and T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()
            SqlCmd = "SELECT icount=0,outamount=0,case T1.Basetype when 202 then '生產收貨' else '一般收貨' end As doctype, " &
                "T0.docdate,T0.docnum,T1.whscode,T1.AcctCode,T1.Quantity As inamount,T0.comments " &
                "FROM OIGN T0 INNER JOIN IGN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,doctype='調撥單(出)',T0.docdate,T0.docnum, " &
                "T0.Filler As whscode,IsNull(T1.AcctCode,''),T1.Quantity As outamount,T0.comments " &
                "FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T0.Filler='" & nowwhs & "' and T1.whscode<>'" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,outamount=0,doctype='調撥單(入)',T0.docdate,T0.docnum,T1.whscode, " &
                "IsNull(T1.AcctCode,''),T1.Quantity As inamount,T0.comments " &
                "FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T0.Filler<>'" & nowwhs & "' and T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,outamount=0,T1.whscode,IsNull(T1.AcctCode,''),doctype='收貨採購', " &
                "T0.docdate,T0.docnum,T1.Quantity As inamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,AcctCode='',doctype='退貨單', " &
                "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM ORPD T0 INNER JOIN RPD1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,AcctCode='',doctype='交貨單', " &
                "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM ODLN T0 INNER JOIN DLN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,IsNull(T1.AcctCode,''),doctype='A/P貸項通知單', " &
                "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,outamount=0,T1.whscode,IsNull(T1.AcctCode,''),doctype='A/R貸項通知單', " &
                "T0.docdate,T0.docnum,T1.Quantity As inamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM ORIN T0 INNER JOIN RIN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()

            SqlCmd = "SELECT icount=0,inamount=0,T1.whscode,AcctCode='',doctype='AR發票', " &
                "T0.docdate,T0.docnum,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM OINV T0 INNER JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.BaseType=-1 and T1.TargetType=-1 and " &
                "T1.Itemcode='" & itemcode & "' and T0.docdate>='" & TxtBeginDate.Text & "' " &
                "and T0.docdate <='" & TxtEndDate.Text & "' and " &
                "T1.whscode='" & nowwhs & "'"
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
            connsap.Close()
        End If

        ds.Tables(0).DefaultView.Sort = "docdate"
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "無任何過帳資料")
        End If
    End Sub

    Sub DisplayCNoSearch(cno As String)
        ds.Reset()
        SetGridViewStyle()
        SetMaterialCNoInOutGridViewFields()
        SqlCmd = "SELECT icount=0,inamount=0,case T1.Basetype when 202 then '生產發貨' else '一般發貨' end As doctype,T0.docdate,T0.docnum,T1.whscode,T1.AcctCode,T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription As Itemname " &
                "FROM OIGE T0 INNER JOIN IGE1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.whscode='" & nowwhs & "' and T0.comments like '%" & cno & "%' order by T1.itemcode"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        SqlCmd = "SELECT icount=0,outamount=0,case T1.Basetype when 202 then '生產收貨' else '一般收貨' end As doctype,T0.docdate,T0.docnum,T1.whscode,T1.AcctCode,T1.Quantity As inamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM OIGN T0 INNER JOIN IGN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.whscode='" & nowwhs & "' and T0.comments like '%" & cno & "%' order by T1.itemcode"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        SqlCmd = "SELECT icount=0,inamount=0,doctype='調撥單(出)',T0.docdate,T0.docnum,T0.Filler As whscode,IsNull(T1.AcctCode,''),T1.Quantity As outamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.whscode<>'" & nowwhs & "' and T0.comments like '%" & cno & "%' and T0.Filler='" & nowwhs & "'"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()

        SqlCmd = "SELECT icount=0,outamount=0,doctype='調撥單(入)',T0.docdate,T0.docnum,T0.Filler As whscode,IsNull(T1.AcctCode,''),T1.Quantity As inamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM OWTR T0 INNER JOIN WTR1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.whscode='" & nowwhs & "' and T0.comments like '%" & cno & "%' and T0.Filler<>'" & nowwhs & "'"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()

        SqlCmd = "SELECT icount=0,outamount=0,whscode='" & nowwhs & "',AcctCode='',doctype='收貨採購',T0.docdate,T0.docnum,T1.Quantity As inamount,T0.comments,T1.Itemcode,T1.dscription " &
                "FROM OPDN T0 INNER JOIN PDN1 T1 ON T0.DocEntry = T1.DocEntry " &
                "where T1.whscode='" & nowwhs & "' and T0.comments like '%" & cno & "%' order by T1.itemcode"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()

        ds.Tables(0).DefaultView.Sort = "docdate"
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "無任何聯絡單過帳資料")
        End If
    End Sub
    Sub FTCreate()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim ce As CalendarExtender
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Left
        '--------------------------------
        Labelx = New Label()
        Labelx.ID = "label_kw"
        Labelx.Text = "料號關鍵字:"
        Labelx.Font.Size = 10
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
        BtnSearch.Font.Size = 10
        AddHandler BtnSearch.Click, AddressOf BtnSearch_Click
        tCell.Controls.Add(BtnSearch)
        tRow.Cells.Add(tCell)

        Labelx = New Label()
        Labelx.ID = "label_adds2"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp或&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnSPM = New Button()
        BtnSPM.ID = "btn_spm"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnSPM.Text = "顯示全部"
        BtnSPM.Font.Size = 10
        AddHandler BtnSPM.Click, AddressOf BtnSPM_Click
        tCell.Controls.Add(BtnSPM)

        Labelx = New Label()
        Labelx.ID = "label_zero"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        ChkWithZero = New CheckBox
        ChkWithZero.ID = "chk_zero"
        ChkWithZero.Text = "含庫存0料件"
        ChkWithZero.Font.Size = 10
        'AddHandler ChkWithZero.Click, AddressOf ChkWithZero_Click
        ChkWithZero.AutoPostBack = True
        tCell.Controls.Add(ChkWithZero)

        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Center
        '--------------------------------
        Labelx = New Label()
        Labelx.ID = "label_cno"
        Labelx.Text = "以聯絡單號查詢:"
        Labelx.Font.Size = 10
        tCell.Controls.Add(Labelx)
        TxtCNo = New TextBox()
        TxtCNo.ID = "txt_cno"
        TxtCNo.Width = 120
        tCell.Controls.Add(TxtCNo)
        '-----------------------------------------
        Labelx = New Label()
        Labelx.ID = "label_cno1"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnCNoSearch = New Button()
        BtnCNoSearch.ID = "btn_cnosearch"
        'CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        BtnCNoSearch.Text = "聯絡單過帳尋找"
        BtnCNoSearch.Font.Size = 10
        AddHandler BtnCNoSearch.Click, AddressOf BtnCNoSearch_Click
        tCell.Controls.Add(BtnCNoSearch)
        tRow.Cells.Add(tCell)

        tCell = New TableCell()
        tCell.HorizontalAlign = HorizontalAlign.Right
        ChkAllWhs = New CheckBox
        ChkAllWhs.ID = "chk_allwhs"
        ChkAllWhs.Text = "過帳所有倉別"
        ChkAllWhs.Font.Size = 10
        AddHandler ChkAllWhs.CheckedChanged, AddressOf ChkAllWhs_CheckedChanged
        ChkAllWhs.AutoPostBack = True
        tCell.Controls.Add(ChkAllWhs)

        Labelx = New Label()
        Labelx.ID = "label_begin"
        Labelx.Text = "&nbsp&nbsp&nbsp過帳期間:"
        Labelx.Font.Size = 10
        tCell.Controls.Add(Labelx)
        TxtBeginDate = New TextBox()
        TxtBeginDate.ID = "txt_begin"
        TxtBeginDate.Width = 70
        TxtBeginDate.AutoPostBack = True
        AddHandler TxtBeginDate.TextChanged, AddressOf TxtBeginDate_TextChanged
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
        Labelx.Text = "~"
        tCell.Controls.Add(Labelx)
        TxtEndDate = New TextBox()
        TxtEndDate.ID = "txt_end"
        TxtEndDate.Width = 70
        TxtEndDate.AutoPostBack = True
        AddHandler TxtEndDate.TextChanged, AddressOf TxtEndDate_TextChanged
        tCell.Controls.Add(TxtEndDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtEndDate.ID
        ce.ID = "ce_enddate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tCell.Controls.Add(TxtEndDate)
        tRow.Cells.Add(tCell)
        FT.Rows.Add(tRow)
    End Sub
    Protected Sub ChkAllWhs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) 'kkk
        If (ChkAllWhs.Checked) Then
            allwhs = "all"
        Else
            allwhs = "single"
        End If
        ViewState("mode") = mode
        ViewState("allwhs") = allwhs
        If (mode = "kwsearch") Then
            DisplayKWSearch(TxtKW.Text)
        ElseIf (mode = "whssearch") Then
            DisplayWhsSearch()
            'MsgBox("KK")
        End If
    End Sub
    Protected Sub TxtBeginDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ViewState("mode") = mode
        ViewState("allwhs") = allwhs
        'gv1.PageIndex = Request.QueryString("indexpage")
        If (mode = "kwsearch") Then
            DisplayKWSearch(TxtKW.Text)
            'MsgBox(TxtBeginDate.Text & "-1")
        ElseIf (mode = "whssearch") Then
            DisplayWhsSearch()
            'MsgBox(TxtBeginDate.Text & "-2")
        End If
        'MsgBox(TxtBeginDate.Text & "-3")
    End Sub
    Protected Sub TxtEndDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ViewState("mode") = mode
        ViewState("allwhs") = allwhs
        'gv1.PageIndex = Request.QueryString("indexpage")
        If (mode = "kwsearch") Then
            DisplayKWSearch(TxtKW.Text)
        ElseIf (mode = "whssearch") Then
            DisplayWhsSearch()
        End If
    End Sub
    Protected Sub BtnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gv1.Visible = True
        If (TxtKW.Text <> "") Then
            mode = "kwsearch"
            ViewState("mode") = mode
            ViewState("allwhs") = allwhs
            gv1.PageIndex = Request.QueryString("indexpage")
            DisplayKWSearch(TxtKW.Text)
        Else
            gv1.Visible = False
            CommUtil.ShowMsg(Me, "需輸入關鍵字")
        End If
    End Sub

    Protected Sub BtnCNoSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gv1.Visible = True
        If (TxtCNo.Text <> "") Then
            mode = "cnosearch"
            ViewState("mode") = mode
            ViewState("allwhs") = allwhs
            gv1.PageIndex = Request.QueryString("indexpage")
            DisplayCNoSearch(TxtCNo.Text)
        Else
            gv1.Visible = False
            CommUtil.ShowMsg(Me, "需輸入聯絡單號")
        End If
    End Sub

    Protected Sub BtnSPM_Click(sender As Object, e As EventArgs)
        gv1.Visible = True
        mode = "whssearch"
        ViewState("mode") = mode
        ViewState("allwhs") = allwhs
        gv1.PageIndex = Request.QueryString("indexpage")
        DisplayWhsSearch()
    End Sub

    'Protected Sub gv1_Sorting(sender As Object, e As GridViewSortEventArgs) Handles gv1.Sorting
    '    If ViewState("mySorting") = Nothing Then
    '        e.SortDirection = SortDirection.Ascending
    '        ViewState("mySorting") = "Ascending"
    '    Else
    '        '-- 如果目前的排序方法，已經是「正排序」，那再度按下排序欄位之後，就變成「反排序」。
    '        If ViewState("mySorting") = "Ascending" Then
    '            e.SortDirection = SortDirection.Descending
    '            ViewState("mySorting") = "Descending"
    '        Else
    '            e.SortDirection = SortDirection.Ascending
    '            ViewState("mySorting") = "Ascending"
    '        End If
    '    End If
    'End Sub

    Sub DisplayKWSearch(kw As String)
        SqlCmd = "SELECT IsNull(Sum(T1.OnHand*T1.AvgPrice),0) " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.onhand<>0 and T1.whscode='" & nowwhs & "' and T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%'"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        TotalPrice = drsap(0)
        drsap.Close()
        connsap.Close()

        ds.Reset()
        SetGridViewStyle()
        SetMaterialGridViewFields()
        If (ChkWithZero.Checked) Then
            If (InStr(permssp100, "p")) Then
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand,T1.AvgPrice,(T1.AvgPrice*T1.Onhand) As tprice " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.whscode='" & nowwhs & "' and T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%' order by T0.itemcode"
            Else
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.whscode='" & nowwhs & "' and T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%' order by T0.itemcode"
            End If
        Else
            If (InStr(permssp100, "p")) Then
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand,T1.AvgPrice,(T1.AvgPrice*T1.Onhand) As tprice " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.onhand<>0 and T1.whscode='" & nowwhs & "' and T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%' order by T0.itemcode"
            Else
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.onhand<>0 and T1.whscode='" & nowwhs & "' and T0.itemcode like '%" & kw & "%' or T0.itemname like '%" & kw & "%' order by T0.itemcode"
            End If
        End If
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        'ds.Tables(0).DefaultView.Sort = "onhand desc"
        connsap.Close()
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "無任何資料")
        End If

    End Sub
    Sub DisplayWhsSearch()
        SqlCmd = "SELECT IsNull(Sum(T1.OnHand*T1.AvgPrice),0) " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.onhand<>0 and T1.whscode='" & nowwhs & "'"
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        drsap.Read()
        TotalPrice = drsap(0)
        drsap.Close()
        connsap.Close()

        ds.Reset()
        SetGridViewStyle()
        SetMaterialGridViewFields()
        If (ChkWithZero.Checked) Then
            If (InStr(permssp100, "p")) Then
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand,T1.AvgPrice,(T1.AvgPrice*T1.Onhand) As tprice " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.whscode='" & nowwhs & "' order by T0.itemcode"
            Else
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.whscode='" & nowwhs & "' order by T0.itemcode"
            End If
        Else
            If (InStr(permssp100, "p")) Then
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand,T1.AvgPrice,(T1.AvgPrice*T1.Onhand) As tprice " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.onhand<>0 and T1.whscode='" & nowwhs & "' order by T0.itemcode"
            Else
                SqlCmd = "SELECT icount=0,status=0,T1.itemcode,T0.itemname,T1.whscode,T1.onhand " &
                "FROM OITM T0 INNER JOIN OITW T1 ON T0.itemcode=T1.Itemcode " &
                "where T1.onhand<>0 and T1.whscode='" & nowwhs & "' order by T0.itemcode"
            End If
        End If
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        'ds.Tables(0).DefaultView.Sort = "onhand desc"
        connsap.Close()
        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
        If (ds.Tables(0).Rows.Count = 0) Then
            CommUtil.ShowMsg(Me, "無任何資料")
        End If
    End Sub

    Sub SetGridViewStyle()
        gv1.AutoGenerateColumns = False
        'gv1.ShowHeader = True
        If (mode <> "inout" And mode <> "cnosearch") Then
            gv1.AllowPaging = True
            gv1.PageSize = 25
            gv1.PagerStyle.HorizontalAlign = HorizontalAlign.Center
        Else
            gv1.AllowPaging = False
        End If
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

        oBoundField = New BoundField
        oBoundField.HeaderText = "倉庫"
        oBoundField.DataField = "whscode"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "數量"
        oBoundField.DataField = "onhand"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)

        If (InStr(permssp100, "p")) Then
            oBoundField = New BoundField
            oBoundField.HeaderText = "平均成本"
            oBoundField.FooterText = "篩選總價"
            oBoundField.DataField = "AvgPrice"
            oBoundField.ShowHeader = True
            oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
            oBoundField.DataFormatString = "{0:N0}"
            gv1.Columns.Add(oBoundField)

            oBoundField = New BoundField
            oBoundField.HeaderText = "總價"
            oBoundField.DataField = "tprice"
            oBoundField.FooterText = Format(TotalPrice, "###,###")
            oBoundField.ShowHeader = True
            oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
            oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
            oBoundField.DataFormatString = "{0:N0}"
            gv1.Columns.Add(oBoundField)
        End If
        oBoundField = New BoundField
        oBoundField.HeaderText = "狀態"
        oBoundField.DataField = "status"
        If (InStr(permssp100, "p")) Then
            oBoundField.FooterText = "NTD"
        End If
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
        oBoundField.HeaderText = "進貨量"
        oBoundField.DataField = "inamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "退貨量"
        oBoundField.DataField = "outamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "備註"
        oBoundField.DataField = "comments"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)
    End Sub
    Sub SetMaterialCNoInOutGridViewFields()
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
        oBoundField.HeaderText = "倉庫"
        oBoundField.DataField = "whscode"
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
        oBoundField.HeaderText = "發貨量"
        oBoundField.DataField = "outamount"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.HorizontalAlign = HorizontalAlign.Center
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        oBoundField.DataFormatString = "{0:F0}"
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)

        oBoundField = New BoundField
        oBoundField.HeaderText = "備註"
        oBoundField.DataField = "comments"
        oBoundField.ShowHeader = True
        oBoundField.ItemStyle.VerticalAlign = VerticalAlign.Middle
        'oBoundField.SortExpression = "onhand"
        gv1.Columns.Add(oBoundField)
    End Sub
End Class