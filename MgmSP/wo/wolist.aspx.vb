Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Partial Public Class wolist
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public connsap As New SqlConnection
    Public SqlCmd As String
    Public modocnum, wsn As String
    Public dr As SqlDataReader
    Public oCompany As New SAPbobsCOM.Company
    Public ret As Long
    Public permsmf202 As String '發料通知
    'Public permsmf203 As String '發料操作
    Public permsmf201 As String
    Public ScriptManager1 As New ScriptManager
    Public nowurl As String
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
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim preact As String
        'MsgBox("HH" & vbCr & vbCr & "KK" & vbCr)
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        If (Session("s_id") = "ron" Or Session("s_id") = "su" Or Session("s_id") = "ltx") Then
            IssuedAutoCheck.Visible = True
            IssuedAutoCheck.Enabled = True
        Else
            IssuedAutoCheck.Visible = False
            IssuedAutoCheck.Enabled = False
        End If
        permsmf202 = CommUtil.GetAssignRight("mf202", Session("s_id"))
        permsmf201 = CommUtil.GetAssignRight("mf201", Session("s_id"))
        If (Not IsPostBack) Then
            DDLWoFun.Items.Clear()
            DDLWoFun.Items.Add("請選取功能")
            DDLWoFun.Items.Add("發料通知")
            DDLWoFun.Items.Add("開工單")
            DDLWoFun.Items.Add("模組退料")
            DDLWoFun.SelectedIndex = 1
            reqdate_text.Enabled = True
            ExecuteBtn.Enabled = True
            DDLAlter.Items.Clear()
            DDLAlter.Items.Add("請選取替代方法")
            DDLAlter.Items.Add("一般替代")
            DDLAlter.Items.Add("必替替代")
            DDLAlter.Items.Add("不替代")
            IssuedAutoCheck.Checked = False
            preact = Request.QueryString("preact")
            If (preact = "退料") Then
                CommUtil.ShowMsg(Me, "退料完成")
            End If
        End If
        GetWoList()
    End Sub

    Sub GetWoList()
        Dim ds As New DataSet
        modocnum = Request.QueryString("modocnum")
        wsn = Request.QueryString("wsn")
        nowurl = "~/wo/wolist.aspx?modocnum=" & modocnum & "*wsn=" & wsn & "*smid=molist*smode=0"

        SqlCmd = "Select OWOR.DueDate, WOR1.ItemCode, OITM.ItemName, WOR1.PlannedQty, WOR1.IssuedQty " &
        ", WOR1.warehouse, OWOR.Comments,OITW.OnHand,OITW.IsCommited,OITW.OnOrder FROM OWOR " &
        "INNER JOIN WOR1 On OWOR.DocEntry = WOR1.DocEntry " &
        "INNER JOIN OITM On WOR1.ItemCode = OITM.ItemCode " &
        "INNER JOIN OITW On OITM.itemcode=OITW.itemcode " &
        "WHERE OITW.Whscode=OWOR.warehouse And WOR1.issuetype = 'M' and OWOR.DocNum = '" & modocnum & "' " &
        "ORDER BY WOR1.ItemCode"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
        connsap.Close()
        If (ds.Tables(0).Rows.Count <> 0) Then
            ds.Tables(0).Columns.Add("docnum")
            ds.Tables(0).Columns.Add("status")
            ds.Tables(0).Columns.Add("shortage")
            ds.Tables(0).Columns.Add("rest")
            ds.Tables(0).Columns.Add("amount")
            ds.Tables(0).Columns.Add("selchk")
            ds.Tables(0).Columns.Add("sysissued")
            gv1.DataSource = ds.Tables(0)
            gv1.DataBind()
        Else
            CommUtil.ShowMsg(Me, "無任何資料")
        End If
    End Sub
    'Sub GetWoList1()
    '    Dim ds As New DataSet
    '    modocnum = Request.QueryString("modocnum")
    '    wsn = Request.QueryString("wsn")
    '    nowurl = "~/wo/wolist.aspx?modocnum=" & modocnum & "*wsn=" & wsn & "*smid=molist*smode=0"

    '    SqlCmd = "SELECT T0.DocNum,T0.Status,T0.CmpltQty,T0.DueDate,T0.Comments,T0.PlannedQty,T0.itemcode,T1.Itemname, " &
    '    "T2.onhand,T0.warehouse,T2.IsCommited,T2.Onorder,IssuedQty=0 " &
    '    "FROM OWOR T0 " &
    '    "INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode " &
    '    "INNER JOIN OITW T2 ON T1.ItemCode=T2.ItemCode " &
    '    "WHERE T0.U_F16 = '" & modocnum & "' and " &
    '    "T0.Status<>'C' and T0.DocNum <> '" & modocnum & "' and T0.warehouse=T2.whscode order by T0.itemcode,T0.docnum"
    '    ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
    '    connsap.Close()

    '    'SqlCmd = "SELECT T0.DocNum=0,status='NA', T1.ItemCode, T2.ItemName,T1.PlannedQty, T1.IssuedQty,T1.BaseQty " &
    '    '", T1.warehouse, T0.Comments='',T3.OnHand,T3.IsCommited,T3.Onorder FROM OWOR T0 " &
    '    '"INNER JOIN WOR1 T1 ON T0.DocEntry = T1.DocEntry " &
    '    '"INNER JOIN OITM T2 ON T1.ItemCode = T2.ItemCode " &
    '    '"INNER JOIN OITW T3 ON T2.ItemCode = T3.ItemCode " &
    '    '"WHERE T0.warehouse=T3.whscode and T1.issuetype = 'M' and T0.DocNum = '" & modocnum & "' " &
    '    '"and T1.Itemcode not in (SELECT T0.itemcode " &
    '    '"FROM OWOR T0 " &
    '    '"WHERE T0.U_F16 = '" & modocnum & "' and " &
    '    '"T0.Status<>'C' and T0.DocNum <> '" & modocnum & "')"
    '    'ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap)
    '    'connsap.Close()
    '    If (ds.Tables(0).Rows.Count <> 0) Then
    '        ds.Tables(0).Columns.Add("shortage")
    '        ds.Tables(0).Columns.Add("rest")
    '        ds.Tables(0).Columns.Add("amount")
    '        ds.Tables(0).Columns.Add("selchk")
    '        ds.Tables(0).Columns.Add("sysissued")
    '        'ds.Tables(0).Columns.Add("issuedqty")
    '        gv1.DataSource = ds.Tables(0)
    '        gv1.DataBind()
    '    Else
    '        CommUtil.ShowMsg(Me, "無任何資料")
    '    End If
    'End Sub
    'Protected Sub gv1_RowDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound
    '    Dim tTxt As TextBox
    '    Dim cChk As CheckBox
    '    Dim Hyper As HyperLink
    '    Dim wor1exist, nowo As Boolean
    '    wor1exist = False
    '    nowo = False
    '    ' Dim TipType As New ToolTip() '宣告引用類別

    '    If (e.Row.RowType = DataControlRowType.Header) Then
    '        'e.Row.Cells.Clear()
    '        AddOneRowSpanCol(sender)
    '        If (DDLWoFun.SelectedIndex = 3) Then
    '            e.Row.Cells(13).Text = "欲退數量"
    '        Else
    '            e.Row.Cells(13).Text = "欲領數量"
    '        End If
    '    End If
    '    If (e.Row.RowType = DataControlRowType.DataRow) Then
    '        'TipType.SetToolTip(e.Row.Cells(2).Text, "執行命令:匯出EXCEL功能鈕")
    '        e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='lightgreen'")
    '        '設定光棒顏色，當滑鼠 onMouseOver 時驅動
    '        e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
    '        '當 onMouseOut 也就是滑鼠移開時，要恢復原本的顏色
    '        If (e.Row.Cells(0).Text <> 0) Then
    '            Hyper = New HyperLink()
    '            Hyper.Text = e.Row.Cells(0).Text
    '            Hyper.Font.Underline = False
    '            Hyper.ID = "hyper_wo_" & e.Row.RowIndex
    '            Hyper.NavigateUrl = "~/commcode/ShowData.aspx?dtype=wo&wo=" & e.Row.Cells(0).Text &
    '                            "&indexpage=" & gv1.PageIndex &
    '                            "&orgurl=" & nowurl
    '            e.Row.Cells(0).Controls.Add(Hyper)
    '            If (e.Row.Cells(3).Text = "P") Then
    '                e.Row.Cells(3).Text = "計畫中"
    '            ElseIf (e.Row.Cells(3).Text = "R") Then
    '                e.Row.Cells(3).Text = "已核發"
    '            ElseIf (e.Row.Cells(3).Text = "L") Then
    '                e.Row.Cells(3).Text = "已結案"
    '            End If
    '            SqlCmd = "SELECT T1.IssuedQty " &
    '                "FROM OWOR T0 " &
    '                "INNER JOIN WOR1 T1 ON T0.DocEntry = T1.DocEntry " &
    '                "WHERE T1.issuetype = 'M' and T0.Docnum='" & modocnum & "' and T1.itemcode = '" & e.Row.Cells(1).Text & "'"
    '            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '            If (dr.HasRows) Then
    '                dr.Read()
    '                e.Row.Cells(11).Text = CInt(dr(0))
    '                wor1exist = True
    '            Else
    '                e.Row.Cells(0).BackColor = Drawing.Color.Red
    '                e.Row.Cells(10).Text = "NA"
    '                e.Row.Cells(11).Text = "NA"
    '                e.Row.Cells(12).Text = "NA"
    '            End If
    '            dr.Close()
    '            connsap.Close()
    '        Else
    '            nowo = True
    '            e.Row.Cells(0).Text = ""
    '            e.Row.Cells(3).Text = "NA"
    '            'e.Row.Cells(4).Text = 0
    '        End If

    '        SqlCmd = "select count(*) from OITT where code='" & e.Row.Cells(1).Text & "'"
    '        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
    '        dr.Read()
    '        If (dr(0) > 0) Then
    '            Hyper = New HyperLink()
    '            Hyper.Text = e.Row.Cells(1).Text
    '            Hyper.Font.Underline = False
    '            Hyper.ID = "hyper_bom_" & e.Row.RowIndex
    '            Hyper.NavigateUrl = "~/commcode/ShowData.aspx?dtype=bom&bomcode=" & e.Row.Cells(1).Text &
    '                            "&indexpage=" & gv1.PageIndex &
    '                            "&orgurl=" & nowurl
    '            e.Row.Cells(1).Controls.Add(Hyper)
    '        End If
    '        dr.Close()
    '        connsap.Close()
    '        e.Row.Cells(9).Text = CStr(CInt(e.Row.Cells(6).Text) + CInt(e.Row.Cells(8).Text) - CInt(e.Row.Cells(7).Text))
    '        If (CInt(e.Row.Cells(9).Text) < 0) Then
    '            e.Row.Cells(9).BackColor = Drawing.Color.Red
    '        End If
    '        If (wor1exist Or nowo) Then
    '            If (wor1exist) Then
    '                SqlCmd = "SELECT sum(T0.req_set) over (partition by itemcode),sum(T0.finish_amount) over (partition by itemcode) " &
    '                 "From omri T0 where T0.modocnum=" & modocnum & " and T0.itemcode='" & e.Row.Cells(1).Text & "' and T0.docnum='" & e.Row.Cells(0).Text & "'"
    '            Else
    '                SqlCmd = "SELECT sum(T0.req_set) over (partition by itemcode),sum(T0.finish_amount) over (partition by itemcode) " &
    '                 "From omri T0 where T0.modocnum=" & modocnum & " and T0.itemcode='" & e.Row.Cells(1).Text & "'"
    '            End If
    '            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
    '            If (dr.HasRows) Then
    '                dr.Read()
    '                e.Row.Cells(10).Text = dr(1)
    '                e.Row.Cells(12).Text = CStr(CInt(e.Row.Cells(4).Text) - dr(0))
    '            Else
    '                e.Row.Cells(10).Text = 0
    '                e.Row.Cells(12).Text = e.Row.Cells(4).Text
    '            End If
    '            conn.Close()

    '            If (DDLWoFun.SelectedIndex = 2 Or DDLWoFun.SelectedIndex = 1) Then
    '                If (CInt(e.Row.Cells(12).Text) > 0) Then
    '                    tTxt = New TextBox()
    '                    tTxt.ID = "tTxt_" & e.Row.RowIndex
    '                    tTxt.Width = 30
    '                    CommUtil.DisableObjectByPermission(tTxt, permsmf202, "n")
    '                    e.Row.Cells(13).Controls.Add(tTxt)
    '                End If
    '            ElseIf (DDLWoFun.SelectedIndex = 3) Then
    '                If (CInt(e.Row.Cells(10).Text) > 0) Then
    '                    tTxt = New TextBox()
    '                    tTxt.ID = "tTxt_" & e.Row.RowIndex
    '                    tTxt.Width = 30
    '                    CommUtil.DisableObjectByPermission(tTxt, permsmf202, "n")
    '                    e.Row.Cells(13).Controls.Add(tTxt)
    '                End If
    '            End If
    '            If ((DDLWoFun.SelectedIndex = 2 And e.Row.Cells(0).Text = "") Or
    '                (DDLWoFun.SelectedIndex = 1 And CInt(e.Row.Cells(12).Text) > 0) Or
    '                (DDLWoFun.SelectedIndex = 3 And CInt(e.Row.Cells(10).Text) > 0)) Then
    '                cChk = New CheckBox()
    '                cChk.ID = "cChk_" & e.Row.RowIndex
    '                If (DDLWoFun.SelectedIndex = 1 And CInt(e.Row.Cells(12).Text) > 0) Then
    '                    CommUtil.DisableObjectByPermission(cChk, permsmf202, "n")
    '                ElseIf (DDLWoFun.SelectedIndex = 2 And e.Row.Cells(0).Text = "") Then
    '                    CommUtil.DisableObjectByPermission(cChk, permsmf201, "n")
    '                ElseIf (DDLWoFun.SelectedIndex = 3 And e.Row.Cells(12).Text > 0) Then
    '                    CommUtil.DisableObjectByPermission(cChk, permsmf202, "n")
    '                End If
    '                e.Row.Cells(14).Controls.Add(cChk)
    '            End If
    '            If (e.Row.Cells(10).Text <> e.Row.Cells(11).Text) Then
    '                e.Row.Cells(10).BackColor = Drawing.Color.Yellow
    '                e.Row.Cells(11).BackColor = Drawing.Color.Yellow
    '            End If
    '        End If
    '    End If
    'End Sub
    Protected Sub gv1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim tTxt As TextBox
        Dim cChk As CheckBox
        Dim Hyper As HyperLink
        ' Dim TipType As New ToolTip() '宣告引用類別

        If (e.Row.RowType = DataControlRowType.Header) Then
            'e.Row.Cells.Clear()
            AddOneRowSpanCol(sender)
            If (DDLWoFun.SelectedIndex = 3) Then
                e.Row.Cells(13).Text = "欲退數量"
            Else
                e.Row.Cells(13).Text = "欲領數量"
            End If
        End If
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            'TipType.SetToolTip(e.Row.Cells(2).Text, "執行命令:匯出EXCEL功能鈕")
            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='lightgreen'")
            '設定光棒顏色，當滑鼠 onMouseOver 時驅動
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            '當 onMouseOut 也就是滑鼠移開時，要恢復原本的顏色
            'InitSAPSQLConnection()
            SqlCmd = "SELECT OWOR.DocNum,OWOR.Status,OWOR.CmpltQty,OWOR.DueDate,OWOR.Comments,OWOR.PlannedQty FROM OWOR WHERE OWOR.U_F16 = '" & modocnum & "' and " &
            "OWOR.Status<>'C' and OWOR.DocNum <> '" & modocnum & "' and OWOR.ItemCode= '" & e.Row.Cells(1).Text & "'"
            'myCommand = New SqlCommand(SqlCmd, connsap)
            'dr = myCommand.ExecuteReader()
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                e.Row.Cells(0).Text = dr(0)
                Hyper = New HyperLink()
                Hyper.Text = e.Row.Cells(0).Text
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_wo_" & e.Row.RowIndex
                Hyper.NavigateUrl = "~/commcode/ShowData.aspx?dtype=wo&wo=" & e.Row.Cells(0).Text &
                                "&indexpage=" & gv1.PageIndex &
                                "&orgurl=" & nowurl
                e.Row.Cells(0).Controls.Add(Hyper)
                If (dr(1) = "P") Then
                    e.Row.Cells(3).Text = "計畫中"
                ElseIf (dr(1) = "R") Then
                    e.Row.Cells(3).Text = "已核發"
                ElseIf (dr(1) = "L") Then
                    e.Row.Cells(3).Text = "已結案"
                End If
            Else
                e.Row.Cells(0).Text = ""
                e.Row.Cells(3).Text = "NA"
            End If
            dr.Close()
            connsap.Close()
            Dim IsBom As Boolean
            IsBom = False
            SqlCmd = "select count(*) from OITT where code='" & e.Row.Cells(1).Text & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            If (dr(0) > 0) Then
                IsBom = True
                Hyper = New HyperLink()
                Hyper.Text = e.Row.Cells(1).Text
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_bom_" & e.Row.RowIndex
                Hyper.NavigateUrl = "~/commcode/ShowData.aspx?dtype=bom&bomcode=" & e.Row.Cells(1).Text &
                                "&indexpage=" & gv1.PageIndex &
                                "&orgurl=" & nowurl
                e.Row.Cells(1).Controls.Add(Hyper)
            End If
            dr.Close()
            connsap.Close()
            e.Row.Cells(9).Text = CStr(CInt(e.Row.Cells(6).Text) + CInt(e.Row.Cells(8).Text) - CInt(e.Row.Cells(7).Text))
            If (CInt(e.Row.Cells(9).Text) < 0) Then
                e.Row.Cells(9).BackColor = Drawing.Color.Red
            End If
            'InitLocalSQLConnection()
            SqlCmd = "SELECT sum(T0.req_set) over (partition by itemcode),sum(T0.finish_amount) over (partition by itemcode) " &
                     "From omri T0 where T0.modocnum=" & modocnum & " and T0.itemcode='" & e.Row.Cells(1).Text & "'"
            'myCommand = New SqlCommand(SqlCmd, conn)
            'dr = myCommand.ExecuteReader()
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            If (dr.HasRows) Then
                dr.Read()
                e.Row.Cells(10).Text = dr(1)
                e.Row.Cells(12).Text = CStr(CInt(e.Row.Cells(4).Text) - dr(0))
            Else
                e.Row.Cells(10).Text = 0
                e.Row.Cells(12).Text = e.Row.Cells(4).Text
            End If
            conn.Close()
            If (DDLWoFun.SelectedIndex = 2 Or DDLWoFun.SelectedIndex = 1) Then
                If (CInt(e.Row.Cells(12).Text) > 0) Then
                    tTxt = New TextBox()
                    tTxt.ID = "tTxt_" & e.Row.RowIndex
                    tTxt.Width = 30
                    CommUtil.DisableObjectByPermission(tTxt, permsmf202, "n")
                    e.Row.Cells(13).Controls.Add(tTxt)
                End If
            ElseIf (DDLWoFun.SelectedIndex = 3) Then
                If (CInt(e.Row.Cells(10).Text) > 0) Then
                    tTxt = New TextBox()
                    tTxt.ID = "tTxt_" & e.Row.RowIndex
                    tTxt.Width = 30
                    CommUtil.DisableObjectByPermission(tTxt, permsmf202, "n")
                    e.Row.Cells(13).Controls.Add(tTxt)
                End If
            End If
            If ((DDLWoFun.SelectedIndex = 2 And e.Row.Cells(0).Text = "" And IsBom) Or
                (DDLWoFun.SelectedIndex = 1 And CInt(e.Row.Cells(12).Text) > 0) Or
                (DDLWoFun.SelectedIndex = 3 And CInt(e.Row.Cells(10).Text) > 0)) Then
                cChk = New CheckBox()
                cChk.ID = "cChk_" & e.Row.RowIndex
                If (DDLWoFun.SelectedIndex = 1 And CInt(e.Row.Cells(12).Text) > 0) Then
                    CommUtil.DisableObjectByPermission(cChk, permsmf202, "n")
                ElseIf (DDLWoFun.SelectedIndex = 2 And e.Row.Cells(0).Text = "") Then
                    CommUtil.DisableObjectByPermission(cChk, permsmf201, "n")
                ElseIf (DDLWoFun.SelectedIndex = 3 And e.Row.Cells(12).Text > 0) Then
                    CommUtil.DisableObjectByPermission(cChk, permsmf202, "n")
                End If
                e.Row.Cells(14).Controls.Add(cChk)
            End If
            If (e.Row.Cells(10).Text <> e.Row.Cells(11).Text) Then
                e.Row.Cells(10).BackColor = Drawing.Color.Yellow
                e.Row.Cells(11).BackColor = Drawing.Color.Yellow
            End If
        End If
    End Sub

    Sub AddOneRowSpanCol(ByVal sender)
        Dim gv As GridView = CType(sender, GridView)
        Dim gvrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert)
        Dim tc0 As TableCell = New TableCell()
        'CommUtil.ShowMsg(Me,modocnum)
        'InitSAPSQLConnection()
        SqlCmd = "Select OWOR.ItemCode, OITM.ItemName,OWOR.DueDate,OWOR.PlannedQty FROM OWOR " & _
        "INNER JOIN OITM On OWOR.ItemCode = OITM.ItemCode " & _
        "WHERE OWOR.DocNum = '" & modocnum & "'"
        'myCommand = New SqlCommand(SqlCmd, connsap)
        'dr = myCommand.ExecuteReader()
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        tc0.ID = "modescri"
        tc0.Text = "母單號:" & modocnum & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp工單號:" & wsn & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp料號:" & dr(0) & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp說明:" &
                    dr(1) & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp數量:" & CInt(dr(3)) & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp出貨日:" & dr(2)
        tc0.ColumnSpan = 15
        tc0.Font.Bold = True
        tc0.BorderWidth = 5
        tc0.HorizontalAlign = HorizontalAlign.Center
        'tc0.Controls.Add(HyperLink)
        tc0.BackColor = System.Drawing.Color.LightSkyBlue
        gvrow.Cells.Add(tc0)
        gv.Controls(0).Controls.AddAt(0, gvrow)
        dr.Close()
        connsap.Close()
    End Sub

    Protected Sub ExecuteBtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExecuteBtn.Click
        Dim chkcount As Integer
        Dim gvr As GridViewRow
        Dim j As Integer
        chkcount = 0
        If (DDLWoFun.SelectedValue = "發料通知") Then
            If (reqdate_text.Text = "") Then
                CommUtil.ShowMsg(Me, "領料需求日期不能空白")
                Exit Sub
            End If
            If (wsn = "NA") Then
                CommUtil.ShowMsg(Me, "請先開立對應sap工單的系統工單")
                Exit Sub
            End If
            For Each gvr In gv1.Rows
                j = gvr.RowIndex

                If (CInt(gv1.Rows(j).Cells(12).Text) > 0) Then
                    If CType(gvr.FindControl("cChk_" & j), CheckBox).Checked = True Then
                        chkcount = chkcount + 1
                        If (CType(gvr.FindControl("tTxt_" & j), TextBox).Text = "") Then
                            CommUtil.ShowMsg(Me, "領料數量不能空白")
                            Exit Sub
                        End If
                    End If
                End If
            Next
            If (chkcount <> 0) Then
                IssuedMaterialInfo()
            Else
                CommUtil.ShowMsg(Me, "沒選擇領料項目之checkbox")
            End If
        ElseIf (DDLWoFun.SelectedValue = "開工單") Then
            'chk = MsgBox(Me, "你確定要開SAP子工單嗎?", vbYesNo)
            'If chk = 7 Then
            'Exit Sub
            'End If
            If (DDLAlter.SelectedIndex = 0) Then
                CommUtil.ShowMsg(Me, "需選擇替代方法")
                Exit Sub
            End If
            CreateWo()
        ElseIf (DDLWoFun.SelectedValue = "模組退料") Then
            For Each gvr In gv1.Rows
                j = gvr.RowIndex

                If (CInt(gv1.Rows(j).Cells(10).Text) > 0) Then
                    If CType(gvr.FindControl("cChk_" & j), CheckBox).Checked = True Then
                        chkcount = chkcount + 1
                        If (CType(gvr.FindControl("tTxt_" & j), TextBox).Text = "") Then
                            CommUtil.ShowMsg(Me, "退料數量不能空白")
                            Exit Sub
                        End If
                        If (CInt(CType(gvr.FindControl("tTxt_" & j), TextBox).Text) > CInt(gv1.Rows(j).Cells(10).Text)) Then
                            CommUtil.ShowMsg(Me, "第" & j & "列之退料數量不能大於已領數量")
                            Exit Sub
                        End If
                    End If
                End If
            Next
            If (chkcount <> 0) Then
                RefundMaterialInfo()
            Else
                CommUtil.ShowMsg(Me, "沒選擇退料項目之checkbox")
            End If

        Else
            CommUtil.ShowMsg(Me,"沒選擇功能")
            'Del()
        End If
    End Sub

    Sub Del()
        'InitLocalSQLConnection()
        SqlCmd = "delete from omri" ' where num=7"
        'myCommand = New SqlCommand(SqlCmd, conn)
        'myCommand.ExecuteNonQuery()
        CommUtil.SqlLocalExecute("del", SqlCmd, conn)
        conn.Close()
        SqlCmd = "ALTER TABLE omri AUTO_INCREMENT=1"
        CommUtil.SqlLocalExecute("alter", SqlCmd, conn)
        conn.Close()
        'myCommand = New SqlCommand(SqlCmd, conn)
        'myCommand.ExecuteNonQuery()
        'CloseLocalSQLConnection()
    End Sub
    Sub IssuedMaterialInfo()
        'Dim KeyName As String '要的关键字,实际就是数据表的主键.需要事先在GridView1的DataKeyNames中设置
        Dim j As Integer
        Dim gvr As GridViewRow
        Dim docnum, itemcode, itemname, ownername, subject, content, title As String

        Dim req_set As Integer
        Dim build_date, upd_date, req_date As Date
        content = ""
        CommUtil.InitLocalSQLConnection(conn)
        build_date = FormatDateTime(Now(), DateFormat.ShortDate)
        upd_date = FormatDateTime(Now(), DateFormat.ShortDate)
        ownername = Session("s_name")
        req_date = reqdate_text.Text
        For Each gvr In gv1.Rows
            j = gvr.RowIndex
            If (CInt(gv1.Rows(j).Cells(12).Text) > 0) Then
                If CType(gvr.FindControl("cChk_" & j), CheckBox).Checked = True Then
                    docnum = gv1.Rows(j).Cells(0).Text
                    itemcode = gv1.Rows(j).Cells(1).Text
                    itemname = gv1.Rows(j).Cells(2).Text
                    req_set = CType(gvr.FindControl("tTxt_" & j), TextBox).Text
                    content = content & j + 1 & "." & docnum & "  " & itemcode & "  " & itemname & "   數量: " & req_set & vbCr
                    If (IssuedAutoCheck.Checked = False) Then
                        SqlCmd = "Insert Into omri (modocnum,docnum,wsn,itemcode,itemname,req_set,build_date,req_date,ownername,comm) " &
                        "Values ('" & modocnum & "','" & docnum & "','" & wsn & "','" & itemcode & "','" & itemname & "'," & req_set & ",'" &
                             build_date & "','" & req_date & "','" & ownername & "',' ')"
                        CommUtil.SqlExecute("ins", SqlCmd, conn)
                    Else
                        SqlCmd = "Insert Into omri (modocnum,docnum,wsn,itemcode,itemname,req_set,build_date,req_date,ownername,comm,finish_amount,stat,finish,upd_date) " &
                            "Values ('" & modocnum & "','" & docnum & "','" & wsn & "','" & itemcode & "','" & itemname & "'," & req_set & ",'" &
                            build_date & "','" & req_date & "','" & ownername & "',' '," & req_set & ",20,1,'" & FormatDateTime(Now(), DateFormat.ShortDate) & "')"
                        CommUtil.SqlExecute("ins", SqlCmd, conn)
                    End If
                    'CommUtil.ShowMsg(Me,CType(gvr.FindControl("tTxt_" & j), TextBox).Text)
                    'i = gvr.RowIndex 'GridView行索引   
                    'KeyName = gv1.DataKeys(i).Value
                    '...根据KeyName想做什么做什么吧.   
                End If
            End If
        Next
        SqlCmd = "Select OWOR.ItemCode, OITM.ItemName,OWOR.DueDate,OWOR.PlannedQty FROM OWOR " &
                        "INNER JOIN OITM On OWOR.ItemCode = OITM.ItemCode " &
                        "WHERE OWOR.DocNum = '" & modocnum & "'"
        'myCommand = New SqlCommand(SqlCmd, connsap)
        'dr = myCommand.ExecuteReader()
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        title = "發料通知===>母單號:" & modocnum & " 工單號:" & wsn & " 料號:" & dr(0) & " 說明:" &
        dr(1) & " 數量:" & CInt(dr(3)) & " 出貨日:" & dr(2)
        'subject = "發料通知===>Sap母單號:" & modocnum & " 說明:" & dr(1) & "多模組發料--子工單:" & firstdocnum & " 模組料號:" & firstitemcode &
        '          " 模組名稱:" & firstitemname & " 數量:" & firstreq_set
        subject = "發料通知===>Sap母單號:" & modocnum & " 說明:" & dr(1)
        content = title & vbCr & vbCr & content
        'MsgBox(content)
        CommUtil.SendMailTxt("ron@jettech.com.tw", subject, content)
        dr.Close()
        connsap.Close()
        Response.Redirect("issuedwoinfo.aspx?smid=molist&smode=5")
        conn.Close()
    End Sub
    Sub RefundMaterialInfo()
        'Dim KeyName As String '要的关键字,实际就是数据表的主键.需要事先在GridView1的DataKeyNames中设置
        Dim j As Integer
        Dim gvr As GridViewRow
        Dim docnum, itemcode, itemname, ownername As String
        Dim req_set As Integer
        Dim build_date, upd_date, req_date As Date
        CommUtil.InitLocalSQLConnection(conn)
        build_date = FormatDateTime(Now(), DateFormat.ShortDate)
        upd_date = FormatDateTime(Now(), DateFormat.ShortDate)
        ownername = Session("s_name")
        req_date = FormatDateTime(Now(), DateFormat.ShortDate)
        For Each gvr In gv1.Rows
            j = gvr.RowIndex
            If (CInt(gv1.Rows(j).Cells(10).Text) > 0) Then
                If CType(gvr.FindControl("cChk_" & j), CheckBox).Checked = True Then
                    docnum = gv1.Rows(j).Cells(0).Text
                    itemcode = gv1.Rows(j).Cells(1).Text
                    itemname = gv1.Rows(j).Cells(2).Text
                    req_set = 0 - CType(gvr.FindControl("tTxt_" & j), TextBox).Text
                    SqlCmd = "Insert Into omri (modocnum,docnum,wsn,itemcode,itemname,req_set,build_date,req_date,ownername,comm,finish_amount,stat,finish,upd_date) " &
                            "Values ('" & modocnum & "','" & docnum & "','" & wsn & "','" & itemcode & "','" & itemname & "'," & req_set & ",'" &
                            build_date & "','" & req_date & "','" & ownername & "',' '," & req_set & ",20,1,'" & FormatDateTime(Now(), DateFormat.ShortDate) & "')"
                    CommUtil.SqlExecute("ins", SqlCmd, conn)
                    'CommUtil.ShowMsg(Me,CType(gvr.FindControl("tTxt_" & j), TextBox).Text)
                    'i = gvr.RowIndex 'GridView行索引   
                    'KeyName = gv1.DataKeys(i).Value
                    '...根据KeyName想做什么做什么吧.   
                End If
            End If
        Next
        Response.Redirect("wolist.aspx?modocnum=" & modocnum & "&wsn=" & wsn & "&smid=molist&smode=0&preact=退料")
        conn.Close()
    End Sub

    Sub CreateWo()
        Dim j As Integer
        Dim gvr As GridViewRow
        Dim vWo As SAPbobsCOM.ProductionOrders
        Dim duedate, remarks As String
        Dim createwo As Boolean
        createwo = True
        'Dim qty As Integer
        '以下取mo資料
        'InitSAPSQLConnection()
        SqlCmd = "Select OWOR.ItemCode, OITM.ItemName, OWOR.DueDate, OWOR.PlannedQty, OWOR.comments FROM OWOR " & _
        "INNER JOIN OITM ON OWOR.ItemCode = OITM.ItemCode " & _
        "WHERE OWOR.DocNum = '" & modocnum & "'"
        'myCommand = New SqlCommand(SqlCmd, connsap)
        'dr = myCommand.ExecuteReader()
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        duedate = dr(2)
        remarks = dr(4)
        dr.Close()
        connsap.Close()
        ret = InitSAPConnection(Session("usingserver"), Session("usingdb"))
        If (ret <> 0) Then
            CommUtil.ShowMsg(Me,"連線失敗")
            Exit Sub
        End If
        For Each gvr In gv1.Rows
            j = gvr.RowIndex
            If (gv1.Rows(j).Cells(0).Text = "") Then
                If CType(gvr.FindControl("cChk_" & j), CheckBox).Checked = True Then
                    '以下create子工單
                    vWo = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    vWo.ItemNo = gv1.Rows(j).Cells(1).Text
                    vWo.PlannedQuantity = gv1.Rows(j).Cells(4).Text
                    vWo.DueDate = duedate
                    vWo.Warehouse = gv1.Rows(j).Cells(5).Text
                    vWo.Remarks = remarks
                    vWo.UserFields.Fields.Item("U_F16").Value = modocnum
                    If (0 <> vWo.Add()) Then
                        CommUtil.ShowMsg(Me, "Failed to add WorkOrder item(可check看交貨日期是否小於今日日期): " & vWo.ItemNo) 'If failed, show a message
                        vWo = Nothing
                        createwo = False
                    Else
                        vWo = Nothing
                        'CommUtil.ShowMsg(Me,"SAP子工單產生成功")
                    End If
                End If
            End If
            'j = j + 1
        Next
        If (createwo) Then
            CommUtil.ShowMsg(Me, "子工單產生完成")
            Response.Redirect("wolist.aspx?modocnum=" & modocnum & "&wsn=" & wsn & "&smid=molist&smode=0")
        Else
            CommUtil.ShowMsg(Me, "有某些子工單產生失敗 , 請check")
        End If
        IssueWoCheck.Checked = False
        CloseSAPConnection()
    End Sub

    Protected Sub IssueWoCheck_CheckedChanged(sender As Object, e As EventArgs) Handles IssueWoCheck.CheckedChanged
        If (IssueWoCheck.Checked = True And DDLWoFun.SelectedIndex = 2) Then
            If (ExecuteBtn.Enabled = False) Then
                ExecuteBtn.Enabled = True
            End If
        ElseIf (IssueWoCheck.Checked = False And DDLWoFun.SelectedIndex = 2) Then
            If (ExecuteBtn.Enabled = True) Then
                ExecuteBtn.Enabled = False
            End If
        End If
    End Sub

    Protected Sub DDLWoFun_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DDLWoFun.SelectedIndexChanged
        If (DDLWoFun.SelectedValue = "開工單") Then
            If (CommUtil.DisableObjectByPermission(ExecuteBtn, permsmf201, "m")) Then
                If (IssueWoCheck.Checked) Then
                    ExecuteBtn.Enabled = True
                Else
                    ExecuteBtn.Enabled = False
                    CommUtil.ShowMsg(Me, "若要開工單,開工單check要打開")
                End If
                'DDLAlter.Enabled = True
                IssueWoCheck.Enabled = True
            Else
                CommUtil.ShowMsg(Me, "無工單操作權限")
                DDLWoFun.SelectedIndex = 0
                IssueWoCheck.Checked = False
                DDLAlter.Enabled = False
                IssueWoCheck.Enabled = False
            End If
            reqdate_text.Enabled = False
            DDLAlter.Items.Clear()
            DDLAlter.Items.Add("請選取替代方法")
            DDLAlter.Items.Add("一般替代")
            DDLAlter.Items.Add("必替替代")
            DDLAlter.Items.Add("不替代")
            DDLAlter.SelectedIndex = 3
            IssuedAutoCheck.Enabled = False
        ElseIf (DDLWoFun.SelectedValue = "發料通知") Then
            If (CommUtil.DisableObjectByPermission(ExecuteBtn, permsmf202, "m")) Then
                ExecuteBtn.Enabled = True
                reqdate_text.Enabled = True
            Else
                CommUtil.ShowMsg(Me, "無領料通知權限")
                DDLWoFun.SelectedIndex = 0
                ExecuteBtn.Enabled = False
                reqdate_text.Enabled = False
            End If
            DDLAlter.Enabled = False
            IssueWoCheck.Enabled = False
            DDLAlter.Items.Clear()
            DDLAlter.Items.Add("不須選取")
            IssuedAutoCheck.Enabled = True
        ElseIf (DDLWoFun.SelectedValue = "模組退料") Then
            ExecuteBtn.Enabled = True
            reqdate_text.Enabled = False
            IssueWoCheck.Enabled = False
            IssuedAutoCheck.Enabled = False
        Else
            ExecuteBtn.Enabled = False
            reqdate_text.Enabled = False
            IssueWoCheck.Enabled = False
            IssuedAutoCheck.Enabled = False
        End If
        'GetWoList()
    End Sub
End Class