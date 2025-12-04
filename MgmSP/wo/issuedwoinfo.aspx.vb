Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Partial Public Class issuedwoinfo
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public connsap As New SqlConnection
    Public myCommand As SqlCommand
    Public SqlCmd As String
    Public modocnum As String
    Public ds As New DataSet
    Public dr As SqlDataReader
    Public oCompany As New SAPbobsCOM.Company
    Public page1 As Integer
    Public ret As Long
    Public permsmf203, permsmf202 As String
    Public ScriptManager1 As New ScriptManager
    Public TxtReqDate, TxtModualSelect, TxtQty As TextBox
    Public DDLModel As DropDownList
    Public BtnReqGen As Button
    Public LBModual As New ListBox

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' Page.Form.Controls.Add(ScriptManager1)
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        permsmf203 = CommUtil.GetAssignRight("mf203", Session("s_id"))
        permsmf202 = CommUtil.GetAssignRight("mf202", Session("s_id"))
        'MsgBox(gv1.PageIndex)
        If (Not IsPostBack) Then
            gv1.PageIndex = Request.QueryString("indexpage")
        End If
        FTCreate()
        IssuedMaterialInfoList()
    End Sub
    Sub FTCreate()
        Dim ce As CalendarExtender
        Dim tRow As New TableRow
        Dim tCell As TableCell
        Dim Labelx As Label
        Dim dde As New DropDownExtender

        tCell = New TableCell
        Labelx = New Label()
        Labelx.ID = "label_1"
        Labelx.Text = "備用模組領料操作:&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        Labelx = New Label()
        Labelx.ID = "label_2"
        Labelx.Text = "需求日期&nbsp"
        tCell.Controls.Add(Labelx)
        TxtReqDate = New TextBox
        TxtReqDate.Width = 85
        TxtReqDate.ID = "txt_reqdate"
        tCell.Controls.Add(TxtReqDate)
        ce = New CalendarExtender
        ce.TargetControlID = TxtReqDate.ID
        ce.ID = "ce_reqdate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        CommUtil.DisableObjectByPermission(TxtReqDate, permsmf202, "n")

        Labelx = New Label()
        Labelx.ID = "label_3"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp機型&nbsp"
        tCell.Controls.Add(Labelx)
        DDLModel = New DropDownList()
        DDLModel.ID = "ddl_model"
        DDLModel.Width = 120
        tCell.Controls.Add(DDLModel)
        CommUtil.DisableObjectByPermission(DDLModel, permsmf202, "n")

        Labelx = New Label()
        Labelx.ID = "label_4"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp發料模組&nbsp"
        tCell.Controls.Add(Labelx)
        LBModual.ID = "lb_modual"
        LBModual.AutoPostBack = True
        LBModual.Rows = 30
        AddHandler LBModual.SelectedIndexChanged, AddressOf LBModual_SelectedIndexChanged
        tCell.Controls.Add(LBModual)
        TxtModualSelect = New TextBox
        TxtModualSelect.ID = "txt_modual"
        TxtModualSelect.Width = 400
        tCell.Controls.Add(TxtModualSelect)
        dde.TargetControlID = TxtModualSelect.ID
        dde.ID = "dde_modual"
        dde.DropDownControlID = LBModual.ID
        tCell.Controls.Add(dde)
        CommUtil.DisableObjectByPermission(TxtModualSelect, permsmf202, "n")
        CommUtil.DisableObjectByPermission(LBModual, permsmf202, "n")

        Labelx = New Label()
        Labelx.ID = "label_qty"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp&nbsp需求數量&nbsp"
        tCell.Controls.Add(Labelx)
        TxtQty = New TextBox
        TxtQty.Width = 30
        TxtQty.ID = "txt_qty"
        tCell.Controls.Add(TxtQty)
        CommUtil.DisableObjectByPermission(TxtQty, permsmf202, "n")

        Labelx = New Label()
        Labelx.ID = "label_req"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        BtnReqGen = New Button()
        BtnReqGen.ID = "btn_reqgen"
        BtnReqGen.Text = "產生需求"
        AddHandler BtnReqGen.Click, AddressOf BtnReqGen_Click
        tCell.Controls.Add(BtnReqGen)
        CommUtil.DisableObjectByPermission(BtnReqGen, permsmf202, "n")
        tRow.Cells.Add(tCell)

        FT.Rows.Add(tRow)
        'gen model and modual list
        SqlCmd = "SELECT T0.u_model,T0.u_mdesc,T0.u_mtype " &
                    "FROM dbo.[@UMMD] T0 where T0.u_mtype='SPI' or T0.u_mtype='AOI' or T0.u_mtype='3DAOI' " &
                    "order by T0.u_model,T0.u_mcode"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        DDLModel.Items.Clear()
        DDLModel.Items.Add("選機型")
        If (dr.HasRows) Then
            Do While (dr.Read())
                DDLModel.Items.Add(dr(0))
            Loop
        End If
        dr.Close()
        connsap.Close()
        SqlCmd = "SELECT T0.u_cspec " &
                    "FROM dbo.[@UMST] T0 order by T0.u_assignmodel"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        LBModual.Items.Clear()
        If (dr.HasRows) Then
            Do While (dr.Read())
                LBModual.Items.Add(dr(0))
            Loop
        End If
        dr.Close()
        connsap.Close()
    End Sub
    Protected Sub BtnReqGen_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim itemname, bdate, reqdate, oname As String
        Dim req_set As Integer

        If (DDLModel.SelectedIndex = 0) Then
            CommUtil.ShowMsg(Me, "機型沒選")
            Exit Sub
        End If
        If (TxtModualSelect.Text = "") Then
            CommUtil.ShowMsg(Me, "模組沒選")
            Exit Sub
        End If
        If (TxtReqDate.Text = "") Then
            CommUtil.ShowMsg(Me, "需求日期空白")
            Exit Sub
        End If
        If (TxtQty.Text = "") Then
            CommUtil.ShowMsg(Me, "需求數量空白")
            Exit Sub
        End If
        oname = Session("s_name")
        itemname = DDLModel.SelectedValue & "_" & TxtModualSelect.Text
        reqdate = TxtReqDate.Text
        bdate = Format(Now(), "yyyy/MM/dd")
        req_set = CInt(TxtQty.Text)
        SqlCmd = "Insert Into omri (itemname,req_set,build_date,req_date,ownername,comm) " &
                "Values ('" & itemname & "'," & req_set & ",'" &
                bdate & "','" & reqdate & "','" & oname & "',' ')"
        CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
        conn.Close()
        Response.Redirect("issuedwoinfo.aspx?smid=molist&smode=5")
    End Sub
    Protected Sub LBModual_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        TxtModualSelect.Text = sender.SelectedValue
    End Sub
    Sub IssuedMaterialInfoList()
        ds.Reset()
        SqlCmd = "SELECT modocnum='NA' ,T0.num,wodocnum='NA' , wsn='NA',cus_name='模組備料',ItemCode='NA', T0.ItemName,T0.req_set,T0.build_date, " &
        "T0.req_date,T0.ownername,T0.Stat,T0.upd_date,T0.comm,T0.prepare_amount,T0.finish_amount,T0.finish FROM omri T0 " &
        "where wsn='' " &
        "ORDER BY T0.finish,T0.wsn,T0.req_date"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        SqlCmd = "SELECT T1.docnum As modocnum ,T0.num,T0.docnum as wodocnum , T0.wsn,T1.cus_name,T0.ItemCode, T0.ItemName,T0.req_set,T0.build_date, " &
        "T0.req_date,T0.ownername,T0.Stat,T0.upd_date,T0.comm,T0.prepare_amount,T0.finish_amount,T0.finish FROM omri T0 " &
        "INNER JOIN worksn T1 ON T0.wsn = T1.wsn " &
        "ORDER BY T0.finish,T0.wsn,T0.req_date"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        If (ds.Tables(0).Rows.Count <> 0) Then
            If ds.Tables(0).Columns.Contains("issueamount") = False Then
                ds.Tables(0).Columns.Add("issueamount")
            End If
            If ds.Tables(0).Columns.Contains("act") = False Then
                ds.Tables(0).Columns.Add("act")
            End If
            If ds.Tables(0).Columns.Contains("issue_count") = False Then
                ds.Tables(0).Columns.Add("issue_count")
            End If
            If ds.Tables(0).Columns.Contains("rtn") = False Then
                ds.Tables(0).Columns.Add("rtn")
            End If
            If ds.Tables(0).Columns.Contains("del") = False Then
                ds.Tables(0).Columns.Add("del")
            End If
            ds.Tables(0).DefaultView.Sort = "finish,wsn,req_date"
            gv1.DataSource = ds.Tables(0)
            gv1.DataBind()
        Else
            CommUtil.ShowMsg(Me, "無任何資料")
        End If
        'MsgBox(ds.Tables(0).DefaultView(0)("num"))
    End Sub

    Protected Sub gv1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles gv1.PageIndexChanging
        ds.Tables(0).Clear()
        gv1.PageIndex = e.NewPageIndex

        IssuedMaterialInfoList()
    End Sub

    Protected Sub gv1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim btn As Button
        Dim tTxt As TextBox
        Dim cChk As CheckBox
        Dim realindex As Integer
        'Dim ce As CalendarExtender
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            'CommUtil.ShowMsg(Me,e.Row.RowIndex)
            realindex = e.Row.RowIndex + gv1.PageIndex * gv1.PageSize
            If (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("stat") <> 20) Then
                tTxt = New TextBox
                tTxt.ID = "issueamounttxt_" & e.Row.RowIndex
                tTxt.Width = 30
                CommUtil.DisableObjectByPermission(tTxt, permsmf203, "m")
                e.Row.Cells(3).Controls.Add(tTxt)

                btn = New Button
                btn.ID = "actbtn_" & e.Row.RowIndex
                If (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("prepare_amount") = 0) Then
                    btn.Text = "領料準備"
                Else
                    btn.Text = "領料完成"
                    btn.BorderColor = Drawing.Color.LightGreen
                End If
                AddHandler btn.Click, AddressOf btn_Click
                CommUtil.DisableObjectByPermission(btn, permsmf203, "m")
                e.Row.Cells(4).Controls.Add(btn)
            End If
            tTxt = New TextBox
            tTxt.ID = "reqdatetxt_" & e.Row.Cells(0).Text
            tTxt.Width = 70
            tTxt.Text = e.Row.Cells(11).Text
            CommUtil.DisableObjectByPermission(tTxt, permsmf203, "m")
            tTxt.AutoPostBack = True
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            e.Row.Cells(11).Controls.Add(tTxt)
            'ce = New CalendarExtender
            'ce.TargetControlID = tTxt.ID
            'ce.ID = "ce_reqdate"
            'ce.Format = "yyyy/MM/dd"
            'e.Row.Cells(11).Controls.Add(ce)

            tTxt = New TextBox
            tTxt.ID = "commtxt_" & e.Row.Cells(0).Text
            tTxt.Width = 100
            tTxt.Text = e.Row.Cells(16).Text
            CommUtil.DisableObjectByPermission(tTxt, permsmf203, "m")
            tTxt.AutoPostBack = True
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            e.Row.Cells(16).Controls.Add(tTxt)

            'e.Row.Cells(13).ForeColor = Drawing.Color.Blue
            e.Row.Cells(13).Text = ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("prepare_amount") & "/" & ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("req_set") - ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("finish_amount")
            'e.Row.Cells(13).ForeColor = Drawing.Color.Black
            'e.Row.Cells(13).Text = e.Row.Cells(13).Text & "/"
            'e.Row.Cells(13).ForeColor = Drawing.Color.Red
            'e.Row.Cells(13).Text = e.Row.Cells(13).Text & ds.Tables(0).Rows(e.Row.RowIndex)("req_set") - ds.Tables(0).Rows(e.Row.RowIndex)("finish_amount")

            If (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("prepare_amount") <> 0) Then
                cChk = New CheckBox
                cChk.ID = "rtnchk_" & e.Row.RowIndex
                cChk.AutoPostBack = True
                AddHandler cChk.CheckedChanged, AddressOf RtnCheck_CheckedChanged
                CommUtil.DisableObjectByPermission(cChk, permsmf203, "m")
                e.Row.Cells(17).Controls.Add(cChk)
            End If

            If (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("finish_amount") = 0 And ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("prepare_amount") = 0) Then
                cChk = New CheckBox
                cChk.ID = "delchk_" & e.Row.RowIndex
                cChk.AutoPostBack = True
                AddHandler cChk.CheckedChanged, AddressOf DelCheck_CheckedChanged
                CommUtil.DisableObjectByPermission(cChk, permsmf203, "d")
                e.Row.Cells(18).Controls.Add(cChk)
            End If

            If (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("prepare_amount") = 0 And ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("stat") = 0) Then
                e.Row.Cells(14).Text = "未領料"
                e.Row.Cells(14).BackColor = Drawing.Color.White
            ElseIf (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("prepare_amount") > 0) Then
                e.Row.Cells(14).Text = "備料中"
                e.Row.Cells(14).BackColor = Drawing.Color.Yellow
            ElseIf (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("stat") = 20) Then
                e.Row.Cells(14).Text = "已完成"
                e.Row.Cells(14).BackColor = Drawing.Color.LightGreen
            ElseIf (ds.Tables(0).DefaultView(e.Row.RowIndex + gv1.PageIndex * gv1.PageSize)("stat") = 10) Then
                e.Row.Cells(14).Text = "部份領"
                e.Row.Cells(14).BackColor = Drawing.Color.MediumSeaGreen
            End If
            e.Row.Cells(14).Font.Size = 10
            If (e.Row.Cells(15).Text = "1900/1/1") Then
                e.Row.Cells(15).Text = ""
            End If
        End If

    End Sub

    Protected Sub DelCheck_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim cb As CheckBox = sender
        Dim index As Integer
        If (cb.Checked) Then
            index = CInt(Split(cb.ID, "_")(1))
            CType(gv1.Rows(index).FindControl("actbtn_" & index), Button).Text = "刪除"
            CType(gv1.Rows(index).FindControl("actbtn_" & index), Button).Enabled = True
            CType(gv1.Rows(index).FindControl("actbtn_" & index), Button).BorderColor = Drawing.Color.Red
            ViewState("editindex") = index
        End If
    End Sub

    Protected Sub RtnCheck_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim cb As CheckBox = sender
        Dim index As Integer
        If (cb.Checked) Then
            index = CInt(Split(cb.ID, "_")(1))
            CType(gv1.Rows(index).FindControl("actbtn_" & index), Button).Text = "退領"
            CType(gv1.Rows(index).FindControl("actbtn_" & index), Button).BorderColor = Drawing.Color.Red
            ViewState("editindex") = index
        End If
    End Sub

    Protected Sub btn_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bt As Button = sender
        Dim index, realindex As Integer
        Dim str As String
        Dim issuecount As Integer
        Dim stat As Integer
        Dim sqltype As String
        Dim reflash As Boolean
        reflash = True
        sqltype = ""
        index = CInt(Split(bt.ID, "_")(1))
        realindex = index + gv1.PageIndex * gv1.PageSize
        SqlCmd = ""
        If (ds.Tables(0).DefaultView(realindex)("prepare_amount") = 0 And bt.Text = "領料準備") Then '準備發料
            str = CType(gv1.Rows(index).FindControl("issueamounttxt_" & index), TextBox).Text
            If (str <> "") Then
                issuecount = CInt(str)
                If (issuecount <= (ds.Tables(0).DefaultView(realindex)("req_set") - ds.Tables(0).DefaultView(realindex)("finish_amount"))) Then
                    SqlCmd = "update omri Set prepare_amount=" & issuecount & " where num=" & gv1.Rows(index).Cells(0).Text
                    sqltype = "upd"
                Else
                    reflash = False
                    CommUtil.ShowMsg(Me, "領料數量已大於可領數量")
                End If
            Else
                reflash = False
                CommUtil.ShowMsg(Me, "需填寫發料數量")
            End If
        ElseIf (ds.Tables(0).DefaultView(realindex)("prepare_amount") > 0 And bt.Text = "領料完成") Then '發料完成
            If ((ds.Tables(0).DefaultView(realindex)("prepare_amount") + ds.Tables(0).DefaultView(realindex)("finish_amount")) = ds.Tables(0).DefaultView(realindex)("req_set")) Then
                stat = 20
                SqlCmd = "update omri Set finish_amount=finish_amount +" & ds.Tables(0).DefaultView(realindex)("prepare_amount") & " ,stat=" & stat & " ," &
                         "prepare_amount=0,finish=1,upd_date='" & FormatDateTime(Now(), DateFormat.ShortDate) & "'where num=" & gv1.Rows(index).Cells(0).Text
            Else
                stat = 10
                SqlCmd = "update omri Set finish_amount=finish_amount +" & ds.Tables(0).DefaultView(realindex)("prepare_amount") & " ,stat=" & stat & " ," &
                         "prepare_amount=0,upd_date='" & FormatDateTime(Now(), DateFormat.ShortDate) & "'where num=" & gv1.Rows(index).Cells(0).Text
            End If
            sqltype = "upd"
        ElseIf (bt.Text = "刪除") Then '刪除動作
            SqlCmd = "delete from omri where num=" & gv1.Rows(index).Cells(0).Text
            sqltype = "del"
        ElseIf (bt.Text = "退領") Then '退領動作
            If (ds.Tables(0).DefaultView(realindex)("prepare_amount") = 0) Then
                SqlCmd = "update omri set prepare_amount=0 ,stat=0 where num=" & gv1.Rows(index).Cells(0).Text
            Else
                SqlCmd = "update omri set prepare_amount=0 where num=" & gv1.Rows(index).Cells(0).Text
            End If
            sqltype = "upd"
        End If
        If (SqlCmd <> "") Then
            CommUtil.SqlLocalExecute(sqltype, SqlCmd, conn)
            conn.Close()
        End If
        If (reflash) Then
            Response.Redirect("issuedwoinfo.aspx?smid=molist&smode=5&indexpage=" & gv1.PageIndex)
        End If
    End Sub

    Sub tTxt_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim Txtx As TextBox = sender
        Dim inum As Long
        Dim txtkw As String
        inum = CInt(Split(Txtx.ID, "_")(1))
        txtkw = Split(Txtx.ID, "_")(0)
        If (txtkw = "reqdatetxt") Then
            SqlCmd = "update dbo.[omri] set req_date= '" & Txtx.Text & "' " &
                 "where num=" & inum
            CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
            conn.Close()
        ElseIf (txtkw = "commtxt") Then
            SqlCmd = "update dbo.[omri] set comm= '" & Txtx.Text & "' " &
                 "where num=" & inum
            CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
            conn.Close()
        End If
    End Sub
End Class