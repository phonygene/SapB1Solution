Imports System.Data
Imports System.Data.SqlClient
Public Class cncitems
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public conn2, connsap As New SqlConnection
    Dim SqlCmd As String
    Public dr As SqlDataReader
    Public oCompany As New SAPbobsCOM.Company
    Public ret As Long
    Public ds As New DataSet
    Public sappo As Long
    Public act As String
    Public permsmf204 As String
    Public num As Long
    Public indexpage As Integer
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
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        permsmf204 = CommUtil.GetAssignRight("mf204", Session("s_id"))
        sappo = Request.QueryString("sappo")
        num = Request.QueryString("num")
        indexpage = Request.QueryString("indexpage")

        If (Not IsPostBack) Then 'Hyper不是 postback , 但button 是 , 故此處可以去除這些action
            act = Request.QueryString("act")
            If (act = "updstat" Or act = "updrawsta") Then
                UpdateStatus()
            ElseIf (act = "del") Then
                DeleteItem()
            End If
            'CommUtil.ShowMsg(Me,"isnot")
        Else
            'CommUtil.ShowMsg(Me,"is")
        End If
        CreateHyperMenu()
        If (num <> 0) Then
            ShowCncItemAddAndInfo()
        End If
        ShowGridView()
    End Sub
    Sub UpdateStatus()
        Dim num As Long
        Dim status As Integer
        num = Request.QueryString("num")
        If (act = "updsta") Then
            status = Request.QueryString("postatus")
            If (status = 0) Then
                status = 10
            ElseIf (status = 10) Then
                status = 90
            ElseIf (status = 90) Then
                status = 0
            End If
            SqlCmd = "update dbo.[cnc1] set stat= " & status & " where num=" & num
        ElseIf (act = "updrawsta") Then
            status = Request.QueryString("rawstatus")
            If (status = 0) Then
                status = 10
            ElseIf (status = 10) Then
                status = 20
            ElseIf (status = 20) Then
                status = 90
            ElseIf (status = 90) Then
                status = 0
            End If
            SqlCmd = "update dbo.[cnc1] set rawstatus= " & status & " where num=" & num
        End If
        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub
    Sub DeleteItem()
        Dim num As Long
        num = Request.QueryString("num")
        'CommUtil.InitLocalSQLConnection(conn)
        SqlCmd = "delete from dbo.[cnc1]  where num=" & num
        'myCommand = New SqlCommand(SqlCmd, conn)
        'count = myCommand.ExecuteNonQuery()
        'If (count = 0) Then
        ' CommUtil.ShowMsg(Me,"刪除失敗")
        'End If
        CommUtil.SqlLocalExecute("del", SqlCmd, conn)
        conn.Close()
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
        Hyper.Width = 150
        Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
        Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
        Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
        j = j + 1
        Hyper = New HyperLink()
        Hyper.ID = "wstatus"
        Hyper.Text = "回加工總表"
        Hyper.NavigateUrl = "cncmain.aspx?act=showlist&smid=molist&smode=7&indexpage=" & indexpage
        Hyper.BackColor = Drawing.Color.Aqua
        Hyper.Font.Underline = False
        Hyper.Width = 150
        Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
        Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
        Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
        Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
    End Sub
    Sub ShowCncItemAddAndInfo()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim Labelx As Label
        Dim Txtx As TextBox
        Dim Btnx As Button
        'Dim dr As SqlDataReader
        Dim total, notf As Integer
        Dim infostr As String
        infostr = ""
        If (num <> 0) Then
            SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & sappo
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            dr.Read()
            total = dr(0)
            dr.Close()
            conn.Close()
            SqlCmd = "Select count(*) from dbo.[cnc1] T0 where T0.sappo=" & sappo & " and T0.stat<>90"
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            dr.Read()
            notf = dr(0)
            dr.Close()
            conn.Close()
            SqlCmd = "Select num,cdate,comm,vender from dbo.[ocnc] T0 where T0.sappo=" & sappo
            dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
            dr.Read()
            infostr = "加工說明:" & dr(2) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &
                  "PO單號:" & sappo & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建立日期:" & dr(1) &
                  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;廠商:" & dr(3) &
                  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;狀態(未完/總數):" & notf & "/" & total
            dr.Close()
            conn.Close()
        Else
            SqlCmd = "select count(*) " &
                "from dbo.OPOR T0 inner join POR1 T1 on T0.docentry=T1.docentry where T0.docnum=" & sappo
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            total = dr(0)
            dr.Close()
            connsap.Close()
            SqlCmd = "select count(*) " &
                "from dbo.OPOR T0 inner join POR1 T1 on T0.docentry=T1.docentry where T0.docnum=" & sappo &
                " and T1.opencreqty<>0"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            dr.Read()
            notf = dr(0)
            dr.Close()
            connsap.Close()
            SqlCmd = "select T0.docdate As cdate,T0.comments As comm,T1.aliasname As vender " &
            "from dbo.OPOR T0 inner join OCRD T1 on T0.cardcode=T1.cardcode where T1.QryGroup10='Y' " &
            "and T0.docnum=" & sappo
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                infostr = "委外說明:" & dr(1) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &
                  "PO單號:" & sappo & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;建立日期:" & dr(0) &
                  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;廠商:" & dr(2) &
                  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;狀態(未完/總數):" & notf & "/" & total
                dr.Close()
            Else
                CommUtil.ShowMsg(Me, "找不到採購單的廠商, 請check此採購單之廠商是否在業務夥伴中gropu10有選取")
            End If
            connsap.Close()
        End If
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Wrap = False
        tCell.Text = infostr
        tCell.Font.Bold = True
        tRow.Cells.Add(tCell)
        InfoT.Rows.Add(tRow)

        'add Table
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Wrap = False
        tCell.Font.Bold = True

        Labelx = New Label()
        Labelx.ID = "label_itemcode"
        Labelx.Text = "料號:"
        tCell.Controls.Add(Labelx)
        Txtx = New TextBox()
        Txtx.ID = "txt_itemcode"
        ViewState("itemcode") = Txtx.ID
        Txtx.Width = 100
        CommUtil.DisableObjectByPermission(Txtx, permsmf204, "n")
        tCell.Controls.Add(Txtx)
        '----------------------------
        Labelx = New Label()
        Labelx.ID = "label_quantity"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp數量:"
        tCell.Controls.Add(Labelx)
        Txtx = New TextBox()
        Txtx.ID = "txt_quantity"
        ViewState("quantity") = Txtx.ID
        Txtx.Width = 30
        CommUtil.DisableObjectByPermission(Txtx, permsmf204, "n")
        tCell.Controls.Add(Txtx)
        '-------------------------------
        Labelx = New Label()
        Labelx.ID = "label_add"
        Labelx.Text = "&nbsp&nbsp&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        Btnx = New Button()
        Btnx.ID = "btn_add"
        CommUtil.DisableObjectByPermission(Btnx, permsmf204, "n")
        Btnx.Text = "新增"
        AddHandler Btnx.Click, AddressOf Btnx_Click
        tCell.Controls.Add(Btnx)

        tRow.Cells.Add(tCell)
        CncAddItemT.Rows.Add(tRow)
    End Sub

    Sub ShowGridView()
        'CommUtil.InitLocalSQLConnection(conn)
        If (num <> 0) Then
            SqlCmd = "Select T0.itemcode,T0.stat,T0.quantity,T0.f_amount,T0.num,T0.sappo,T0.rawstatus " &
                 "from dbo.cnc1 T0 where T0.sappo=" & sappo
            ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        Else
            SqlCmd = "select T1.Itemcode,stat=0,T1.quantity,f_amount=0,num=0,T0.docnum As sappo " &
            "from dbo.OPOR T0 inner join POR1 T1 on T0.docentry=T1.docentry where T0.docnum=" & sappo
            ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, conn)
            ds.Tables(0).Columns.Add("rawstatus")
        End If
        conn.Close()
        ds.Tables(0).Columns.Add("no")
        ds.Tables(0).Columns.Add("unitno")
        ds.Tables(0).Columns.Add("itemname")
        ds.Tables(0).Columns.Add("rawspec")
        ds.Tables(0).Columns.Add("notin")
        ds.Tables(0).Columns.Add("whours")
        ds.Tables(0).Columns.Add("action")

        gv1.DataSource = ds.Tables(0)
        gv1.DataBind()
    End Sub
    Protected Sub gv1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gv1.RowDataBound
        'Dim btn As Button
        Dim tTxt As TextBox
        Dim Hyper As HyperLink
        Dim unitno, whours As Integer
        Dim itemname, rawspec As String
        Dim strA() As String
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            e.Row.Cells(0).Text = e.Row.RowIndex + 1
            SqlCmd = "select T0.itemname,IsNull(T0.U_F5,0),IsNull(T0.U_F6,0),IsNull(T0.UserText,'') " &
                    "from dbo.OITM T0 where T0.itemcode='" & e.Row.Cells(2).Text & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                itemname = dr(0)
                whours = dr(1)
                unitno = dr(2)
                If (dr(3) <> "") Then
                    strA = Split(dr(3), vbCr)
                    If (UBound(strA) = 1) Then 'UBOUND: 查詢陣列有幾個 1:表示2個
                        rawspec = strA(1)
                    Else
                        rawspec = ""
                    End If
                Else
                    rawspec = ""
                End If
            Else
                CommUtil.ShowMsg(Me, "在SAP找不到此料號--" & e.Row.Cells(2).Text)
                connsap.Close()
                Exit Sub
            End If
            dr.Close()
            connsap.Close()
            e.Row.Cells(1).Text = unitno
            e.Row.Cells(3).Text = itemname
            e.Row.Cells(4).Text = rawspec
            '入庫數量
            SqlCmd = "select T1.quantity,T1.opencreqty " &
                    "from dbo.OPOR T0 inner join POR1 T1 on T0.docentry=T1.docentry where T0.docnum=" &
                    ds.Tables(0).Rows(e.Row.RowIndex)("sappo") & "and T1.itemcode='" & ds.Tables(0).Rows(e.Row.RowIndex)("itemcode") & "'"
            dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
            If (dr.HasRows) Then
                dr.Read()
                e.Row.Cells(8).Text = CInt(CInt(e.Row.Cells(6).Text) - dr(0))
            Else
                e.Row.Cells(8).Text = CInt(e.Row.Cells(6).Text)
            End If
            dr.Close()
            connsap.Close()
            If (CInt(e.Row.Cells(6).Text) = CInt(e.Row.Cells(8).Text)) Then
                e.Row.Cells(8).BackColor = Drawing.Color.LightGreen
            ElseIf (CInt(e.Row.Cells(8).Text) > 0) Then
                e.Row.Cells(8).BackColor = Drawing.Color.Yellow
            End If
            e.Row.Cells(6).Text = CInt(e.Row.Cells(6).Text)
            '以上是共有

            If (num <> 0) Then '自製
                tTxt = New TextBox
                tTxt.ID = "txtfamount_" & e.Row.RowIndex
                tTxt.Width = 30
                tTxt.AutoPostBack = True
                tTxt.Text = e.Row.Cells(7).Text
                CommUtil.DisableObjectByPermission(tTxt, permsmf204, "m")
                AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
                e.Row.Cells(7).Controls.Add(tTxt)

                '總工時
                tTxt = New TextBox
                tTxt.ID = "txtwhours_" & e.Row.RowIndex
                tTxt.Width = 30
                tTxt.Text = whours
                tTxt.AutoPostBack = True
                CommUtil.DisableObjectByPermission(tTxt, permsmf204, "m")
                AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
                e.Row.Cells(10).Controls.Add(tTxt)
                '狀態
                Hyper = New HyperLink()
                If (ds.Tables(0).Rows(e.Row.RowIndex)("stat") = 0) Then '用e.Row.Cells(XX)判斷也行 , 但為了之後如換欄位可不用改,故用ds
                    Hyper.Text = "未開工"
                    e.Row.Cells(5).BackColor = Drawing.Color.White
                ElseIf (ds.Tables(0).Rows(e.Row.RowIndex)("stat") = 10) Then
                    Hyper.Text = "進行中"
                    e.Row.Cells(5).BackColor = Drawing.Color.Yellow
                ElseIf (ds.Tables(0).Rows(e.Row.RowIndex)("stat") = 90) Then
                    Hyper.Text = "已完工"
                    e.Row.Cells(5).BackColor = Drawing.Color.LightGreen
                End If
                Hyper.Font.Underline = False
                CommUtil.DisableObjectByPermission(Hyper, permsmf204, "m")
                Hyper.ID = "hyperstat_" & e.Row.RowIndex
                Hyper.NavigateUrl = "cncitems.aspx?act=updstat&num=" & ds.Tables(0).Rows(e.Row.RowIndex)("num") &
                                "&postatus=" & ds.Tables(0).Rows(e.Row.RowIndex)("stat") & "&sappo=" & sappo & "&indexpage=" & indexpage
                Hyper.Enabled = False '由其填入數量決定狀態,故先disable操作
                e.Row.Cells(5).Controls.Add(Hyper)

                Hyper = New HyperLink()
                If (e.Row.Cells(9).Text = 0) Then
                    Hyper.Text = "未採購"
                    e.Row.Cells(9).BackColor = Drawing.Color.White
                ElseIf (e.Row.Cells(9).Text = 10) Then
                    Hyper.Text = "已採購"
                    e.Row.Cells(9).BackColor = Drawing.Color.Yellow
                ElseIf (e.Row.Cells(9).Text = 20) Then
                    Hyper.Text = "已來料"
                    e.Row.Cells(9).BackColor = Drawing.Color.MediumSeaGreen
                ElseIf (e.Row.Cells(9).Text = 90) Then
                    Hyper.Text = "料確認"
                    e.Row.Cells(9).BackColor = Drawing.Color.LightGreen
                End If
                Hyper.Font.Underline = False
                Hyper.ID = "hyper_rawsta_" & ds.Tables(0).Rows(e.Row.RowIndex)("num")
                Hyper.NavigateUrl = "cncitems.aspx?act=updrawsta&num=" & ds.Tables(0).Rows(e.Row.RowIndex)("num") &
                                "&rawstatus=" & ds.Tables(0).Rows(e.Row.RowIndex)("rawstatus") & "&sappo=" & sappo & "&indexpage=" & indexpage
                CommUtil.DisableObjectByPermission(Hyper, permsmf204, "m")
                e.Row.Cells(9).Controls.Add(Hyper)

                '動作(刪除)
                Hyper = New HyperLink()
                Hyper.Text = "刪除"
                Hyper.Font.Underline = False
                Hyper.ID = "hyperdel_" & e.Row.RowIndex
                CommUtil.DisableObjectByPermission(Hyper, permsmf204, "d")
                Hyper.NavigateUrl = "cncitems.aspx?act=del&num=" & ds.Tables(0).Rows(e.Row.RowIndex)("num") & "&sappo=" & sappo & "&indexpage=" & indexpage
                e.Row.Cells(11).Controls.Add(Hyper)
            Else '外發
                e.Row.Cells(5).Text = "NA"
                e.Row.Cells(7).Text = "NA"
                e.Row.Cells(9).Text = "NA"
                If (whours <> 0) Then
                    e.Row.Cells(10).Text = whours
                End If
            End If
        End If
    End Sub
    Protected Sub Btnx_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim itemcode, itemname As String
        Dim Txtx As TextBox
        Dim quantity As Integer
        'CommUtil.ShowMsg(Me,"but")
        Txtx = CncAddItemT.FindControl(ViewState("itemcode"))
        If (Txtx.Text = "") Then
            CommUtil.ShowMsg(Me,"沒料號")
            Exit Sub
        End If
        itemcode = Txtx.Text

        Txtx = CncAddItemT.FindControl(ViewState("quantity"))
        If (Txtx.Text = "") Then
            CommUtil.ShowMsg(Me,"沒數量")
            Exit Sub
        End If
        quantity = CInt(Txtx.Text)
        'CommUtil.InitSAPSQLConnection(connsap)
        SqlCmd = "select T0.itemname " &
        "from dbo.OITM T0 where T0.itemcode='" & itemcode & "'"
        'myCommand = New SqlCommand(SqlCmd, connsap)
        'dr = myCommand.ExecuteReader()
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            itemname = dr(0)
        Else
            CommUtil.ShowMsg(Me,"在SAP找不到此料號--" & itemcode)
            connsap.Close()
            Exit Sub
        End If
        dr.Close()
        connsap.Close()

        'CommUtil.InitLocalSQLConnection(conn)
        SqlCmd = "Insert into dbo.[cnc1] (sappo,itemcode,quantity) " &
                                    "values(" & sappo & ",'" & itemcode & "'," & quantity & ")"
        'myCommand = New SqlCommand(SqlCmd, conn)
        'count = myCommand.ExecuteNonQuery()
        'If (count = 0) Then
        'CommUtil.ShowMsg(Me,"新增失敗")
        'End If
        CommUtil.SqlLocalExecute("ins", SqlCmd, conn)
        conn.Close()
        Response.Redirect("cncitems.aspx?sappo=" & sappo & "&indexpage=" & indexpage & "&num=" & num)
    End Sub
    Sub tTxt_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim Txtx As TextBox = sender
        Dim inum As Long
        Dim index As Integer
        Dim txtkw As String
        Dim stat As Integer
        Dim ok As Boolean
        ok = True
        stat = 0
        index = CInt(Split(Txtx.ID, "_")(1))
        txtkw = Split(Txtx.ID, "_")(0)
        If (txtkw = "txtfamount") Then
            If (CInt(Txtx.Text) > ds.Tables(0).Rows(index)("quantity")) Then
                CommUtil.ShowMsg(Me,"完成數量不能大於加工數量")
                ok = False
            ElseIf (CInt(Txtx.Text) = ds.Tables(0).Rows(index)("quantity")) Then
                stat = 90
            ElseIf (CInt(Txtx.Text) <> 0) Then
                stat = 10
            End If
            If (ok) Then
                inum = ds.Tables(0).Rows(index)("num")
                'CommUtil.InitLocalSQLConnection(conn)
                SqlCmd = "update dbo.[cnc1] set f_amount= " & CInt(Txtx.Text) & ",stat=" & stat &
                 " where num=" & inum
                'myCommand = New SqlCommand(SqlCmd, conn)
                'count = myCommand.ExecuteNonQuery()
                'If (count = 0) Then
                'CommUtil.ShowMsg(Me,"更新失敗")
                'End If
                CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
            End If
            Response.Redirect("cncitems.aspx?sappo=" & sappo & "&num=" & num)
        ElseIf (txtkw = "txtwhours") Then
            'CommUtil.InitSAPSQLConnection(connsap)
            SqlCmd = "update dbo.[oitm] set U_F5= " & CInt(Txtx.Text) &
                     " where itemcode='" & ds.Tables(0).Rows(index)("itemcode") & "'"
            'myCommand = New SqlCommand(SqlCmd, connsap)
            'count = myCommand.ExecuteNonQuery()
            'If (count = 0) Then
            'CommUtil.ShowMsg(Me,"更新失敗")
            'End If
            CommUtil.SqlSapExecute("upd", SqlCmd, connsap)
            connsap.Close()
        End If
    End Sub
End Class