Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class wostamodify
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public connsap As New SqlConnection
    Public SqlCmd As String
    Public oCompany As New SAPbobsCOM.Company
    Public ret As Long
    Public ds As New DataSet
    Public dr As SqlDataReader
    Public wsn As String
    Public k As Integer
    Public permsmf205 As String
    Public ScriptManager1 As New ScriptManager

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

    Sub WoStatusModify()
        Dim num As Long
        Dim status As String
        Dim dpart, iseq, wseq As Integer
        Dim datestr As String
        Dim sqlresult As Boolean
        datestr = Format(Now(), "yyyy/MM/dd")
        num = Request.QueryString("num")
        status = Request.QueryString("status")
        dpart = Request.QueryString("dpart")
        iseq = Request.QueryString("iseq")
        wseq = Request.QueryString("wseq")
        'InitLocalSQLConnection()
        If (dpart = 5 And iseq = 3) Then
            If (status = "未出貨") Then
                status = "已完工"
                SqlCmd = "update dbo.[worksn] set f_set= f_set+1 " &
                "where wsn='" & wsn & "'"
                'myCommand = New SqlCommand(SqlCmd, conn)
                'count = myCommand.ExecuteNonQuery()
                sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
                If (sqlresult = False) Then
                    CommUtil.ShowMsg(Me, "更新工單完工數量加一失敗")
                Else
                    SqlCmd = "Select T0.f_set,T0.model_set " &
                            "from dbo.[worksn] T0 where T0.wsn='" & wsn & "'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    If (dr(0) = dr(1)) Then
                        dr.Close()
                        conn.Close()
                        SqlCmd = "update dbo.[worksn] set f_stat= 60 " &
                                "where wsn='" & wsn & "'"
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                    Else
                        dr.Close()
                        conn.Close()
                    End If
                End If
            ElseIf (status = "已完工") Then
                status = "已包裝"
            ElseIf (status = "已包裝") Then
                status = "已出貨"
                SqlCmd = "update dbo.[work_records] set status= '" & datestr & "' " &
                "where dpart=" & dpart & " and iseq=" & (iseq + 1) & " and wseq=" & wseq & " and wsn='" & wsn & "'"
                sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
                If (sqlresult = False) Then
                    CommUtil.ShowMsg(Me, "更新出貨日期為今日日期失敗")
                End If
                SqlCmd = "update dbo.[worksn] set ship_set= ship_set+1 " &
                "where wsn='" & wsn & "'"
                sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
                If (sqlresult = False) Then
                    CommUtil.ShowMsg(Me, "更新工單出貨數量加一失敗")
                Else
                    SqlCmd = "Select T0.model_set,T0.ship_set " &
                            "from dbo.[worksn] T0 where T0.wsn='" & wsn & "'"
                    dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                    dr.Read()
                    If (dr(0) = dr(1)) Then
                        dr.Close()
                        conn.Close()
                        SqlCmd = "update dbo.[worksn] set f_stat= 70 " &
                                "where wsn='" & wsn & "'"
                        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                        conn.Close()
                    Else
                        dr.Close()
                        conn.Close()
                    End If
                End If
            ElseIf (status = "已出貨") Then
                status = "未出貨"
                SqlCmd = "update dbo.[work_records] set status= '' " &
                 "where dpart=" & dpart & " and iseq=" & (iseq + 1) & " and wseq=" & wseq & " and wsn='" & wsn & "'"
                sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
                If (sqlresult = False) Then
                    CommUtil.ShowMsg(Me, "更新狀態為空白失敗")
                End If

                SqlCmd = "update dbo.[worksn] set ship_set= ship_set-1,f_set=f_set-1 " &
                "where wsn='" & wsn & "'"
                sqlresult = CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                conn.Close()
                If (sqlresult = False) Then
                    CommUtil.ShowMsg(Me, "更新工單出貨數量減一失敗")
                Else
                    SqlCmd = "update dbo.[worksn] set f_stat= 10 " &
                              "where wsn='" & wsn & "'"
                    CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
                    conn.Close()
                End If
            End If
        Else
            If (status = "未進行") Then
                status = "進行中"
            ElseIf (status = "進行中") Then
                status = "已完成"
            ElseIf (status = "已完成") Then
                status = "未進行"
            End If
        End If

        SqlCmd = "update dbo.[work_records] set status= '" & status & "' " &
                 "where num=" & num
        'myCommand = New SqlCommand(SqlCmd, conn)
        'count = myCommand.ExecuteNonQuery()
        'If (count = 0) Then
        'CommUtil.ShowMsg(Me,"更新失敗")
        'End If
        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        Page.Form.Controls.Add(ScriptManager1)
        'If (Not IsPostBack) Then
        permsmf205 = CommUtil.GetAssignRight("mf205", Session("s_id"))
        Dim mode As String
        k = 0
        wsn = Request.QueryString("wsn")
        mode = Request.QueryString("mode")
        If (mode = "modify") Then
            If (Not IsPostBack) Then 'Hyper 才會執行 , 不然更新出貨日期時出貨數都會加一
                WoStatusModify()
            End If
        End If
        CreateHyperMenu()
        ShowWoStatus()
        'CommUtil.ShowMsg(Me,"end")
        'End If
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
        If (Request.QueryString("source") = "fromwostatus") Then
            j = j + 1
            Hyper = New HyperLink()
            Hyper.ID = "wstatus"
            Hyper.Text = "回工單狀態表"
            Hyper.NavigateUrl = "~/wo/wostatus.aspx?smid=molist&smode=6"
            Hyper.BackColor = Drawing.Color.Aqua
            Hyper.Font.Underline = False
            Hyper.Width = 150
            Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
            Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
            Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
        ElseIf (Request.QueryString("source") = "frommolist") Then
            j = j + 1
            Hyper = New HyperLink()
            Hyper.ID = "molist"
            Hyper.Text = "回機台總表"
            Hyper.NavigateUrl = "~/wo/molist.aspx?smid=molist&smode=1&indexpage=" & Request.QueryString("indexpage")
            Hyper.BackColor = Drawing.Color.Aqua
            Hyper.Font.Underline = False
            Hyper.Width = 150
            Hyper.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='AliceBlue'")
            Hyper.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            Me.HyperMenuT.Rows(0).Cells(j).HorizontalAlign = HorizontalAlign.Center
            Me.HyperMenuT.Rows(0).Cells(j).Controls.Add(Hyper)
        End If
    End Sub
    Sub ShowWoStatus()
        Dim itemcount As Integer
        Dim cus_name, model As String

        'InitLocalSQLConnection()
        SqlCmd = "Select T0.docnum, T0.cus_name ,T0.model , T0.model_set " &
                 "from dbo.[worksn] T0 " &
                 "where T0.wsn='" & wsn & "'"
        'myCommand = New SqlCommand(SqlCmd, conn)
        'dr = myCommand.ExecuteReader()
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        itemcount = dr(3)
        cus_name = dr(1)
        model = dr(2)
        dr.Close()
        conn.Close()
        WoStatusT.BackColor = Drawing.Color.PaleGoldenrod
        CreateHeadTable(WoStatusT, 1, itemcount, cus_name, model)
        CreateMaterialStatusTable(WoStatusT, 5, itemcount)
        CreateAssemblyStatusTable(WoStatusT, 5, itemcount)
        CreateWiringStatusTable(WoStatusT, 5, itemcount)
        CreateAdjustStatusTable(WoStatusT, 5, itemcount)
        CreateQCStatusTable(WoStatusT, 4, itemcount)
    End Sub
    Public Sub CreateHeadTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer, cus_name As String, model As String)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i As Integer
        tRow = New TableRow()
        tRow.BorderWidth = 1
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Wrap = False
        tCell.Text = wsn & " " & cus_name & " " & model
        tCell.ColumnSpan = rcount + 3
        tCell.Font.Bold = True
        tRow.Cells.Add(tCell)
        tTable.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BorderWidth = 1
        For i = 0 To (rcount + 2)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.BackColor = Drawing.Color.DeepSkyBlue
            tCell.Font.Size = 10
            If (i = 0) Then
                tCell.Font.Bold = True
                tCell.Text = "總類"
            ElseIf (i = 1) Then
                tCell.Font.Bold = True
                tCell.Text = "工作項目"
            ElseIf (i > 1 And (i < rcount + 2)) Then
                tCell.Font.Bold = False
                tCell.Text = i - 1
            Else
                tCell.Font.Bold = True
                tCell.Text = "備註"
            End If
            tRow.Cells.Add(tCell)
        Next
        tTable.Rows.Add(tRow)
    End Sub
    Public Sub CreateMaterialStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tTxt As TextBox
        Dim i, j As Integer
        Dim col As Integer
        Dim Hyper As HyperLink
        'InitLocalSQLConnection()
        SqlCmd = "Select T0.num,T0.status,T0.dpart,T0.iseq,T0.wseq " &
                 "from dbo.[work_records] T0 " &
                 "where T0.wsn='" & wsn & "' and T0.dpart=1 order by T0.iseq ,T0.wseq"
        'Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        'da1.Fill(ds)
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Font.Size = 10
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.PaleGoldenrod
                tCell.Text = "料件"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                tCell.Text = "骨架"
                tCell.Font.Bold = True
            Else
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                    tCell.BackColor = Drawing.Color.Yellow
                ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                    tCell.BackColor = Drawing.Color.GreenYellow
                End If
                tCell.Text = ds.Tables(0).Rows(k)("status")
                Hyper = New HyperLink
                Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                Hyper.Font.Underline = False
                CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                tCell.Controls.Add(Hyper)
                Hyper.Dispose()
                k = k + 1
            End If
            tRow.Cells.Add(tCell)
        Next
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Font.Size = 10
        tCell.HorizontalAlign = HorizontalAlign.Center
        'tCell.Wrap = False
        tCell.BackColor = Drawing.Color.PapayaWhip
        tCell.Text = ds.Tables(0).Rows(k)("status")
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
        tTxt.Text = ds.Tables(0).Rows(k)("status")
        tTxt.AutoPostBack = True
        CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
        AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
        tCell.Controls.Add(tTxt)
        tRow.Cells.Add(tCell)

        k = k + 1

        tTable.Rows.Add(tRow)

        row = row - 1
        col = col - 1
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.Font.Size = 10
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i Mod 2) Then
                    tCell.BackColor = Drawing.Color.PapayaWhip
                Else
                    tCell.BackColor = Drawing.Color.Lavender
                End If
                'tCell.BackColor = Drawing.Color.DeepSkyBlue
                If (i = 0 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "鈑金"
                ElseIf (i = 1 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "市購"
                ElseIf (i = 2 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "車件"
                ElseIf (i = 3 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "銑件"
                Else
                    If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                        tCell.BackColor = Drawing.Color.Yellow
                    ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                        tCell.BackColor = Drawing.Color.GreenYellow
                    End If
                    tCell.Wrap = False
                    tCell.Text = ds.Tables(0).Rows(k)("status")
                    Hyper = New HyperLink
                    Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                    Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                    Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                    Hyper.Font.Underline = False
                    CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                    tCell.Controls.Add(Hyper)
                    Hyper.Dispose()
                    k = k + 1
                End If
                tRow.Cells.Add(tCell)
            Next
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Wrap = False
            If (i Mod 2) Then
                tCell.BackColor = Drawing.Color.PapayaWhip
            Else
                tCell.BackColor = Drawing.Color.Lavender
            End If
            tCell.Text = ds.Tables(0).Rows(k)("status")
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
            tTxt.Text = ds.Tables(0).Rows(k)("status")
            tTxt.AutoPostBack = True
            CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            tCell.Controls.Add(tTxt)
            tRow.Cells.Add(tCell)
            k = k + 1
            tTable.Rows.Add(tRow)
        Next
        ' CloseLocalSQLConnection()
    End Sub
    Public Sub CreateAssemblyStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tTxt As TextBox
        Dim i, j As Integer
        Dim col As Integer
        Dim Hyper As HyperLink
        'InitLocalSQLConnection()
        SqlCmd = "Select T0.num,T0.status,T0.dpart,T0.iseq,T0.wseq " &
                 "from dbo.[work_records] T0 " &
                 "where T0.wsn='" & wsn & "' and T0.dpart=2 order by T0.iseq ,T0.wseq"
        'Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        'da1.Fill(ds)
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Font.Size = 10
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.PaleGoldenrod
                tCell.Text = "機構組裝"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                tCell.Text = "發料"
                tCell.Font.Bold = True
            Else
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                    tCell.BackColor = Drawing.Color.Yellow
                ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                    tCell.BackColor = Drawing.Color.GreenYellow
                End If
                tCell.Text = ds.Tables(0).Rows(k)("status")
                Hyper = New HyperLink
                Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                Hyper.Font.Underline = False
                CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                tCell.Controls.Add(Hyper)
                Hyper.Dispose()
                k = k + 1
            End If
            tRow.Cells.Add(tCell)
        Next
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Font.Size = 10
        tCell.HorizontalAlign = HorizontalAlign.Center
        'tCell.Wrap = False
        tCell.BackColor = Drawing.Color.PapayaWhip
        tCell.Text = ds.Tables(0).Rows(k)("status")
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
        tTxt.Text = ds.Tables(0).Rows(k)("status")
        tTxt.AutoPostBack = True
        CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
        AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
        tCell.Controls.Add(tTxt)
        tRow.Cells.Add(tCell)

        k = k + 1

        tTable.Rows.Add(tRow)

        row = row - 1
        col = col - 1
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.Font.Size = 10
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i Mod 2) Then
                    tCell.BackColor = Drawing.Color.PapayaWhip
                Else
                    tCell.BackColor = Drawing.Color.Lavender
                End If
                'tCell.BackColor = Drawing.Color.DeepSkyBlue
                If (i = 0 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "XYZ"
                ElseIf (i = 1 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "軌道"
                ElseIf (i = 2 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "鏡頭"
                ElseIf (i = 3 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "驗證"
                Else
                    If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                        tCell.BackColor = Drawing.Color.Yellow
                    ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                        tCell.BackColor = Drawing.Color.GreenYellow
                    End If
                    tCell.Wrap = False
                    tCell.Text = ds.Tables(0).Rows(k)("status")
                    Hyper = New HyperLink
                    Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                    Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                    Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                    Hyper.Font.Underline = False
                    CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                    tCell.Controls.Add(Hyper)
                    Hyper.Dispose()
                    k = k + 1
                End If
                tRow.Cells.Add(tCell)
            Next
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Wrap = False
            If (i Mod 2) Then
                tCell.BackColor = Drawing.Color.PapayaWhip
            Else
                tCell.BackColor = Drawing.Color.Lavender
            End If
            tCell.Text = ds.Tables(0).Rows(k)("status")
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
            tTxt.Text = ds.Tables(0).Rows(k)("status")
            tTxt.AutoPostBack = True
            CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            tCell.Controls.Add(tTxt)
            tRow.Cells.Add(tCell)
            k = k + 1
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Public Sub CreateWiringStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tTxt As TextBox
        Dim i, j As Integer
        Dim col As Integer
        Dim Hyper As HyperLink
        'InitLocalSQLConnection()
        SqlCmd = "Select T0.num,T0.status,T0.dpart,T0.iseq,T0.wseq " &
                 "from dbo.[work_records] T0 " &
                 "where T0.wsn='" & wsn & "' and T0.dpart=3 order by T0.iseq ,T0.wseq"
        'Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        'da1.Fill(ds)
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Font.Size = 10
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.PaleGoldenrod
                tCell.Text = "機台佈線"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                tCell.Text = "發料"
                tCell.Font.Bold = True
            Else
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                    tCell.BackColor = Drawing.Color.Yellow
                ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                    tCell.BackColor = Drawing.Color.GreenYellow
                End If
                tCell.Text = ds.Tables(0).Rows(k)("status")
                Hyper = New HyperLink
                Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                Hyper.Font.Underline = False
                CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                tCell.Controls.Add(Hyper)
                Hyper.Dispose()
                k = k + 1
            End If
            tRow.Cells.Add(tCell)
        Next
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Font.Size = 10
        tCell.HorizontalAlign = HorizontalAlign.Center
        'tCell.Wrap = False
        tCell.BackColor = Drawing.Color.PapayaWhip
        tCell.Text = ds.Tables(0).Rows(k)("status")
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
        tTxt.Text = ds.Tables(0).Rows(k)("status")
        tTxt.AutoPostBack = True
        CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
        AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
        tCell.Controls.Add(tTxt)
        tRow.Cells.Add(tCell)

        k = k + 1

        tTable.Rows.Add(tRow)

        row = row - 1
        col = col - 1
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.Font.Size = 10
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i Mod 2) Then
                    tCell.BackColor = Drawing.Color.PapayaWhip
                Else
                    tCell.BackColor = Drawing.Color.Lavender
                End If
                'tCell.BackColor = Drawing.Color.DeepSkyBlue
                If (i = 0 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "XYZ"
                ElseIf (i = 1 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "配電盤"
                ElseIf (i = 2 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "外罩"
                ElseIf (i = 3 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "罩外罩"
                Else
                    If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                        tCell.BackColor = Drawing.Color.Yellow
                    ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                        tCell.BackColor = Drawing.Color.GreenYellow
                    End If
                    tCell.Wrap = False
                    tCell.Text = ds.Tables(0).Rows(k)("status")
                    Hyper = New HyperLink
                    Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                    Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                    Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                    Hyper.Font.Underline = False
                    CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                    tCell.Controls.Add(Hyper)
                    Hyper.Dispose()
                    k = k + 1
                End If
                tRow.Cells.Add(tCell)
            Next
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Wrap = False
            If (i Mod 2) Then
                tCell.BackColor = Drawing.Color.PapayaWhip
            Else
                tCell.BackColor = Drawing.Color.Lavender
            End If
            tCell.Text = ds.Tables(0).Rows(k)("status")
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
            tTxt.Text = ds.Tables(0).Rows(k)("status")
            tTxt.AutoPostBack = True
            CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            tCell.Controls.Add(tTxt)
            tRow.Cells.Add(tCell)
            k = k + 1
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Public Sub CreateAdjustStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tTxt As TextBox
        Dim i, j As Integer
        Dim col As Integer
        Dim Hyper As HyperLink
        'InitLocalSQLConnection()
        SqlCmd = "Select T0.num,T0.status,T0.dpart,T0.iseq,T0.wseq " &
                 "from dbo.[work_records] T0 " &
                 "where T0.wsn='" & wsn & "' and T0.dpart=4 order by T0.iseq ,T0.wseq"
        'Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        'da1.Fill(ds)
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Font.Size = 10
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.PaleGoldenrod
                tCell.Text = "機台調適"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                tCell.Text = "正交"
                tCell.Font.Bold = True
            Else
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                    tCell.BackColor = Drawing.Color.Yellow
                ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                    tCell.BackColor = Drawing.Color.GreenYellow
                End If
                tCell.Text = ds.Tables(0).Rows(k)("status")
                Hyper = New HyperLink
                Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                Hyper.Font.Underline = False
                CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                tCell.Controls.Add(Hyper)
                Hyper.Dispose()
                k = k + 1
            End If
            tRow.Cells.Add(tCell)
        Next
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Font.Size = 10
        tCell.HorizontalAlign = HorizontalAlign.Center
        'tCell.Wrap = False
        tCell.BackColor = Drawing.Color.PapayaWhip
        tCell.Text = ds.Tables(0).Rows(k)("status")
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
        tTxt.Text = ds.Tables(0).Rows(k)("status")
        tTxt.AutoPostBack = True
        CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
        AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
        tCell.Controls.Add(tTxt)
        tRow.Cells.Add(tCell)

        k = k + 1

        tTable.Rows.Add(tRow)

        row = row - 1
        col = col - 1
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.Font.Size = 10
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i Mod 2) Then
                    tCell.BackColor = Drawing.Color.PapayaWhip
                Else
                    tCell.BackColor = Drawing.Color.Lavender
                End If
                'tCell.BackColor = Drawing.Color.DeepSkyBlue
                If (i = 0 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "燈盤"
                ElseIf (i = 1 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "參數設置"
                ElseIf (i = 2 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "校正檔"
                ElseIf (i = 3 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "側門安裝"
                Else
                    If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                        tCell.BackColor = Drawing.Color.Yellow
                    ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                        tCell.BackColor = Drawing.Color.GreenYellow
                    End If
                    tCell.Wrap = False
                    tCell.Text = ds.Tables(0).Rows(k)("status")
                    Hyper = New HyperLink
                    Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                    Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                    Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                    Hyper.Font.Underline = False
                    CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                    tCell.Controls.Add(Hyper)
                    Hyper.Dispose()
                    k = k + 1
                End If
                tRow.Cells.Add(tCell)
            Next
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Wrap = False
            If (i Mod 2) Then
                tCell.BackColor = Drawing.Color.PapayaWhip
            Else
                tCell.BackColor = Drawing.Color.Lavender
            End If
            tCell.Text = ds.Tables(0).Rows(k)("status")
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
            tTxt.Text = ds.Tables(0).Rows(k)("status")
            tTxt.AutoPostBack = True
            CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            tCell.Controls.Add(tTxt)
            tRow.Cells.Add(tCell)
            k = k + 1
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Public Sub CreateQCStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tTxt, tTxt1 As TextBox
        Dim i, j As Integer
        Dim col As Integer
        Dim Hyper As HyperLink
        Dim ce As CalendarExtender
        'InitLocalSQLConnection()
        SqlCmd = "Select T0.num,T0.status,T0.dpart,T0.iseq,T0.wseq " &
                 "from dbo.[work_records] T0 " &
                 "where T0.wsn='" & wsn & "' and T0.dpart=5 order by T0.iseq ,T0.wseq"
        'Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        'da1.Fill(ds)
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Font.Size = 10
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.PaleGoldenrod
                tCell.Text = "QC出貨"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                tCell.Text = "CheckList"
                tCell.Font.Bold = True
            Else
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.PapayaWhip
                If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                    tCell.BackColor = Drawing.Color.Yellow
                ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                    tCell.BackColor = Drawing.Color.GreenYellow
                End If
                tCell.Text = ds.Tables(0).Rows(k)("status")
                Hyper = New HyperLink
                Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                Hyper.Font.Underline = False
                CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                tCell.Controls.Add(Hyper)
                Hyper.Dispose()
                k = k + 1
            End If
            tRow.Cells.Add(tCell)
        Next
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Font.Size = 10
        tCell.HorizontalAlign = HorizontalAlign.Center
        'tCell.Wrap = False
        tCell.BackColor = Drawing.Color.PapayaWhip
        tCell.Text = ds.Tables(0).Rows(k)("status")
        tTxt = New TextBox()
        tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
        tTxt.Text = ds.Tables(0).Rows(k)("status")
        tTxt.AutoPostBack = True
        CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
        AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
        tCell.Controls.Add(tTxt)
        tRow.Cells.Add(tCell)

        k = k + 1

        tTable.Rows.Add(tRow)

        row = row - 1
        col = col - 1
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.Font.Size = 10
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i Mod 2) Then
                    tCell.BackColor = Drawing.Color.PapayaWhip
                Else
                    tCell.BackColor = Drawing.Color.Lavender
                End If
                'tCell.BackColor = Drawing.Color.DeepSkyBlue
                If (i = 0 And j = 0) Then
                    tCell.Font.Bold = True
                    'tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "測試"
                ElseIf (i = 1 And j = 0) Then
                    tCell.Font.Bold = True
                    tCell.Text = "出貨狀態"
                ElseIf (i = 2 And j = 0) Then
                    tCell.Font.Bold = True
                    tCell.Text = "出貨日期"
                ElseIf (i = 2 And j <> 0) Then '出貨日期 Text
                    If (ds.Tables(0).Rows(k)("status") <> "") Then
                        tCell.BackColor = Drawing.Color.SkyBlue
                        tTxt1 = New TextBox()
                        tTxt1.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
                        tTxt1.Text = ds.Tables(0).Rows(k)("status")
                        tTxt1.Width = 70
                        CommUtil.DisableObjectByPermission(tTxt1, permsmf205, "m")
                        tTxt1.AutoPostBack = True
                        tCell.Controls.Add(tTxt1)

                        ce = New CalendarExtender
                        ce.TargetControlID = tTxt1.ID
                        ce.ID = "ce_shipdate" & "_" & j
                        ce.Format = "yyyy/MM/dd"
                        tCell.Controls.Add(ce)
                        AddHandler tTxt1.TextChanged, AddressOf tTxt_TextChanged

                    End If
                    k = k + 1
                Else
                    If (i <> 1) Then
                        If (ds.Tables(0).Rows(k)("status") = "進行中") Then
                            tCell.BackColor = Drawing.Color.Yellow
                        ElseIf (ds.Tables(0).Rows(k)("status") = "已完成") Then
                            tCell.BackColor = Drawing.Color.GreenYellow
                        End If
                    Else
                        If (ds.Tables(0).Rows(k)("status") = "已完工") Then
                            tCell.BackColor = Drawing.Color.GreenYellow
                        ElseIf (ds.Tables(0).Rows(k)("status") = "已包裝") Then
                            tCell.BackColor = Drawing.Color.LightGreen
                        ElseIf (ds.Tables(0).Rows(k)("status") = "已出貨") Then
                            tCell.BackColor = Drawing.Color.SkyBlue
                        End If
                    End If
                    tCell.Wrap = False
                    tCell.Text = ds.Tables(0).Rows(k)("status")
                    Hyper = New HyperLink
                    Hyper.ID = "hyper_" & ds.Tables(0).Rows(k)("num")
                    Hyper.Text = ds.Tables(0).Rows(k)("status") 'tCell.Text
                    Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & wsn & "&mode=modify" &
                                            "&num=" & ds.Tables(0).Rows(k)("num") &
                                            "&status=" & ds.Tables(0).Rows(k)("status") &
                                            "&dpart=" & ds.Tables(0).Rows(k)("dpart") &
                                            "&iseq=" & ds.Tables(0).Rows(k)("iseq") &
                                            "&wseq=" & ds.Tables(0).Rows(k)("wseq") &
                                            "&source=" & Request.QueryString("source") &
                                            "&indexpage=" & Request.QueryString("indexpage")
                    Hyper.Font.Underline = False
                    CommUtil.DisableObjectByPermission(Hyper, permsmf205, "m")
                    tCell.Controls.Add(Hyper)
                    tCell.Wrap = True
                    If (i = 1 And ds.Tables(0).Rows(k)("status") = "已出貨") Then
                        tCell.BackColor = Drawing.Color.SkyBlue
                    End If
                    Hyper.Dispose()
                    k = k + 1
                End If
                tRow.Cells.Add(tCell)
            Next
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tCell.Wrap = False
            If (i Mod 2) Then
                tCell.BackColor = Drawing.Color.PapayaWhip
            Else
                tCell.BackColor = Drawing.Color.Lavender
            End If
            tCell.Text = ds.Tables(0).Rows(k)("status")
            tTxt = New TextBox()
            tTxt.ID = "tTxt_" & ds.Tables(0).Rows(k)("num")
            tTxt.Text = ds.Tables(0).Rows(k)("status")
            tTxt.AutoPostBack = True
            CommUtil.DisableObjectByPermission(tTxt, permsmf205, "m")
            AddHandler tTxt.TextChanged, AddressOf tTxt_TextChanged
            tCell.Controls.Add(tTxt)
            tRow.Cells.Add(tCell)
            k = k + 1
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Sub tTxt_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim tTxt As TextBox
        Dim num As Long
        Dim txtid As String
        Dim str() As String
        txtid = DirectCast(sender, TextBox).ID
        str = Split(txtid, "_")
        num = CLng(str(1))
        tTxt = WoStatusT.FindControl(txtid)
        SqlCmd = "update dbo.[work_records] set status= '" & tTxt.Text & "' " &
                 "where num=" & num
        CommUtil.SqlLocalExecute("upd", SqlCmd, conn)
        conn.Close()
    End Sub
End Class