Imports System.Data
Imports System.Data.SqlClient
Partial Public Class WebForm2
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn, connsap As New SqlConnection
    Public oCompany As New SAPbobsCOM.Company
    Public ret As Long
    Public ds As New DataSet
    Public kd As String
    Public gi As Integer = 0
    Public formtype As String
    Public fnamearr() As String = {"銷售單", "工單", "料號", "說明", "需領", "未領", "單價", "倉庫", "本庫", "本需", "本供", "它庫", "它需", "它供", "不足"}
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

    Public Sub InitSAPSQLConnection()
        connsap.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
        connsap.Open()
    End Sub

    Public Sub InitLocalSQLConnection()
        'Dim myCommand As SqlCommand 'a
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("MyConnection").ConnectionString
        conn.Open()
        'myCommand = New SqlCommand("Create DATABASE MyDatabase", conn)'a
        'myCommand.ExecuteNonQuery()'a
        'CommUtil.ShowMsg(Me,"Database is Create Successfully")'a
        CommUtil.ShowMsg(Me,conn.State.ToString())
        conn.Close()
        CommUtil.ShowMsg(Me,conn.State.ToString())
    End Sub

    Public Sub CloseSAPSQLConnection()
        connsap.Close()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        If (Not Page.IsPostBack) Then
            'InitLocalSQLConnection()
        End If
    End Sub
    Sub GetSpareInfoFromSO(ByVal seq As Integer)
        Dim i As Integer
        Dim dstemp As DataSet
        Dim idstr As String
        Dim SelectCmd As String
        If (seq <> 8) Then
            SelectCmd = "SELECT T0.DocNum,T0.[CardName],T0.TaxDate,T0.DocDueDate ,T0.Comments " & _
                    "FROM ORDR T0 " & _
                    "WHERE T0.DocStatus='O' order by T0.docnum"
        Else
            idstr = InputBox("單號:")
            SelectCmd = "SELECT T0.DocNum,T0.[CardName],T0.TaxDate,T0.DocDueDate ,T0.Comments " & _
                "FROM ORDR T0 " & _
                "WHERE T0.DocNum='" & idstr & "' " & _
                "order by T0.docnum"
        End If
        Dim da1 As New SqlDataAdapter(SelectCmd, conn)
        da1.Fill(ds, "SOList")
        i = 0
        For i = 0 To ds.Tables(0).Rows.Count - 1
            SelectCmd = "SELECT T0.DocNum,T0.[CardName], " & _
                "T1.ItemCode,T2.ItemName,T1.Quantity,T1.OpenCreQty,T1.WhsCode,T5.Onhand,T5.IsCommited,T5.OnOrder, " & _
                "T0.TaxDate,T0.DocDueDate ,T0.Comments,T4.SlpName,T3.lastname,T3.FirstName,T1.Price, " & _
                "(T2.Onhand-T5.Onhand) as OnHand1,(T2.IsCommited-T5.IsCommited) As IsCommited,(T2.OnOrder-T5.OnOrder) As OnOrder1,T1.U_F1 " & _
                "FROM ORDR T0 INNER JOIN RDR1 T1 ON T0.DocEntry = T1.DocEntry " & _
                "INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] " & _
                "INNER JOIN OHEM T3 ON T0.OwnerCode=T3.empID " & _
                "INNER JOIN OSLP T4 ON T0.SlpCode=T4.SlpCode " & _
                "INNER JOIN OITW T5 ON T1.Itemcode=T5.ItemCode " & _
                "WHERE T0.Docnum='" & ds.Tables(0).Rows(i)(0) & "' " & _
                "And T1.Whscode=T5.Whscode " & _
                "order by T1.price desc"
            da1 = New SqlDataAdapter(SelectCmd, conn)
            dstemp = New DataSet
            da1.Fill(dstemp)
            If (seq = 1 Or seq = 8) Then
                da1.Fill(ds, "SODetail")
            ElseIf (seq = 2 And dstemp.Tables(0).Rows(0)(6) = "C01") Then
                da1.Fill(ds, "SODetail")
            ElseIf (seq = 3 And (dstemp.Tables(0).Rows(0)(6) = "C02" Or dstemp.Tables(0).Rows(0)(6) = "S02")) Then
                da1.Fill(ds, "SODetail")
            ElseIf (seq = 4 And dstemp.Tables(0).Rows(0)(6) <> "C02" And dstemp.Tables(0).Rows(0)(6) <> "S02" And dstemp.Tables(0).Rows(0)(6) <> "C01") Then
                da1.Fill(ds, "SODetail")
            End If
        Next
        If (ds.Tables(1).Rows.Count <> 0) Then
            ds.Tables("SODetail").Columns.Add("Status")
            gv1.DataSource = ds.Tables("SODetail")
            gv1.DataBind()
        Else
            CommUtil.ShowMsg(Me,"無任何資料")
        End If
        'ds.Dispose()
        'dstemp.Dispose()
        'da1.Dispose()
    End Sub

    Sub GetSpareInfoFromWO(ByVal seq As Integer)
        Dim i As Integer
        Dim dstemp As DataSet
        Dim idstr As String
        Dim SelectCmd As String
        Dim da1 As SqlDataAdapter
        SelectCmd = ""
        If (seq = 5) Then
            SelectCmd = "SELECT T0.DocNum,T0.PostDate,T0.DueDate ,T0.Comments, T0.U_F16 , T0.Status , T0.CardCode,T0.Itemcode " & _
                "FROM OWOR T0 " & _
                "WHERE T0.Status<>'L' and T0.Status<>'C' and (T0.Itemcode='AOIS' or T0.itemcode='ICTS') " & _
                "order by T0.docnum"
        ElseIf (seq = 6) Then
            SelectCmd = "SELECT T0.DocNum,T0.PostDate,T0.DueDate ,T0.Comments, T0.U_F16 , T0.Status , T0.CardCode,T0.Itemcode " & _
                "FROM OWOR T0 " & _
                "WHERE T0.Status<>'L' and T0.Status<>'C' and T0.itemcode='ICTS' " & _
                "order by T0.docnum"
        ElseIf (seq = 7) Then
            SelectCmd = "SELECT T0.DocNum,T0.PostDate,T0.DueDate ,T0.Comments, T0.U_F16 , T0.Status , T0.CardCode,T0.Itemcode " & _
                "FROM OWOR T0 " & _
                "WHERE T0.Status<>'L' and T0.Status<>'C' and T0.Itemcode='AOIS' " & _
                "order by T0.docnum"
        ElseIf (seq = 9) Then
            idstr = InputBox("單號:")
            SelectCmd = "SELECT T0.DocNum,T0.PostDate,T0.DueDate ,T0.Comments, T0.U_F16 , T0.Status , T0.CardCode,T0.Itemcode " & _
                "FROM OWOR T0 " & _
                "WHERE T0.DocNum='" & idstr & "' " & _
                "order by T0.docnum"
        End If
        da1 = New SqlDataAdapter(SelectCmd, conn)
        da1.Fill(ds, "SOList")
        i = 0
        For i = 0 To ds.Tables(0).Rows.Count - 1
            SelectCmd = "SELECT T0.DocNum,T5.AvgPrice, " & _
                "T1.ItemCode,T2.ItemName,T1.PlannedQty As Quantity,(T1.PlannedQty-T1.IssuedQty) As OpenCreQty,T1.warehouse As WhsCode,T5.Onhand,T5.IsCommited,T5.OnOrder, " & _
                "T0.PostDate,T0.DueDate ,T0.Comments,T3.U_Name, " & _
                "T2.Onhand,T2.IsCommited,T2.OnOrder " & _
                "FROM OWOR T0 INNER JOIN WOR1 T1 ON T0.DocEntry = T1.DocEntry " & _
                "INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] " & _
                "INNER JOIN OUSR T3 ON T0.UserSign = T3.INTERNAL_K " & _
                "INNER JOIN OITW T5 ON T1.Itemcode=T5.ItemCode " & _
                "WHERE T0.Docnum='" & ds.Tables(0).Rows(i)(0) & "' " & _
                "And T1.Warehouse=T5.Whscode"
            da1 = New SqlDataAdapter(SelectCmd, conn)
            dstemp = New DataSet
            da1.Fill(dstemp)
            If (seq = 5 Or seq = 8) Then
                da1.Fill(ds, "SODetail")
            ElseIf (seq = 6 And dstemp.Tables(0).Rows(0)(6) = "C01") Then
                da1.Fill(ds, "SODetail")
            ElseIf (seq = 7 And (dstemp.Tables(0).Rows(0)(6) = "C02" Or dstemp.Tables(0).Rows(0)(6) = "S02")) Then
                da1.Fill(ds, "SODetail")
            End If
        Next
        If (ds.Tables(1).Rows.Count <> 0) Then
            ds.Tables("SODetail").Columns.Add("Status")
            gv1.DataSource = ds.Tables("SODetail")
            gv1.DataBind()
        Else
            CommUtil.ShowMsg(Me,"無任何資料")
        End If
    End Sub
    Protected Sub gv1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowDataBound
        Dim Index As Integer
        Dim dr() As DataRow
        Dim cno, cname, wostatus As String
        If (e.Row.RowType = DataControlRowType.Header) Then
            e.Row.Cells.Clear()
        End If
        If (e.Row.RowType = DataControlRowType.DataRow) Then
            ' Dim Hyper As New HyperLink
            ' Hyper.Text = e.Row.Cells(11).Text
            ' Hyper.NavigateUrl = "http://www.yahoo.com.tw"
            ' e.Row.Cells(11).Controls.Add(Hyper)
            'AddFieldHyper(ds.Tables(1), e.Row.RowIndex)

            e.Row.Attributes.Add("onmouseover", "c=this.style.backgroundColor;this.style.backgroundColor='lightgreen'")
            '設定光棒顏色，當滑鼠 onMouseOver 時驅動
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor=c")
            '當 onMouseOut 也就是滑鼠移開時，要恢復原本的顏色

            If (CInt(e.Row.Cells(8).Text) + CInt(e.Row.Cells(10).Text) >= CInt(e.Row.Cells(9).Text)) Then
                e.Row.Cells(14).Text = "OK"
                e.Row.Cells(14).BackColor = System.Drawing.Color.LightGreen
            Else
                e.Row.Cells(14).Text = CInt(e.Row.Cells(8).Text) + CInt(e.Row.Cells(10).Text) - CInt(e.Row.Cells(9).Text)
                e.Row.Cells(14).BackColor = System.Drawing.Color.Red
            End If
            If (kd <> e.Row.Cells(0).Text) Then
                gi = gi + 1
                Dim str As String
                If (formtype = "so") Then
                    str = "銷售單號:" & ds.Tables(1).Rows(e.Row.RowIndex)("docnum") & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp客戶名:" & ds.Tables(1).Rows(e.Row.RowIndex)("cardname") & _
                          "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp文件日期:" & ds.Tables(1).Rows(e.Row.RowIndex)("TaxDate") & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp預交日期:" & ds.Tables(1).Rows(e.Row.RowIndex)("DocDueDate")
                    AddOneRowSpanCol(sender, e.Row.RowIndex, e.Row.Cells.Count, str, True)
                    gi = gi + 1
                    str = "業務人員:" & ds.Tables(1).Rows(e.Row.RowIndex)("SlpName") & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp製單人員:" & ds.Tables(1).Rows(e.Row.RowIndex)("lastname") & _
                    ds.Tables(1).Rows(e.Row.RowIndex)("FirstName")
                    AddOneRowSpanCol(sender, e.Row.RowIndex, e.Row.Cells.Count, str, False)
                Else
                    dr = ds.Tables(0).Select("DocNum='" & e.Row.Cells(0).Text & "'")
                    Index = ds.Tables(0).Rows.IndexOf(dr(0))
                    cno = ds.Tables(0).Rows(Index)(4)
                    If (ds.Tables(0).Rows(Index)("Status") = "P") Then
                        wostatus = "計畫中"
                    ElseIf (ds.Tables(0).Rows(Index)("Status") = "R") Then
                        wostatus = "核發中"
                    ElseIf (ds.Tables(0).Rows(Index)("Status") = "L") Then
                        wostatus = "已結案"
                    ElseIf (ds.Tables(0).Rows(Index)("Status") = "C") Then
                        wostatus = "已取消"
                    Else
                        wostatus = "不明"
                    End If
                    Dim da1 As SqlDataAdapter
                    Dim dstemp As DataSet
                    Dim SelectCmd As String
                    SelectCmd = "SELECT T0.CardName " & _
                                "FROM OCRD T0 " & _
                                "WHERE T0.CardCode='" & ds.Tables(0).Rows(Index)("CardCode") & "'"
                    da1 = New SqlDataAdapter(SelectCmd, conn)
                    dstemp = New DataSet
                    da1.Fill(dstemp)
                    If (dstemp.Tables(0).Rows.Count = 0) Then
                        cname = "無指定"
                    Else
                        cname = dstemp.Tables(0).Rows(0)(0)
                    End If
                    If (ds.Tables(0).Rows(Index)("ItemCode") = "AOIS" Or ds.Tables(0).Rows(Index)("ItemCode") = "ICTS") Then
                        str = "工單號:" & ds.Tables(1).Rows(e.Row.RowIndex)("docnum") & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp聯絡單號:" & cno & _
                              "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp客戶名:" & cname & _
                              "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp文件日期:" & ds.Tables(1).Rows(e.Row.RowIndex)("PostDate") & _
                              "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp預交日期:" & ds.Tables(1).Rows(e.Row.RowIndex)("DueDate")
                    Else
                        str = "工單號:" & ds.Tables(1).Rows(e.Row.RowIndex)("docnum") & "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp聯絡單號:" & cno & _
                              "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp客戶名:" & cname & _
                              "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp文件日期:" & ds.Tables(1).Rows(e.Row.RowIndex)("PostDate") & _
                              "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp預交日期:" & ds.Tables(1).Rows(e.Row.RowIndex)("DueDate")
                    End If
                    AddOneRowSpanCol(sender, e.Row.RowIndex, e.Row.Cells.Count, str, True)
                    gi = gi + 1
                    str = "製單人員:" & ds.Tables(1).Rows(e.Row.RowIndex)("U_Name")
                    AddOneRowSpanCol(sender, e.Row.RowIndex, e.Row.Cells.Count, str, False)
                End If

                gi = gi + 1
                AddGridFieldName(sender, e.Row.RowIndex)
                kd = e.Row.Cells(0).Text
                'tc0.HorizontalAlign = HorizontalAlign.Center
            End If
        End If
    End Sub

    Sub AddFieldHyper(ByVal dt As DataTable, ByVal currentrow As Integer)
        Dim Hyper As New HyperLink
        'Dim gv As GridView = CType(sender, GridView)
        Hyper.Text = dt.Rows(currentrow)("OnHand1")
        Hyper.NavigateUrl = "http://www.yahoo.com.tw"
        If (currentrow = 2) Then
            'CommUtil.ShowMsg(Me,Hyper.Text)
        End If
        gv1.Rows(currentrow).Cells(11).Controls.Add(Hyper)
    End Sub
    Sub AddOneRowSpanCol(ByVal sender As Object, ByVal currentrow As Integer, ByVal spancol As Integer, ByVal descrip As String, ByVal judge As Boolean)
        Dim gv As GridView = CType(sender, GridView)
        Dim gvrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert)
        Dim tc0 As TableCell = New TableCell()
        tc0.Text = descrip
        tc0.ColumnSpan = spancol
        tc0.Font.Bold = True
        tc0.BorderWidth = 5
        'tc0.Controls.Add(HyperLink)
        tc0.BackColor = System.Drawing.Color.LightSkyBlue
        If (judge) Then
            If (ds.Tables(1).Rows(currentrow)(11) < Date.Today) Then 'duedate
                tc0.BackColor = System.Drawing.Color.Red
            End If
        End If
        gvrow.Cells.Add(tc0)
        gv.Controls(0).Controls.AddAt(currentrow + gi, gvrow)
    End Sub

    Sub AddGridFieldName(ByVal sender As Object, ByVal currentrow As Integer)

        Dim gv As GridView = CType(sender, GridView)
        Dim gvrow1 As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert)
        Dim tc0 As TableCell
        For Each fname As String In fnamearr
            tc0 = New TableCell
            tc0.Text = fname
            tc0.Font.Bold = True
            tc0.BorderWidth = 5
            tc0.HorizontalAlign = HorizontalAlign.Center
            tc0.BackColor = System.Drawing.Color.LightSkyBlue
            gvrow1.Cells.Add(tc0)
        Next
        gv.Controls(0).Controls.AddAt(currentrow + gi, gvrow1)
    End Sub

    Protected Sub gv1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gv1.RowCreated
        'If (e.Row.RowType = DataControlRowType.Header) Then
        '    e.Row.Cells.Clear()
        'ElseIf (e.Row.RowType = DataControlRowType.DataRow) Then
        '    If (e.Row.RowIndex = 0) Then
        '        '加heater
        '        '加欄位名
        '    Else
        '        kd = ds.Tables(1).Rows(e.Row.RowIndex - 1)(0)
        '        If (kd <> ds.Tables(1).Rows(e.Row.RowIndex)(0)) Then
        '            gi = gi + 1
        '            Dim Index As Integer
        '            Dim dr() As DataRow
        '            Dim str As String
        '            dr = ds.Tables(0).Select("DocNum=" & ds.Tables(1).Rows(e.Row.RowIndex)(0))
        '            Index = ds.Tables(0).Rows.IndexOf(dr(0))
        '            str = "銷售單號:" & ds.Tables(0).Rows(Index)("docnum") & "    客戶名:" & ds.Tables(0).Rows(Index)("cardname")
        '            Dim gv As GridView = CType(sender, GridView)
        '            Dim gvrow As GridViewRow = New GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert)
        '            Dim tc0 As TableCell = New TableCell()
        '            tc0.Text = str
        '            tc0.ColumnSpan = e.Row.Cells.Count
        '            gvrow.Cells.Add(tc0)
        '            gv.Controls(0).Controls.AddAt(e.Row.RowIndex + gi, gvrow)
        '            'CommUtil.ShowMsg(Me,e.Row.RowIndex)
        '        End If
        '    End If
        'End If
    End Sub
End Class