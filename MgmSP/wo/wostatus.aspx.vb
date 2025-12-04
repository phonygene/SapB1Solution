Imports System.Data
Imports System.Data.SqlClient
Public Class wostatus
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public conn As New SqlConnection
    Public connsap As New SqlConnection
    Public SqlCmd As String
    Public ds As New DataSet
    Public dr As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If (Session("s_id") = "") Then
            Response.Redirect("~\index.aspx?smid=index&timeout=1")
        End If
        ShowOnLineWoStatus()
        'ShowOnLineWoStatus_Shipped()
    End Sub
    Sub ShowOnLineWoStatus()
        Dim itemcount As Integer
        'Dim model_set As Integer
        'InitLocalSQLConnection()
        SqlCmd = "Select T0.wsn ,T0.docnum, T0.cus_name ,T0.model , T0.model_set,T0.f_set,T0.ship_set,T0.ship_date " &
                 "from dbo.[worksn] T0 " &
                 "where T0.f_stat<>70 and T0.model_set<>0 and T0.getpo=20 order by T0.model , T0.docnum" 'getpo=20 表示聯絡單
        'Dim da1 As New SqlDataAdapter(SqlCmd, conn)
        'da1.Fill(ds)
        'CloseLocalSQLConnection()
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        itemcount = ds.Tables(0).Rows.Count
        Table1.BackColor = Drawing.Color.AntiqueWhite
        CreateHeadTable(Table1, 8, itemcount)
        CreateMaterialStatusTable(Table1, 5, itemcount)
        CreateAssemblyStatusTable(Table1, 5, itemcount)
        CreateWiringStatusTable(Table1, 5, itemcount)
        CreateAdjustStatusTable(Table1, 5, itemcount)
        CreateQCStatusTable(Table1, 3, itemcount)
        CreateTimeUpdateStatusTable(Table1, 1, itemcount)
    End Sub

    Sub ShowOnLineWoStatus_Shipped()
        Dim itemcount As Integer
        'SqlCmd = "Select T0.wsn ,T0.docnum, T0.cus_name ,T0.model , T0.model_set,T0.f_set,T0.ship_set,T0.ship_date " &
        '         "from dbo.[worksn] T0 " &
        '         "where T0.ship_set<>0 order by T0.model , T0.docnum"
        SqlCmd = "Select T0.wsn ,T0.docnum, T0.cus_name ,T0.model , T0.model_set,T0.f_set,T0.ship_set,T0.ship_date " &
                 "from dbo.[worksn] T0 " &
                 "order by T0.model , T0.docnum"
        ds = CommUtil.SelectLocalSqlUsingDataSet(ds, SqlCmd, conn)
        conn.Close()
        itemcount = ds.Tables(0).Rows.Count
        Table1.BackColor = Drawing.Color.AntiqueWhite
        CreateHeadTable(Table1, 8, itemcount)
        CreateMaterialStatusTable(Table1, 5, itemcount)
        CreateAssemblyStatusTable(Table1, 5, itemcount)
        CreateWiringStatusTable(Table1, 5, itemcount)
        CreateAdjustStatusTable(Table1, 5, itemcount)
        CreateQCStatusTable(Table1, 3, itemcount)
        CreateTimeUpdateStatusTable(Table1, 1, itemcount)
    End Sub
    Public Sub CreateHeadTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim col As Integer
        Dim Hyper As HyperLink
        'Dim Hyper As New HyperLink
        col = rcount + 1
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i = 0) Then
                    If (j = 0) Then
                        tCell.Wrap = False
                        tCell.Text = "客戶名稱"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("cus_name")
                    End If
                ElseIf (i = 1) Then
                    If (j = 0) Then
                        tCell.Text = "工單號"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("wsn")
                        Hyper = New HyperLink
                        Hyper.Text = ds.Tables(0).Rows(j - 1)("wsn") 'tCell.Text
                        Hyper.NavigateUrl = "wostamodify.aspx?wsn=" & tCell.Text & "&mode=show&source=fromwostatus"
                        Hyper.Font.Underline = False
                        tCell.Controls.Add(Hyper)
                        Hyper.Dispose()
                    End If
                ElseIf (i = 2) Then
                    If (j = 0) Then
                        tCell.Text = "SAP號"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("docnum")
                    End If
                ElseIf (i = 3) Then
                    If (j = 0) Then
                        tCell.Text = "機型"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("model")
                    End If
                ElseIf (i = 4) Then
                    If (j = 0) Then
                        tCell.Text = "訂單台數"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("model_set")
                    End If
                ElseIf (i = 5) Then
                    If (j = 0) Then
                        tCell.Text = "完成台數"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("f_set")
                        If (ds.Tables(0).Rows(j - 1)("f_set") = ds.Tables(0).Rows(j - 1)("model_set")) Then
                            tCell.BackColor = Drawing.Color.GreenYellow
                        ElseIf (ds.Tables(0).Rows(j - 1)("f_set") > 0) Then
                            tCell.BackColor = Drawing.Color.Yellow
                        End If
                    End If
                ElseIf (i = 6) Then
                    If (j = 0) Then
                        tCell.Text = "出貨台數"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("ship_set")
                        If (ds.Tables(0).Rows(j - 1)("ship_set") = ds.Tables(0).Rows(j - 1)("model_set")) Then
                            tCell.BackColor = Drawing.Color.DeepSkyBlue
                        ElseIf (ds.Tables(0).Rows(j - 1)("ship_set") > 0) Then
                            tCell.BackColor = Drawing.Color.Yellow
                        End If
                    End If
                ElseIf (i = 7) Then
                    If (j = 0) Then
                        tCell.Text = "預計出貨"
                    Else
                        tCell.Text = ds.Tables(0).Rows(j - 1)("ship_date")
                        If (CDate(ds.Tables(0).Rows(j - 1)("ship_date")) < CDate(Format(Now(), "yyyy/MM/dd"))) Then
                            tCell.BackColor = Drawing.Color.Red
                        ElseIf (CDate(ds.Tables(0).Rows(j - 1)("ship_date")) < DateAdd("d", 7, Format(Now(), "yyyy/MM/dd"))) Then
                            tCell.BackColor = Drawing.Color.Yellow
                        End If
                    End If
                End If
                If (j = 0) Then
                    tCell.ColumnSpan = 2
                    tCell.Font.Bold = True
                End If
                tRow.Cells.Add(tCell)
            Next
            tTable.Rows.Add(tRow)
            tTable.Rows(i).Font.Bold = True
        Next
    End Sub

    Function GetItemStatus(j As Integer, dpart As Integer, iseq As Integer)
        Dim tCell As TableCell
        Dim fcount, pcount As Integer
        Dim str0, str1, str2, str3 As String
        'InitLocalSQLConnection()
        If (dpart = 5 And iseq = 3) Then
            str1 = "未出貨"
            str0 = "已完工"
            str2 = "已包裝"
            str3 = "已出貨"
        Else
            str1 = "未進行"
            str2 = "進行中"
            str3 = "已完成"
        End If
        SqlCmd = "Select count(*) " &
                "from dbo.[work_records] T0 " &
                "where T0.wsn='" & ds.Tables(0).Rows(j - 2)("wsn") & "' " &
                "and dpart=" & dpart & " and iseq=" & iseq & " and status='" & str3 & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        fcount = dr(0)
        dr.Close()
        conn.Close()
        SqlCmd = "Select count(*) " &
                "from dbo.[work_records] T0 " &
                "where T0.wsn='" & ds.Tables(0).Rows(j - 2)("wsn") & "' " &
                "and dpart=" & dpart & " and iseq=" & iseq & " and status='" & str2 & "'"
        'myCommand = New SqlCommand(SqlCmd, conn)
        'dr = myCommand.ExecuteReader()
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        dr.Read()
        pcount = dr(0)
        dr.Close()
        conn.Close()
        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (fcount = ds.Tables(0).Rows(j - 2)("model_set")) Then
            tCell.Text = str3 & fcount & "/" & ds.Tables(0).Rows(j - 2)("model_set")
            tCell.BackColor = Drawing.Color.GreenYellow
        ElseIf (pcount > 0 Or fcount > 0) Then
            If (dpart = 5 And iseq = 3) Then
                tCell.Text = str2 & pcount & "/" & (ds.Tables(0).Rows(j - 2)("model_set") - fcount) '
            Else
                tCell.Text = str2 & fcount & "/" & ds.Tables(0).Rows(j - 2)("model_set")
            End If
            tCell.BackColor = Drawing.Color.Yellow
        Else
            tCell.Text = str1
            tCell.BackColor = Drawing.Color.WhiteSmoke
        End If
        'CloseLocalSQLConnection()
        Return tCell
    End Function
    Public Sub CreateMaterialStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim col As Integer
        'InitLocalSQLConnection()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1

        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.DeepSkyBlue
                tCell.Text = "料件"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.DeepSkyBlue
                tCell.Text = "骨架"
                tCell.Font.Bold = True
            Else
                tCell = GetItemStatus(j, 1, 1)
            End If
            tRow.Cells.Add(tCell)
        Next
        tTable.Rows.Add(tRow)

        'row = row - 1
        'col = col - 1
        For i = 2 To row
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 1 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i = 2 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "鈑金"
                ElseIf (i = 3 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "市購"
                ElseIf (i = 4 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "車件"
                ElseIf (i = 5 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "銑件"
                Else '一般
                    tCell = GetItemStatus(j, 1, i)
                End If
                tRow.Cells.Add(tCell)
            Next
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub

    Public Sub CreateAssemblyStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim col As Integer
        'InitLocalSQLConnection()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.Khaki
                tCell.Text = "機構組裝"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.Khaki
                tCell.Text = "發料"
                tCell.Font.Bold = True
            Else
                tCell = GetItemStatus(j, 2, 1)
            End If
            tRow.Cells.Add(tCell)
        Next
        tTable.Rows.Add(tRow)

        'row = row - 1
        'col = col - 1
        For i = 2 To row
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 1 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i = 2 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "XYZ"
                ElseIf (i = 3 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "軌道"
                ElseIf (i = 4 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "鏡頭"
                ElseIf (i = 5 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "驗證"
                Else
                    tCell = GetItemStatus(j, 2, i)
                End If
                tRow.Cells.Add(tCell)
            Next
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Public Sub CreateWiringStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim col As Integer
        'InitLocalSQLConnection()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.DeepSkyBlue
                tCell.Text = "機台佈線"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.DeepSkyBlue
                tCell.Text = "發料"
                tCell.Font.Bold = True
            Else
                tCell = GetItemStatus(j, 3, 1)
            End If
            tRow.Cells.Add(tCell)
        Next
        tTable.Rows.Add(tRow)

        'row = row - 1
        'col = col - 1
        For i = 2 To row
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 1 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i = 2 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "XYZ"
                ElseIf (i = 3 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "配電盤"
                ElseIf (i = 4 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "外罩"
                ElseIf (i = 5 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "罩外罩"
                Else
                    tCell = GetItemStatus(j, 3, i)
                End If
                tRow.Cells.Add(tCell)
            Next
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Public Sub CreateAdjustStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim col As Integer
        'InitLocalSQLConnection()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (j = 0) Then
                tCell.RowSpan = 5
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.Khaki
                tCell.Text = "機台調試"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.Khaki
                tCell.Text = "正交"
                tCell.Font.Bold = True
            Else
                tCell = GetItemStatus(j, 4, 1)
            End If
            tRow.Cells.Add(tCell)
        Next
        tTable.Rows.Add(tRow)

        'row = row - 1
        'col = col - 1
        For i = 2 To row
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 1 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i = 2 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "燈盤"
                ElseIf (i = 3 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.Wrap = False
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "參數設置"
                ElseIf (i = 4 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "校正檔"
                ElseIf (i = 5 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.Khaki
                    tCell.Text = "側門安裝"
                Else
                    tCell = GetItemStatus(j, 4, i)
                End If
                tRow.Cells.Add(tCell)
            Next
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Public Sub CreateQCStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim col As Integer
        'InitLocalSQLConnection()
        col = rcount + 2
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To (col - 1)
            tCell = New TableCell()
            tCell.BorderWidth = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (j = 0) Then
                tCell.RowSpan = 3
                tCell.Font.Bold = True
                tCell.BackColor = Drawing.Color.DeepSkyBlue
                tCell.Text = "出貨QC"
            ElseIf (j = 1) Then
                tCell.Wrap = False
                tCell.BackColor = Drawing.Color.DeepSkyBlue
                tCell.Text = "CheckList"
                tCell.Font.Bold = True
            Else
                tCell = GetItemStatus(j, 5, 1)
            End If
            tRow.Cells.Add(tCell)
        Next
        tTable.Rows.Add(tRow)

        'row = row - 1CreateTimeUpdateStatusTable
        'col = col - 1
        For i = 2 To row
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 1 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i = 2 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "測試"
                ElseIf (i = 3 And j = 1) Then
                    tCell.Font.Bold = True
                    tCell.BackColor = Drawing.Color.DeepSkyBlue
                    tCell.Text = "出貨狀態"
                Else
                    tCell = GetItemStatus(j, 5, i)
                End If
                tRow.Cells.Add(tCell)
            Next
            tTable.Rows.Add(tRow)
        Next
        'CloseLocalSQLConnection()
    End Sub
    Public Sub CreateTimeUpdateStatusTable(ByVal tTable As Table, ByVal row As Integer, ByVal rcount As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim i, j As Integer
        Dim col As Integer
        col = rcount + 1
        For i = 0 To (row - 1)
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To (col - 1)
                tCell = New TableCell()
                tCell.BorderWidth = 1
                tCell.HorizontalAlign = HorizontalAlign.Center
                If (i = 0 And j = 0) Then
                    tCell.Wrap = False
                    tCell.ColumnSpan = 2
                    tCell.Font.Bold = True
                    tCell.Text = "資料更新"
                End If
                tRow.Cells.Add(tCell)
            Next
            tTable.Rows.Add(tRow)
        Next
    End Sub
    Sub SetRowBackColor(ByVal tTable As Table, ByVal row As Integer)
        If (row Mod 2) Then
            tTable.Rows(row).BackColor = Drawing.Color.LightBlue
        Else
            tTable.Rows(row).BackColor = Drawing.Color.Azure
        End If
    End Sub

    Sub TableSample()
        'Dim mytable As Table
        Dim MyRow As TableRow
        Dim MyCell As TableCell

        '#動態表格 
        'mytable = New Table
        'MyRow = New TableRow
        'MyCell = New TableCell
        'Table1.CellPadding = "2"
        'Table1.CellSpacing = "0"


        '下方for..next為第一列顯示(選項名稱),columnspan是mycell欄位向右平移語法 
        'MyCell.ColumnSpan = 2 
        'MyCell.HorizontalAlign = HorizontalAlign.Justify 


        'MyCell.BorderWidth = "1"
        'MyCell.BorderColor = Drawing.Color.Black
        'MyCell.Width = "50"
        'MyCell.Height = "20"



        MyRow = New TableRow
        MyCell = New TableCell
        MyCell.Text = "區域"
        MyCell.RowSpan = "2"
        MyCell.ColumnSpan = "2"
        MyCell.BorderWidth = "1"
        MyCell.BorderColor = Drawing.Color.Black
        MyCell.HorizontalAlign = HorizontalAlign.Center
        MyRow.Cells.Add(MyCell)


        MyCell = New TableCell
        MyCell.Text = "月份"
        MyCell.ColumnSpan = "12"
        MyCell.BorderWidth = "1"
        MyCell.BorderColor = Drawing.Color.Black
        MyCell.HorizontalAlign = HorizontalAlign.Center
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "合計"
        MyCell.RowSpan = "2"
        MyCell.BorderWidth = "1"
        MyCell.BorderColor = Drawing.Color.Black
        MyCell.HorizontalAlign = HorizontalAlign.Center
        MyRow.Cells.Add(MyCell)

        Table1.Rows.Add(MyRow)

        MyRow = New TableRow
        For b = 1 To 12

            MyCell = New TableCell
            MyCell.Text = b
            MyRow.Cells.Add(MyCell)

            MyCell.BorderWidth = "1"
            MyCell.BorderColor = Drawing.Color.Black
            MyCell.HorizontalAlign = HorizontalAlign.Right
            MyCell.Width = "50"
            MyCell.Height = "20"
        Next
        Table1.Rows.Add(MyRow)




        '第二列向下(選定第一行題目 + 選項) 

        For c = 1 To 2
            MyRow = New TableRow
            MyCell = New TableCell
            MyCell.Text = "北區"
            MyRow.Cells.Add(MyCell)

            MyCell = New TableCell
            MyCell.Text = "收入"
            MyRow.Cells.Add(MyCell)



            MyCell = New TableCell
            MyCell.Text = "支出"
            MyRow.Cells.Add(MyCell)


            MyCell.BorderWidth = "1"
            MyCell.BorderColor = Drawing.Color.Black


            For d = 1 To 13

                MyCell = New TableCell
                MyRow.Cells.Add(MyCell)
                MyCell.Text = "0"
                MyCell.BorderWidth = "1"
                MyCell.BorderColor = Drawing.Color.Black
            Next

            Table1.Rows.Add(MyRow)

        Next
    End Sub

End Class