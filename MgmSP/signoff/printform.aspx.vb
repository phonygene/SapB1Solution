Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Public Class printform
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn, connsap1, connsap2 As New SqlConnection
    Public SqlCmd As String
    Public dr, dr1, drsap As SqlDataReader
    Public ds As New DataSet
    Public ScriptManager1 As New ScriptManager
    Public docnum As Long
    Public docstatus, url, usingwhs As String
    Public sfid As Integer
    Public TxtReason As TextBox
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        docnum = Request.QueryString("docnum")
        sfid = Request.QueryString("sfid")
        Page.Form.Controls.Add(ScriptManager1)
        url = Application("http")
        usingwhs = Request.QueryString("usingwhs")
        ContentTCreate() '料件List Table
        SqlCmd = "select seq from  [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " order by seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            Do While (dr.Read())
                CreateSignOffFlowField(dr(0))
            Loop
        End If
        dr.Close()
        connsap.Close()
        'End If
        PutDataToSignOffFlow()
    End Sub
    Sub ContentTCreate()
        If (sfid = 51 Or sfid = 50) Then

        ElseIf (sfid = 16) Then '未刷卡單

        ElseIf (sfid = 12) Then '機台聯絡單
            InitTableToSfid12()
            ShowXMSCT()
        ElseIf (sfid = 1) Then '通用聯絡單
            InitTableToSfid1()
            ShowXGCT()
        ElseIf (sfid = 22 Or sfid = 3) Then '客戶機台問題反應單
            InitTableToSfid3_22()
        ElseIf (sfid = 23 Or sfid = 24) Then
            InitTableToSfid23_24()
        ElseIf (sfid = 100) Then '
            InitTableToSfid100()
        ElseIf (sfid = 101) Then '
            InitTableToSfid101()
        End If
    End Sub
    Sub InitTableToSfid100()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor, WhiteColor As Drawing.Color
        Dim tImage As Image
        Dim colcountadjust As Integer
        Dim totalprice As Double
        Dim i As Integer
        colcountadjust = 2
        BColor = System.Drawing.Color.LightBlue
        WhiteColor = Drawing.Color.White
        contentT.Font.Name = "標楷體"
        contentT.Font.Size = 12

        'logo
        tRow = New TableRow()
        For j = 1 To (16 - colcountadjust) '10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 150
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow)
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 3
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 10 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 49) Then
            tCell.Text = "捷智科技 生產料件報廢單"
        ElseIf (sfid = 50) Then
            tCell.Text = "捷智科技 備品需求聯絡單"
        ElseIf (sfid = 51) Then
            tCell.Text = "捷智科技 料件入出庫單"
        ElseIf (sfid = 100) Then
            tCell.Text = "捷智科技 已簽核單據之補充單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 3
        Dim maindoc As Long
        If (docnum <> 0) Then
            If (sfid = 100) Then
                SqlCmd = "Select attadoc from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    maindoc = drL(0)
                End If
                drL.Close()
                connL.Close()

                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & " 母單:" & maindoc & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow)

        'tRow = New TableRow()
        'For j = 1 To 10
        '    tCell = New TableCell
        '    tCell.BorderWidth = 0
        '    tCell.Width = 200
        '    tCell.HorizontalAlign = HorizontalAlign.Center
        '    tRow.Controls.Add(tCell)
        'Next
        'ContentT.Rows.Add(tRow)
        Dim descripreason, reasontitle As String
        Dim itemlabelwidth As Integer = 200
        SqlCmd = "Select descrip FROM [dbo].[@XSMLS] T0 WHERE head=1 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            descripreason = drL(0)
        Else
            CommUtil.ShowMsg(Me, "找不到此簽核單說明資料(@XSMLS),中斷產生Pdf")
            Exit Sub
        End If
        drL.Close()
        connL.Close()
        If (sfid <> 100) Then
            reasontitle = "事由說明"
        Else
            reasontitle = "補充說明"
        End If
        '事由說明表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.Controls.Add(CellSet(reasontitle, 1, 16 - colcountadjust, False, itemlabelwidth, 0, "center", BColor))
        contentT.Rows.Add(tRow)

        tRow = New TableRow()
        tCell = CellSet(CommUtil.TextTransToHtmlFormat(descripreason), 1, 16 - colcountadjust, False, itemlabelwidth, 100, "left", WhiteColor)
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        '料件表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.Controls.Add(CellSet("補充加簽之增減料件表列", 1, 16 - colcountadjust, False, itemlabelwidth, 0, "center", BColor))
        contentT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 0 To 15 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Width = 40
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (i = 0) Then
                tCell.Text = "項次"
                tCell.Width = 40
            ElseIf (i = 1) Then
                tCell.Text = "料號"
                tCell.Width = 120
            ElseIf (i = 2) Then
                tCell.Text = "說明"
                tCell.Width = 300
            ElseIf (i = 3) Then
                tCell.Text = "需求"
                tCell.Width = 40
            ElseIf (i = 4) Then
                tCell.Text = "單價"
                tCell.Width = 40
            ElseIf (i = 5) Then
                tCell.Text = "總價"
                tCell.Width = 40
                '-------------------------------------
            ElseIf (i = 6) Then
                tCell.Text = "本庫"
                tCell.Width = 40
            ElseIf (i = 7) Then
                tCell.Text = "本需"
                tCell.Width = 40
            ElseIf (i = 8) Then
                tCell.Text = "本供"
                tCell.Width = 40
            ElseIf (i = 9) Then
                tCell.Text = "它庫"
                tCell.Width = 40
            ElseIf (i = 10) Then
                tCell.Text = "它需"
                tCell.Width = 40
            ElseIf (i = 11) Then
                tCell.Text = "它供"
                tCell.Width = 40
            ElseIf (i = 12) Then '6
                If (sfid = 51 Or sfid = 50 Or sfid = 100) Then
                    tCell.Text = "處置"
                ElseIf (sfid = 49) Then
                    tCell.Text = "報廢原因"
                End If
                tCell.Width = 160
            ElseIf (i = 13) Then
                tCell.Text = "備註"
                tCell.Width = 250
            End If
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow)

        SqlCmd = "Select IsNull(sum(quantity*price),0) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        totalprice = drL(0)
        drL.Close()
        connL.Close()
        SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.price,T0.method,T0.comment,T0.num, " &
                "T2.Onhand,T2.IsCommited,T2.OnOrder,T1.OnHand-T2.OnHand,T1.IsCommited-T2.IsCommited,T1.OnOrder-T2.OnOrder " &
                "FROM [dbo].[@XSMLS] T0 " &
                "Inner Join OITM T1 On T0.itemcode=T1.Itemcode Inner Join OITW T2 on T1.Itemcode=T2.Itemcode " &
                "WHERE T0.head=0 And T0.[docentry] =" & docnum & " And T2.Whscode='" & usingwhs & "' ORDER BY T0.num"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        'For i = 4 To tablerow + 2 '原 i=1 to tablerow , 但此之前有幾列row , 故需加上 , 以便用i命名之id 能與showmaterial 一致
        i = 1
        If (drL.HasRows) Then
            Do While (drL.Read())
                'MsgBox(drL(0))
                tRow = New TableRow()
                tRow.BorderWidth = 1
                For j = 0 To 15 - colcountadjust
                    tCell = New TableCell
                    tCell.BorderWidth = 1
                    tCell.Height = 20
                    If (j = 0 Or j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 9 Or j = 10 Or j = 11) Then
                        tCell.HorizontalAlign = HorizontalAlign.Center
                    End If
                    If (j = 0) Then '項次
                        tCell.Text = i
                    ElseIf (j = 1) Then '料號
                        tCell.Text = drL(0)
                    ElseIf (j = 2) Then '說明
                        tCell.Text = drL(1)
                    ElseIf (j = 3) Then '數量
                        tCell.Text = CLng(drL(2))
                    ElseIf (j = 4) Then '單價
                        tCell.Text = CLng(drL(3))
                    ElseIf (j = 5) Then '總價
                        tCell.Text = CLng(drL(2) * drL(3))
                    ElseIf (j = 6) Then '本庫
                        tCell.Text = CLng(drL(7))
                    ElseIf (j = 7) Then '本需
                        tCell.Text = CLng(drL(8))
                    ElseIf (j = 8) Then '本供
                        tCell.Text = CLng(drL(9))
                    ElseIf (j = 9) Then '它庫
                        tCell.Text = CLng(drL(10))
                    ElseIf (j = 10) Then '它需
                        tCell.Text = CLng(drL(11))
                    ElseIf (j = 11) Then '它供
                        tCell.Text = CLng(drL(12))
                    ElseIf (j = 12) Then '處置
                        tCell.Text = drL(4)
                    ElseIf (j = 13) Then '備註
                        tCell.Text = drL(5)
                    End If
                    tRow.Controls.Add(tCell)
                Next
                contentT.Rows.Add(tRow)
                i = i + 1
            Loop
        End If
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To 15 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Height = 20
            If (j = 0 Or j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 9 Or j = 10 Or j = 11) Then
                tCell.HorizontalAlign = HorizontalAlign.Center
            End If
            If (j = 4) Then
                tCell.Text = "Total"
            ElseIf (j = 5) Then
                tCell.Text = Format(totalprice, "###,###.##")
            End If
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow)
        'Next
    End Sub
    Sub InitTableToSfid101()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim tablerow As Integer
        Dim mcount, i As Integer
        Dim BColor As Drawing.Color
        Dim tImage As Image
        Dim colcountadjust As Integer
        Dim attadoc As String
        colcountadjust = 0
        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then '因送審後 , 有 2個欄位會刪除
        '    colcountadjust = 2
        'End If
        SqlCmd = "Select attadoc from [dbo].[@XASCH] where docnum=" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            attadoc = drL(0)
            If (attadoc = "" Or attadoc = "NA") Then
                CommUtil.ShowMsg(Me, "找不到此單號:" & docnum & "之所屬母單號(顯示之母單號為:" & attadoc & ")")
                Exit Sub
            End If
        Else
            CommUtil.ShowMsg(Me, "找不到此單號:" & docnum)
            Exit Sub
        End If
        drL.Close()
        connL.Close()

        BColor = System.Drawing.Color.LightBlue
        contentT.Font.Name = "標楷體"
        contentT.Font.Size = 12
        SqlCmd = "Select count(*) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & CLng(attadoc)
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        mcount = drL(0)
        If (drL(0) <= 5) Then
            tablerow = 6
        Else
            tablerow = drL(0) + 1
        End If
        drL.Close()
        connL.Close()
        'logo
        tRow = New TableRow()
        For j = 1 To (8 - colcountadjust) '10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 300
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow)
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        'tCell.ColumnSpan = 3
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 6 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "捷智科技 料件返還單(對應離倉單號:" & attadoc & ")"
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        'tCell.ColumnSpan = 3
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow)

        '料件表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 8 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "料件返還表列"
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 0 To 7 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Width = 40
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (i = 0) Then
                tCell.Text = "項次"
                tCell.Width = 40
            ElseIf (i = 1) Then
                tCell.Text = "料號"
                tCell.Width = 200 '120
            ElseIf (i = 2) Then
                tCell.Text = "說明"
                tCell.Width = 500 '300
            ElseIf (i = 3) Then
                tCell.Text = "離倉數量"
                tCell.Width = 40
            ElseIf (i = 4) Then
                tCell.Text = "已還數量"
                tCell.Width = 40
            ElseIf (i = 5) Then
                tCell.Text = "此次返還"
                tCell.Width = 40
            ElseIf (i = 6) Then
                tCell.Text = "當初離倉原因"
                tCell.Width = 40
            ElseIf (i = 7) Then
                tCell.Text = "備註"
                tCell.Width = 300 '250
            End If
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow)
        For i = 2 To tablerow + 1 '原 i=1 to tablerow , 但此之前有幾列row , 故需加上 , 以便用i命名之id 能與showmaterial 一致
            tRow = New TableRow()
            tRow.BorderWidth = 1
            For j = 0 To 7 - colcountadjust
                tCell = New TableCell
                tCell.BorderWidth = 1
                tCell.Height = 20
                If (j = 5) Then
                    tCell.Width = 40
                ElseIf (j = 7) Then
                    tCell.Width = 200
                End If
                If (j = 0 Or j = 3 Or j = 4 Or j = 5 Or j = 6) Then
                    tCell.HorizontalAlign = HorizontalAlign.Center
                End If
                tRow.Controls.Add(tCell)
            Next
            contentT.Rows.Add(tRow)
        Next
        i = 2 '料件在Table 之起始列
        SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.rtnqty,T0.method,T0.comment,T0.num,T0.nowrtnqty " &
            "FROM [dbo].[@XSMLS] T0 " &
            "WHERE T0.head=0 And T0.[docentry] =" & CLng(attadoc) & " ORDER BY T0.num"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                contentT.Rows(i).Cells(0).Text = i - 1 '項次
                contentT.Rows(i).Cells(1).Text = drL(0) '料號
                contentT.Rows(i).Cells(2).Text = drL(1) '說明
                contentT.Rows(i).Cells(3).Text = drL(2) '離倉數量
                contentT.Rows(i).Cells(4).Text = drL(3) '已還數量
                contentT.Rows(i).Cells(5).Text = drL(7)
                contentT.Rows(i).Cells(6).Text = drL(4) '當初離倉原因
                contentT.Rows(i).Cells(7).Text = drL(5)
                i = i + 1
            Loop
        End If
        drL.Close()
        connL.Close()
    End Sub

    Sub InitTableToSfid3_22() 'kkkkk
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor, NeedInputColor, WhiteColor As Drawing.Color
        Dim tImage As Image
        Dim tChk As CheckBox
        NeedInputColor = Drawing.Color.AntiqueWhite
        WhiteColor = Drawing.Color.White
        tRow = New TableRow()
        For j = 1 To 16
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 1
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 14
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 22) Then
            tCell.Text = "捷智科技 客戶機台問題反應單"
        ElseIf (sfid = 3) Then
            tCell.Text = "捷智科技 廠內機台問題反應單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 1
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow) 'row=1

        tRow = New TableRow()
        For j = 1 To 16
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow) 'row=0
        BColor = System.Drawing.Color.LightBlue
        contentT.Font.Name = "標楷體"
        contentT.Font.Size = 14
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.Controls.Add(CommUtil.CellSet("機台基本資訊", 1, 16, False, 0, 0, "center", BColor))
        contentT.Rows.Add(tRow)  'row=1
        Dim itemlabelwidth As Integer = 200
        SqlCmd = "Select T0.reportdate,T0.machinetype,T0.cusname,cusfactoryOrmo,model,machineserialOrwo,installdateOrshipdate, " &
             "problemtype,typedescrip,verandspec,faeperson, " &
             "problemdescrip,processdescrip,verifydescrip,problemnote,firstinstallOrnoassign,inwarranty,qcperson " &
             "FROM [dbo].[@XCMRT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

        If (drL.HasRows) Then
            drL.Read()
            '聯絡表頭 
            tRow = New TableRow()
            'CellSet(Text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer, align As String)
            tRow.Controls.Add(CellSet("回報日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CellSet(drL(0), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CellSet("產品別", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            'tRow.Controls.Add(CellSet("", 1, 2, False, 0, 0, "center", False))
            tRow.Controls.Add(CellSet(drL(1), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CellSet("客戶名稱", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CellSet(drL(2), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            If (sfid = 22) Then
                tRow.Controls.Add(CommUtil.CellSet("客戶廠區", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(3), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            ElseIf (sfid = 3) Then
                tRow.Controls.Add(CommUtil.CellSet("機台料號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(3), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            End If
            contentT.Rows.Add(tRow) 'row=2
            '第三列
            tRow = New TableRow()
            tRow.Controls.Add(CellSet("機台型號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CellSet(drL(4), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            If (sfid = 22) Then
                tRow.Controls.Add(CommUtil.CellSet("機台序號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tCell = CommUtil.CellSet(drL(5), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor)
                tCell.Font.Size = 10
                tRow.Controls.Add(tCell)

                tRow.Controls.Add(CommUtil.CellSet("裝機日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(6), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

                'tCell = New TableCell
                tCell = CommUtil.CellSet("", 1, 4, False, itemlabelwidth * 2, 0, "center", WhiteColor)
                tChk = New CheckBox
                tChk.ID = "chk_firstinstallOrnoassign"
                tChk.Text = "新安裝&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                If (drL(15) = 1) Then
                    tChk.Checked = True
                Else
                    tChk.Checked = False
                End If
                tCell.Controls.Add(tChk)
                tChk = New CheckBox
                tChk.ID = "chk_inwarranty"
                tChk.Text = "保固期內"
                If (drL(16) = 1) Then
                    tChk.Checked = True
                Else
                    tChk.Checked = False
                End If
                tCell.Controls.Add(tChk)
                tRow.Controls.Add(tCell)
            ElseIf (sfid = 3) Then
                tRow.Controls.Add(CommUtil.CellSet("工單號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tCell = CommUtil.CellSet(drL(5), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor)
                tCell.Font.Size = 10
                tRow.Controls.Add(tCell)

                tRow.Controls.Add(CommUtil.CellSet("出貨日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(6), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

                'tCell = New TableCell
                tCell = CommUtil.CellSet("", 1, 4, False, itemlabelwidth * 2, 0, "center", WhiteColor)
                tChk = New CheckBox
                tChk.ID = "chk_firstinstallOrnoassign"
                tChk.Text = "無指定機台"
                If (drL(15) = 1) Then
                    tChk.Checked = True
                Else
                    tChk.Checked = False
                End If
                tCell.Controls.Add(tChk)
                tRow.Controls.Add(tCell)
            End If
            contentT.Rows.Add(tRow) 'row=3

            tRow = New TableRow()
            tRow.Font.Size = 16
            tRow.Controls.Add(CellSet("資訊記錄", 1, 16, False, 0, 0, "center", BColor))
            contentT.Rows.Add(tRow)  'row=4

            '第五列
            tRow = New TableRow()
            tRow.Controls.Add(CellSet("問題類型", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CellSet(drL(7), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CellSet("分類說明", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CellSet(drL(8), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            tRow.Controls.Add(CellSet("版本/規格/序號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
            tRow.Controls.Add(CellSet(drL(9), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

            If (sfid = 22) Then
                tRow.Controls.Add(CommUtil.CellSet("當責FAE", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(10), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            ElseIf (sfid = 3) Then
                tRow.Controls.Add(CommUtil.CellSet("品管", 1, 2, False, itemlabelwidth, 0, "center", BColor))
                tRow.Controls.Add(CommUtil.CellSet(drL(17), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
            End If
            contentT.Rows.Add(tRow) 'row=5

            '第六列
            tRow = New TableRow()
            tRow.Controls.Add(CellSet("問題描述", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CellSet(CommUtil.TextTransToHtmlFormat(drL(11)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))

            tRow.Controls.Add(CellSet("現場臨時故障排除的處理過程", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CellSet(CommUtil.TextTransToHtmlFormat(drL(12)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))
            contentT.Rows.Add(tRow) 'row=6

            '第七列
            tRow = New TableRow()
            tRow.Controls.Add(CellSet("故障品的驗證過程", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CellSet(CommUtil.TextTransToHtmlFormat(drL(13)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))

            tRow.Controls.Add(CellSet("備註", 1, 2, False, itemlabelwidth, 100, "center", BColor))
            tRow.Controls.Add(CellSet(CommUtil.TextTransToHtmlFormat(drL(14)), 1, 6, False, itemlabelwidth, 100, "left", WhiteColor))
            contentT.Rows.Add(tRow) 'row=7
        End If
        drL.Close()
        connL.Close()
    End Sub
    'Sub InitTableToSfid3()
    '    Dim tCell As TableCell
    '    Dim tRow As TableRow
    '    Dim connL As New SqlConnection
    '    Dim drL As SqlDataReader
    '    Dim BColor, NeedInputColor, WhiteColor As Drawing.Color
    '    Dim tImage As Image
    '    Dim tChk As CheckBox
    '    NeedInputColor = Drawing.Color.AntiqueWhite
    '    WhiteColor = Drawing.Color.White
    '    tRow = New TableRow()
    '    For j = 1 To 16
    '        tCell = New TableCell
    '        tCell.BorderWidth = 0
    '        tCell.Width = 200
    '        tCell.HorizontalAlign = HorizontalAlign.Center
    '        tRow.Controls.Add(tCell)
    '    Next
    '    FormLogoTitleT.Rows.Add(tRow) 'row=0
    '    BColor = System.Drawing.Color.LightBlue
    '    FormLogoTitleT.Font.Name = "標楷體"
    '    FormLogoTitleT.Font.Size = 12
    '    tRow = New TableRow()
    '    tRow.Font.Bold = True
    '    tCell = New TableCell
    '    tCell.BorderWidth = 0
    '    tCell.HorizontalAlign = HorizontalAlign.Left
    '    tCell.ColumnSpan = 1
    '    tImage = New Image
    '    tImage.ID = "image_logo"
    '    tImage.ImageUrl = "~/image/jetlog80%.jpg"
    '    tCell.Controls.Add(tImage)
    '    tRow.Controls.Add(tCell)
    '    tCell = New TableCell
    '    tCell.BorderWidth = 0
    '    tCell.Font.Size = 24
    '    tCell.ColumnSpan = 14
    '    tCell.HorizontalAlign = HorizontalAlign.Center
    '    tCell.Text = "捷智科技 廠內機台問題反應單"
    '    tRow.Controls.Add(tCell)
    '    tCell = New TableCell
    '    tCell.Font.Size = 12
    '    tCell.BorderWidth = 0
    '    tCell.HorizontalAlign = HorizontalAlign.Right
    '    tCell.VerticalAlign = VerticalAlign.Bottom
    '    tCell.ColumnSpan = 1
    '    If (docnum <> 0) Then
    '        If (docstatus = "E" Or docstatus = "D") Then
    '            SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
    '            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
    '            If (drL.HasRows) Then
    '                drL.Read()
    '                tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
    '            End If
    '            drL.Close()
    '            connL.Close()
    '        Else
    '            SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
    '            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
    '            If (drL.HasRows) Then
    '                drL.Read()
    '                tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
    '            End If
    '            drL.Close()
    '            connL.Close()
    '        End If
    '    End If
    '    tRow.Controls.Add(tCell)
    '    FormLogoTitleT.Rows.Add(tRow) 'row=1

    '    tRow = New TableRow()
    '    For j = 1 To 16
    '        tCell = New TableCell
    '        tCell.BorderWidth = 0
    '        tCell.Width = 200
    '        tCell.HorizontalAlign = HorizontalAlign.Center
    '        tRow.Controls.Add(tCell)
    '    Next
    '    contentT.Rows.Add(tRow) 'row=0
    '    BColor = System.Drawing.Color.LightBlue
    '    contentT.Font.Name = "標楷體"
    '    contentT.Font.Size = 14
    '    tRow = New TableRow()
    '    tRow.Controls.Add(CommUtil.CellSet("機台基本資訊", 1, 16, False, 0, 0, "center", BColor))
    '    contentT.Rows.Add(tRow)  'row=1
    '    Dim itemlabelwidth As Integer = 200
    '    SqlCmd = "Select T0.reportdate,T0.machinetype,T0.cusname,mo,model,wo,shipdate, " &
    '         "problemtype,typedescrip,verandspec,qcperson, " &
    '         "problemdescrip,processdescrip,verifydescrip,problemnote,noassign " &
    '         "FROM [dbo].[@XFMRT] T0 WHERE T0.[docentry] =" & docnum
    '    drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

    '    If (drL.HasRows) Then
    '        drL.Read()
    '        '聯絡表頭 
    '        tRow = New TableRow()
    '        'CellSet(Text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, txtid As String, width As Integer, height As Integer, align As String)
    '        tRow.Controls.Add(CellSet("回報日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(0), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CellSet("產品別", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        'tRow.Controls.Add(CellSet("", 1, 2, False, 0, 0, "center", False))
    '        tRow.Controls.Add(CellSet(drL(1), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CellSet("客戶名稱", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(2), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CellSet("機台料號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(3), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
    '        contentT.Rows.Add(tRow) 'row=2
    '        '第三列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CellSet("機台型號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(4), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CommUtil.CellSet("工單號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tCell = CommUtil.CellSet(drL(5), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor)
    '        tCell.Font.Size = 10
    '        tRow.Controls.Add(tCell)

    '        tRow.Controls.Add(CellSet("出貨日期", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(6), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        'tCell = New TableCell
    '        tCell = CellSet("", 1, 4, False, itemlabelwidth * 2, 0, "center", WhiteColor)
    '        tChk = New CheckBox
    '        tChk.ID = "chk_noassign"
    '        tChk.Text = "無指定機台"
    '        If (drL(15) = 1) Then
    '            tChk.Checked = True
    '        Else
    '            tChk.Checked = False
    '        End If
    '        tCell.Controls.Add(tChk)
    '        tRow.Controls.Add(tCell)

    '        contentT.Rows.Add(tRow) 'row=3

    '        tRow = New TableRow()
    '        tRow.Controls.Add(CellSet("資訊記錄", 1, 16, False, 0, 0, "center", BColor))
    '        contentT.Rows.Add(tRow)  'row=4

    '        '第五列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CellSet("問題類型", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(7), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CellSet("分類說明", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(8), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CellSet("版本/規格/序號", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(9), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))

    '        tRow.Controls.Add(CellSet("品管", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(10), 1, 2, False, itemlabelwidth, 0, "center", WhiteColor))
    '        contentT.Rows.Add(tRow) 'row=5

    '        '第六列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CellSet("問題描述", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(11), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))

    '        tRow.Controls.Add(CellSet("現場臨時故障排除的處理過程", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(12), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))
    '        contentT.Rows.Add(tRow) 'row=6

    '        '第七列
    '        tRow = New TableRow()
    '        tRow.Controls.Add(CellSet("故障品的驗證過程", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(13), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))

    '        tRow.Controls.Add(CellSet("備註", 1, 2, False, itemlabelwidth, 0, "center", BColor))
    '        tRow.Controls.Add(CellSet(drL(14), 1, 6, False, itemlabelwidth, 0, "left", WhiteColor))
    '        contentT.Rows.Add(tRow) 'row=7
    '    End If
    '    drL.Close()
    '    connL.Close()
    'End Sub
    Sub InitTableToSfid12()
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim tImage As Image
        'Dim HyperBtn As LinkButton
        Dim tTxt As TextBox
        Dim Labelx As Label
        Dim ce As CalendarExtender
        Dim rRBL As RadioButtonList
        Dim cChk As CheckBox
        Dim BColor As Drawing.Color
        Dim fontstyle As String = "標楷體"
        Dim fontsize As Integer = 16
        BColor = System.Drawing.Color.LightBlue
        contentT.Font.Name = "標楷體"
        contentT.Font.Size = 14
        tRow = New TableRow()
        tRow.Font.Bold = True
        For j = 1 To 6
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow)
        'row=0 Title
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 36
        tCell.ColumnSpan = 4
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "捷智科技 AOI/SPI 內部發包單"
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        'Row = 1 cell 1-3
        tRow = New TableRow()
        'tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "發包單位"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'Row = 1 cell 2-3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_area"
        rRBL.Items.Add("台北捷智")
        rRBL.Items.Add("深圳捷智通")
        rRBL.Items.Add("昆山捷豐")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=1 cell 4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "負責業務"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=1 cell 5-6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_sales"
        'tTxt.Font.Name = fontstyle
        'tTxt.Font.Size = fontsize
        tTxt.Width = 100
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=2 cell=1
        tRow = New TableRow()
        'tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "捷智機型"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=2 cell 2-3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox
        tTxt.ID = "txt_model"
        'tTxt.Font.Name = fontstyle
        'tTxt.Font.Size = fontsize
        tTxt.Width = 400
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)
        'row=2 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "客戶名稱"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=2 cell=5
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_customer"
        'tTxt.Font.Name = fontstyle
        'tTxt.Font.Size = fontsize
        tTxt.Width = 130
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)
        'row=2 cell=6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        Labelx = New Label
        Labelx.ID = "label_2_6"
        Labelx.Text = "數量:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_amount"
        tTxt.Width = 30
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=3 cell=1
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "出貨型號"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=3 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_shipmodel"
        tTxt.Width = 180
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)
        'row=3 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "工廠出貨日期"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=3 cell=5-6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_shipdate"
        tTxt.Width = 120
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        ce = New CalendarExtender
        ce.TargetControlID = tTxt.ID
        ce.ID = "ce_shipdate"
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=4 cell=1-6
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.ColumnSpan = 6
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "系統規格要求"
        tCell.Font.Bold = True
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=5 cell=1
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "產品大小/重量/厚度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=5 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label
        Labelx.ID = "label_5_2_1"
        Labelx.Text = "客戶端產品大小:長"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutlength"
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_2"
        Labelx.Text = "mm/寬"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutwidth"
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_3"
        Labelx.Text = "mm<br>產品重量:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutweight"
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_4"
        Labelx.Text = "KG<br>產品厚度:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_uutthick"
        tTxt.Width = 60
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_5"
        Labelx.Text = "mm(與皮帶型式有關)"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        'row=5 cell 4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "生產線高度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=5 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_plheight"
        rRBL.Items.Add("900+-20 mm")
        rRBL.Items.Add("其它特殊高度")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tTxt = New TextBox()
        tTxt.ID = "txt_plotherheight"
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_6"
        Labelx.Text = "&nbsp&nbsp+-&nbsp&nbsp"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_plotherheighttol"
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_5_2_7"
        Labelx.Text = "&nbsp&nbspmm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=6 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "是否有載具"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=6 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_withfixture"
        rRBL.Items.Add("無載具")
        rRBL.Items.Add("有載具")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        Labelx = New Label
        Labelx.ID = "label_6_2_1"
        Labelx.Text = "載具大小:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_fixturesize"
        tTxt.Width = 80
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_6_2_2"
        Labelx.Text = "&nbsp&nbspmm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        'row=6 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "載具樣式"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=6 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Text = "請詳細確認載具樣式,尤其是與皮帶接觸邊之幾何構形;必要時拍圖片;或向客戶取得圖檔;或寄回實板參考"
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=7 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "進板方向"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=7 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_pcbdir"
        rRBL.Items.Add("由左 --> 右")
        rRBL.Items.Add("由右 --> 左")
        rRBL.Items.Add("雙向")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor ' System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=7 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "Cycle Time"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=7 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label
        Labelx.ID = "label_7_5_1"
        Labelx.Text = "客戶具體板子大小:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_pcbsizeX"
        tTxt.Width = 40
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_7_5_2"
        Labelx.Text = "&nbsp*&nbsp"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_pcbsizeY"
        tTxt.Width = 40
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_7_5_3"
        Labelx.Text = "mm<br>"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_cycletime"
        tTxt.Width = 40
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_7_5_4"
        Labelx.Text = "秒<br>"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=8 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "作業系統"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=8 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_oslang"
        rRBL.Items.Add("繁體版")
        rRBL.Items.Add("簡體版")
        rRBL.Items.Add("英文版")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=8 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "加裝Z軸"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=8 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        cChk = New CheckBox
        cChk.ID = "chk_upz"
        cChk.Text = ""
        tCell.Controls.Add(cChk)
        Labelx = New Label
        Labelx.ID = "label_8_5_1"
        Labelx.Text = "上Z軸行程:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_zmm"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 40
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_8_5_2"
        Labelx.Text = "mm&nbsp&nbsp"
        tCell.Controls.Add(Labelx)

        cChk = New CheckBox
        cChk.ID = "chk_downz"
        cChk.Text = ""
        tCell.Controls.Add(cChk)
        Labelx = New Label
        Labelx.ID = "label_8_5_3"
        Labelx.Text = "下Z軸行程:"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_dzmm"
        tTxt.Font.Name = "標楷體"
        tTxt.Font.Size = 14
        tTxt.Width = 40
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_8_5_4"
        Labelx.Text = "mm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=9 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "待測板上下預留空間<br>電路板上下方零件高度"
        tCell.Font.Bold = True
        tCell.Font.Size = 15
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=9 cell 2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        Labelx = New Label
        Labelx.ID = "label_9_2_1"
        Labelx.Text = "設備規格標準:上"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_topspace"
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_9_2_2"
        Labelx.Text = "mm/下"
        tCell.Controls.Add(Labelx)
        tTxt = New TextBox()
        tTxt.ID = "txt_botspace"
        tTxt.Width = 50
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_9_2_3"
        Labelx.Text = "mm"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        'row=9 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "上下鏡組需求"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=9 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_tblens"
        rRBL.Items.Add("上鏡組")
        rRBL.Items.Add("下鏡組")
        rRBL.Items.Add("上下鏡組")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        cChk = New CheckBox
        cChk.ID = "chk_sidecamera"
        cChk.Text = "側面相機"
        tCell.Controls.Add(cChk)
        tRow.Controls.Add(tCell)

        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=10 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "相機規格"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=10 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_camerapixel"
        rRBL.Items.Add("6.5M")
        rRBL.Items.Add("12M")
        rRBL.Items.Add("21M")
        rRBL.Items.Add("25M")
        rRBL.Items.Add("37M")
        rRBL.Items.Add("其它")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=10 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "RGB光盤規格"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=10 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_rgb"
        rRBL.Items.Add("自製(如V5/V7)")
        rRBL.Items.Add("外購(如OPT)")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=11 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "系統解析度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=11 cell=2-5
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 4
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_resolution"
        rRBL.Items.Add("20um")
        rRBL.Items.Add("15um")
        rRBL.Items.Add("12um")
        rRBL.Items.Add("10um")
        rRBL.Items.Add("8um")
        rRBL.Items.Add("7um")
        rRBL.Items.Add("6um")
        rRBL.Items.Add("5.5um")
        rRBL.Items.Add("5um")
        rRBL.Items.Add("3um")
        rRBL.Items.Add("2.5um")
        rRBL.Items.Add("其它")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=11 cell=6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt = New TextBox()
        tTxt.ID = "txt_otherresolution"
        tTxt.Width = 50
        tCell.Controls.Add(tTxt)
        Labelx = New Label
        Labelx.ID = "label_11_2"
        Labelx.Text = "&nbsp&nbspum"
        tCell.Controls.Add(Labelx)
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=12 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "光源控制器"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=12 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_rbgcontrol"
        rRBL.Items.Add("自製(如V5/V7)")
        rRBL.Items.Add("外購(如OPT)")
        rRBL.RepeatDirection = RepeatDirection.Vertical
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=12 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "同軸光"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=12 cell=5
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_coaxialinstall"
        rRBL.Items.Add("加裝")
        rRBL.Items.Add("不加裝")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        'row=12 cell=6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_coaxialcolor"
        rRBL.Items.Add("紅")
        rRBL.Items.Add("白")
        rRBL.Items.Add("紅白")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        tCell.Controls.Add(rRBL)
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True
        'row=13 cell=1
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "軌道皮帶型式與抗靜電需求"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=13 cell=2,3
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Left
        rRBL = New RadioButtonList()
        rRBL.ID = "rbl_belttype"
        rRBL.Items.Add("平面皮帶(產品重量 < 5 KG)")
        rRBL.Items.Add("時規皮帶(產品重量 > 5 KG)")
        rRBL.RepeatDirection = RepeatDirection.Horizontal
        rRBL.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(rRBL)
        cChk = New CheckBox
        cChk.ID = "chk_flux"
        cChk.Text = "考慮助焊劑沾黏問題"
        tCell.Controls.Add(cChk)
        cChk = New CheckBox
        cChk.ID = "chk_anti"
        cChk.Text = "具備抗靜電"
        tCell.Controls.Add(cChk)
        tRow.Controls.Add(tCell)
        'row=13 cell=4
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.BackColor = Drawing.Color.Beige
        tCell.Text = "軌道皮帶露出寬度(與電路板或載具的接處寬度"
        tCell.Font.Bold = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        'row=13 cell=5,6
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.Text = "設備規格為3.5mm以上"
        tCell.HorizontalAlign = HorizontalAlign.Center
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'row=14 cell=1-6
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.ColumnSpan = 6
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.Text = "備註:"
        tCell.Font.Bold = True
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        tRow = New TableRow()
        'tRow.Font.Bold = True

        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 6
        tTxt = New TextBox()
        tTxt.ID = "txt_memo"
        tTxt.Width = 1250
        'tTxt.Height = 200
        tTxt.TextMode = TextBoxMode.MultiLine
        tTxt.Rows = 6
        tTxt.BackColor = BColor 'System.Drawing.Color.LightGreen
        tCell.Controls.Add(tTxt)
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)

    End Sub
    Sub ShowXMSCT()
        Dim connL, connL2 As New SqlConnection
        Dim drL, drL2 As SqlDataReader
        Dim ddl_model As String
        Dim BColor As Drawing.Color
        BColor = System.Drawing.Color.LightBlue
        SqlCmd = "Select rbl_area, txt_amount, rbl_plheight, rbl_withfixture, rbl_pcbdir, rbl_oslang, rbl_camerapixel, rbl_rgb, " &
                         "rbl_resolution, rbl_rbgcontrol, rbl_coaxialinstall, rbl_coaxialcolor, rbl_belttype, chk_flux, chk_anti, " &
                         "txt_sales, ddl_model, txt_customer, txt_shipmodel, txt_shipdate, txt_uutlength, txt_uutwidth, txt_uutweight, " &
                         "txt_uutthick, txt_plotherheight, txt_plotherheighttol, txt_fixturesize, txt_pcbsizeX, txt_pcbsizeY, txt_cycletime, " &
                         "txt_zmm, txt_topspace, txt_botspace, txt_otherresolution, txt_memo,txt_dzmm,rbl_tblens,chk_sidecamera,chk_upz,chk_downz " &
                     "FROM [dbo].[@XMSCT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            If (drL(16) <> "請選擇") Then
                SqlCmd = "SELECT T0.u_mdesc " &
                     "FROM dbo.[@UMMD] T0 where T0.u_model='" & drL(16) & "'"
                drL2 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL2)
                If (drL2.HasRows) Then
                    drL2.Read()
                    ddl_model = drL(16) & "-" & drL2(0)
                Else
                    ddl_model = drL(16)
                End If
                drL2.Close()
                connL2.Close()
            Else
                ddl_model = drL(16)
            End If
            CType(contentT.FindControl("rbl_area"), RadioButtonList).SelectedIndex = drL(0)
            If (drL(1) <> 0) Then
                CType(contentT.FindControl("txt_amount"), TextBox).Text = drL(1)
            Else
                CType(contentT.FindControl("txt_amount"), TextBox).Text = ""
            End If
            CType(contentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex = drL(2)
            CType(contentT.FindControl("rbl_withfixture"), RadioButtonList).SelectedIndex = drL(3)
            CType(contentT.FindControl("rbl_pcbdir"), RadioButtonList).SelectedIndex = drL(4)
            CType(contentT.FindControl("rbl_oslang"), RadioButtonList).SelectedIndex = drL(5)
            CType(contentT.FindControl("rbl_camerapixel"), RadioButtonList).SelectedIndex = drL(6)
            CType(contentT.FindControl("rbl_rgb"), RadioButtonList).SelectedIndex = drL(7)
            CType(contentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex = drL(8)
            CType(contentT.FindControl("rbl_rbgcontrol"), RadioButtonList).SelectedIndex = drL(9)
            CType(contentT.FindControl("rbl_coaxialinstall"), RadioButtonList).SelectedIndex = drL(10)
            CType(contentT.FindControl("rbl_coaxialcolor"), RadioButtonList).SelectedIndex = drL(11)
            CType(contentT.FindControl("rbl_belttype"), RadioButtonList).SelectedIndex = drL(12)
            CType(contentT.FindControl("chk_flux"), CheckBox).Checked = drL(13)
            CType(contentT.FindControl("chk_anti"), CheckBox).Checked = drL(14)

            CType(contentT.FindControl("txt_sales"), TextBox).Text = drL(15)
            CType(contentT.FindControl("txt_model"), TextBox).Text = ddl_model
            CType(contentT.FindControl("txt_customer"), TextBox).Text = drL(17)
            CType(contentT.FindControl("txt_shipmodel"), TextBox).Text = drL(18)
            CType(contentT.FindControl("txt_shipdate"), TextBox).Text = drL(19)
            CType(contentT.FindControl("txt_uutlength"), TextBox).Text = drL(20)
            CType(contentT.FindControl("txt_uutwidth"), TextBox).Text = drL(21)
            CType(contentT.FindControl("txt_uutweight"), TextBox).Text = drL(22)
            CType(contentT.FindControl("txt_uutthick"), TextBox).Text = drL(23)
            CType(contentT.FindControl("txt_plotherheight"), TextBox).Text = drL(24)
            CType(contentT.FindControl("txt_plotherheighttol"), TextBox).Text = drL(25)
            CType(contentT.FindControl("txt_fixturesize"), TextBox).Text = drL(26)
            CType(contentT.FindControl("txt_pcbsizeX"), TextBox).Text = drL(27)
            CType(contentT.FindControl("txt_pcbsizeY"), TextBox).Text = drL(28)
            CType(contentT.FindControl("txt_cycletime"), TextBox).Text = drL(29)
            CType(contentT.FindControl("txt_zmm"), TextBox).Text = drL(30)
            CType(contentT.FindControl("txt_topspace"), TextBox).Text = drL(31)
            CType(contentT.FindControl("txt_botspace"), TextBox).Text = drL(32)
            CType(contentT.FindControl("txt_otherresolution"), TextBox).Text = drL(33)
            If ((System.Text.RegularExpressions.Regex.Matches(drL(34), "\r\n").Count + 1) <= 6) Then
                CType(contentT.FindControl("txt_memo"), TextBox).Rows = 6
            Else
                CType(contentT.FindControl("txt_memo"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(34), "\r\n").Count + 1
            End If
            CType(contentT.FindControl("txt_memo"), TextBox).Text = drL(34)
            If (CType(contentT.FindControl("rbl_plheight"), RadioButtonList).SelectedIndex = 1) Then
                CType(contentT.FindControl("txt_plotherheight"), TextBox).BackColor = BColor 'System.Drawing.Color.LightGreen
                CType(contentT.FindControl("txt_plotherheighttol"), TextBox).BackColor = BColor 'System.Drawing.Color.LightGreen
            End If
            If (CType(contentT.FindControl("rbl_resolution"), RadioButtonList).SelectedIndex = 11) Then
                CType(contentT.FindControl("txt_otherresolution"), TextBox).BackColor = BColor 'System.Drawing.Color.LightGreen
            End If
            CType(contentT.FindControl("txt_dzmm"), TextBox).Text = drL(35)
            CType(contentT.FindControl("rbl_tblens"), RadioButtonList).SelectedIndex = drL(36)
            CType(contentT.FindControl("chk_sidecamera"), CheckBox).Checked = drL(37)
            CType(contentT.FindControl("chk_upz"), CheckBox).Checked = drL(38)
            CType(contentT.FindControl("chk_downz"), CheckBox).Checked = drL(39)
        End If
        drL.Close()
        connL.Close()
        If (docstatus <> "A" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "B") Then
            contentT.Enabled = False
        End If
    End Sub
    Sub CreateSignOffFlowField(i As Integer)
        Dim tCell As TableCell
        Dim tRow As TableRow
        'Dim TxtComm As TextBox
        Dim tImage As Image
        Dim width As Integer
        If (i Mod 2) Then
            SignT.Font.Name = "標楷體"
            SignT.Font.Size = 12
            width = 80
            tRow = New TableRow()
            'If (i Mod 2) Then
            'tRow.BackColor = Drawing.Color.Cornsilk
            'End If
            tCell = New TableCell() '職位資訊_L
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() 'image_L
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.RowSpan = 3
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tImage = New Image
            tImage.ID = "image_signL_" & i
            tCell.Controls.Add(tImage)
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '意見_L
            tCell.BorderWidth = 1
            tCell.RowSpan = 3
            tCell.ColumnSpan = 2
            tCell.Wrap = True
            tCell.Width = width
            tCell.HorizontalAlign = HorizontalAlign.Left
            'TxtComm = New TextBox
            'TxtComm.ID = "txt_commL_" & i
            'TxtComm.TextMode = TextBoxMode.MultiLine
            'TxtComm.Rows = 4
            'TxtComm.Width = 300 'pixel
            'TxtComm.Font.Size = 12 'point
            'TxtComm.BorderWidth = 0
            'TxtComm.Enabled = False
            'tCell.Controls.Add(TxtComm)
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '職位資訊_R '因此列此Cell為此列第四個create , 故序號為3
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() 'image_R
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.RowSpan = 3
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.HorizontalAlign = HorizontalAlign.Center
            tImage = New Image
            tImage.ID = "image_signR_" & i + 1
            tCell.Controls.Add(tImage)
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '意見_R
            tCell.BorderWidth = 1
            tCell.RowSpan = 3
            tCell.ColumnSpan = 2
            tCell.Wrap = True
            tCell.Width = width
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.HorizontalAlign = HorizontalAlign.Left
            'TxtComm = New TextBox
            'TxtComm.ID = "txt_commR_" & i + 1
            'TxtComm.TextMode = TextBoxMode.MultiLine
            'TxtComm.Rows = 4
            'TxtComm.Width = 300 'pixel
            'TxtComm.Font.Size = 12 'point
            'TxtComm.BorderWidth = 0
            'TxtComm.Enabled = False
            'TxtComm.BackColor = Drawing.Color.Cornsilk
            'tCell.Controls.Add(TxtComm)
            tRow.Cells.Add(tCell)
            SignT.Rows.Add(tRow)

            'row=1
            tRow = New TableRow()
            tCell = New TableCell() '簽核人資訊L'因此列此Cell為此列第一個create , 故序號為0
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)

            tCell = New TableCell() '簽核人資訊R '因此列此Cell為此列第二個create , 故序號為1
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)
            SignT.Rows.Add(tRow)
            'row=2
            tRow = New TableRow()
            tCell = New TableCell() '簽核日期資訊L
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)
            tCell = New TableCell() '簽核日期資訊R
            tCell.BorderWidth = 1
            tCell.Width = width
            tCell.Wrap = False
            tCell.BackColor = Drawing.Color.Cornsilk
            tCell.ColumnSpan = 1
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Cells.Add(tCell)
            SignT.Rows.Add(tRow)
        End If
    End Sub
    Sub PutDataToSignOffFlow()
        Dim uid, uname, upos, comment, deptdesc, areadesc As String
        Dim seq, status, row As Integer
        Dim signdate, agnid As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim rowsbychara, rowsbynewline, memorows As Integer
        row = 0
        deptdesc = ""
        areadesc = ""
        'If (docstatus = "F" Or docstatus = "T") Then
        SqlCmd = "select uid,uname,upos,comment,seq,status,IsNull(signdate,''),agnid from [dbo].[@XSPWT] where signprop=0 and docentry=" & docnum & " order by seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            Do While (dr.Read())
                uid = dr(0)
                uname = dr(1)
                If (dr(2) <> "NA") Then
                    upos = dr(2)
                Else
                    upos = ""
                End If
                comment = dr(3)
                seq = dr(4)
                status = dr(5)
                agnid = dr(7)
                If (dr(6) = "1900/01/01") Then
                    signdate = "NA"
                Else
                    signdate = dr(6)
                End If
                'row = (seq - 1) * 3
                'row = (seq - 1) * 3
                SqlCmd = "select T1.deptdesc,T2.areadesc,T0.position from dbo.[user] T0 Inner join dbo.[dept] T1 on T0.grp=T1.deptcode " &
                        "Inner Join dbo.[branch] T2 on T0.branch=T2.areacode where T0.id='" & uid & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr1.HasRows) Then
                    dr1.Read()
                    deptdesc = dr1(0)
                    areadesc = dr1(1)
                Else
                    CommUtil.ShowMsg(Me, "沒找到id為" & uid & "之資料,請檢查")
                End If
                dr1.Close()
                conn.Close()
                If (seq Mod 2) Then
                    SignT.Rows(row).Cells(0).Text = areadesc & "  " & deptdesc 'upos
                    'SignT.Rows(row + 1).Cells(0).Text = uname
                    SignT.Rows(row + 2).Cells(0).Text = signdate
                    If (agnid = "") Then
                        SignT.Rows(row + 1).Cells(0).Text = uname & " " & upos '"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    Else
                        Dim agnname As String
                        agnname = ""
                        SqlCmd = "select name from dbo.[user] where id='" & agnid & "'"
                        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            agnname = drL(0)
                        End If
                        drL.Close()
                        connL.Close()
                        'SignT.Rows(row).Cells(1).Text = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>代理人:" & agnname
                        SignT.Rows(row + 1).Cells(0).Text = uname & "&nbsp;&nbsp;待理人:" & agnname
                    End If
                    If (sfid = 16 And seq = 1) Then
                        SqlCmd = "Select T0.id,T0.createid,T0.idname FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            If (drL(0) <> drL(1)) Then
                                SignT.Rows(row + 1).Cells(0).Text = SignT.Rows(row + 1).Cells(0).Text & "(替 " & drL(2) & " 發送)"
                            End If
                        End If
                        drL.Close()
                        connL.Close()
                    End If


                    SignT.Rows(row).Cells(2).Text = CommUtil.TextTransToHtmlFormat(comment) 'abcde

                    'rowsbynewline = System.Text.RegularExpressions.Regex.Matches(comment, "\r\n").Count + 1
                    'rowsbychara = (comment.Length / 40) + 1
                    'If (rowsbynewline >= rowsbychara) Then
                    '    memorows = rowsbynewline
                    'Else
                    '    memorows = rowsbychara
                    'End If
                    'If (memorows <= 4) Then
                    '    CType(SignT.FindControl("txt_commL_" & seq), TextBox).Rows = 4
                    'Else
                    '    CType(SignT.FindControl("txt_commL_" & seq), TextBox).Rows = memorows
                    'End If

                    'CType(SignT.FindControl("txt_commL_" & seq), TextBox).Text = comment

                    If (status = 2 Or status = 100) Then
                        CType(SignT.FindControl("image_signL_" & seq), Image).ImageUrl = "~/image/ok1.jpg"
                    ElseIf (status = 3 Or status = 5) Then
                        CType(SignT.FindControl("image_signL_" & seq), Image).ImageUrl = "~/image/rj1.jpg"
                    ElseIf (status = 10) Then
                        CType(SignT.FindControl("image_signL_" & seq), Image).ImageUrl = "~/image/skip1.jpg"
                    End If
                Else
                    SignT.Rows(row).Cells(3).Text = areadesc & "  " & deptdesc 'upos '因此列此Cell為此列第四個create , 故序號為3
                    'SignT.Rows(row + 1).Cells(1).Text = uname '因此列此Cell為此列第二個create , 故序號為1
                    SignT.Rows(row + 2).Cells(1).Text = signdate
                    If (agnid = "") Then
                        SignT.Rows(row + 1).Cells(1).Text = uname & " " & upos '"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
                    Else
                        Dim agnname As String
                        agnname = ""
                        SqlCmd = "select name from dbo.[user] where id='" & agnid & "'"
                        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            agnname = drL(0)
                        End If
                        drL.Close()
                        connL.Close()
                        'SignT.Rows(row).Cells(1).Text = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & uname & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>代理人:" & agnname
                        SignT.Rows(row + 1).Cells(1).Text = uname & "&nbsp;&nbsp;待理人:" & agnname
                    End If
                    If (sfid = 16 And seq = 1) Then
                        SqlCmd = "Select T0.id,T0.createid,T0.idname FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
                        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                        If (drL.HasRows) Then
                            drL.Read()
                            If (drL(0) <> drL(1)) Then
                                SignT.Rows(row + 1).Cells(1).Text = SignT.Rows(row + 1).Cells(1).Text & "(替 " & drL(2) & " 發送)"
                            End If
                        End If
                        drL.Close()
                        connL.Close()
                    End If
                    SignT.Rows(row).Cells(5).Text = CommUtil.TextTransToHtmlFormat(comment)
                    'rowsbynewline = System.Text.RegularExpressions.Regex.Matches(comment, "\r\n").Count + 1
                    'rowsbychara = (comment.Length / 40) + 1
                    'If (rowsbynewline >= rowsbychara) Then
                    '    memorows = rowsbynewline
                    'Else
                    '    memorows = rowsbychara
                    'End If
                    'If (memorows <= 4) Then
                    '    CType(SignT.FindControl("txt_commR_" & seq), TextBox).Rows = 4
                    'Else
                    '    CType(SignT.FindControl("txt_commR_" & seq), TextBox).Rows = memorows
                    'End If
                    'CType(SignT.FindControl("txt_commR_" & seq), TextBox).Text = comment
                    If (status = 2 Or status = 100) Then
                        CType(SignT.FindControl("image_signR_" & seq), Image).ImageUrl = "~/image/ok1.jpg"
                    ElseIf (status = 3 Or status = 5) Then
                        CType(SignT.FindControl("image_signR_" & seq), Image).ImageUrl = "~/image/rj1.jpg"
                    ElseIf (status = 10) Then
                        CType(SignT.FindControl("image_signR_" & seq), Image).ImageUrl = "~/image/skip1.jpg"
                    End If
                    row = row + 3
                End If
            Loop
        End If
        If (seq Mod 2) Then
            SignT.Rows(row).Cells(3).Text = ""
            SignT.Rows(row + 1).Cells(1).Text = ""
            SignT.Rows(row + 2).Cells(1).Text = ""
            SignT.Rows(row).Cells(4).Text = ""
            SignT.Rows(row).Cells(5).Text = ""
        End If
        dr.Close()
        connsap.Close()
        'End If
    End Sub
    'Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
    '    Const OUTPUT_FILENAME As String = "renderedpage.html"
    '    Dim renderedOutput As StringBuilder = Nothing
    '    Dim strWriter As IO.StringWriter = Nothing
    '    Dim tWriter As HtmlTextWriter = Nothing
    '    Dim outputStream As IO.FileStream = Nothing
    '    Dim sWriter As IO.StreamWriter = Nothing
    '    Dim filename As String
    '    Dim nextPage As String

    '    Try
    '        'create a HtmlTextWriter to use for rendering the page
    '        renderedOutput = New StringBuilder
    '        strWriter = New IO.StringWriter(renderedOutput)
    '        tWriter = New HtmlTextWriter(strWriter)

    '        MyBase.Render(tWriter)

    '        'save the rendered output to a file
    '        filename = Server.MapPath("~\") & "renderdir\" & OUTPUT_FILENAME
    '        'filename = "c:\data\" & OUTPUT_FILENAME
    '        outputStream = New IO.FileStream(filename,
    '                              IO.FileMode.Create)
    '        sWriter = New IO.StreamWriter(outputStream)
    '        sWriter.Write(renderedOutput.ToString())
    '        sWriter.Flush()

    '        ' redirect to another page
    '        '  NOTE: Continuing with the display of this page will result in the
    '        '       page being rendered a second time which will cause an exception
    '        '        to be thrown
    '        nextPage = "DisplayMessage.aspx?" &
    '                   "PageHeader=Information" & "&" &
    '                   "Message1=HTML Output Saved To " & OUTPUT_FILENAME
    '        'Response.Redirect(nextPage)
    '        'Response.Write(renderedOutput.ToString())

    '        writer.Write(renderedOutput.ToString())
    '    Finally

    '        'clean up
    '        If (Not IsNothing(outputStream)) Then
    '            outputStream.Close()
    '        End If

    '        If (Not IsNothing(tWriter)) Then
    '            tWriter.Close()
    '        End If

    '        If (Not IsNothing(strWriter)) Then
    '            strWriter.Close()
    '        End If
    '    End Try

    '    # 将本地 source.html 文件转换为 target.pdf
    'wkhtmltopdf source.html target.pdf

    '# 如有乱码需指定编码
    'wkhtmltopdf --encoding utf-8 source.html target.pdf
    'End Sub
    Sub InitTableToSfid1()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim BColor As Drawing.Color
        Dim tImage As Image
        Dim colcountadjust As Integer
        colcountadjust = 0
        'If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
        'colcountadjust = 2
        'End If
        tRow = New TableRow()
        For j = 1 To 10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow)
        BColor = System.Drawing.Color.LightBlue
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 1
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 32
        tCell.ColumnSpan = 8
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "捷智科技 內部聯絡單"
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        tCell.ColumnSpan = 1
        If (docnum <> 0) Then
            If (docstatus = "E" Or docstatus = "D") Then
                SqlCmd = "Select convert(varchar(12), docdate, 111) from [dbo].[@XASCH] where docnum=" & docnum
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>建單日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            Else
                SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
                drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
                End If
                drL.Close()
                connL.Close()
            End If
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow)
        'contentT
        tRow = New TableRow()
        For j = 1 To 10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 200
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow) 'row=0
        contentT.Font.Name = "標楷體"
        contentT.Font.Size = 18
        '聯絡表頭 
        tRow = New TableRow()
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "送審日期"

        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        tRow.Cells.Add(tCell)

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "送審人"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        tRow.Cells.Add(tCell)
        contentT.Rows.Add(tRow) 'row=1
        '第二列
        tRow = New TableRow()
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "聯絡部門"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        tRow.Cells.Add(tCell)

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "聯絡人員"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell()
        tCell.BorderWidth = 1
        tCell.Wrap = True
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = 3
        tRow.Cells.Add(tCell)
        contentT.Rows.Add(tRow) 'row=2
        '主旨列
        tRow = New TableRow()
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 2
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "主旨"
        tCell.BackColor = Drawing.Color.LightBlue
        tRow.Controls.Add(tCell)

        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.Wrap = True
        tCell.ColumnSpan = 8
        tCell.HorizontalAlign = HorizontalAlign.Left
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow) 'row=3
        '事由說明表頭列
        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 10 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.Text = "事由說明"
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow) 'row=4
        '事由輸入列 hhhhhh
        tRow = New TableRow()
        tCell = New TableCell()
        tCell.BorderWidth = 0
        tCell.Wrap = False
        tCell.HorizontalAlign = HorizontalAlign.Left
        tCell.ColumnSpan = 10 - colcountadjust
        tRow.Cells.Add(tCell)
        TxtReason = New TextBox
        TxtReason.ID = "txt_reason"
        TxtReason.TextMode = TextBoxMode.MultiLine
        TxtReason.Font.Size = 18
        TxtReason.Rows = 7
        TxtReason.Width = 1150
        'TxtReason.BackColor = Drawing.Color.Cornsilk
        tCell.Controls.Add(TxtReason)
        tRow.Cells.Add(tCell)
        contentT.Rows.Add(tRow) 'row=5

    End Sub
    Sub ShowXGCT()
        Dim connL, connL2 As New SqlConnection
        Dim drL, drL2 As SqlDataReader
        Dim BColor As Drawing.Color
        Dim senddate, sendperson, subject As String
        BColor = System.Drawing.Color.LightBlue
        senddate = ""
        sendperson = ""
        subject = ""
        SqlCmd = "Select convert(char(12),T0.docdate,111) ,sname,subject from [dbo].[@XASCH] T0 " &
        "where docnum=" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            senddate = drL(0)
            sendperson = drL(1)
            subject = drL(2)
        End If
        drL.Close()
        connL.Close()

        SqlCmd = "Select T0.ctdept,T0.ctperson,T0.ctdescrip " &
                     "FROM [dbo].[@XGCT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

        If (drL.HasRows) Then
            drL.Read()
            contentT.Rows(2).Cells(1).Text = drL(0)
            contentT.Rows(2).Cells(3).Text = drL(1)
            'contentT.Rows(6).Cells(0).Text = drL(2)
            TxtReason.Text = drL(2)
            If (docstatus = "A" Or docstatus = "E" Or docstatus = "D") Then

            Else
                SqlCmd = "Select convert(char(12),T0.signdate,111),T0.uname " &
                        "FROM [dbo].[@XSPWT] T0 WHERE T0.[docentry] =" & docnum & "and seq=1"
                drL2 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL2)
                If (drL2.HasRows) Then
                    drL2.Read()
                    senddate = drL2(0)
                    sendperson = drL2(1)
                End If
                drL2.Close()
                connL2.Close()
            End If
            If ((System.Text.RegularExpressions.Regex.Matches(drL(2), "\r\n").Count + 1) <= 7) Then
                TxtReason.Rows = 7
            Else
                TxtReason.Rows = System.Text.RegularExpressions.Regex.Matches(drL(2), "\r\n").Count + 1
            End If
        End If
        contentT.Rows(1).Cells(1).Text = senddate
        contentT.Rows(1).Cells(3).Text = sendperson
        contentT.Rows(3).Cells(1).Text = subject
        drL.Close()
        connL.Close()
    End Sub
    Sub ShowXCMRT()
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        SqlCmd = "Select T0.reportdate,T0.machinetype,T0.cusname,cusfactory,model,machineserial,installdate, " &
                     "problemtype,typedescrip,verandspec,faeperson, " &
                     "problemdescrip,processdescrip,verifydescrip,problemnote,firstinstall,inwarranty " &
                     "FROM [dbo].[@XCMRT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)

        If (drL.HasRows) Then
            drL.Read()
            CType(contentT.FindControl("txt_reportdate"), TextBox).Text = drL(0)
            CType(contentT.FindControl("txt_machinetype"), TextBox).Text = drL(1)
            CType(contentT.FindControl("txt_cusname"), TextBox).Text = drL(2)
            CType(contentT.FindControl("txt_cusfactory"), TextBox).Text = drL(3)
            CType(contentT.FindControl("txt_model"), TextBox).Text = drL(4)
            CType(contentT.FindControl("txt_machineserial"), TextBox).Text = drL(5)
            CType(contentT.FindControl("txt_installdate"), TextBox).Text = drL(6)
            CType(contentT.FindControl("txt_problemtype"), TextBox).Text = drL(7)
            CType(contentT.FindControl("txt_typedescrip"), TextBox).Text = drL(8)
            CType(contentT.FindControl("txt_verandspec"), TextBox).Text = drL(9)
            CType(contentT.FindControl("txt_faeperson"), TextBox).Text = drL(10)
            CType(contentT.FindControl("txt_problemdescrip"), TextBox).Text = drL(11)
            CType(contentT.FindControl("txt_processdescrip"), TextBox).Text = drL(12)
            CType(contentT.FindControl("txt_verifydescrip"), TextBox).Text = drL(13)
            CType(contentT.FindControl("txt_problemnote"), TextBox).Text = drL(14)
            If (drL(15) = 1) Then
                CType(contentT.FindControl("chk_firstinstall"), CheckBox).Checked = True
            End If
            If (drL(16) = 1) Then
                CType(contentT.FindControl("chk_inwarranty"), CheckBox).Checked = True
            End If
            Dim multirow As Integer
            multirow = Math.Max((System.Text.RegularExpressions.Regex.Matches(drL(11), "\r\n").Count + 1), (System.Text.RegularExpressions.Regex.Matches(drL(12), "\r\n").Count + 1))
            If (multirow >= 5) Then
                CType(contentT.FindControl("txt_problemdescrip"), TextBox).Rows = multirow
                CType(contentT.FindControl("txt_processdescrip"), TextBox).Rows = multirow
            Else
                CType(contentT.FindControl("txt_problemdescrip"), TextBox).Rows = 5
                CType(contentT.FindControl("txt_processdescrip"), TextBox).Rows = 5
            End If
            'If ((System.Text.RegularExpressions.Regex.Matches(drL(11), "\r\n").Count + 1) <= 5) Then
            '    CType(contentT.FindControl("txt_problemdescrip"), TextBox).Rows = 5
            'Else
            '    CType(contentT.FindControl("txt_problemdescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(11), "\r\n").Count + 1
            'End If

            'If ((System.Text.RegularExpressions.Regex.Matches(drL(12), "\r\n").Count + 1) <= 5) Then
            '    CType(contentT.FindControl("txt_processdescrip"), TextBox).Rows = 5
            'Else
            '    CType(contentT.FindControl("txt_processdescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(12), "\r\n").Count + 1
            'End If

            multirow = Math.Max((System.Text.RegularExpressions.Regex.Matches(drL(13), "\r\n").Count + 1), (System.Text.RegularExpressions.Regex.Matches(drL(14), "\r\n").Count + 1))
            If (multirow >= 5) Then
                CType(contentT.FindControl("txt_problemdescrip"), TextBox).Rows = multirow
                CType(contentT.FindControl("txt_processdescrip"), TextBox).Rows = multirow
            Else
                CType(contentT.FindControl("txt_verifydescrip"), TextBox).Rows = 5
                CType(contentT.FindControl("txt_problemnote"), TextBox).Rows = 5
            End If

            'If ((System.Text.RegularExpressions.Regex.Matches(drL(13), "\r\n").Count + 1) <= 5) Then
            '    CType(contentT.FindControl("txt_verifydescrip"), TextBox).Rows = 5
            'Else
            '    CType(contentT.FindControl("txt_verifydescrip"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(13), "\r\n").Count + 1
            'End If

            'If ((System.Text.RegularExpressions.Regex.Matches(drL(14), "\r\n").Count + 1) <= 5) Then
            '    CType(contentT.FindControl("txt_problemnote"), TextBox).Rows = 5
            'Else
            '    CType(contentT.FindControl("txt_problemnote"), TextBox).Rows = System.Text.RegularExpressions.Regex.Matches(drL(14), "\r\n").Count + 1
            'End If
        End If

        drL.Close()
        connL.Close()
    End Sub
    Function CellSetWithExtender(rowspan As Integer, colspan As Integer, LBxid As String, txtid As String, ddeid As String, BColor As Drawing.Color)
        Dim tCell As New TableCell
        Dim dde As New DropDownExtender
        Dim tTxt As New TextBox
        Dim LBx As New ListBox
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        LBx.ID = LBxid
        LBx.AutoPostBack = True
        LBx.Rows = 30
        AddHandler LBx.SelectedIndexChanged, AddressOf LB_SelectedIndexChanged
        tCell.Controls.Add(LBx)
        tTxt.ID = txtid
        tTxt.BackColor = BColor
        tTxt.Width = 150
        tCell.Controls.Add(tTxt)
        dde.TargetControlID = txtid
        dde.ID = ddeid
        dde.DropDownControlID = LBxid
        tCell.Controls.Add(dde)
        Return tCell
    End Function

    Protected Sub LB_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim tTxt As TextBox
        Dim str() As String
        Dim id As String
        'Dim LBx As ListBox
        'Dim model, mdesc, mtype As String
        str = Split(sender.ID, "_")
        id = str(1)
        tTxt = contentT.FindControl("txt_" & id)
        str = Split(sender.SelectedValue, "-")
        If (id <> "model") Then
            tTxt.Text = sender.SelectedValue
        Else
            tTxt.Text = str(1)
        End If
    End Sub
    Function CellSet(text As String, rowspan As Integer, colspan As Integer, FondBold As Boolean, width As Integer, height As Integer, align As String, BColor As Drawing.Color)
        Dim tCell As TableCell
        tCell = New TableCell()
        tCell.BorderWidth = 1
        If (align = "right") Then
            tCell.HorizontalAlign = HorizontalAlign.Right
        ElseIf (align = "center") Then
            tCell.HorizontalAlign = HorizontalAlign.Center
        Else
            tCell.HorizontalAlign = HorizontalAlign.Left
        End If
        'If (color) Then
        tCell.BackColor = BColor
        'End If
        tCell.Wrap = True
        If (text <> "") Then
            tCell.Text = text
        End If
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        If (width <> 0) Then
            tCell.Width = width 'stdwidth * colspan * 0.95
        End If
        If (height <> 0) Then
            tCell.Height = height '20 * rowspan
        End If
        tCell.Font.Bold = FondBold
        Return tCell
    End Function

    Function CellSetWithTextBox(rowspan As Integer, colspan As Integer, txtid As String, multiline As Integer, fonesize As Integer, width As Integer, BColor As Drawing.Color)
        '如果textbox要以預設大小適合 Cell , 則width設為0 , 如果要調大小 , try 看數值多少
        Dim tCell As TableCell
        Dim tTxt As New TextBox
        tCell = New TableCell()
        tCell.Wrap = True
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        tTxt.ID = txtid
        tTxt.BackColor = BColor
        If (fonesize <> 0) Then
            tTxt.Font.Size = fonesize
        End If
        'tTxt.Width = tCell.Width 'stdwidth * colspan * 0.95
        'tTxt.Height = tCell.Height '20 * rowspan
        If (width <> 0) Then
            tTxt.Width = width
        End If
        If (multiline <> 0) Then
            tTxt.TextMode = TextBoxMode.MultiLine
            tTxt.Rows = multiline
        End If
        tCell.Controls.Add(tTxt)
        Return tCell
    End Function
    Function CellSetWithCalenderExtender(rowspan As Integer, colspan As Integer, txtid As String, ceid As String, BColor As Drawing.Color, width As Integer)
        Dim tCell As New TableCell
        Dim ce As New CalendarExtender
        Dim tTxt As New TextBox
        tCell.ColumnSpan = colspan
        tCell.RowSpan = rowspan
        tCell.BorderWidth = 1
        tCell.HorizontalAlign = HorizontalAlign.Center
        tTxt.ID = txtid
        If (width <> 0) Then
            tTxt.Width = width
        End If
        tTxt.Height = tCell.Height
        tTxt.BackColor = BColor
        tCell.Controls.Add(tTxt)
        ce.TargetControlID = txtid
        ce.ID = ceid
        ce.Format = "yyyy/MM/dd"
        tCell.Controls.Add(ce)
        Return tCell
    End Function
    Sub InitTableToSfid23_24()
        Dim tCell As TableCell
        Dim tRow As TableRow
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim tablerow As Integer
        Dim mcount As Integer
        Dim BColor As Drawing.Color
        Dim tImage As Image
        Dim colcountadjust As Integer
        Dim i, j As Integer
        colcountadjust = 2
        BColor = System.Drawing.Color.LightBlue
        contentT.Font.Name = "標楷體"
        contentT.Font.Size = 12
        SqlCmd = "Select count(*) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        mcount = drL(0)
        If (drL(0) <= 5) Then
            tablerow = 6
        Else
            tablerow = drL(0) + 1
        End If
        drL.Close()
        connL.Close()
        'logo
        tRow = New TableRow()
        For j = 1 To (8 - colcountadjust) '10
            tCell = New TableCell
            tCell.BorderWidth = 0
            tCell.Width = 300
            tCell.HorizontalAlign = HorizontalAlign.Center
            tRow.Controls.Add(tCell)
        Next
        FormLogoTitleT.Rows.Add(tRow)
        FormLogoTitleT.Font.Name = "標楷體"
        FormLogoTitleT.Font.Size = 12
        tRow = New TableRow()
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Left
        'tCell.ColumnSpan = 3
        tImage = New Image
        tImage.ID = "image_logo"
        tImage.ImageUrl = "~/image/jetlog80%.jpg"
        tCell.Controls.Add(tImage)
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.BorderWidth = 0
        tCell.Font.Size = 24
        tCell.ColumnSpan = 6 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 23) Then
            tCell.Text = "捷智科技 離倉料件管制單"
        ElseIf (sfid = 24) Then
            tCell.Text = "捷智科技 借入料件管制單"
        End If
        tRow.Controls.Add(tCell)
        tCell = New TableCell
        tCell.Font.Size = 12
        tCell.BorderWidth = 0
        tCell.HorizontalAlign = HorizontalAlign.Right
        tCell.VerticalAlign = VerticalAlign.Bottom
        'tCell.ColumnSpan = 3
        If (docnum <> 0) Then
            SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                drL.Read()
                tCell.Text = "單號:" & docnum & "<br>送審日期:" & drL(0)
            End If
            drL.Close()
            connL.Close()
        End If
        tRow.Controls.Add(tCell)
        FormLogoTitleT.Rows.Add(tRow)



        Dim descripreason As String
        Dim itemlabelwidth As Integer = 200
        SqlCmd = "Select descrip FROM [dbo].[@XSMLS] T0 WHERE head=1 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            descripreason = drL(0)
        Else
            CommUtil.ShowMsg(Me, "找不到此簽核單說明資料(@XSMLS),中斷產生Pdf")
            Exit Sub
        End If
        drL.Close()
        connL.Close()
        '事由說明表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        If (sfid = 23) Then
            tRow.Controls.Add(CellSet("離倉料件說明", 1, 8 - colcountadjust, False, itemlabelwidth, 0, "center", BColor))
        ElseIf (sfid = 24) Then
            tRow.Controls.Add(CellSet("借入料件說明", 1, 8 - colcountadjust, False, itemlabelwidth, 0, "center", BColor))
        End If
        contentT.Rows.Add(tRow)

        tRow = New TableRow()
        tCell = CellSet(CommUtil.TextTransToHtmlFormat(descripreason), 1, 8 - colcountadjust, False, itemlabelwidth, 100, "left", Drawing.Color.White)
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)
        '料件表頭列
        tRow = New TableRow()
        tRow.Font.Size = 16
        tRow.BackColor = Drawing.Color.LightBlue
        tRow.Font.Bold = True
        tCell = New TableCell
        tCell.BorderWidth = 1
        tCell.ColumnSpan = 8 - colcountadjust
        tCell.HorizontalAlign = HorizontalAlign.Center
        If (sfid = 23) Then
            If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                tCell.Text = "離倉料件表列"
            Else
                tCell.Text = "離倉料件表列(請由上述橘色功能列加入)"
            End If
        ElseIf (sfid = 24) Then
            If (docstatus <> "B" And docstatus <> "E" And docstatus <> "D" And docstatus <> "R" And docstatus <> "A") Then
                tCell.Text = "借入料件表列"
            Else
                tCell.Text = "借入料件表列(請由上述橘色功能列加入)"
            End If
        End If
        tRow.Controls.Add(tCell)
        contentT.Rows.Add(tRow)

        tRow = New TableRow()
        tRow.BackColor = Drawing.Color.LightGreen
        tRow.Font.Bold = True
        For i = 0 To 7 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Width = 40
            tCell.HorizontalAlign = HorizontalAlign.Center
            If (i = 0) Then
                tCell.Text = "項次"
                tCell.Width = 40
            ElseIf (i = 1) Then
                tCell.Text = "料號"
                tCell.Width = 200 '120
            ElseIf (i = 2) Then
                tCell.Text = "說明"
                tCell.Width = 500 '300
            ElseIf (i = 3) Then
                tCell.Text = "數量"
                tCell.Width = 40
            ElseIf (i = 4) Then '6
                If (sfid = 23) Then
                    tCell.Text = "離倉原因"
                ElseIf (sfid = 24) Then
                    tCell.Text = "借出原因"
                End If
                tCell.Width = 200 '160
            ElseIf (i = 5) Then
                tCell.Text = "備註"
                tCell.Width = 300 '250
            End If
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow)
        SqlCmd = "Select T0.itemcode,T0.itemname,T0.quantity,T0.price,T0.method,T0.comment,T0.num " &
                "FROM [dbo].[@XSMLS] T0 " &
                "WHERE T0.head=0 And T0.[docentry] =" & docnum & " ORDER BY T0.num"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            i = 1
            Do While (drL.Read())
                'MsgBox(drL(0))
                tRow = New TableRow()
                tRow.BorderWidth = 1
                For j = 0 To 7 - colcountadjust
                    tCell = New TableCell
                    tCell.BorderWidth = 1
                    tCell.Height = 20
                    If (j = 0 Or j = 3 Or j = 4) Then
                        tCell.HorizontalAlign = HorizontalAlign.Center
                    End If
                    If (j = 0) Then '項次
                        tCell.Text = i
                    ElseIf (j = 1) Then '料號
                        tCell.Text = drL(0)
                    ElseIf (j = 2) Then '說明
                        tCell.Text = drL(1)
                    ElseIf (j = 3) Then '數量
                        tCell.Text = CLng(drL(2))
                    ElseIf (j = 4) Then '離倉原因
                        tCell.Text = drL(4)
                    ElseIf (j = 5) Then '備註(何人何處)
                        tCell.Text = drL(5)
                    End If
                    tRow.Controls.Add(tCell)
                Next
                contentT.Rows.Add(tRow)
                i = i + 1
            Loop
        End If
        drL.Close()
        connL.Close()
        tRow = New TableRow()
        tRow.BorderWidth = 1
        For j = 0 To 7 - colcountadjust
            tCell = New TableCell
            tCell.BorderWidth = 1
            tCell.Height = 20
            tRow.Controls.Add(tCell)
        Next
        contentT.Rows.Add(tRow)
    End Sub
End Class