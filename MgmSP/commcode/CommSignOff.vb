Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Public Class CommSignOff
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Public connsap, conn As New SqlConnection
    Public SqlCmd As String
    Public dr, dr1, drsap As SqlDataReader
    Public Function createProcessDtlPDF(ByVal docnum As Long, ByVal vFormName As String, ByVal fileName As String, ByVal fpath As String, ByRef vAmount As String, ByVal vReason As String, ByVal vCurreny As String)
        Dim processDtlPDFFileName As String = ""
        Dim comment As String
        Dim t2status As Integer
        SqlCmd = "Select sname,subject,sapno,price,T1.sfname,priceunit,dept,area,T2.signdate,T2.uname,T2.uid,T2.comment,T2.status " &
        "from [dbo].[@XASCH] T0 Inner Join [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid Inner Join [dbo].[@XSPWT] T2 On T0.docnum=T2.docentry " &
        "where docnum=" & docnum & " and T2.signprop=0 order by T2.seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            processDtlPDFFileName = fpath & fileName
            Dim doc1 = New Document(PageSize.A4, 50, 50, 50, 50) '設定pagesize, 邊界left, right, top, bottom
            Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(doc1, New FileStream(fpath & fileName, FileMode.Create))
            '指定使用中文字型, 中文字才能顯示
            Dim svrPath As String = System.Web.HttpContext.Current.Server.MapPath("/") '--網站根目錄實體目錄路徑
            Dim bfChinese As BaseFont = BaseFont.CreateFont(svrPath & "/jFonts/kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
            Dim chFont As Font = New Font(bfChinese, 10)
            Dim chFont_header As Font = New Font(bfChinese, 16)
            Dim chFont_blue As Font = New Font(bfChinese, 12, Font.NORMAL, New Color(51, 0, 153))
            Dim chFont_red As Font = New Font(bfChinese, 12, Font.ITALIC, Color.RED)
            doc1.Open()

            '表格區塊
            Dim intTblWidth() As Integer = {6, 4, 4, 4}
            Dim pTable As PdfPTable = New PdfPTable(intTblWidth(0))
            pTable.TotalWidth = 550.0F
            pTable.LockedWidth = True

            Dim header As PdfPCell = New PdfPCell(New Paragraph(20.0F, vFormName, New Font(bfChinese, 20)))
            header.HorizontalAlignment = 1
            header.Padding = 10
            header.Colspan = 6
            pTable.AddCell(header)

            header = New PdfPCell(New Paragraph(20.0F, "表單編號:" & docnum, chFont_header))
            header.Padding = 5
            header.Colspan = 6
            pTable.AddCell(header)

            header = New PdfPCell(New Paragraph(20.0F, "主旨:" & vReason, chFont_header))
            header.Padding = 5
            header.Colspan = 6
            pTable.AddCell(header)

            header = New PdfPCell(New Paragraph(20.0F, "核准簽條", chFont_header))
            header.Padding = 5
            header.Colspan = 6
            pTable.AddCell(header)

            Dim columnHeader1 As PdfPCell = New PdfPCell(New Phrase(14.0F, "部門/人員/簽核時間", chFont))
            columnHeader1.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader1.Padding = 5
            columnHeader1.Colspan = 2
            pTable.AddCell(columnHeader1)
            'Dim columnHeader2 As PdfPCell = New PdfPCell(New Phrase(14.0F, "人員", chFont))
            'columnHeader2.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            'columnHeader2.Padding = 5
            'pTable.AddCell(columnHeader2)
            Dim columnHeader3 As PdfPCell = New PdfPCell(New Phrase(14.0F, "核准章", chFont))
            columnHeader3.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader3.Padding = 5
            columnHeader3.Colspan = 1
            pTable.AddCell(columnHeader3)
            'Dim columnHeader4 As PdfPCell = New PdfPCell(New Phrase(14.0F, "簽核時間", chFont))
            'columnHeader4.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            'columnHeader4.Padding = 5
            'pTable.AddCell(columnHeader4)
            Dim columnHeader4 As PdfPCell = New PdfPCell(New Phrase(14.0F, "簽核意見", chFont))
            columnHeader4.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader4.Padding = 5
            columnHeader4.Colspan = 3
            pTable.AddCell(columnHeader4)
            '簽核明細
            Dim rowCell As PdfPCell
            Dim strImageUrl As String
            Dim signatureJpg As iTextSharp.text.Image
            Dim jpgCell As PdfPCell
            Dim deptdesc, areadesc, posi, personinfo As String
            deptdesc = ""
            areadesc = ""
            posi = ""
            Do While (dr.Read())
                t2status = dr(12)
                SqlCmd = "select T1.deptdesc,T2.areadesc,T0.position from dbo.[user] T0 Inner join dbo.[dept] T1 on T0.grp=T1.deptcode " &
                        "Inner Join dbo.[branch] T2 on T0.branch=T2.areacode where T0.id='" & dr(10) & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr1.HasRows) Then
                    dr1.Read()
                    deptdesc = dr1(0)
                    areadesc = dr1(1)
                    posi = dr1(2)
                Else
                    CommUtil.ShowMsg(Me, "沒找到id為" & dr(9) & "之資料,請檢查")
                    'Exit Function
                End If
                dr1.Close()
                conn.Close()

                personinfo = areadesc & " " & deptdesc & " " & posi & vbLf & dr(9) & vbLf & dr(8)
                rowCell = New PdfPCell(New Phrase(12.0F, personinfo, chFont)) '部門
                rowCell.Colspan = 2
                pTable.AddCell(rowCell)

                '簽名圖檔
                If (t2status = 3) Then
                    strImageUrl = svrPath & "image/rjpdf.jpg"
                ElseIf (t2status = 10) Then
                    strImageUrl = svrPath & "image/skip1.jpg"
                Else
                    strImageUrl = svrPath & "image/okpdf.jpg"
                End If
                signatureJpg = iTextSharp.text.Image.GetInstance(New Uri(strImageUrl))
                ''''''''signatureJpg.ScaleToFit(100.0F, 33.0F)
                signatureJpg.ScalePercent(70.0F)
                jpgCell = New PdfPCell(signatureJpg, False)
                jpgCell.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
                jpgCell.Padding = 5
                rowCell.Colspan = 1
                pTable.AddCell(jpgCell)

                comment = dr(11)
                rowCell = New PdfPCell(New Phrase(12.0F, comment, chFont))
                rowCell.Colspan = 3
                pTable.AddCell(rowCell)
                'Next
            Loop
            doc1.Add(pTable)
            doc1.Close()
            'doc1.Dispose()
        End If
        dr.Close()
        connsap.Close()
        Return processDtlPDFFileName
    End Function
    Public Function createProcessDtlWithPricePDF(ByVal docnum As Long, ByVal vFormName As String, ByVal fileName As String, ByVal fpath As String, ByRef vAmount As String, ByVal vReason As String, ByVal vCurreny As String)
        Dim processDtlPDFFileName As String = ""
        Dim t2status As Integer
        SqlCmd = "Select sname,subject,sapno,price,T1.sfname,priceunit,dept,area,T2.signdate,T2.uname,T2.uid,T2.status " &
        "from [dbo].[@XASCH] T0 Inner Join [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid Inner Join [dbo].[@XSPWT] T2 On T0.docnum=T2.docentry " &
        "where docnum=" & docnum & " and T2.signprop=0 order by T2.seq"
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            processDtlPDFFileName = fpath & fileName
            Dim doc1 = New Document(PageSize.A4, 50, 50, 50, 50) '設定pagesize, 邊界left, right, top, bottom
            Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(doc1, New FileStream(fpath & fileName, FileMode.Create))
            '指定使用中文字型, 中文字才能顯示
            Dim svrPath As String = System.Web.HttpContext.Current.Server.MapPath("/") '--網站根目錄實體目錄路徑
            Dim bfChinese As BaseFont = BaseFont.CreateFont(svrPath & "/jFonts/kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
            Dim chFont As Font = New Font(bfChinese, 10)
            Dim chFont_header As Font = New Font(bfChinese, 16)
            Dim chFont_blue As Font = New Font(bfChinese, 12, Font.NORMAL, New Color(51, 0, 153))
            Dim chFont_red As Font = New Font(bfChinese, 12, Font.ITALIC, Color.RED)
            doc1.Open()

            '表格區塊
            Dim intTblWidth() As Integer = {4, 4, 4, 4}
            Dim pTable As PdfPTable = New PdfPTable(intTblWidth(0))
            pTable.TotalWidth = 550.0F
            pTable.LockedWidth = True

            Dim header As PdfPCell = New PdfPCell(New Paragraph(20.0F, vFormName, New Font(bfChinese, 20)))
            header.HorizontalAlignment = 1
            header.Padding = 10
            header.Colspan = 4
            pTable.AddCell(header)

            header = New PdfPCell(New Paragraph(20.0F, "表單編號:" & docnum, chFont_header))
            header.Padding = 5
            header.Colspan = 4
            pTable.AddCell(header)

            header = New PdfPCell(New Paragraph(20.0F, "金額:" & vCurreny & " " & vAmount & " 元 ", chFont_header))
            header.Padding = 5
            header.Colspan = 4
            pTable.AddCell(header)

            header = New PdfPCell(New Paragraph(20.0F, "主旨:" & vReason, chFont_header))
            header.Padding = 5
            header.Colspan = 4
            pTable.AddCell(header)

            header = New PdfPCell(New Paragraph(20.0F, "核准簽條", chFont_header))
            header.Padding = 5
            header.Colspan = 4
            pTable.AddCell(header)

            Dim columnHeader1 As PdfPCell = New PdfPCell(New Phrase(14.0F, "部門", chFont))
            columnHeader1.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader1.Padding = 5
            pTable.AddCell(columnHeader1)
            Dim columnHeader2 As PdfPCell = New PdfPCell(New Phrase(14.0F, "人員", chFont))
            columnHeader2.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader2.Padding = 5
            pTable.AddCell(columnHeader2)
            Dim columnHeader3 As PdfPCell = New PdfPCell(New Phrase(14.0F, "核准章", chFont))
            columnHeader3.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader3.Padding = 5
            pTable.AddCell(columnHeader3)
            Dim columnHeader4 As PdfPCell = New PdfPCell(New Phrase(14.0F, "簽核時間", chFont))
            columnHeader4.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader4.Padding = 5
            pTable.AddCell(columnHeader4)
            '簽核明細
            Dim rowCell As PdfPCell
            Dim strImageUrl As String
            Dim signatureJpg As iTextSharp.text.Image
            Dim jpgCell As PdfPCell
            Dim deptdesc, areadesc, posi As String
            deptdesc = ""
            areadesc = ""
            posi = ""
            Do While (dr.Read())
                t2status = dr(11)
                SqlCmd = "select T1.deptdesc,T2.areadesc,T0.position from dbo.[user] T0 Inner join dbo.[dept] T1 on T0.grp=T1.deptcode " &
                        "Inner Join dbo.[branch] T2 on T0.branch=T2.areacode where T0.id='" & dr(10) & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr1.HasRows) Then
                    dr1.Read()
                    deptdesc = dr1(0)
                    areadesc = dr1(1)
                    posi = dr1(2)
                Else
                    CommUtil.ShowMsg(Me, "沒找到id為" & dr(9) & "之資料,請檢查")
                    'Exit Function
                End If
                dr1.Close()
                conn.Close()
                'For i = 0 To dtRowsConut - 1
                rowCell = New PdfPCell(New Phrase(12.0F, areadesc & vbLf & deptdesc & vbLf & posi, chFont)) '部門
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, dr(9), chFont)) '人員
                pTable.AddCell(rowCell)

                '簽名圖檔
                'MsgBox(areadesc & " " & deptdesc & " " & t2status)
                If (t2status = 3) Then
                    strImageUrl = svrPath & "image/rjpdf.jpg"
                ElseIf (t2status = 10) Then
                    strImageUrl = svrPath & "image/skip1.jpg"
                Else
                    strImageUrl = svrPath & "image/okpdf.jpg"
                End If
                signatureJpg = iTextSharp.text.Image.GetInstance(New Uri(strImageUrl))
                ''''''''signatureJpg.ScaleToFit(100.0F, 33.0F)
                signatureJpg.ScalePercent(70.0F)
                jpgCell = New PdfPCell(signatureJpg, False)
                'pTable.AddCell(signatureJpg)
                jpgCell.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
                jpgCell.Padding = 5
                pTable.AddCell(jpgCell)

                rowCell = New PdfPCell(New Phrase(12.0F, dr(8), chFont)) '簽核時間
                pTable.AddCell(rowCell)
                'Next
            Loop
            doc1.Add(pTable)
            doc1.Close()
            'doc1.Dispose()
        End If
        dr.Close()
        connsap.Close()
        Return processDtlPDFFileName
    End Function
    Public Function createMaterialInOutPDF(ByVal docnum As Long, ByVal vFormName As String, ByVal fileName As String, ByVal fpath As String, ByRef vAmount As Double, ByVal vReason As String, ByVal vCurreny As String, sfid As Integer)
        Dim processDtlPDFFileName As String = ""
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim totalprice As Double
        SqlCmd = "Select IsNull(sum(quantity*price),0) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        totalprice = drL(0)
        drL.Close()
        connL.Close()
        'If (dr.HasRows) Then
        processDtlPDFFileName = fpath & fileName
        Dim doc1 = New Document(PageSize.A4, 50, 50, 50, 50) '設定pagesize, 邊界left, right, top, bottom
        Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(doc1, New FileStream(fpath & fileName, FileMode.Create))
        '指定使用中文字型, 中文字才能顯示
        Dim svrPath As String = System.Web.HttpContext.Current.Server.MapPath("/") '--網站根目錄實體目錄路徑
        Dim bfChinese As BaseFont = BaseFont.CreateFont(svrPath & "/jFonts/kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
        Dim chFont_col As Font = New Font(bfChinese, 10)
        Dim chFont_content As Font = New Font(bfChinese, 9)
        Dim chFont_rightinfo As Font = New Font(bfChinese, 8)
        Dim chFont_header As Font = New Font(bfChinese, 14)
        Dim chFont_blue As Font = New Font(bfChinese, 12, Font.NORMAL, New Color(51, 0, 153))
        Dim chFont_red As Font = New Font(bfChinese, 12, Font.ITALIC, Color.RED)

        'Dim HeaderFontSize As Single = 20.0F
        Dim colHeaderFontSize As Single = 12.0F
        doc1.Open()
        '表格區塊
        Dim TblWidth() As Single = {1.5, 3.5, 7, 1.5, 1.5, 1.5, 3, 3.5}
        'Dim pTable As PdfPTable = New PdfPTable(intTblWidth(0))
        Dim pTable As PdfPTable = New PdfPTable(TblWidth)
        pTable.TotalWidth = 550.0F
        'pTable.SetWidths(intTblWidth)
        pTable.LockedWidth = True
        Dim JetlogoJpg As iTextSharp.text.Image
        Dim strJetlogoUrl As String
        strJetlogoUrl = svrPath & "image/jetlog80%.jpg"
        JetlogoJpg = iTextSharp.text.Image.GetInstance(New Uri(strJetlogoUrl))
        JetlogoJpg.ScalePercent(70.0F)
        Dim jetlogoCell = New PdfPCell(JetlogoJpg, False)
        jetlogoCell.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        jetlogoCell.BorderWidth = 0
        jetlogoCell.Padding = 5
        jetlogoCell.Colspan = 2
        pTable.AddCell(jetlogoCell)

        Dim header As PdfPCell = New PdfPCell(New Paragraph(20.0F, vFormName, New Font(bfChinese, 20)))
        header.HorizontalAlignment = 1
        header.BorderWidth = 0
        header.Padding = 10
        header.Colspan = 5
        pTable.AddCell(header)

        Dim rightinfo As String
        rightinfo = ""
        SqlCmd = "Select convert(varchar(12), signdate, 111) from [dbo].[@XSPWT] where docentry=" & docnum & " and seq=1"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            rightinfo = "單號:" & docnum & vbCr & "送審日期:" & vbCr & drL(0)
        End If
        drL.Close()
        connL.Close()
        header = New PdfPCell(New Paragraph(12.0F, rightinfo, chFont_rightinfo))
        header.HorizontalAlignment = 2
        header.BorderWidth = 0
        header.Padding = 10
        header.Colspan = 1
        pTable.AddCell(header)

        header = New PdfPCell(New Paragraph(20.0F, "表單編號", chFont_header))
        header.Padding = 5
        header.Colspan = 2
        header.HorizontalAlignment = 1
        pTable.AddCell(header)
        header = New PdfPCell(New Paragraph(20.0F, docnum, chFont_header))
        header.Padding = 5
        header.Colspan = 6
        pTable.AddCell(header)

        header = New PdfPCell(New Paragraph(20.0F, "金額", chFont_header))
        header.Padding = 5
        header.Colspan = 2
        header.HorizontalAlignment = 1
        pTable.AddCell(header)
        header = New PdfPCell(New Paragraph(20.0F, vCurreny & " " & Format(vAmount, "###,###.##") & " 元 ", chFont_header))
        header.Padding = 5
        header.Colspan = 6
        pTable.AddCell(header)

        header = New PdfPCell(New Paragraph(20.0F, "主旨", chFont_header))
        header.Padding = 5
        header.Colspan = 2
        header.HorizontalAlignment = 1
        pTable.AddCell(header)
        header = New PdfPCell(New Paragraph(20.0F, vReason, chFont_col))
        header.Padding = 5
        header.Colspan = 6
        pTable.AddCell(header)

        header = New PdfPCell(New Paragraph(20.0F, "事由", chFont_header))
        header.Padding = 5
        header.Colspan = 2
        header.HorizontalAlignment = 1
        header.VerticalAlignment = 1
        pTable.AddCell(header)

        Dim descrip As String
        descrip = ""
        SqlCmd = "Select descrip FROM [dbo].[@XSMLS] T0 WHERE head=1 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            descrip = drL(0)
        End If
        drL.Close()
        connL.Close()
        header = New PdfPCell(New Paragraph(20.0F, descrip, chFont_col))
        header.Padding = 5
        header.Colspan = 6
        pTable.AddCell(header)

        Dim mstr As String
        mstr = "料件明細"
        If (sfid = 51) Then
            mstr = "領用(入庫)料件明細"
        ElseIf (sfid = 50) Then
            mstr = "備品料件明細"
        ElseIf (sfid = 49) Then
            mstr = "報廢料件明細"
        End If
        header = New PdfPCell(New Paragraph(20.0F, mstr, chFont_header))
        header.HorizontalAlignment = 1
        header.Padding = 5
        header.Colspan = 8
        pTable.AddCell(header)

        Dim columnHeader1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "項次", chFont_col))
        columnHeader1.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader1.Padding = 5
        pTable.AddCell(columnHeader1)
        Dim columnHeader2 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "料號", chFont_col))
        columnHeader2.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader2.Padding = 5
        pTable.AddCell(columnHeader2)
        Dim columnHeader3 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "說明", chFont_col))
        columnHeader3.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader3.Padding = 5
        pTable.AddCell(columnHeader3)
        Dim columnHeader4 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "數量", chFont_col))
        columnHeader4.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader4.Padding = 5
        pTable.AddCell(columnHeader4)

        Dim columnHeader5 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "單價", chFont_col))
        columnHeader5.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader5.Padding = 5
        pTable.AddCell(columnHeader5)
        Dim columnHeader6 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "總價", chFont_col))
        columnHeader6.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader6.Padding = 5
        pTable.AddCell(columnHeader6)
        mstr = "處置"
        If (sfid = 51 Or sfid = 50) Then
            mstr = "處置"
        ElseIf (sfid = 49) Then
            mstr = "報廢原因"
        End If
        Dim columnHeader7 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, mstr, chFont_col))
        columnHeader7.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader7.Padding = 5
        pTable.AddCell(columnHeader7)
        Dim columnHeader8 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "備註", chFont_col))
        columnHeader8.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader8.Padding = 5
        pTable.AddCell(columnHeader8)
        Dim i As Integer
        Dim itemcount As Integer
        Dim rowCell As PdfPCell
        SqlCmd = "Select count(*) FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        If (drL(0) <= 5) Then
            itemcount = 6
        Else
            itemcount = drL(0) + 1
        End If
        drL.Close()
        connL.Close()
        i = 1
        SqlCmd = "Select itemcode,itemname,quantity,price,method,comment,num FROM [dbo].[@XSMLS] T0 WHERE head=0 and T0.[docentry] =" & docnum & " ORDER BY num"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                rowCell = New PdfPCell(New Phrase(12.0F, i, chFont_content)) '項次
                rowCell.HorizontalAlignment = 1
                rowCell.VerticalAlignment = 1
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, drL(0), chFont_content)) '料號
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, drL(1), chFont_content)) '說明
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, drL(2), chFont_content)) '數量
                rowCell.HorizontalAlignment = 1
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, Format(drL(3), "###,###.##"), chFont_content)) '單價
                rowCell.HorizontalAlignment = 1
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, Format(drL(2) * drL(3), "###,###.##"), chFont_content)) '總價
                rowCell.HorizontalAlignment = 1
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, drL(4), chFont_content)) '處置
                rowCell.HorizontalAlignment = 1
                pTable.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, drL(5), chFont_content)) '備註
                pTable.AddCell(rowCell)
                i = i + 1
            Loop
        End If
        drL.Close()
        connL.Close()
        Dim j, k As Integer
        For j = i To itemcount
            For k = 1 To 8
                rowCell = New PdfPCell(New Phrase(12.0F, " ", chFont_content)) '空白
                pTable.AddCell(rowCell)
            Next
        Next
        doc1.Add(pTable)

        'doc1.Add(Chunk.NEWLINE)
        doc1.Add(New Paragraph(" "))
        '簽核明細
        Dim t2status As Integer
        Dim pTable1 As PdfPTable = New PdfPTable(6)
        pTable1.TotalWidth = 550.0F
        pTable1.LockedWidth = True

        'Dim columnHeader11 As PdfPCell = New PdfPCell(New Phrase(14.0F, "部門", chFont_col))
        'columnHeader11.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        'columnHeader11.Padding = 5
        'pTable1.AddCell(columnHeader11)
        'Dim columnHeader12 As PdfPCell = New PdfPCell(New Phrase(14.0F, "人員", chFont_col))
        'columnHeader12.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        'columnHeader12.Padding = 5
        'pTable1.AddCell(columnHeader12)
        'Dim columnHeader13 As PdfPCell = New PdfPCell(New Phrase(14.0F, "核准章", chFont_col))
        'columnHeader13.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        'columnHeader13.Padding = 5
        'pTable1.AddCell(columnHeader13)
        'Dim columnHeader14 As PdfPCell = New PdfPCell(New Phrase(14.0F, "簽核時間", chFont_col))
        'columnHeader14.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        'columnHeader14.Padding = 5
        'pTable1.AddCell(columnHeader14)

        Dim columnHeader11 As PdfPCell = New PdfPCell(New Phrase(14.0F, "部門/人員/時間", chFont_col))
        columnHeader11.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader11.Padding = 5
        pTable1.AddCell(columnHeader11)
        Dim columnHeader12 As PdfPCell = New PdfPCell(New Phrase(14.0F, "核准章", chFont_col))
        columnHeader12.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader12.Padding = 5
        pTable1.AddCell(columnHeader12)
        Dim columnHeader13 As PdfPCell = New PdfPCell(New Phrase(14.0F, "簽核意見", chFont_col))
        columnHeader13.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader13.Padding = 5
        pTable1.AddCell(columnHeader13)
        Dim columnHeader14 As PdfPCell = New PdfPCell(New Phrase(14.0F, "部門/人員/時間", chFont_col))
        columnHeader14.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader14.Padding = 5
        pTable1.AddCell(columnHeader14)
        Dim columnHeader15 As PdfPCell = New PdfPCell(New Phrase(14.0F, "核准章", chFont_col))
        columnHeader15.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader15.Padding = 5
        pTable1.AddCell(columnHeader15)
        Dim columnHeader16 As PdfPCell = New PdfPCell(New Phrase(14.0F, "簽核意見", chFont_col))
        columnHeader16.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader16.Padding = 5
        pTable1.AddCell(columnHeader16)
        '簽核明細

        Dim strImageUrl As String
        Dim signatureJpg As iTextSharp.text.Image
        Dim jpgCell As PdfPCell
        Dim deptdesc, areadesc, posi, comment As String
        deptdesc = ""
        areadesc = ""
        posi = ""
        comment = ""
        SqlCmd = "Select sname,subject,sapno,price,T1.sfname,priceunit,dept,area,IsNull(T2.signdate,''),T2.uname,T2.uid,T2.status,T2.comment " &
        "from [dbo].[@XASCH] T0 Inner Join [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid Inner Join [dbo].[@XSPWT] T2 On T0.docnum=T2.docentry " &
        "where docnum=" & docnum & " and T2.signprop=0 order by T2.seq"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        Dim signcount As Integer
        signcount = 0
        If (drL.HasRows) Then
            Do While (drL.Read())
                signcount = signcount + 1
                t2status = drL(11)
                comment = drL(12)
                SqlCmd = "select T1.deptdesc,T2.areadesc,T0.position from dbo.[user] T0 Inner join dbo.[dept] T1 on T0.grp=T1.deptcode " &
                        "Inner Join dbo.[branch] T2 on T0.branch=T2.areacode where T0.id='" & drL(10) & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr1.HasRows) Then
                    dr1.Read()
                    deptdesc = dr1(0)
                    areadesc = dr1(1)
                    posi = dr1(2)
                Else
                    CommUtil.ShowMsg(Me, "沒找到id為" & drL(9) & "之資料,請檢查")
                    'Exit Function
                End If
                dr1.Close()
                conn.Close()
                'For i = 0 To dtRowsConut - 1
                'rowCell = New PdfPCell(New Phrase(12.0F, areadesc & vbLf & deptdesc & vbLf & posi, chFont_col)) '部門
                rowCell = New PdfPCell(New Phrase(12.0F, areadesc & "  " & deptdesc & vbLf & drL(9) & " " & posi & vbLf & drL(8), chFont_col)) '部門
                pTable1.AddCell(rowCell)

                'rowCell = New PdfPCell(New Phrase(12.0F, drL(9), chFont_col)) '人員
                'pTable1.AddCell(rowCell)

                '簽名圖檔
                'MsgBox(areadesc & " " & deptdesc & " " & t2status)
                If (t2status = 3) Then
                    strImageUrl = svrPath & "image/rjpdf.jpg"
                ElseIf (t2status = 10) Then
                    strImageUrl = svrPath & "image/skip1.jpg"
                Else
                    strImageUrl = svrPath & "image/okpdf.jpg"
                End If
                signatureJpg = iTextSharp.text.Image.GetInstance(New Uri(strImageUrl))
                ''''''''signatureJpg.ScaleToFit(100.0F, 33.0F)
                signatureJpg.ScalePercent(70.0F)
                jpgCell = New PdfPCell(signatureJpg, False)
                'pTable.AddCell(signatureJpg)
                jpgCell.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
                jpgCell.Padding = 5
                pTable1.AddCell(jpgCell)

                'rowCell = New PdfPCell(New Phrase(12.0F, drL(8), chFont_col)) '簽核時間
                'pTable1.AddCell(rowCell)

                rowCell = New PdfPCell(New Phrase(12.0F, comment, chFont_col)) '簽核意見
                pTable1.AddCell(rowCell)
                'Next
            Loop
        End If
        If (signcount Mod 2) Then
            rowCell = New PdfPCell(New Phrase(12.0F, "", chFont_col)) '部門
            pTable1.AddCell(rowCell)
            rowCell = New PdfPCell(New Phrase(12.0F, "", chFont_col)) '簽名圖檔
            pTable1.AddCell(rowCell)
            rowCell = New PdfPCell(New Phrase(12.0F, "", chFont_col)) '簽核意見
            pTable1.AddCell(rowCell)
        End If
        drL.Close()
        connL.Close()
        doc1.Add(pTable1)
        doc1.Close()
        'doc1.Dispose()
        Return processDtlPDFFileName
    End Function
    Public Function createRSCPDF(ByVal docnum As Long, ByVal vFormName As String, ByVal fileName As String, ByVal fpath As String, sfid As Integer)
        Dim processDtlPDFFileName As String = ""
        Dim connL, connL1 As New SqlConnection
        Dim drL, drL1 As SqlDataReader
        Dim BColor As New Color(System.Drawing.Color.LightBlue)
        'If (dr.HasRows) Then
        processDtlPDFFileName = fpath & fileName
        Dim doc1 = New Document(PageSize.A4, 50, 50, 50, 50) '設定pagesize, 邊界left, right, top, bottom
        Dim pdfWrite As PdfWriter = PdfWriter.GetInstance(doc1, New FileStream(fpath & fileName, FileMode.Create))
        '指定使用中文字型, 中文字才能顯示
        Dim svrPath As String = System.Web.HttpContext.Current.Server.MapPath("/") '--網站根目錄實體目錄路徑
        Dim bfChinese As BaseFont = BaseFont.CreateFont(svrPath & "/jFonts/kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
        Dim chFont_col As Font = New Font(bfChinese, 10)
        Dim chFont_content As Font = New Font(bfChinese, 9)
        Dim chFont_12 As Font = New Font(bfChinese, 12)
        Dim chFont_header As Font = New Font(bfChinese, 14)
        Dim chFont_blue As Font = New Font(bfChinese, 12, Font.NORMAL, New Color(51, 0, 153))
        Dim chFont_red As Font = New Font(bfChinese, 12, Font.ITALIC, Color.RED)

        'Dim HeaderFontSize As Single = 20.0F
        Dim colHeaderFontSize As Single = 12.0F
        doc1.Open()
        '表格區塊
        Dim TblWidth() As Single = {1, 1, 1, 1, 1, 1}
        'Dim pTable As PdfPTable = New PdfPTable(intTblWidth(0))
        Dim pTable As PdfPTable = New PdfPTable(TblWidth)
        pTable.TotalWidth = 550.0F
        'pTable.SetWidths(intTblWidth)
        pTable.LockedWidth = True

        Dim JetlogoJpg As iTextSharp.text.Image
        Dim strJetlogoUrl As String
        strJetlogoUrl = svrPath & "image/jetlog80%.jpg"
        JetlogoJpg = iTextSharp.text.Image.GetInstance(New Uri(strJetlogoUrl))
        JetlogoJpg.ScalePercent(70.0F)
        Dim jetlogoCell = New PdfPCell(JetlogoJpg, False)
        jetlogoCell.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        jetlogoCell.BorderWidth = 0
        jetlogoCell.Padding = 5
        jetlogoCell.Colspan = 2
        pTable.AddCell(jetlogoCell)

        Dim header As PdfPCell = New PdfPCell(New Paragraph(20.0F, vFormName, New Font(bfChinese, 20)))
        header.HorizontalAlignment = 1
        header.BorderWidth = 0
        header.Padding = 10
        header.Colspan = 3
        pTable.AddCell(header)

        header = New PdfPCell(New Paragraph(20.0F, "", chFont_header))
        header.HorizontalAlignment = 1
        header.BorderWidth = 0
        header.Padding = 10
        header.Colspan = 1
        pTable.AddCell(header)

        header = New PdfPCell(New Paragraph(20.0F, "表單編號:" & docnum, chFont_header))
        header.Padding = 5
        header.Colspan = 6
        pTable.AddCell(header)

        'header = New PdfPCell(New Paragraph(20.0F, "主旨:" & vReason, chFont_header))
        'header.Padding = 5
        'header.Colspan = 6
        'pTable.AddCell(header)

        Dim str() As String
        Dim str1() As String
        Dim datestr As String
        SqlCmd = "Select T0.idname,convert(char(12),T0.cdate,111) as cdate,convert(char(12),T0.albdate,111) as albdate, " &
                     "T0.albhour,T0.albmin,convert(char(12),T0.aledate,111) as aledate, " &
                     "T0.alehour,T0.alemin,T0.createname,T0.rsreason,T0.id,T0.createid,T0.v5id,T0.id " &
                     "FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            str = Split(drL(1), "/")
            Dim columnHeader1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "填寫日期", chFont_col))
            columnHeader1.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader1.Padding = 5
            columnHeader1.BackgroundColor = BColor
            pTable.AddCell(columnHeader1)
            datestr = str(0) & " 年 " & str(1) & " 月 " & Trim(str(2)) & " 日"
            Dim columnHeader1_1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, datestr, chFont_col))
            columnHeader1_1.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
            columnHeader1_1.Padding = 5
            columnHeader1_1.Colspan = 5
            pTable.AddCell(columnHeader1_1)

            Dim columnHeader2 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "補刷日期", chFont_col))
            columnHeader2.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader2.Padding = 5
            columnHeader2.BackgroundColor = BColor
            pTable.AddCell(columnHeader2)
            str = Split(drL(2), "/")
            str1 = Split(drL(5), "/")
            datestr = str(0) & " 年 " & str(1) & " 月 " & Trim(str(2)) & " 日  至  " & str1(0) & " 年 " & str1(1) & " 月 " & Trim(str1(2)) & " 日  止"
            Dim columnHeader2_1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, datestr, chFont_col))
            columnHeader2_1.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
            columnHeader2_1.Padding = 5
            columnHeader2_1.Colspan = 5
            pTable.AddCell(columnHeader2_1)

            Dim columnHeader3 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "補刷時間", chFont_col))
            columnHeader3.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader3.Padding = 5
            columnHeader3.BackgroundColor = BColor
            pTable.AddCell(columnHeader3)
            datestr = "起始    " & drL(3).ToString.PadLeft(2, "0") & " 時 " & drL(4).ToString.PadLeft(2, "0") &
                                                " 分  至  " & drL(6).ToString.PadLeft(2, "0") & " 時 " &
                                                drL(7).ToString.PadLeft(2, "0") & " 分  止"
            Dim columnHeader3_1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, datestr, chFont_col))
            columnHeader3_1.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
            columnHeader3_1.Padding = 5
            columnHeader3_1.Colspan = 5
            pTable.AddCell(columnHeader3_1)

            Dim columnHeader4 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "未刷事由", chFont_col))
            columnHeader4.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader4.Padding = 5
            columnHeader4.BackgroundColor = BColor
            pTable.AddCell(columnHeader4)
            Dim columnHeader4_1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, drL(9), chFont_col))
            columnHeader4_1.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
            columnHeader4_1.Padding = 5
            columnHeader4_1.Colspan = 5
            pTable.AddCell(columnHeader4_1)

            Dim columnHeader5 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "補刷卡人", chFont_col))
            columnHeader5.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader5.Padding = 5
            columnHeader5.Colspan = 3
            columnHeader5.BackgroundColor = BColor
            pTable.AddCell(columnHeader5)
            Dim columnHeader5_1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, "正航代號", chFont_col))
            columnHeader5_1.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            columnHeader5_1.Colspan = 3
            columnHeader5_1.Padding = 5
            columnHeader5_1.BackgroundColor = BColor
            pTable.AddCell(columnHeader5_1)

            Dim columnHeader6 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, drL(0), chFont_12))
            columnHeader6.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            'columnHeader6.VerticalAlignment = 1
            columnHeader6.Padding = 5
            columnHeader6.Colspan = 3
            columnHeader6.FixedHeight = 48.0F
            pTable.AddCell(columnHeader6)
            Dim columnHeader6_1 As PdfPCell = New PdfPCell(New Phrase(colHeaderFontSize, drL(12), chFont_12))
            columnHeader6_1.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
            'columnHeader6_1.VerticalAlignment = 0
            columnHeader6_1.Colspan = 3
            columnHeader6_1.Padding = 5
            columnHeader6_1.FixedHeight = 48.0F
            pTable.AddCell(columnHeader6_1)
        End If
        drL.Close()
        connL.Close()
        doc1.Add(pTable)

        doc1.Add(New Paragraph(" "))
        '簽核明細
        Dim t2status As Integer
        Dim pTable1 As PdfPTable = New PdfPTable(4)
        pTable1.TotalWidth = 550.0F
        pTable1.LockedWidth = True

        'Dim fields As AcroFields
        'fields.SetField("kk","Yes")
        Dim columnHeader11 As PdfPCell = New PdfPCell(New Phrase(14.0F, "部門", chFont_col))
        columnHeader11.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader11.Padding = 5
        pTable1.AddCell(columnHeader11)
        Dim columnHeader12 As PdfPCell = New PdfPCell(New Phrase(14.0F, "人員", chFont_col))
        columnHeader12.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader12.Padding = 5
        pTable1.AddCell(columnHeader12)
        Dim columnHeader13 As PdfPCell = New PdfPCell(New Phrase(14.0F, "核准章", chFont_col))
        columnHeader13.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader13.Padding = 5
        pTable1.AddCell(columnHeader13)
        Dim columnHeader14 As PdfPCell = New PdfPCell(New Phrase(14.0F, "簽核時間", chFont_col))
        columnHeader14.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        columnHeader14.Padding = 5
        pTable1.AddCell(columnHeader14)
        '簽核明細

        Dim strImageUrl As String
        Dim signatureJpg As iTextSharp.text.Image
        Dim jpgCell As PdfPCell
        Dim deptdesc, areadesc, posi As String
        Dim rowCell As PdfPCell
        Dim seq As Integer
        Dim showstr As String
        seq = 1
        deptdesc = ""
        areadesc = ""
        posi = ""
        SqlCmd = "Select sname,subject,sapno,price,T1.sfname,priceunit,dept,area,IsNull(T2.signdate,''),T2.uname,T2.uid,T2.status " &
        "from [dbo].[@XASCH] T0 Inner Join [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid Inner Join [dbo].[@XSPWT] T2 On T0.docnum=T2.docentry " &
        "where docnum=" & docnum & " and T2.signprop=0 order by T2.seq"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                t2status = drL(11)
                SqlCmd = "select T1.deptdesc,T2.areadesc,T0.position from dbo.[user] T0 Inner join dbo.[dept] T1 on T0.grp=T1.deptcode " &
                        "Inner Join dbo.[branch] T2 on T0.branch=T2.areacode where T0.id='" & drL(10) & "'"
                dr1 = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
                If (dr1.HasRows) Then
                    dr1.Read()
                    deptdesc = dr1(0)
                    areadesc = dr1(1)
                    posi = dr1(2)
                Else
                    CommUtil.ShowMsg(Me, "沒找到id為" & drL(9) & "之資料,請檢查")
                    'Exit Function
                End If
                dr1.Close()
                conn.Close()
                'For i = 0 To dtRowsConut - 1
                rowCell = New PdfPCell(New Phrase(12.0F, areadesc & vbLf & deptdesc & vbLf & posi, chFont_col)) '部門
                pTable1.AddCell(rowCell)

                showstr = drL(9)
                If (sfid = 16 And seq = 1) Then
                    SqlCmd = "Select T0.id,T0.createid,T0.idname FROM [dbo].[@XRSCT] T0 WHERE T0.[docentry] =" & docnum
                    drL1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL1)
                    If (drL1.HasRows) Then
                        drL1.Read()
                        If (drL1(0) <> drL1(1)) Then
                            showstr = showstr & vbLf & "(代替 " & drL1(2) & " 發送)"
                        End If
                    End If
                    drL1.Close()
                    connL1.Close()
                End If
                rowCell = New PdfPCell(New Phrase(14.0F, showstr, chFont_col)) '人員
                pTable1.AddCell(rowCell)

                '簽名圖檔
                'MsgBox(areadesc & " " & deptdesc & " " & t2status)
                If (t2status = 3) Then
                    strImageUrl = svrPath & "image/rjpdf.jpg"
                ElseIf (t2status = 10) Then
                    strImageUrl = svrPath & "image/skip1.jpg"
                Else
                    strImageUrl = svrPath & "image/okpdf.jpg"
                End If
                signatureJpg = iTextSharp.text.Image.GetInstance(New Uri(strImageUrl))
                ''''''''signatureJpg.ScaleToFit(100.0F, 33.0F)
                signatureJpg.ScalePercent(70.0F)
                jpgCell = New PdfPCell(signatureJpg, False)
                'pTable.AddCell(signatureJpg)
                jpgCell.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
                jpgCell.Padding = 5
                pTable1.AddCell(jpgCell)

                rowCell = New PdfPCell(New Phrase(12.0F, drL(8), chFont_col)) '簽核時間
                pTable1.AddCell(rowCell)
                'Next
                seq = seq + 1
            Loop
        End If
        drL.Close()
        connL.Close()
        doc1.Add(pTable1)
        doc1.Close()
        'doc1.Dispose()
        Return processDtlPDFFileName
    End Function
    Sub GenPdfStamper1(ByVal fileNameApproved As String, ByVal fileNameStamped As String, ByVal fpath As String, ByVal localfpath As String, specialsfid As Integer, mainfile As Boolean) 'use
        Dim reader As PdfReader = New PdfReader(fpath & fileNameApproved)
        Dim out1 As FileStream = New FileStream(fpath & fileNameStamped, FileMode.Create, FileAccess.Write)
        Dim pdfStamper As PdfStamper = New PdfStamper(reader, out1)
        Dim svrPath As String = System.Web.HttpContext.Current.Server.MapPath("/")
        Dim bfChinese As BaseFont = BaseFont.CreateFont(svrPath & "/jFonts/kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED)
        Dim i As Integer
        Dim gState As PdfGState = New PdfGState()
        gState.FillOpacity = 0.2F
        gState.StrokeOpacity = 0.2F

        For i = 1 To reader.NumberOfPages
            Dim pageSize As Rectangle = reader.GetPageSizeWithRotation(i)
            Dim ctitle As Chunk = New Chunk("Page-" + i.ToString().Trim() + "/" + reader.NumberOfPages.ToString(), FontFactory.GetFont("Futura", 10.0F, New Color(0, 0, 0)))
            Dim ptitle As Phrase = New Phrase(ctitle)
            '浮水印
            Dim imageUrl As String = svrPath & "/image/jetlog.jpg" ' Logo
            Dim img As Image = iTextSharp.text.Image.GetInstance(imageUrl)
            img.ScalePercent(100.0F)  '//縮放比例 org 20.0F
            img.RotationDegrees = 40 '//旋轉角度
            img.SetAbsolutePosition(pageSize.Width / 2 - img.Width / 2, pageSize.Height / 2 - img.Height / 2) '()(10, pageSize.Height - 40) '//設定圖片每頁的絕對位置

            'PdfContentByte類，用設置圖像像和文本的絕對位置
            Dim over As PdfContentByte = pdfStamper.GetOverContent(i)
            ColumnText.ShowTextAligned(over, Element.ALIGN_LEFT, ptitle, pageSize.Width / 2, 10, 0) '//設定頁尾的絕對位置
            over.SetGState(gState) ' //寫入入設定的透明度
            over.AddImage(img) '//圖片印上去
        Next
        pdfStamper.FormFlattening = True 'enable this if you want the pdf flattened
        pdfStamper.Close() 'always close the stamper or you'll have a 0 byte stream
        reader.Close()
        out1.Close()
        File.Copy(fpath & fileNameStamped, localfpath & fileNameStamped)
        Dim str() As String
        If (mainfile) Then
            str = Split(fileNameApproved, "_")
            'If File.Exists(fpath & str(0) & ".pdf") Then
            If (specialsfid = 0) Then '有import 主檔 ,保留原始檔
                File.Delete(fpath & str(0) & "_sign.pdf") '刪除簽核條(已merge 到原主檔成 _Approved.pdf)
                File.Delete(fpath & fileNameApproved) '刪除Approved , 因上述已產生Stamped.pdf
            Else '內建表單,主檔自行產生
                'File.Delete(fpath & str(0) & "_ApprovedTemp.pdf")
                File.Delete(fpath & fileNameApproved)
            End If
        Else
            File.Delete(fpath & fileNameApproved)
            File.Delete(localfpath & fileNameApproved)
        End If
        'Else '內建表單 ,主檔自行產生
        'File.Move(fpath & fileNameApproved, fpath & str(0) & ".pdf")
        'File.Copy(fpath & fileNameApproved, localfpath & str(0) & ".pdf")
        'File.Delete(fpath & str(0) & "_ApprovedTemp.pdf")
        'End If
    End Sub
    Public Function mergePDF(ByVal pdfFiles() As String, ByVal savefileName As String, ByVal fpath As String) As Boolean
        'Dim svrPath As String = System.Web.HttpContext.Current.Server.MapPath(fpath)
        Dim result As Boolean = False
        'Dim result As String = ""
        Dim pdfCount As Integer = 0     'total input pdf file count
        Dim f As Integer = 0    'pointer to current input pdf file
        Dim fileName As String
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pageCount As Integer = 0
        Dim pdfDoc As iTextSharp.text.Document = Nothing    'the output pdf document
        Dim writer As PdfWriter = Nothing
        Dim cb As PdfContentByte = Nothing

        Dim page As PdfImportedPage = Nothing
        Dim rotation As Integer = 0

        Try
            pdfCount = pdfFiles.Length
            If pdfCount > 1 Then
                'Open the 1st item in the array PDFFiles
                fileName = fpath & pdfFiles(f)
                reader = New iTextSharp.text.pdf.PdfReader(fileName)
                'Get page count
                pageCount = reader.NumberOfPages

                pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1), 18, 18, 18, 18)

                writer = PdfWriter.GetInstance(pdfDoc, New FileStream(fpath & savefileName, FileMode.OpenOrCreate))


                With pdfDoc
                    .Open()
                End With
                'Instantiate a PdfContentByte object
                cb = writer.DirectContent
                'Now loop thru the input pdfs
                While f < pdfCount
                    'Declare a page counter variable
                    Dim i As Integer = 0
                    'Loop thru the current input pdf's pages starting at page 1
                    While i < pageCount
                        i += 1
                        'Get the input page size
                        pdfDoc.SetPageSize(reader.GetPageSizeWithRotation(i))
                        'Create a new page on the output document
                        pdfDoc.NewPage()
                        'If it is the 1st page, we add bookmarks to the page
                        'Now we get the imported page
                        page = writer.GetImportedPage(reader, i)
                        'Read the imported page's rotation
                        rotation = reader.GetPageRotation(i)
                        'Then add the imported page to the PdfContentByte object as a template based on the page's rotation
                        If rotation = 90 Then
                            cb.AddTemplate(page, 0, -1.0F, 1.0F, 0, 0, reader.GetPageSizeWithRotation(i).Height)
                        ElseIf rotation = 270 Then
                            cb.AddTemplate(page, 0, 1.0F, -1.0F, 0, reader.GetPageSizeWithRotation(i).Width + 60, -30)
                        Else
                            cb.AddTemplate(page, 1.0F, 0, 0, 1.0F, 0, 0)
                        End If
                    End While
                    'Increment f and read the next input pdf file
                    f += 1
                    If f < pdfCount Then
                        fileName = fpath & pdfFiles(f)
                        reader = New iTextSharp.text.pdf.PdfReader(fileName)
                        pageCount = reader.NumberOfPages
                    End If
                End While
                'When all done, we close the document so that the pdfwriter object can write it to the output file
                pdfDoc.Close()
                'pdfDoc.Dispose()
                reader.Close()
                'File.Delete(fpath & pdfFiles(1)) '在GenPdfStamper1處理
                'File.Delete(fpath & pdfFiles(0)) '在GenPdfStamper1處理
                result = True
            End If
        Catch ex As Exception
            pdfDoc.Close()
            'pdfDoc.Dispose()
            'result = svrPath & "<br>" & ex.ToString()
            Return False
        End Try
        Return result
    End Function
    Public Function mergePDFOfDiffDir(ByVal pdfFiles() As String, ByVal savefileName As String, ByVal subformdir As String, ByVal mainformdir As String) As Boolean
        'Dim svrPath As String = System.Web.HttpContext.Current.Server.MapPath(fpath)
        Dim result As Boolean = False
        'Dim result As String = ""
        Dim pdfCount As Integer = 0     'total input pdf file count
        Dim f As Integer = 0    'pointer to current input pdf file
        Dim fileName As String
        Dim reader As iTextSharp.text.pdf.PdfReader = Nothing
        Dim pageCount As Integer = 0
        Dim pdfDoc As iTextSharp.text.Document = Nothing    'the output pdf document
        Dim writer As PdfWriter = Nothing
        Dim cb As PdfContentByte = Nothing

        Dim page As PdfImportedPage = Nothing
        Dim rotation As Integer = 0

        Try
            pdfCount = pdfFiles.Length
            If pdfCount > 1 Then
                'Open the 1st item in the array PDFFiles
                fileName = subformdir & pdfFiles(f)
                reader = New iTextSharp.text.pdf.PdfReader(fileName)
                'Get page count
                pageCount = reader.NumberOfPages

                pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1), 18, 18, 18, 18)

                writer = PdfWriter.GetInstance(pdfDoc, New FileStream(mainformdir & savefileName, FileMode.OpenOrCreate)) ''設定把merge後檔案, 以原母單檔名存在加簽目錄中,待存檔後,把母工單原主檔刪除,再把此merge後主檔copy至母工單目錄

                With pdfDoc
                    .Open()
                End With
                'Instantiate a PdfContentByte object
                cb = writer.DirectContent
                'Now loop thru the input pdfs
                While f < pdfCount
                    'Declare a page counter variable
                    Dim i As Integer = 0
                    'Loop thru the current input pdf's pages starting at page 1
                    While i < pageCount
                        i += 1
                        'Get the input page size
                        pdfDoc.SetPageSize(reader.GetPageSizeWithRotation(i))
                        'Create a new page on the output document
                        pdfDoc.NewPage()
                        'If it is the 1st page, we add bookmarks to the page
                        'Now we get the imported page
                        page = writer.GetImportedPage(reader, i)
                        'Read the imported page's rotation
                        rotation = reader.GetPageRotation(i)
                        'Then add the imported page to the PdfContentByte object as a template based on the page's rotation
                        If rotation = 90 Then
                            cb.AddTemplate(page, 0, -1.0F, 1.0F, 0, 0, reader.GetPageSizeWithRotation(i).Height)
                        ElseIf rotation = 270 Then
                            cb.AddTemplate(page, 0, 1.0F, -1.0F, 0, reader.GetPageSizeWithRotation(i).Width + 60, -30)
                        Else
                            cb.AddTemplate(page, 1.0F, 0, 0, 1.0F, 0, 0)
                        End If
                    End While
                    'Increment f and read the next input pdf file
                    f += 1
                    If f < pdfCount Then
                        fileName = mainformdir & pdfFiles(f)
                        reader = New iTextSharp.text.pdf.PdfReader(fileName)
                        pageCount = reader.NumberOfPages
                    End If
                End While
                'When all done, we close the document so that the pdfwriter object can write it to the output file
                pdfDoc.Close()
                'pdfDoc.Dispose()
                reader.Close()
                result = True
            End If
        Catch ex As Exception
            pdfDoc.Close()
            'pdfDoc.Dispose()
            'result = svrPath & "<br>" & ex.ToString()
            Return False
        End Try
        Return result
    End Function
    Function AgencySet(id As String)
        Dim conn As New SqlConnection
        Dim dr As SqlDataReader
        Dim agnid As String
        Dim nowday As String
        nowday = Format(Now(), "yyyy/MM/dd")
        agnid = ""
        SqlCmd = "Select T0.agnen,T0.agnid,T0.agndatefrom,T0.agndateto From dbo.[User] T0 where T0.id='" & id & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            If (dr(0) = 0) Then
                agnid = ""
            Else
                If (dr(2) <> "" And dr(3) <> "" And dr(1) <> "") Then
                    If (nowday >= CDate(dr(2)) And CDate(nowday) <= dr(3)) Then
                        agnid = dr(1)
                    Else
                        agnid = ""
                    End If
                Else
                    agnid = ""
                    CommUtil.ShowMsg(Me, "!!!!!代理啟動有打勾, 但其他欄位有空白 , 代理無法啟動!!!!!")
                End If
            End If
        Else
            CommUtil.ShowMsg(Me, "在AgencySet 比對不到id=" & id)
        End If
        dr.Close()
        conn.Close()
        Return agnid
    End Function
    Sub SignOffPush(httpid As String, pushtype As Integer) '1: all 2: 只發代理mail
        Dim connsap, connsap1 As New SqlConnection
        Dim SqlCmd As String
        Dim drsap As SqlDataReader
        Dim ds As New DataSet

        Dim body As String
        Dim subject1 As String
        Dim urlpara As String
        Dim signid As String
        Dim ds1 As New DataSet
        Dim connsaplocal As New SqlConnection
        Dim drsaplocal As SqlDataReader
        Dim nrule As String
        Dim row As Integer = 0
        Dim i As Integer
        Dim agnid As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim emailadd As String
        emailadd = ""
        i = 0
        nrule = ""
        '因未送審(狀態為 E ,D者 , 因尚未出現在XSPWT) , 故需獨立篩選出 , 在與其他簽核篩選合併 , 故需採用 dataset
        SqlCmd = "Select distinct T1.sid " &
                 "FROM dbo.[@XASCH] T1 where T1.status='E' or T1.status='D'"  '目前做法是把 B,R 狀態列入 XSPWT來處理 , 若要與E,D理同,加入並於其下要相應處理
        drsap = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (drsap.HasRows) Then
            '下面Do Loop 是在建立有E,D表單的是否也會在"已進入簽和狀態"列表內 , 如果有就把id 列入 rule 內
            Do While (drsap.Read())
                SqlCmd = "Select count(*) " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where T0.uid='" & drsap(0) & "' and ((T0.signprop = 1 And T1.status ='F') or " &
                 "(T0.status=1) or " &
                 "(T0.signprop=2 And T0.status=1))" ' 是否只需留 T0.ststus=1即可
                drsaplocal = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsaplocal)
                drsaplocal.Read()
                If (drsaplocal(0) <> 0) Then
                    nrule = nrule & " and T1.sid <> '" & drsap(0) & "'"
                End If
                drsaplocal.Close()
                connsaplocal.Close()
            Loop
            '下面是若有列入 rule之id , 則在下述 E,D 狀態時 , 不處理 , 其名單改由 XSPWT 獲得 , 以避免重覆
            SqlCmd = "Select distinct T1.sid As uid " &
                     "FROM dbo.[@XASCH] T1 where (T1.status='E' or T1.status='D')" & nrule
            ds1 = CommUtil.SelectSapSqlUsingDataSet(ds1, SqlCmd, connsaplocal)
            connsaplocal.Close()
            '以下由 XSPWT獲得需簽核名單
            SqlCmd = "Select distinct T0.uid " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where (T0.signprop = 1 And T0.status =1) or " &
                 "(T0.signprop=0 and T0.status=1) or " &
                 "(T0.signprop=2 And T0.status=1)"
            ds1 = CommUtil.SelectSapSqlUsingDataSet(ds1, SqlCmd, connsaplocal)
            connsaplocal.Close()
        Else
            SqlCmd = "Select distinct T0.uid " &
                 "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum " &
                 "where (T0.signprop = 1 And T0.status =1) or " &
                 "(T0.signprop=0 and T0.status=1) or " &
                 "(T0.signprop=2 And T0.status=1)"
            ds1 = CommUtil.SelectSapSqlUsingDataSet(ds1, SqlCmd, connsaplocal)
            connsaplocal.Close()
        End If
        drsap.Close()
        connsap.Close()
        '上述是在獲得需催簽之名單

        If (ds1.Tables(0).Rows.Count <> 0) Then
            For i = 0 To ds1.Tables(0).Rows.Count - 1
                signid = ds1.Tables(0).Rows(i)("uid")
                SqlCmd = "select email from dbo.[user] where id='" & signid & "'"
                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    emailadd = drL(0)
                Else
                    CommUtil.ShowMsg(Me, "無法在User資料表中找到" & signid & "資料")
                End If
                drL.Close()
                connL.Close()
                'SqlCmd = "SELECT T1.sname As uname,T1.spos As upos,T1.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname ,T1.sfid,T1.status,T1.docdate," &
                '         "formdesc='被退回-再送審' " &
                '         "FROM dbo.[@XASCH] T1 Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                '         "where T1.sid='" & signid & "' and (T1.status='E' or T1.status='D' or T1.status='B' or T1.status='R')" &
                '         " order by T1.sfid,T1.docnum desc"
                SqlCmd = "SELECT T1.sname As uname,T1.spos As upos,T1.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname ,T1.sfid,T1.status,T1.docdate," &
                         "case when T1.status='B' then '被退回-再送審' when T1.status='R' then '抽回-再送審' else '未送審' end As formdesc " &
                         "FROM dbo.[@XASCH] T1 Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T1.sid='" & signid & "' and (T1.status='E' or T1.status='D' or T1.status='B' or T1.status='R')" &
                         " order by T1.sfid,T1.docnum desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '未送審
                connsap1.Close()
                SqlCmd = "Select T0.uname,T0.upos,T0.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname,T1.sfid,T1.status,T1.docdate,formdesc='關卡簽核' " &
                         "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T0.signprop=0 and T0.status=1 and T1.status<>'B' and T1.status<>'R' and T0.uid='" & signid & "' " &
                         " order by T0.signprop,T1.sfid,T1.docnum desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '關卡簽核
                connsap1.Close()
                SqlCmd = "Select T0.uname,T0.upos,T0.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname,T1.sfid,T1.status,T1.docdate,formdesc='待歸檔' " &
                         "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T0.signprop=1 And T0.status=1 and T0.uid='" & signid & "' " &
                         " order by T0.signprop,T1.sfid,T1.docnum desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待歸檔
                connsap1.Close()
                SqlCmd = "Select T0.uname,T0.upos,T0.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname,T1.sfid,T1.status,T1.docdate,formdesc='待知悉' " &
                         "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T0.signprop=2 And T0.status=1 and T0.uid ='" & signid & "' " &
                         " order by T1.sfid,T1.docnum desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待知悉
                connsap1.Close()
                'If (ds.Tables(0).Rows.Count = 0) Then
                '    ds.Reset()
                '    Exit Sub
                'End If
                Dim href, title, title_text As String
                href = httpid & "usermgm/login.aspx"
                If (pushtype = 1) Then
                    body = ""
                    title = "之催簽通知"
                    title_text = "之催簽通知"
                    If (ds.Tables(0).Rows(0)("upos") <> "NA") Then
                        title = ds.Tables(0).Rows(0)("uname") & "&nbsp;" & ds.Tables(0).Rows(0)("upos") & "&nbsp;" & title
                        title_text = ds.Tables(0).Rows(0)("uname") & " " & ds.Tables(0).Rows(0)("upos") & " " & title_text
                    Else
                        title = ds.Tables(0).Rows(0)("uname") & "&nbsp;" & title
                        title_text = ds.Tables(0).Rows(0)("uname") & " " & title_text
                    End If
                    body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                        "<table border=1 width=800 border-collapse:collapse>" &
                        "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                        "<tr>" &
                            "<td align=center width=60>簽核原因</td><td align=center width=100>表單名稱</td>" &
                            "<td align=center>主旨</td><td align=center width=200>接收時間</td>" &
                            "</tr>"
                    For row = 0 To ds.Tables(0).Rows.Count - 1
                        'RecordPushSignFlowHistoty(signid, ds.Tables(0).Rows(row)("docnum"))
                        If (ds.Tables(0).Rows(row)("status") = "E" Or ds.Tables(0).Rows(row)("status") = "D" Or ds.Tables(0).Rows(row)("status") = "B" Or ds.Tables(0).Rows(row)("status") = "R") Then
                            urlpara = "?actmode=single_signoff&uid=" & signid & "&status=" & ds.Tables(0).Rows(row)("status") & "&formtypeindex=0" &
                                        "&formstatusindex=0&docnum=" & ds.Tables(0).Rows(row)("docnum") & "&sfid=" & ds.Tables(0).Rows(row)("sfid")
                        Else
                            urlpara = "?actmode=signoff&uid=" & signid & "&status=" & ds.Tables(0).Rows(row)("status") & "&formtypeindex=0" &
                                        "&formstatusindex=0&docnum=" & ds.Tables(0).Rows(row)("docnum") & "&sfid=" & ds.Tables(0).Rows(row)("sfid")
                        End If
                        body = body & "<tr>" &
                            "<tr>" &
                            "<td align=center><a href=" & href & urlpara & ">" & ds.Tables(0).Rows(row)("formdesc") & "</a></td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                            "<td align=center>" & ds.Tables(0).Rows(row)("receivedate") & "</td>" &
                            "</tr>"
                    Next
                    body = body & "</table>"
                    subject1 = "(催簽通知) " & title_text
                    'emailadd = "ron@jettech.com.tw" 'temp test
                    CommUtil.SendMail(emailadd, subject1, body)
                End If

                Dim agnname, pos As String
                agnname = ""
                pos = ""
                emailadd = ""
                agnid = AgencySet(signid)
                If (agnid <> "") Then '啟動代理人郵件通知
                    SqlCmd = "select name,email,position from dbo.[user] where id='" & agnid & "'"
                    drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                    If (drL.HasRows) Then
                        drL.Read()
                        agnname = drL(0)
                        pos = drL(2)
                        emailadd = drL(1)
                    Else
                        CommUtil.ShowMsg(Me, "無法在User資料表中找到" & agnid & "代理人資料")
                    End If
                    body = ""
                    subject1 = ""
                    title = "之代理催簽通知"
                    title_text = "之代理催簽通知"
                    If (pos <> "NA") Then
                        title = agnname & "&nbsp;" & pos & "&nbsp;" & title
                        title_text = agnname & " " & pos & " " & title_text
                    Else
                        title = agnname & "&nbsp;" & title
                        title_text = agnname & " " & title_text
                    End If
                    body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                            "<table border=1 width=800 border-collapse:collapse>" &
                            "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                            "<tr>" &
                                "<td align=center width=60>簽核原因</td><td align=center width=100>表單名稱</td>" &
                                "<td align=center>主旨</td><td align=center width=200>接收時間</td>" &
                                "</tr>"
                    For row = 0 To ds.Tables(0).Rows.Count - 1
                        urlpara = "?actmode=signoff&uid=" & signid & "&status=" & ds.Tables(0).Rows(row)("status") & "&formtypeindex=0" &
                            "&formstatusindex=0&docnum=" & ds.Tables(0).Rows(row)("docnum") & "&sfid=" & ds.Tables(0).Rows(row)("sfid") & "&agnid=" & agnid
                        body = body & "<tr>" &
                                "<tr>" &
                                "<td align=center><a href=" & href & urlpara & ">" & ds.Tables(0).Rows(row)("formdesc") & "</a></td>" &
                                "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                                "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                                "<td align=center>" & ds.Tables(0).Rows(row)("receivedate") & "</td>" &
                                "</tr>"
                    Next
                    body = body & "</table>"
                    subject1 = "(代理催簽通知) " & title_text
                    'emailadd = "ron@jettech.com.tw" 'temp test
                    CommUtil.SendMail(emailadd, subject1, body)
                    drL.Close()
                    connL.Close()
                End If
                ds.Reset()
            Next
        End If
        ds1.Reset()
    End Sub
    Sub ReplaceSignOffInform(httpid As String, signid As String, signreason As String)
        Dim connsap, connsap1 As New SqlConnection
        Dim SqlCmd As String
        Dim ds As New DataSet
        Dim body As String
        Dim subject1 As String
        Dim urlpara As String
        Dim connsaplocal As New SqlConnection
        Dim row As Integer = 0
        Dim agnid As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim emailadd As String
        emailadd = ""
        SqlCmd = "select email from dbo.[user] where id='" & signid & "'"
        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            emailadd = drL(0)
        Else
            CommUtil.ShowMsg(Me, "無法在User資料表中找到" & signid & "資料")
        End If
        drL.Close()
        connL.Close()
        'SqlCmd = "SELECT T1.sname As uname,T1.spos As upos,T1.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname ,T1.sfid,T1.status,T1.docdate," &
        '                 "case when T1.status='B' then '被退回-再送審' when T1.status='R' then '抽回-再送審' else '未送審' end As formdesc " &
        '                 "FROM dbo.[@XASCH] T1 Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
        '                 "where T1.sid='" & signid & "' and (T1.status='E' or T1.status='D' or T1.status='B' or T1.status='R')" &
        '                 " order by T1.sfid,T1.docnum desc"
        'ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '未送審
        'connsap1.Close()
        SqlCmd = "Select T0.uname,T0.upos,T0.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname,T1.sfid,T1.status,T1.docdate,formdesc='關卡簽核' " &
                         "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T0.signprop=0 and T0.status=1 and T1.status<>'B' and T1.status<>'R' and T0.uid='" & signid & "' " &
                         " order by T0.signprop,T1.sfid,T1.docnum desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '關卡簽核
        connsap1.Close()
        SqlCmd = "Select T0.uname,T0.upos,T0.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname,T1.sfid,T1.status,T1.docdate,formdesc='待歸檔' " &
                         "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T0.signprop=1 And T0.status=1 and T0.uid='" & signid & "' " &
                         " order by T0.signprop,T1.sfid,T1.docnum desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待歸檔
        connsap1.Close()
        SqlCmd = "Select T0.uname,T0.upos,T0.receivedate,T2.sfname,T1.docnum,T1.subject,T1.sname,T1.sfid,T1.status,T1.docdate,formdesc='待知悉' " &
                         "FROM dbo.[@XSPWT] T0 INNER JOIN dbo.[@XASCH] T1 ON T0.docentry=T1.docnum Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T0.signprop=2 And T0.status=1 and T0.uid ='" & signid & "' " &
                         " order by T1.sfid,T1.docnum desc"
        ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1) '待知悉
        connsap1.Close()
        If (ds.Tables(0).Rows.Count <> 0) Then
            Dim href, title, title_text As String
            href = httpid & "usermgm/login.aspx"
            body = ""
            title = "之" & signreason & "通知"
            title_text = "之" & signreason & "通知"
            If (ds.Tables(0).Rows(0)("upos") <> "NA") Then
                title = ds.Tables(0).Rows(0)("uname") & "&nbsp;" & ds.Tables(0).Rows(0)("upos") & "&nbsp;" & title
                title_text = ds.Tables(0).Rows(0)("uname") & " " & ds.Tables(0).Rows(0)("upos") & " " & title_text
            Else
                title = ds.Tables(0).Rows(0)("uname") & "&nbsp;" & title
                title_text = ds.Tables(0).Rows(0)("uname") & " " & title_text
            End If
            body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                    "<table border=1 width=800 border-collapse:collapse>" &
                    "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                    "<tr>" &
                        "<td align=center width=60>簽核原因</td><td align=center width=100>表單名稱</td>" &
                        "<td align=center>主旨</td><td align=center width=200>接收時間</td>" &
                        "</tr>"
            For row = 0 To ds.Tables(0).Rows.Count - 1
                'RecordPushSignFlowHistoty(signid, ds.Tables(0).Rows(row)("docnum"))
                If (ds.Tables(0).Rows(row)("status") = "E" Or ds.Tables(0).Rows(row)("status") = "D" Or ds.Tables(0).Rows(row)("status") = "B" Or ds.Tables(0).Rows(row)("status") = "R") Then
                    urlpara = "?actmode=single_signoff&uid=" & signid & "&status=" & ds.Tables(0).Rows(row)("status") & "&formtypeindex=0" &
                                    "&formstatusindex=0&docnum=" & ds.Tables(0).Rows(row)("docnum") & "&sfid=" & ds.Tables(0).Rows(row)("sfid")
                Else
                    urlpara = "?actmode=signoff&uid=" & signid & "&status=" & ds.Tables(0).Rows(row)("status") & "&formtypeindex=0" &
                                    "&formstatusindex=0&docnum=" & ds.Tables(0).Rows(row)("docnum") & "&sfid=" & ds.Tables(0).Rows(row)("sfid")
                End If
                body = body & "<tr>" &
                        "<tr>" &
                        "<td align=center><a href=" & href & urlpara & ">" & ds.Tables(0).Rows(row)("formdesc") & "</a></td>" &
                        "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                        "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                        "<td align=center>" & ds.Tables(0).Rows(row)("receivedate") & "</td>" &
                        "</tr>"
            Next
            body = body & "</table>"
            subject1 = "(" & signreason & "通知) " & title_text
            'emailadd = "ron@jettech.com.tw" 'temp test
            CommUtil.SendMail(emailadd, subject1, body)


            Dim agnname, pos As String
            agnname = ""
            pos = ""
            emailadd = ""
            agnid = AgencySet(signid)
            If (agnid <> "") Then '啟動代理人郵件通知
                SqlCmd = "select name,email,position from dbo.[user] where id='" & agnid & "'"
                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    agnname = drL(0)
                    pos = drL(2)
                    emailadd = drL(1)
                Else
                    CommUtil.ShowMsg(Me, "無法在User資料表中找到" & agnid & "代理人資料")
                End If
                body = ""
                subject1 = ""
                title = "之代理" & signreason & "通知"
                title_text = "之代理" & signreason & "通知"
                If (pos <> "NA") Then
                    title = agnname & "&nbsp;" & pos & "&nbsp;" & title
                    title_text = agnname & " " & pos & " " & title_text
                Else
                    title = agnname & "&nbsp;" & title
                    title_text = agnname & " " & title_text
                End If
                body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                            "<table border=1 width=800 border-collapse:collapse>" &
                            "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                            "<tr>" &
                                "<td align=center width=60>簽核原因</td><td align=center width=100>表單名稱</td>" &
                                "<td align=center>主旨</td><td align=center width=200>接收時間</td>" &
                                "</tr>"
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    urlpara = "?actmode=signoff&uid=" & signid & "&status=" & ds.Tables(0).Rows(row)("status") & "&formtypeindex=0" &
                            "&formstatusindex=0&docnum=" & ds.Tables(0).Rows(row)("docnum") & "&sfid=" & ds.Tables(0).Rows(row)("sfid") & "&agnid=" & agnid
                    body = body & "<tr>" &
                                "<tr>" &
                                "<td align=center><a href=" & href & urlpara & ">" & ds.Tables(0).Rows(row)("formdesc") & "</a></td>" &
                                "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                                "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                                "<td align=center>" & ds.Tables(0).Rows(row)("receivedate") & "</td>" &
                                "</tr>"
                Next
                body = body & "</table>"
                subject1 = "(代理" & signreason & "通知) " & title_text
                'emailadd = "ron@jettech.com.tw" 'temp test
                CommUtil.SendMail(emailadd, subject1, body)
                drL.Close()
                connL.Close()
            End If
        End If
        ds.Reset()
    End Sub
    Sub ToDoListPush(httpid As String)
        Dim nowdate As String
        Dim connsap1 As New SqlConnection
        Dim SqlCmd As String
        Dim ds, dss As New DataSet
        Dim pushreason As String
        Dim body As String
        Dim subject1 As String
        Dim urlpara As String
        Dim inchargepersonid, inchargepersonname, position As String
        Dim ds1, ds2 As New DataSet
        Dim connsaplocal As New SqlConnection
        Dim row As Integer
        Dim i As Integer
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim emailadd As String
        Dim errstr As String
        errstr = ""
        nowdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        emailadd = ""
        inchargepersonname = ""
        inchargepersonid = ""
        position = ""
        i = 0
        SqlCmd = "Select distinct T0.incharge " &
                 "FROM dbo.[@XSTDT] T0 " &
                 "where tdate < '" & nowdate & "' and incharge <> '' and status <> 100"
        ds1 = CommUtil.SelectSapSqlUsingDataSet(ds1, SqlCmd, connsaplocal)
        connsaplocal.Close()
        '上述是在獲得需催簽之名單

        If (ds1.Tables(0).Rows.Count <> 0) Then
            For i = 0 To ds1.Tables(0).Rows.Count - 1
                inchargepersonid = ds1.Tables(0).Rows(i)("incharge")
                SqlCmd = "select email,position,name from dbo.[user] where id='" & inchargepersonid & "'"
                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    emailadd = drL(0)
                    inchargepersonname = drL(2)
                    position = drL(1)
                Else
                    'CommUtil.ShowMsg(Me, "無法在User資料表中找到" & inchargepersonid & "資料")
                    errstr = "無法在User資料表中找到" & inchargepersonid & "資料"
                End If
                drL.Close()
                connL.Close()
                SqlCmd = "SELECT T2.sfname,T1.docentry,T1.subject,T1.sfid,T1.tdate,T1.num " &
                         "FROM dbo.[@XSTDT] T1 Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T1.incharge='" & ds1.Tables(0).Rows(i)("incharge") & "' and (tdate < '" & nowdate & "' or tdate='1900/01/01') " &
                         " order by T1.sfid,T1.docentry desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
                connsap1.Close()
                'If (ds.Tables(0).Rows.Count = 0) Then
                'ds.Reset()
                'Exit Sub
                'End If
                Dim href, title, title_text As String
                href = httpid & "usermgm/login.aspx"
                body = ""
                title = "之待辦進度催促通知"
                title_text = "之待辦進度催促通知"
                If (position <> "NA") Then
                    title = inchargepersonname & "&nbsp;" & position & "&nbsp;" & title
                    title_text = inchargepersonname & " " & position & " " & title_text
                Else
                    title = inchargepersonname & "&nbsp;" & title
                    title_text = inchargepersonname & " " & title_text
                End If
                body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                        "<table border=1 width=800 border-collapse:collapse>" &
                        "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                        "<tr>" &
                            "<td align=center width=100>催促原因</td><td align=center width=100>表單名稱</td>" &
                            "<td align=center>主旨</td><td align=center width=200>催促時間</td>" &
                            "</tr>"
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    If (ds.Tables(0).Rows(row)("tdate") = "1900/01/01") Then
                        pushreason = "未設定完成日期"
                    Else
                        pushreason = "進度逾期"
                    End If
                    urlpara = "?actmode=todoitem&uid=" & inchargepersonid & "&inchargeid=" & inchargepersonid &
                                        "&docentry=" & ds.Tables(0).Rows(row)("docentry") & "&sfid=" & ds.Tables(0).Rows(row)("sfid") &
                                        "&num=" & ds.Tables(0).Rows(row)("num")
                    body = body & "<tr>" &
                            "<tr>" &
                            "<td align=center><a href=" & href & urlpara & ">" & pushreason & "</a></td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                            "<td align=center>" & nowdate & "</td>" &
                            "</tr>"
                Next
                body = body & "</table>"
                subject1 = "(待辦進度催促通知) " & title_text
                'emailadd = "ron@jettech.com.tw" 'temp test
                CommUtil.SendMail(emailadd, subject1, body)
            Next
        End If
        ds.Reset()
        ds1.Reset()

        SqlCmd = "Select distinct T0.incharge " &
                 "FROM dbo.[@XSTDT] T0 " &
                 "where tdate >= '" & nowdate & "' and updflag <> 1 and incharge <> ''"
        ds1 = CommUtil.SelectSapSqlUsingDataSet(ds1, SqlCmd, connsaplocal)
        connsaplocal.Close()
        '上述是在獲得未更新需催簽之名單

        If (ds1.Tables(0).Rows.Count <> 0) Then
            For i = 0 To ds1.Tables(0).Rows.Count - 1
                inchargepersonid = ds1.Tables(0).Rows(i)("incharge")
                SqlCmd = "select email,position,name from dbo.[user] where id='" & inchargepersonid & "'"
                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    emailadd = drL(0)
                    inchargepersonname = drL(2)
                    position = drL(1)
                Else
                    'CommUtil.ShowMsg(Me, "無法在User資料表中找到" & inchargepersonid & "資料")
                    errstr = "無法在User資料表中找到" & inchargepersonid & "資料"
                End If
                drL.Close()
                connL.Close()
                SqlCmd = "SELECT T2.sfname,T1.docentry,T1.subject,T1.sfid,T1.tdate,T1.updflag,T1.num " &
                         "FROM dbo.[@XSTDT] T1 Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T1.incharge='" & ds1.Tables(0).Rows(i)("incharge") & "' and updflag <> 1 and tdate >= '" & nowdate & "' " &
                         " order by T1.sfid,T1.docentry desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
                connsap1.Close()
                'If (ds.Tables(0).Rows.Count = 0) Then
                'ds.Reset()
                'Exit Sub
                'End If
                Dim href, title, title_text As String
                href = httpid & "usermgm/login.aspx"
                body = ""
                title = "之待辦進度催促通知"
                title_text = "之待辦進度催促通知"
                If (position <> "NA") Then
                    title = inchargepersonname & "&nbsp;" & position & "&nbsp;" & title
                    title_text = inchargepersonname & " " & position & " " & title_text
                Else
                    title = inchargepersonname & "&nbsp;" & title
                    title_text = inchargepersonname & " " & title_text
                End If
                body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                        "<table border=1 width=800 border-collapse:collapse>" &
                        "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                        "<tr>" &
                            "<td align=center width=100>催促原因</td><td align=center width=100>表單名稱</td>" &
                            "<td align=center>主旨</td><td align=center width=200>催促時間</td>" &
                            "</tr>"
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    If (ds.Tables(0).Rows(row)("updflag") = 0) Then
                        pushreason = "您上周無任何更新或點擊"
                    ElseIf (ds.Tables(0).Rows(row)("updflag") = 2) Then
                        pushreason = "因您無反應,每日催促"
                    ElseIf (ds.Tables(0).Rows(row)("updflag") = 3) Then
                        pushreason = "您尚未第一次更新"
                    Else
                        pushreason = "不明"
                    End If
                    urlpara = "?actmode=todoitem&uid=" & inchargepersonid & "&inchargeid=" & inchargepersonid &
                                        "&docentry=" & ds.Tables(0).Rows(row)("docentry") & "&sfid=" & ds.Tables(0).Rows(row)("sfid") &
                                        "&num=" & ds.Tables(0).Rows(row)("num")
                    body = body & "<tr>" &
                            "<tr>" &
                            "<td align=center><a href=" & href & urlpara & ">" & pushreason & "</a></td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                            "<td align=center>" & nowdate & "</td>" &
                            "</tr>"
                    'SqlCmd = "update dbo.[@XSTDT] set updflag= 2 where num=" & ds.Tables(0).Rows(row)("num")
                    'CommUtil.SqlSapExecute("upd", SqlCmd, connsap1)
                    'connsap1.Close()
                Next
                body = body & "</table>"
                subject1 = "(待辦進度催促通知) " & title_text
                'emailadd = "ron@jettech.com.tw" 'temp test
                CommUtil.SendMail(emailadd, subject1, body)
            Next
            'SqlCmd = "update dbo.[@XSTDT] set updflag= 0,updcount=0 where updflag=1 and status<>100"
            'CommUtil.SqlSapExecute("upd", SqlCmd, connsap1)
            'connsap1.Close()
        End If
        ds.Reset()
        ds1.Reset()
        'Return errstr
    End Sub

    Sub InformByMail(httpid As String, title As String, reason As String, ds As DataSet, urlpara As String)
        Dim nowdate As String
        Dim body As String
        Dim subject1 As String
        Dim mailtoid, mailtoidname, position As String
        Dim row As Integer
        Dim i As Integer
        Dim emailadd As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim errstr As String
        Dim href, title_text As String

        nowdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        emailadd = ""
        mailtoidname = ""
        mailtoid = ""
        position = ""
        If (ds.Tables(0).Rows.Count <> 0) Then
            For i = 0 To ds.Tables(0).Rows.Count - 1
                mailtoid = ds.Tables(0).Rows(i)("mailtoid")
                SqlCmd = "select email,position,name from dbo.[user] where id='" & mailtoid & "'"
                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    emailadd = drL(0)
                    mailtoidname = drL(2)
                    position = drL(1)
                Else
                    errstr = "無法在User資料表中找到" & ds.Tables(0).Rows(i)("mailtoid") & "資料"
                End If
                href = httpid & "usermgm/login.aspx"
                body = ""
                title_text = title
                If (position <> "NA") Then
                    title = mailtoidname & "&nbsp;" & position & "&nbsp;" & title
                    title_text = mailtoidname & " " & position & " " & title_text
                Else
                    title = mailtoidname & "&nbsp;" & title
                    title_text = mailtoidname & " " & title_text
                End If
                body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                        "<table border=1 width=800 border-collapse:collapse>" &
                        "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                        "<tr>" &
                            "<td align=center width=100>通知原因</td><td align=center width=100>表單名稱</td>" &
                            "<td align=center>主旨</td><td align=center width=200>通知時間</td>" &
                            "</tr>"
                body = body & "<tr>" &
                        "<tr>" &
                        "<td align=center><a href=" & href & urlpara & ">" & reason & "</a></td>" &
                        "<td align=left>" & ds.Tables(0).Rows(i)("sfname") & "</td>" &
                        "<td align=left>" & ds.Tables(0).Rows(i)("subject") & "</td>" &
                        "<td align=center>" & nowdate & "</td>" &
                        "</tr>"
                body = body & "</table>"
                subject1 = "(事件通知) " & title_text
                'emailadd = "ron@jettech.com.tw" 'temp test
                CommUtil.SendMail(emailadd, subject1, body)
            Next
        End If
    End Sub
    Sub InformTracePersonForItemDone(httpid As String)
        Dim nowdate As String
        Dim connsap1 As New SqlConnection
        Dim SqlCmd As String
        Dim ds, dss As New DataSet
        Dim pushreason As String
        Dim body As String
        Dim subject1 As String
        Dim urlpara As String
        Dim inchargepersonid, inchargepersonname, position As String
        Dim ds1, ds2 As New DataSet
        Dim connsaplocal As New SqlConnection
        Dim row As Integer
        Dim i As Integer
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim emailadd As String
        Dim errstr As String
        errstr = ""
        nowdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        emailadd = ""
        inchargepersonname = ""
        inchargepersonid = ""
        position = ""
        i = 0
        SqlCmd = "Select distinct T0.incharge " &
                 "FROM dbo.[@XSTDT] T0 " &
                 "where tdate < '" & nowdate & "' and incharge <> ''"
        ds1 = CommUtil.SelectSapSqlUsingDataSet(ds1, SqlCmd, connsaplocal)
        connsaplocal.Close()
        '上述是在獲得需催簽之名單

        If (ds1.Tables(0).Rows.Count <> 0) Then
            For i = 0 To ds1.Tables(0).Rows.Count - 1
                inchargepersonid = ds1.Tables(0).Rows(i)("incharge")
                SqlCmd = "select email,position,name from dbo.[user] where id='" & inchargepersonid & "'"
                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    emailadd = drL(0)
                    inchargepersonname = drL(2)
                    position = drL(1)
                Else
                    'CommUtil.ShowMsg(Me, "無法在User資料表中找到" & inchargepersonid & "資料")
                    errstr = "無法在User資料表中找到" & inchargepersonid & "資料"
                End If
                drL.Close()
                connL.Close()
                SqlCmd = "SELECT T2.sfname,T1.docentry,T1.subject,T1.sfid,T1.tdate,T1.num " &
                         "FROM dbo.[@XSTDT] T1 Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T1.incharge='" & ds1.Tables(0).Rows(i)("incharge") & "' and (tdate < '" & nowdate & "' or tdate='1900/01/01') " &
                         " order by T1.sfid,T1.docentry desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
                connsap1.Close()
                'If (ds.Tables(0).Rows.Count = 0) Then
                'ds.Reset()
                'Exit Sub
                'End If
                Dim href, title, title_text As String
                href = httpid & "usermgm/login.aspx"
                body = ""
                title = "之待辦進度催促通知"
                title_text = "之待辦進度催促通知"
                If (position <> "NA") Then
                    title = inchargepersonname & "&nbsp;" & position & "&nbsp;" & title
                    title_text = inchargepersonname & " " & position & " " & title_text
                Else
                    title = inchargepersonname & "&nbsp;" & title
                    title_text = inchargepersonname & " " & title_text
                End If
                body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                        "<table border=1 width=800 border-collapse:collapse>" &
                        "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                        "<tr>" &
                            "<td align=center width=100>催促原因</td><td align=center width=100>表單名稱</td>" &
                            "<td align=center>主旨</td><td align=center width=200>催促時間</td>" &
                            "</tr>"
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    If (ds.Tables(0).Rows(row)("tdate") = "1900/01/01") Then
                        pushreason = "未設定完成日期"
                    Else
                        pushreason = "進度逾期"
                    End If
                    urlpara = "?actmode=todoitem&uid=" & inchargepersonid &
                                        "&docentry=" & ds.Tables(0).Rows(row)("docentry") & "&sfid=" & ds.Tables(0).Rows(row)("sfid") &
                                        "&num=" & ds.Tables(0).Rows(row)("num")
                    body = body & "<tr>" &
                            "<tr>" &
                            "<td align=center><a href=" & href & urlpara & ">" & pushreason & "</a></td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                            "<td align=center>" & nowdate & "</td>" &
                            "</tr>"
                Next
                body = body & "</table>"
                subject1 = "(待辦進度催促通知) " & title_text
                'emailadd = "ron@jettech.com.tw" 'temp test
                CommUtil.SendMail(emailadd, subject1, body)
            Next
        End If
        ds.Reset()
        ds1.Reset()

        SqlCmd = "Select distinct T0.incharge " &
                 "FROM dbo.[@XSTDT] T0 " &
                 "where tdate >= '" & nowdate & "' and updflag <> 1 and incharge <> ''"
        ds1 = CommUtil.SelectSapSqlUsingDataSet(ds1, SqlCmd, connsaplocal)
        connsaplocal.Close()
        '上述是在獲得未更新需催簽之名單

        If (ds1.Tables(0).Rows.Count <> 0) Then
            For i = 0 To ds1.Tables(0).Rows.Count - 1
                inchargepersonid = ds1.Tables(0).Rows(i)("incharge")
                SqlCmd = "select email,position,name from dbo.[user] where id='" & inchargepersonid & "'"
                drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
                If (drL.HasRows) Then
                    drL.Read()
                    emailadd = drL(0)
                    inchargepersonname = drL(2)
                    position = drL(1)
                Else
                    'CommUtil.ShowMsg(Me, "無法在User資料表中找到" & inchargepersonid & "資料")
                    errstr = "無法在User資料表中找到" & inchargepersonid & "資料"
                End If
                drL.Close()
                connL.Close()
                SqlCmd = "SELECT T2.sfname,T1.docentry,T1.subject,T1.sfid,T1.tdate,T1.updflag,T1.num " &
                         "FROM dbo.[@XSTDT] T1 Inner Join [dbo].[@XSFTT] T2 ON T1.sfid=T2.sfid " &
                         "where T1.incharge='" & ds1.Tables(0).Rows(i)("incharge") & "' and updflag <> 1 and tdate >= '" & nowdate & "' " &
                         " order by T1.sfid,T1.docentry desc"
                ds = CommUtil.SelectSapSqlUsingDataSet(ds, SqlCmd, connsap1)
                connsap1.Close()
                'If (ds.Tables(0).Rows.Count = 0) Then
                'ds.Reset()
                'Exit Sub
                'End If
                Dim href, title, title_text As String
                href = httpid & "usermgm/login.aspx"
                body = ""
                title = "之待辦進度催促通知"
                title_text = "之待辦進度催促通知"
                If (position <> "NA") Then
                    title = inchargepersonname & "&nbsp;" & position & "&nbsp;" & title
                    title_text = inchargepersonname & " " & position & " " & title_text
                Else
                    title = inchargepersonname & "&nbsp;" & title
                    title_text = inchargepersonname & " " & title_text
                End If
                body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                        "<table border=1 width=800 border-collapse:collapse>" &
                        "<tr bgcolor=#add8e6><td colspan=4><h1 align=center>" & title & "</h1></td></tr>" &
                        "<tr>" &
                            "<td align=center width=100>催促原因</td><td align=center width=100>表單名稱</td>" &
                            "<td align=center>主旨</td><td align=center width=200>催促時間</td>" &
                            "</tr>"
                For row = 0 To ds.Tables(0).Rows.Count - 1
                    If (ds.Tables(0).Rows(row)("updflag") = 0) Then
                        pushreason = "您上周無任何更新或點擊"
                    ElseIf (ds.Tables(0).Rows(row)("updflag") = 2) Then
                        pushreason = "因您無反應,每日催促"
                    ElseIf (ds.Tables(0).Rows(row)("updflag") = 3) Then
                        pushreason = "您尚未第一次更新"
                    Else
                        pushreason = "不明"
                    End If
                    urlpara = "?actmode=todoitem&uid=" & inchargepersonid &
                                        "&docentry=" & ds.Tables(0).Rows(row)("docentry") & "&sfid=" & ds.Tables(0).Rows(row)("sfid") &
                                        "&num=" & ds.Tables(0).Rows(row)("num")
                    body = body & "<tr>" &
                            "<tr>" &
                            "<td align=center><a href=" & href & urlpara & ">" & pushreason & "</a></td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("sfname") & "</td>" &
                            "<td align=left>" & ds.Tables(0).Rows(row)("subject") & "</td>" &
                            "<td align=center>" & nowdate & "</td>" &
                            "</tr>"
                    'SqlCmd = "update dbo.[@XSTDT] set updflag= 2 where num=" & ds.Tables(0).Rows(row)("num")
                    'CommUtil.SqlSapExecute("upd", SqlCmd, connsap1)
                    'connsap1.Close()
                Next
                body = body & "</table>"
                subject1 = "(待辦進度催促通知) " & title_text
                'emailadd = "ron@jettech.com.tw" 'temp test
                CommUtil.SendMail(emailadd, subject1, body)
            Next
            'SqlCmd = "update dbo.[@XSTDT] set updflag= 0,updcount=0 where updflag=1 and status<>100"
            'CommUtil.SqlSapExecute("upd", SqlCmd, connsap1)
            'connsap1.Close()
        End If
        ds.Reset()
        ds1.Reset()
        'Return errstr
    End Sub
    Sub RecordPushSignFlowHistoty(signid As String, docnum As Long)
        Dim SqlCmd As String
        Dim flowseq As Integer
        Dim signdate As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim agnname As String
        Dim signname, comment, reason As String
        signname = ""
        comment = "from localhost"
        reason = "催簽"
        SqlCmd = "select name from dbo.[user] where id='" & signid & "'"
        drL = CommUtil.SelectLocalSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            signname = drL(0)
        End If
        drL.Close()
        connL.Close()
        signdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        SqlCmd = "Select IsNull(Max(flowseq),0) from [dbo].[@XSPHT] where docentry=" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        drL.Read()
        flowseq = drL(0) + 1
        drL.Close()
        connL.Close()
        agnname = ""
        SqlCmd = "insert into [dbo].[@XSPHT] (docentry,uid,uname,flowseq,signdate,status,comment,agnname) " &
        "values(" & docnum & ",'" & signid & "','" & signname & "'," & flowseq &
        ",'" & signdate & "','" & reason & "','" & comment & "','" & agnname & "')"
        CommUtil.SqlSapExecute("ins", SqlCmd, connL)
        connL.Close()
    End Sub

    Sub HtmlToPdfGen(url As String, docnum As Long, sfid As Integer, genfile As String)
        Dim p, p1 As New Process()
        p.StartInfo.FileName = "C:\program files\wkhtmltopdf\bin\wkhtmltopdf.exe"
        p.StartInfo.Arguments = url & "signoff/printform.aspx?sfid=" & sfid & "&docnum=" & docnum & "&usingwhs=" & Session("usingwhs") & " " & genfile 'localdir & "gencltemp.pdf"
        p.StartInfo.WindowStyle = ProcessWindowStyle.Maximized 'WindowStyle可以設定開啟視窗的大小
        p.StartInfo.Verb = "open"
        p.StartInfo.CreateNoWindow = False
        p.Start()
        p.WaitForExit(1000)
        p.Close()
        p.Dispose()
    End Sub
    Function GetCurrentAttachedFileName(docnum As Long)
        Dim currentfilename As String
        Dim connsap As New SqlConnection
        Dim dr As SqlDataReader
        currentfilename = ""
        SqlCmd = "Select attachfileno " &
        "from [dbo].[@XASCH] " &
        "where docnum=" & docnum
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            currentfilename = "(" & CStr(dr(0)) & ")"
        End If
        dr.Close()
        connsap.Close()
        Return currentfilename
    End Function
    Sub Email_InformInCharge(docentry As Long, sfid As Integer, url As String)
        Dim inchargepersonid As String
        Dim inchargepersonname, inchargepersonemail As String
        Dim urlpara, body, href, subject, signmode, infostr, informdate, formname As String
        inchargepersonid = ""
        inchargepersonname = ""
        inchargepersonemail = ""
        subject = ""
        signmode = "todoitem"
        infostr = "追蹤單據日期設定通知"
        informdate = Format(Now(), "yyyy/MM/dd HH:mm:ss")
        SqlCmd = "Select incharge,subject " &
         "FROM dbo.[@XSTDT] where docentry=" & docentry
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        If (dr.HasRows) Then
            dr.Read()
            inchargepersonid = dr(0)
            subject = dr(1)
        End If
        dr.Close()
        connsap.Close()
        SqlCmd = "select name,email from dbo.[user] where id='" & inchargepersonid & "'"
        dr = CommUtil.SelectLocalSqlUsingDr(SqlCmd, conn)
        If (dr.HasRows) Then
            dr.Read()
            inchargepersonname = dr(0)
            inchargepersonemail = dr(1)
        End If
        dr.Close()
        conn.Close()
        SqlCmd = "select T0.subject,T1.sfname from  [dbo].[@XASCH] T0 INNER JOIN [dbo].[@XSFTT] T1 ON T0.sfid=T1.sfid where T0.docnum=" & docentry
        dr = CommUtil.SelectSapSqlUsingDr(SqlCmd, connsap)
        dr.Read()
        formname = dr(1)
        dr.Close()
        connsap.Close()
        href = url & "usermgm/login.aspx"
        urlpara = "?actmode=" & signmode & "&uid=" & inchargepersonid & "&inchargeid=" & inchargepersonid & "&docentry=" & docentry & "&sfid=" & sfid
        body = "<span><h5>此信件為系統發出信件，請勿直接回覆，感謝您的配合!</h5></span>" &
                "<table border=1 width=300 border-collapse:collapse>" &
                "<tr bgcolor=#add8e6><td><h1 align=center>" & infostr & "</h1></td></tr>" &
                "<tr><td align=center><a href=" & href & urlpara & ">前往追蹤單據設定</a></td></tr>" &
                "<tr><td>待通知人&nbsp;:&nbsp;" & inchargepersonname & "</td></tr>" &
                "<tr><td>單據編號&nbsp;:&nbsp;" & docentry & "</td></tr>" &
                "<tr><td>單據名稱&nbsp;:&nbsp;" & formname & "</td></tr>" &
                "<tr><td>主旨&nbsp;:&nbsp;" & subject & "</td></tr>" &
                "<tr><td>通知日期&nbsp;:&nbsp;" & informdate & "</td></tr>" &
                "</table>"
        subject = "(捷智通知) " & infostr & ":" & formname & " - " & subject
        'inchargepersonemail = "ron@jettech.com.tw" 'temp test
        CommUtil.SendMail(inchargepersonemail, subject, body)
    End Sub

    Function ArchiveCheck()
        Dim connL, connL1 As New SqlConnection
        Dim drL, drL1 As SqlDataReader
        Dim sfid, count, propc As Integer
        Dim mes As String
        sfid = 0
        count = 0
        propc = 0
        SqlCmd = "select count(*),sfid from [@XSPMT] T0 where T0.prop=1 group by sfid"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            Do While (drL.Read())
                propc = drL(0)
                If (drL(0) > 1) Then
                    sfid = drL(1)
                    Exit Do
                End If
            Loop
        End If
        drL.Close()
        connL.Close()
        If (sfid <> 0) Then
            SqlCmd = "select uid from [@XSPMT] where prop=1 and sfid=" & sfid
            drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
            If (drL.HasRows) Then
                Do While (drL.Read())
                    SqlCmd = "select count(*) from [dbo].[@XSDET] where uid='" & drL(0) & "' and sfid=" & sfid
                    drL1 = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL1)
                    If (drL1.HasRows) Then
                        drL1.Read()
                        If (drL1(0) = 0) Then
                            count = count + 1
                        End If
                    End If
                    drL1.Close()
                    connL1.Close()
                Loop
            End If
        End If
        mes = CStr(sfid) & "-" & CStr(count) & "-" & CStr(propc)
        Return mes
    End Function
    Function IsSelfForm(sfid As Integer)
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim selfformflag As Integer
        selfformflag = 0
        SqlCmd = "select selfform from [@XSFTT] T0 where sfid=" & sfid
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            selfformflag = drL(0)
        End If
        drL.Close()
        connL.Close()
        Return selfformflag
    End Function

    Function FormStatusMes(docnum As Long, docstatus As String)
        Dim actionstatus, nowformstatus As String
        Dim connL As New SqlConnection
        Dim drL As SqlDataReader
        Dim info(2) As String
        Dim delflag As Boolean
        delflag = False

        actionstatus = ""
        nowformstatus = ""
        SqlCmd = "SELECT status " &
                "FROM dbo.[@XSPHT] where uid='" & Session("s_id") & "' and docentry=" & docnum & " order by flowseq desc"
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            actionstatus = drL(0)
        Else
            'If (docstatus <> "E" And docstatus <> "D") Then
            '    actionstatus = "刪除"
            'End If
        End If
        drL.Close()
        connL.Close()
        SqlCmd = "SELECT status " &
                        "FROM dbo.[@XASCH] where docnum=" & docnum
        drL = CommUtil.SelectSapSqlUsingDr(SqlCmd, connL)
        If (drL.HasRows) Then
            drL.Read()
            If (drL(0) = "O") Then
                nowformstatus = "且此單目前狀態還在簽核中"
            ElseIf (drL(0) = "F") Then
                nowformstatus = "且此單目前狀態已簽核結案"
            ElseIf (drL(0) = "C") Then
                nowformstatus = "且此單目前狀態已被作廢"
                delflag = True
            ElseIf (drL(0) = "R") Then
                nowformstatus = "且此單目前狀態被抽回"
            ElseIf (drL(0) = "B") Then
                nowformstatus = "且此單目前狀態已被退回"
            ElseIf (drL(0) = "T") Then
                nowformstatus = "且此單目前狀態已簽核結案並歸檔"
                docstatus = "T" '當其用email 之link 點選時 , 若之前已被歸檔過 , 因之前docstatus=F , 此處要改為T , 才能使歸檔button隱藏
            End If
        Else
            nowformstatus = "且此單目前狀態被刪除"
        End If
        drL.Close()
        connL.Close()
        'row = 1000 '
        info(0) = docstatus
        If (actionstatus <> "") Then
            If (actionstatus <> "跳過簽核") Then
                info(1) = "此筆簽核資料曾被您(或代理人) '" & actionstatus & "' 覆核過," & nowformstatus
            Else
                info(1) = "此筆簽核資料已被管理者 '" & actionstatus & "'," & nowformstatus
            End If

        Else
            If (delflag) Then
                info(1) = "此筆簽核資料已被刪除," & nowformstatus
            Else
                info(1) = "此筆簽核資料已被前一人抽單," & nowformstatus
            End If
        End If
        Return info
    End Function
End Class
