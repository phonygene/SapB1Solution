'Imports Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Data.SqlClient
Imports AjaxControlToolkit
Imports Microsoft.Office.Interop
Imports Microsoft.Win32
Imports System.IO
Imports System.Threading
'Imports System.Windows.Forms

Public Class workevent
    Inherits System.Web.UI.Page
    Public CommUtil As New CommUtil
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim oExcel As Excel.Application
        Dim oBook As Excel.Workbook
        Dim oBooks As Excel.Workbooks
        Dim oSheet As Excel.Worksheet
        Dim filepath, str() As String
        Try
            'Dim Dialog As OpenFileDialog
            'Dialog = New OpenFileDialog()
            'Dialog.InitialDirectory = "C:\"
            'Dialog.Filter = "xls files (*.xls)|*.xls"
            ''System.IO.Path.GetFullPath(dialog.FileName)
            'If Dialog.ShowDialog() Then
            '    'MsgBox(dialog.FileName)
            'End If
            If (FileUpload1.HasFile) Then
                str = Split(FileUpload1.FileName, ".")
                If (str(1) <> "xls" And str(1) <> "xlsx") Then
                    CommUtil.ShowMsg(Me, "要上傳檔案需為xls or xlsx")
                    Exit Sub
                End If
                FileUpload1.SaveAs(Application("localdir") & FileUpload1.FileName)
            Else
                CommUtil.ShowMsg(Me, "未指定上傳檔案")
            End If
            '建立Excel物件並開啟C:\01.xls中的Sheet1
            oExcel = CreateObject("Excel.Application")
            oExcel.Visible = True
            oBooks = oExcel.Workbooks
            oBook = oBooks.Open(Application("localdir") & FileUpload1.FileName)
            Try
                oSheet = oBook.Worksheets("工作表1")
            Catch ex As Exception
                CommUtil.ShowMsg(Me, "無工作表1")
            End Try
            Dim i As Integer
            i = 2
            Do While oSheet.Cells(i, 1).value <> ""

            Loop
            '讀取Sheet1中的A1儲存格
            Dim test As String
            'test = oSheet.Cells(1, 1).value.ToString
            'MsgBox(test)
            oBook.SaveAs("C:\test\02.xlsx")
            '關閉並釋放Excel物件
            oBook.Close(False)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook)
            oBook = Nothing
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks)
            oBooks = Nothing
            oExcel.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel)
            oExcel = Nothing
            IO.File.Delete(Application("localdir") & FileUpload1.FileName)
        Catch ex As Exception
            CommUtil.ShowMsg(Me, "操作Excel遇到問題")
        End Try
        'Dim oExcel As Excel.Application = Nothing
        'Dim oWorkBook As Excel.Workbook = Nothing
        'Dim oSheet As Excel.Worksheet = Nothing

        'oExcel = New Excel.Application()
        'oExcel.Interactive = False
        'oExcel.DisplayAlerts = False
        'oExcel.Visible = False
        'oWorkBook = oExcel.Workbooks.Add
        'oSheet = oWorkBook.Worksheets(1)
        ''指定 cell(1,1) 的 Formula
        'oSheet.Range(oSheet.Cells(1, 1), oSheet.Cells(1, 1)).Formula = "=MyDDHHMM(A1)"
        'Dim app As New Excel.Application

        'Dim book As Excel.Workbook

        'Dim sheet As Excel.Worksheet

        'Dim range As Excel.Range


        'app.DisplayAlerts = False

        'app.Visible = False


        ''建立一個新的 Workbooks

        'book = app.Workbooks.Add


        ''將資料放置 Excel 中的 Excel

        'sheet = book.Worksheets(1)

        'sheet.Cells(1, 1).Value = "天數"

        'sheet.Cells(1, 2).Value = "人數"

        'For i As Integer = 1 To 10

        '    sheet.Cells(i + 1, 1).Value = i

        '    sheet.Cells(i + 1, 2).Value = i * 10

        'Next


        ''建立圖表

        'Dim chart As Excel.Chart

        'Dim myChart As Excel.ChartObject


        'myChart = book.Sheets(1).ChartObjects.Add(10, 10, 400, 300)

        'chart = myChart.Chart


        ''設定 Y 軸資料

        'range = book.Sheets(1).Range("B1", "B11")

        'chart.SetSourceData(Source:=range)


        ''設定 X 軸

        'chart.SeriesCollection(1).XValues = "=Sheet2!$A$2:$A$11"


        ''Chart Type 設為折線圖

        'chart.ChartType = Excel.XlChartType.xlLine


        ''另存檔案到程式目錄中的 test.xlsx

        'book.SaveAs("c:\test.xlsx")

        'book.Close()
    End Sub
End Class