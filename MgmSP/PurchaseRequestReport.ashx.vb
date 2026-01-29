Imports System.Data.SqlClient
Imports System.IO
Imports System.Web
Imports System.Web.Configuration
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports iTextSharp.text
Imports iTextSharp.text.pdf

''' <summary>
''' 請購單 Crystal Report PDF 匯出 Handler
''' 使用方式：PurchaseRequestReport.ashx?jID={平台單號}
''' 合併附件：PurchaseRequestReport.ashx?jID={平台單號}&mergeAttach={附件ID列表，逗號分隔}
''' </summary>
Public Class PurchaseRequestReport
    Implements IHttpHandler, System.Web.SessionState.IRequiresSessionState

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    Public Sub ProcessRequest(context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim report As ReportDocument = Nothing

        Try
            ' 檢查登入狀態
            If context.Session("s_id") Is Nothing Then
                context.Response.StatusCode = 401
                context.Response.Write("未授權的存取")
                Return
            End If

            ' 取得 jID 參數
            Dim jID As String = context.Request.QueryString("jID")
            If String.IsNullOrEmpty(jID) Then
                context.Response.StatusCode = 400
                context.Response.Write("缺少必要參數 jID")
                Return
            End If

            ' 驗證 jID 是否存在
            If Not ValidateJID(jID) Then
                context.Response.StatusCode = 404
                context.Response.Write("找不到指定的請購單")
                Return
            End If

            ' 載入 Crystal Report
            Dim reportPath As String = context.Server.MapPath("~/CrystalReport/PurchaseRequest.rpt")
            If Not File.Exists(reportPath) Then
                context.Response.StatusCode = 500
                context.Response.Write("報表檔案不存在: " & reportPath)
                Return
            End If

            report = New ReportDocument()
            report.Load(reportPath)

            ' 設定資料庫連線
            SetDatabaseLogon(report)

            ' 設定參數
            report.SetParameterValue("myjid", jID)

            ' 檢查是否有合併附件參數
            Dim mergeAttach As String = context.Request.QueryString("mergeAttach")
            If Not String.IsNullOrEmpty(mergeAttach) Then
                ' 有合併參數，執行 PDF 合併
                ExportMergedPdf(report, context, jID, mergeAttach)
            Else
                ' 無合併參數，直接匯出
                ExportToPdfResponse(report, context, jID)
            End If

        Catch ex As Exception
            context.Response.StatusCode = 500
            Dim errorMsg As String = "報表產生失敗: " & ex.Message
            If ex.InnerException IsNot Nothing Then
                errorMsg &= "<br/><br/>內部錯誤: " & ex.InnerException.Message
                If ex.InnerException.InnerException IsNot Nothing Then
                    errorMsg &= "<br/><br/>詳細: " & ex.InnerException.InnerException.Message
                End If
            End If
            context.Response.Write(errorMsg)
        Finally
            ' 釋放報表資源
            If report IsNot Nothing Then
                report.Close()
                report.Dispose()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 設定 Crystal Report 資料庫登入資訊
    ''' 根據資料表原本的資料庫名稱，對應到正確的連線字串：
    ''' - jtdb → jtdbConnectionString
    ''' - JTSTD/JTSTD1031/jettech → SapSQLConnection
    ''' </summary>
    Private Sub SetDatabaseLogon(report As ReportDocument)
        Try
            ' 取得兩個連線字串
            Dim jtdbConnStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
            Dim sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString

            Dim jtdbBuilder As New SqlConnectionStringBuilder(jtdbConnStr)
            Dim sapBuilder As New SqlConnectionStringBuilder(sapConnStr)

            ' 設定主報表的每個資料表連線
            For Each table As CrystalDecisions.CrystalReports.Engine.Table In report.Database.Tables
                Dim logonInfo As CrystalDecisions.Shared.TableLogOnInfo = table.LogOnInfo
                Dim originalDbName As String = logonInfo.ConnectionInfo.DatabaseName.ToUpper()

                ' 根據原本的資料庫名稱判斷使用哪個連線
                If originalDbName.Contains("JTSTD") OrElse originalDbName.Contains("JETTECH") Then
                    ' SAP 資料庫
                    logonInfo.ConnectionInfo.ServerName = sapBuilder.DataSource
                    logonInfo.ConnectionInfo.DatabaseName = sapBuilder.InitialCatalog
                    logonInfo.ConnectionInfo.UserID = sapBuilder.UserID
                    logonInfo.ConnectionInfo.Password = sapBuilder.Password
                Else
                    ' jtdb 資料庫 (預設)
                    logonInfo.ConnectionInfo.ServerName = jtdbBuilder.DataSource
                    logonInfo.ConnectionInfo.DatabaseName = jtdbBuilder.InitialCatalog
                    logonInfo.ConnectionInfo.UserID = jtdbBuilder.UserID
                    logonInfo.ConnectionInfo.Password = jtdbBuilder.Password
                End If

                logonInfo.ConnectionInfo.IntegratedSecurity = False
                table.ApplyLogOnInfo(logonInfo)
            Next

            ' 處理子報表（如果有的話）
            For Each section As CrystalDecisions.CrystalReports.Engine.Section In report.ReportDefinition.Sections
                For Each reportObject As CrystalDecisions.CrystalReports.Engine.ReportObject In section.ReportObjects
                    If reportObject.Kind = CrystalDecisions.Shared.ReportObjectKind.SubreportObject Then
                        Dim subReportObj As CrystalDecisions.CrystalReports.Engine.SubreportObject = CType(reportObject, CrystalDecisions.CrystalReports.Engine.SubreportObject)
                        Dim subReport As ReportDocument = subReportObj.OpenSubreport(subReportObj.SubreportName)

                        For Each table As CrystalDecisions.CrystalReports.Engine.Table In subReport.Database.Tables
                            Dim logonInfo As CrystalDecisions.Shared.TableLogOnInfo = table.LogOnInfo
                            Dim originalDbName As String = logonInfo.ConnectionInfo.DatabaseName.ToUpper()

                            If originalDbName.Contains("JTSTD") OrElse originalDbName.Contains("JETTECH") Then
                                logonInfo.ConnectionInfo.ServerName = sapBuilder.DataSource
                                logonInfo.ConnectionInfo.DatabaseName = sapBuilder.InitialCatalog
                                logonInfo.ConnectionInfo.UserID = sapBuilder.UserID
                                logonInfo.ConnectionInfo.Password = sapBuilder.Password
                            Else
                                logonInfo.ConnectionInfo.ServerName = jtdbBuilder.DataSource
                                logonInfo.ConnectionInfo.DatabaseName = jtdbBuilder.InitialCatalog
                                logonInfo.ConnectionInfo.UserID = jtdbBuilder.UserID
                                logonInfo.ConnectionInfo.Password = jtdbBuilder.Password
                            End If

                            logonInfo.ConnectionInfo.IntegratedSecurity = False
                            table.ApplyLogOnInfo(logonInfo)
                        Next
                    End If
                Next
            Next

        Catch ex As Exception
            Throw New Exception("資料庫連線設定失敗: " & ex.Message, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 匯出為 PDF 並寫入 Response
    ''' </summary>
    Private Sub ExportToPdfResponse(report As ReportDocument, context As HttpContext, jID As String)
        ' 匯出為 PDF Stream
        Dim pdfStream As Stream = report.ExportToStream(ExportFormatType.PortableDocFormat)

        ' 輸出 PDF
        context.Response.Clear()
        context.Response.Buffer = True
        context.Response.ContentType = "application/pdf"

        ' 設定檔名
        Dim safeFileName As String = "PurchaseRequest_" & SanitizeFileName(jID) & ".pdf"
        context.Response.AddHeader("Content-Disposition", "inline; filename=""" & safeFileName & """")

        ' 將 Stream 寫入 Response
        Dim buffer(4096) As Byte
        Dim bytesRead As Integer
        pdfStream.Position = 0
        Do
            bytesRead = pdfStream.Read(buffer, 0, buffer.Length)
            If bytesRead > 0 Then
                context.Response.OutputStream.Write(buffer, 0, bytesRead)
            End If
        Loop While bytesRead > 0

        context.Response.Flush()
    End Sub

    ''' <summary>
    ''' 匯出合併後的 PDF（主報表 + 附件 PDF）
    ''' </summary>
    Private Sub ExportMergedPdf(report As ReportDocument, context As HttpContext, jID As String, attachIds As String)
        Dim tempFolder As String = context.Server.MapPath("~/Temp/")
        Dim mainPdfPath As String = ""
        Dim mergedPdfPath As String = ""

        Try
            ' 確保暫存資料夾存在
            If Not Directory.Exists(tempFolder) Then
                Directory.CreateDirectory(tempFolder)
            End If

            ' 1. 匯出主報表到暫存檔
            mainPdfPath = Path.Combine(tempFolder, "PR_" & jID & "_" & Guid.NewGuid().ToString("N") & ".pdf")
            report.ExportToDisk(ExportFormatType.PortableDocFormat, mainPdfPath)

            ' 2. 取得附件 PDF 檔案路徑
            Dim attachPdfPaths As New List(Of String)
            Dim ids() As String = attachIds.Split(","c)

            For Each idStr As String In ids
                Dim attachId As Integer
                If Integer.TryParse(idStr.Trim(), attachId) Then
                    Dim attachPath As String = GetAttachmentPath(attachId)
                    If Not String.IsNullOrEmpty(attachPath) Then
                        Dim fullPath As String = context.Server.MapPath("~/" & attachPath)
                        If File.Exists(fullPath) AndAlso Path.GetExtension(fullPath).ToLower() = ".pdf" Then
                            attachPdfPaths.Add(fullPath)
                        End If
                    End If
                End If
            Next

            ' 3. 如果沒有有效的附件 PDF，直接輸出主報表
            If attachPdfPaths.Count = 0 Then
                OutputPdfFile(context, mainPdfPath, jID)
                Return
            End If

            ' 4. 合併 PDF
            mergedPdfPath = Path.Combine(tempFolder, "PR_Merged_" & jID & "_" & Guid.NewGuid().ToString("N") & ".pdf")

            ' 建立要合併的 PDF 清單（主報表 + 附件）
            Dim allPdfFiles As New List(Of String)
            allPdfFiles.Add(mainPdfPath)
            allPdfFiles.AddRange(attachPdfPaths)

            MergePdfFiles(allPdfFiles, mergedPdfPath)

            ' 5. 輸出合併後的 PDF
            OutputPdfFile(context, mergedPdfPath, jID & "_Combined")

        Finally
            ' 6. 清理暫存檔
            Try
                If Not String.IsNullOrEmpty(mainPdfPath) AndAlso File.Exists(mainPdfPath) Then
                    File.Delete(mainPdfPath)
                End If
                If Not String.IsNullOrEmpty(mergedPdfPath) AndAlso File.Exists(mergedPdfPath) Then
                    File.Delete(mergedPdfPath)
                End If
            Catch
                ' 忽略清理失敗
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' 從資料庫取得附件路徑
    ''' </summary>
    Private Function GetAttachmentPath(attachId As Integer) As String
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT FilePath FROM jAttach WHERE ID = @ID AND (IsDeleted = 0 OR IsDeleted IS NULL)"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ID", attachId)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End Using
        Catch
            ' 忽略錯誤
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' 使用 iTextSharp 合併多個 PDF 檔案
    ''' </summary>
    Private Sub MergePdfFiles(pdfFiles As List(Of String), outputPath As String)
        Dim document As iTextSharp.text.Document = Nothing
        Dim writer As PdfCopy = Nothing

        Try
            document = New iTextSharp.text.Document()
            writer = New PdfCopy(document, New FileStream(outputPath, FileMode.Create))
            document.Open()

            For Each pdfFile As String In pdfFiles
                Dim reader As PdfReader = Nothing
                Try
                    reader = New PdfReader(pdfFile)
                    Dim pageCount As Integer = reader.NumberOfPages

                    For i As Integer = 1 To pageCount
                        Dim page As PdfImportedPage = writer.GetImportedPage(reader, i)
                        writer.AddPage(page)
                    Next
                Finally
                    If reader IsNot Nothing Then
                        reader.Close()
                    End If
                End Try
            Next

        Finally
            If document IsNot Nothing Then
                document.Close()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 輸出 PDF 檔案到 Response
    ''' </summary>
    Private Sub OutputPdfFile(context As HttpContext, pdfPath As String, fileNamePrefix As String)
        context.Response.Clear()
        context.Response.Buffer = True
        context.Response.ContentType = "application/pdf"

        Dim safeFileName As String = "PurchaseRequest_" & SanitizeFileName(fileNamePrefix) & ".pdf"
        context.Response.AddHeader("Content-Disposition", "inline; filename=""" & safeFileName & """")

        context.Response.TransmitFile(pdfPath)
        context.Response.Flush()
    End Sub

    ''' <summary>
    ''' 驗證 jID 是否存在
    ''' </summary>
    Private Function ValidateJID(jID As String) As Boolean
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT COUNT(1) FROM jOPRQ WHERE jID = @jID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@jID", jID)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 清理檔名中的非法字元
    ''' </summary>
    Private Function SanitizeFileName(fileName As String) As String
        Dim invalidChars As Char() = Path.GetInvalidFileNameChars()
        Dim result As String = fileName
        For Each c As Char In invalidChars
            result = result.Replace(c, "_"c)
        Next
        Return result
    End Function

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class
