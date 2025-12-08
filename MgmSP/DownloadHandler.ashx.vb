Imports System.Data.SqlClient
Imports System.IO
Imports System.Web
Imports System.Web.Configuration

''' <summary>
''' [C] 附件下載 Handler
''' 用途：安全下載附件，避免路徑洩漏
''' 使用方式：DownloadHandler.ashx?id={AttachmentID}
''' </summary>
Public Class DownloadHandler
    Implements IHttpHandler, System.Web.SessionState.IRequiresSessionState

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    Public Sub ProcessRequest(context As HttpContext) Implements IHttpHandler.ProcessRequest
        Try
            ' 檢查登入狀態
            If context.Session("s_id") Is Nothing Then
                context.Response.StatusCode = 401
                context.Response.Write("未授權的存取")
                Return
            End If

            Dim userId As String = context.Session("s_id").ToString()
            Dim attachId As Integer = 0

            If Not Integer.TryParse(context.Request.QueryString("id"), attachId) OrElse attachId <= 0 Then
                context.Response.StatusCode = 400
                context.Response.Write("無效的附件 ID")
                Return
            End If

            ' 從資料庫取得附件資訊
            Dim filePath As String = ""
            Dim fileName As String = ""
            Dim docEntry As Integer = 0

            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT a.FilePath, a.FileName, a.DocEntry, h.CreateBy " &
                                    "FROM jAttach a " &
                                    "INNER JOIN jOPCH h ON a.DocEntry = h.DocEntry " &
                                    "WHERE a.ID = @ID AND (a.IsDeleted = 0 OR a.IsDeleted IS NULL)"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ID", attachId)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            filePath = dr("FilePath").ToString()
                            fileName = If(IsDBNull(dr("FileName")), "", dr("FileName").ToString())
                            docEntry = Convert.ToInt32(dr("DocEntry"))
                            ' 可選：檢查使用者是否有權限查看此附件
                            ' Dim createBy As String = dr("CreateBy").ToString()
                        Else
                            context.Response.StatusCode = 404
                            context.Response.Write("找不到附件")
                            Return
                        End If
                    End Using
                End Using
            End Using

            ' 組合完整路徑
            Dim fullPath As String = context.Server.MapPath("~/" & filePath)

            If Not File.Exists(fullPath) Then
                context.Response.StatusCode = 404
                context.Response.Write("檔案不存在")
                Return
            End If

            ' 若 FileName 為空，從 FilePath 取得
            If String.IsNullOrEmpty(fileName) Then
                fileName = Path.GetFileName(filePath)
            End If

            ' 設定 Response Header
            context.Response.Clear()
            context.Response.ContentType = GetContentType(fileName)
            context.Response.AddHeader("Content-Disposition", "attachment; filename=""" & HttpUtility.UrlEncode(fileName) & """")
            context.Response.AddHeader("Content-Length", New FileInfo(fullPath).Length.ToString())

            ' 輸出檔案
            context.Response.TransmitFile(fullPath)
            context.Response.Flush()
            context.Response.End()

        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write("下載失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 根據檔案副檔名取得 MIME Type
    ''' </summary>
    Private Function GetContentType(fileName As String) As String
        Dim ext As String = Path.GetExtension(fileName).ToLower()
        Select Case ext
            Case ".pdf" : Return "application/pdf"
            Case ".doc" : Return "application/msword"
            Case ".docx" : Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case ".xls" : Return "application/vnd.ms-excel"
            Case ".xlsx" : Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Case ".ppt" : Return "application/vnd.ms-powerpoint"
            Case ".pptx" : Return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            Case ".jpg", ".jpeg" : Return "image/jpeg"
            Case ".png" : Return "image/png"
            Case ".gif" : Return "image/gif"
            Case ".bmp" : Return "image/bmp"
            Case ".txt" : Return "text/plain"
            Case ".csv" : Return "text/csv"
            Case ".zip" : Return "application/zip"
            Case ".rar" : Return "application/x-rar-compressed"
            Case ".7z" : Return "application/x-7z-compressed"
            Case Else : Return "application/octet-stream"
        End Select
    End Function

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class
