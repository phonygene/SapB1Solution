'=============================================================================
' SAP B1 AP 發票 (A/P Invoice) 匯入功能 - Service Layer 版本
'
' 用途：將 jOPCH/jPCH1 資料透過 Service Layer API 匯入 SAP Business One
' 日期：2025/12/15
' 版本：1.0
'
' Service Layer vs DI API 比較：
' ┌─────────────────┬────────────────────────┬────────────────────────┐
' │ 功能            │ DI API                 │ Service Layer          │
' ├─────────────────┼────────────────────────┼────────────────────────┤
' │ 通訊協定        │ COM (本機/DCOM)        │ REST/HTTPS             │
' │ 跨平台          │ 僅 Windows             │ 任何平台               │
' │ Excel VBA 支援  │ 需安裝 SAP Client      │ ✓ 原生 HTTP 支援       │
' │ 連線管理        │ 手動管理 COM 物件      │ Session Cookie 管理    │
' │ 批次處理        │ 逐筆處理               │ Batch Request 支援     │
' │ 併發控制        │ 需自行處理             │ Session 序列化         │
' │ 寫入佇列        │ ✗ 無                   │ ✗ 無 (需自行實作)      │
' │ 效能            │ 較快 (直接記憶體)      │ 較慢 (HTTP 開銷)       │
' └─────────────────┴────────────────────────┴────────────────────────┘
'
' 注意事項：
' 1. 需要 SAP B1 10.0 PL03 以上版本
' 2. Service Layer 預設 Port: 50000 (HTTP) 或 50001 (HTTPS)
' 3. 建議使用 HTTPS 確保安全性
' 4. Session 預設 30 分鐘逾時，可透過 SessionTimeout 設定
'=============================================================================

Imports System.Net
Imports System.Net.Http
Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports System.Configuration
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' SAP AP 發票匯入類別 - Service Layer 版本
''' </summary>
Public Class SapAPInvoiceImporterSL

    Private httpClient As HttpClient
    Private sessionId As String = ""
    Private baseUrl As String = ""
    Private connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    ' 本幣代碼
    Private Const LOCAL_CURRENCY As String = "TWD"

#Region "Web.config 設定讀取"

    ''' <summary>
    ''' 從 Web.config 讀取 Service Layer 連線設定
    ''' </summary>
    Private Class SLConfig
        ''' <summary>Service Layer 基礎 URL (例如: https://192.168.1.219:50001/b1s/v1)</summary>
        Public Shared ReadOnly Property BaseUrl As String
            Get
                Return ConfigurationManager.AppSettings("SL_BaseUrl")
            End Get
        End Property

        ''' <summary>SAP 公司資料庫名稱</summary>
        Public Shared ReadOnly Property CompanyDB As String
            Get
                Return ConfigurationManager.AppSettings("SL_CompanyDB")
            End Get
        End Property

        ''' <summary>SAP 登入帳號</summary>
        Public Shared ReadOnly Property UserName As String
            Get
                Return ConfigurationManager.AppSettings("SL_UserName")
            End Get
        End Property

        ''' <summary>SAP 登入密碼</summary>
        Public Shared ReadOnly Property Password As String
            Get
                Return ConfigurationManager.AppSettings("SL_Password")
            End Get
        End Property
    End Class

#End Region

#Region "Service Layer 連線"

    ''' <summary>
    ''' 登入 Service Layer，取得 Session ID
    ''' </summary>
    Public Async Function LoginAsync() As Task
        Try
            ' 初始化 HttpClient (忽略 SSL 憑證驗證 - 僅供測試)
            Dim handler As New HttpClientHandler()
            handler.ServerCertificateCustomValidationCallback = Function(sender, cert, chain, sslPolicyErrors) True

            httpClient = New HttpClient(handler)
            httpClient.Timeout = TimeSpan.FromMinutes(5)
            baseUrl = SLConfig.BaseUrl

            ' 建立登入請求
            Dim loginUrl As String = baseUrl & "/Login"
            Dim loginData As New JObject()
            loginData("CompanyDB") = SLConfig.CompanyDB
            loginData("UserName") = SLConfig.UserName
            loginData("Password") = SLConfig.Password

            Dim content As New StringContent(loginData.ToString(), Encoding.UTF8, "application/json")
            Dim response As HttpResponseMessage = Await httpClient.PostAsync(loginUrl, content)

            If Not response.IsSuccessStatusCode Then
                Dim errorContent As String = Await response.Content.ReadAsStringAsync()
                Throw New Exception("Service Layer 登入失敗: " & errorContent)
            End If

            ' 從 Cookie 取得 Session ID
            Dim cookies As IEnumerable(Of String) = Nothing
            If response.Headers.TryGetValues("Set-Cookie", cookies) Then
                For Each cookie As String In cookies
                    If cookie.StartsWith("B1SESSION=") Then
                        sessionId = cookie.Split(";"c)(0).Replace("B1SESSION=", "")
                        Exit For
                    End If
                Next
            End If

            If String.IsNullOrEmpty(sessionId) Then
                Throw New Exception("無法取得 Session ID")
            End If

            ' 設定後續請求的 Cookie
            httpClient.DefaultRequestHeaders.Add("Cookie", "B1SESSION=" & sessionId)

        Catch ex As Exception
            Throw New Exception("Service Layer 連線失敗: " & ex.Message, ex)
        End Try
    End Function

    ''' <summary>
    ''' 登出 Service Layer
    ''' </summary>
    Public Async Function LogoutAsync() As Task
        Try
            If httpClient IsNot Nothing AndAlso Not String.IsNullOrEmpty(sessionId) Then
                Dim logoutUrl As String = baseUrl & "/Logout"
                Await httpClient.PostAsync(logoutUrl, Nothing)
            End If
        Catch
            ' 忽略登出錯誤
        Finally
            If httpClient IsNot Nothing Then
                httpClient.Dispose()
                httpClient = Nothing
            End If
            sessionId = ""
        End Try
    End Function

    ''' <summary>
    ''' 檢查是否已登入
    ''' </summary>
    Public ReadOnly Property IsLoggedIn As Boolean
        Get
            Return httpClient IsNot Nothing AndAlso Not String.IsNullOrEmpty(sessionId)
        End Get
    End Property

#End Region

#Region "AP 發票匯入"

    ''' <summary>
    ''' 匯入單一 AP 發票到 SAP (非同步版本)
    ''' </summary>
    ''' <param name="docEntry">jOPCH 的 DocEntry</param>
    ''' <returns>成功回傳 SAP DocEntry，失敗拋出 Exception</returns>
    Public Async Function ImportAPInvoiceAsync(docEntry As Integer) As Task(Of Integer)
        Dim sapDocEntry As Integer = -1

        Try
            ' 檢查連線
            If Not IsLoggedIn Then
                Throw New Exception("Service Layer 未登入")
            End If

            ' 建立 AP Invoice JSON 物件
            Dim invoice As New JObject()

            ' 文件幣別和匯率
            Dim docCurrency As String = LOCAL_CURRENCY
            Dim docRate As Double = 1.0
            Dim isForeignCurrency As Boolean = False

            ' 從 jtdb 讀取資料並建立 JSON
            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 讀取 Header
                Dim sqlH As String = "SELECT * FROM jOPCH WHERE DocEntry = @DocEntry"
                Using cmdH As New SqlCommand(sqlH, conn)
                    cmdH.Parameters.AddWithValue("@DocEntry", docEntry)
                    Using drH As SqlDataReader = cmdH.ExecuteReader()
                        If Not drH.Read() Then
                            Throw New Exception("找不到文件: " & docEntry)
                        End If

                        ' ===== 表頭欄位 =====
                        invoice("CardCode") = drH("CardCode").ToString()
                        invoice("CardName") = drH("CardName").ToString()

                        ' 供應商參考號 (發票號碼)
                        If Not IsDBNull(drH("NumAtCard")) Then
                            invoice("NumAtCard") = drH("NumAtCard").ToString()
                        End If

                        ' 文件日期 (格式: yyyy-MM-dd)
                        invoice("DocDate") = Convert.ToDateTime(drH("DocDate")).ToString("yyyy-MM-dd")

                        ' 到期日
                        If Not IsDBNull(drH("DocDueDate")) Then
                            invoice("DocDueDate") = Convert.ToDateTime(drH("DocDueDate")).ToString("yyyy-MM-dd")
                        End If

                        ' 過帳日期 (Tax Date)
                        If Not IsDBNull(drH("TaxDate")) Then
                            invoice("TaxDate") = Convert.ToDateTime(drH("TaxDate")).ToString("yyyy-MM-dd")
                        End If

                        ' 幣別
                        If Not IsDBNull(drH("DocCurrency")) Then
                            docCurrency = drH("DocCurrency").ToString()
                            invoice("DocCurrency") = docCurrency
                            isForeignCurrency = (docCurrency <> LOCAL_CURRENCY)
                        End If

                        ' 匯率 (若非本幣)
                        If Not IsDBNull(drH("DocRate")) Then
                            docRate = Convert.ToDouble(drH("DocRate"))
                            If isForeignCurrency Then
                                invoice("DocRate") = docRate
                            End If
                        End If

                        ' 付款條件
                        If Not IsDBNull(drH("GroupNum")) Then
                            invoice("PaymentGroupCode") = Convert.ToInt32(drH("GroupNum"))
                        End If

                        ' 備註
                        If Not IsDBNull(drH("Comments")) Then
                            invoice("Comments") = drH("Comments").ToString()
                        End If

                        ' 日記帳備註
                        If Not IsDBNull(drH("JrnlMemo")) Then
                            invoice("JournalMemo") = drH("JrnlMemo").ToString()
                        End If

                        ' 收貨地址
                        If Not IsDBNull(drH("Address")) Then
                            invoice("Address") = drH("Address").ToString()
                        End If

                        ' 採購人員
                        If Not IsDBNull(drH("SlpCode")) Then
                            invoice("SalesPersonCode") = Convert.ToInt32(drH("SlpCode"))
                        End If

                        drH.Close()
                    End Using
                End Using

                ' 讀取 Lines (明細)
                Dim lines As New JArray()
                Dim sqlL As String = "SELECT * FROM jPCH1 WHERE DocEntry = @DocEntry ORDER BY LineNum"
                Using cmdL As New SqlCommand(sqlL, conn)
                    cmdL.Parameters.AddWithValue("@DocEntry", docEntry)
                    Using drL As SqlDataReader = cmdL.ExecuteReader()
                        While drL.Read()
                            Dim line As New JObject()

                            ' 會計科目 (費用類)
                            line("AccountCode") = drL("AcctCode").ToString()

                            ' 說明
                            If Not IsDBNull(drL("Dscription")) Then
                                line("ItemDescription") = drL("Dscription").ToString()
                            End If

                            ' ===== 金額處理 =====
                            Dim lineTotal As Double = 0
                            If Not IsDBNull(drL("LineTotal")) Then
                                lineTotal = Convert.ToDouble(drL("LineTotal"))
                            End If

                            If isForeignCurrency Then
                                ' 外幣單據
                                line("TotalForeignCurrency") = lineTotal
                                line("LineTotal") = Math.Round(lineTotal * docRate, 2, MidpointRounding.AwayFromZero)
                            Else
                                ' 本幣單據
                                line("LineTotal") = lineTotal
                            End If

                            ' 稅碼
                            If Not IsDBNull(drL("VatGroup")) Then
                                line("VatGroup") = drL("VatGroup").ToString()
                            End If

                            ' 成本中心1 (產品/專案)
                            If Not IsDBNull(drL("CostingCode")) AndAlso drL("CostingCode").ToString() <> "" Then
                                line("CostingCode") = drL("CostingCode").ToString()
                            End If

                            ' 成本中心2 (部門)
                            If Not IsDBNull(drL("CostingCode2")) AndAlso drL("CostingCode2").ToString() <> "" Then
                                line("CostingCode2") = drL("CostingCode2").ToString()
                            End If

                            lines.Add(line)
                        End While

                        drL.Close()
                    End Using
                End Using

                invoice("DocumentLines") = lines
                conn.Close()
            End Using

            ' ===== 發送 POST 請求到 Service Layer =====
            Dim url As String = baseUrl & "/PurchaseInvoices"
            Dim content As New StringContent(invoice.ToString(), Encoding.UTF8, "application/json")

            Dim response As HttpResponseMessage = Await httpClient.PostAsync(url, content)
            Dim responseContent As String = Await response.Content.ReadAsStringAsync()

            If response.IsSuccessStatusCode Then
                ' 成功，解析回傳的 DocEntry
                Dim result As JObject = JObject.Parse(responseContent)
                sapDocEntry = Convert.ToInt32(result("DocEntry"))

                ' 更新 jtdb 的狀態
                UpdatePostStatus(docEntry, sapDocEntry, "Y", "")

                Return sapDocEntry
            Else
                ' 失敗，解析錯誤訊息
                Dim errMsg As String = responseContent
                Try
                    Dim errorObj As JObject = JObject.Parse(responseContent)
                    If errorObj("error") IsNot Nothing Then
                        errMsg = errorObj("error")("message")("value").ToString()
                    End If
                Catch
                    ' 保持原始錯誤訊息
                End Try

                ' 更新 jtdb 的錯誤狀態
                UpdatePostStatus(docEntry, 0, "E", errMsg)

                Throw New Exception("SAP 新增失敗: " & errMsg)
            End If

        Catch ex As Exception
            ' 記錄錯誤
            UpdatePostStatus(docEntry, 0, "E", ex.Message)
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 更新 jtdb 的過帳狀態
    ''' </summary>
    Private Sub UpdatePostStatus(docEntry As Integer, sapDocEntry As Integer, status As String, errMsg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE jOPCH SET " &
                                   "DocEntry = CASE WHEN @SapDocEntry > 0 THEN @SapDocEntry ELSE DocEntry END, " &
                                   "B1PostStatus = @Status, " &
                                   "B1PostDate = GETDATE(), " &
                                   "B1ErrMsg = @ErrMsg " &
                                   "WHERE DocEntry = @DocEntry"

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                    cmd.Parameters.AddWithValue("@SapDocEntry", sapDocEntry)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@ErrMsg", If(String.IsNullOrEmpty(errMsg), DBNull.Value, errMsg))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
            ' 忽略更新狀態的錯誤
        End Try
    End Sub

#End Region

#Region "批次匯入 (使用 Batch Request)"

    ''' <summary>
    ''' 批次匯入多筆 AP 發票 (使用 Service Layer Batch Request)
    ''' 這是 Service Layer 相對於 DI API 的一個優勢 - 可以將多個操作合併成一個 HTTP 請求
    ''' </summary>
    ''' <param name="docEntries">要匯入的 DocEntry 清單</param>
    ''' <returns>匯入結果 (成功數, 失敗數)</returns>
    Public Async Function ImportBatchAsync(docEntries As List(Of Integer)) As Task(Of Tuple(Of Integer, Integer))
        Dim successCount As Integer = 0
        Dim failCount As Integer = 0

        ' Service Layer Batch Request 限制每批最多 100 筆
        Const BATCH_SIZE As Integer = 100

        For i As Integer = 0 To docEntries.Count - 1 Step BATCH_SIZE
            Dim batch = docEntries.Skip(i).Take(BATCH_SIZE).ToList()
            Dim batchResult = Await ProcessBatchAsync(batch)
            successCount += batchResult.Item1
            failCount += batchResult.Item2
        Next

        Return New Tuple(Of Integer, Integer)(successCount, failCount)
    End Function

    ''' <summary>
    ''' 處理單一批次
    ''' </summary>
    Private Async Function ProcessBatchAsync(docEntries As List(Of Integer)) As Task(Of Tuple(Of Integer, Integer))
        Dim successCount As Integer = 0
        Dim failCount As Integer = 0

        Try
            Dim boundary As String = "batch_" & Guid.NewGuid().ToString()
            Dim changesetBoundary As String = "changeset_" & Guid.NewGuid().ToString()

            Dim sb As New StringBuilder()
            sb.AppendLine("--" & boundary)
            sb.AppendLine("Content-Type: multipart/mixed; boundary=" & changesetBoundary)
            sb.AppendLine()

            Dim contentId As Integer = 1
            For Each docEntry As Integer In docEntries
                ' 讀取並建立 Invoice JSON
                Dim invoiceJson As String = BuildInvoiceJson(docEntry)
                If String.IsNullOrEmpty(invoiceJson) Then
                    failCount += 1
                    Continue For
                End If

                sb.AppendLine("--" & changesetBoundary)
                sb.AppendLine("Content-Type: application/http")
                sb.AppendLine("Content-Transfer-Encoding: binary")
                sb.AppendLine("Content-ID: " & contentId)
                sb.AppendLine()
                sb.AppendLine("POST /b1s/v1/PurchaseInvoices HTTP/1.1")
                sb.AppendLine("Content-Type: application/json")
                sb.AppendLine()
                sb.AppendLine(invoiceJson)

                contentId += 1
            Next

            sb.AppendLine("--" & changesetBoundary & "--")
            sb.AppendLine("--" & boundary & "--")

            ' 發送 Batch Request
            Dim url As String = baseUrl & "/$batch"
            Dim content As New StringContent(sb.ToString(), Encoding.UTF8, "multipart/mixed")
            content.Headers.Remove("Content-Type")
            content.Headers.Add("Content-Type", "multipart/mixed; boundary=" & boundary)

            Dim response As HttpResponseMessage = Await httpClient.PostAsync(url, content)
            Dim responseContent As String = Await response.Content.ReadAsStringAsync()

            ' 解析 Batch 回應 (簡化版，實際需要更完整的解析)
            If response.IsSuccessStatusCode Then
                ' 計算成功與失敗數 (簡化處理)
                Dim successMatches = System.Text.RegularExpressions.Regex.Matches(responseContent, """DocEntry"":\s*\d+")
                successCount = successMatches.Count
                failCount = docEntries.Count - successCount
            Else
                failCount = docEntries.Count
            End If

        Catch ex As Exception
            failCount = docEntries.Count
        End Try

        Return New Tuple(Of Integer, Integer)(successCount, failCount)
    End Function

    ''' <summary>
    ''' 建立 Invoice JSON 字串
    ''' </summary>
    Private Function BuildInvoiceJson(docEntry As Integer) As String
        Try
            Dim invoice As New JObject()

            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 讀取 Header
                Dim sqlH As String = "SELECT * FROM jOPCH WHERE DocEntry = @DocEntry"
                Using cmdH As New SqlCommand(sqlH, conn)
                    cmdH.Parameters.AddWithValue("@DocEntry", docEntry)
                    Using drH As SqlDataReader = cmdH.ExecuteReader()
                        If Not drH.Read() Then Return Nothing

                        invoice("CardCode") = drH("CardCode").ToString()
                        invoice("DocDate") = Convert.ToDateTime(drH("DocDate")).ToString("yyyy-MM-dd")

                        If Not IsDBNull(drH("DocDueDate")) Then
                            invoice("DocDueDate") = Convert.ToDateTime(drH("DocDueDate")).ToString("yyyy-MM-dd")
                        End If

                        If Not IsDBNull(drH("NumAtCard")) Then
                            invoice("NumAtCard") = drH("NumAtCard").ToString()
                        End If

                        If Not IsDBNull(drH("DocCurrency")) Then
                            invoice("DocCurrency") = drH("DocCurrency").ToString()
                        End If

                        drH.Close()
                    End Using
                End Using

                ' 讀取 Lines
                Dim lines As New JArray()
                Dim sqlL As String = "SELECT * FROM jPCH1 WHERE DocEntry = @DocEntry ORDER BY LineNum"
                Using cmdL As New SqlCommand(sqlL, conn)
                    cmdL.Parameters.AddWithValue("@DocEntry", docEntry)
                    Using drL As SqlDataReader = cmdL.ExecuteReader()
                        While drL.Read()
                            Dim line As New JObject()
                            line("AccountCode") = drL("AcctCode").ToString()

                            If Not IsDBNull(drL("LineTotal")) Then
                                line("LineTotal") = Convert.ToDouble(drL("LineTotal"))
                            End If

                            If Not IsDBNull(drL("Dscription")) Then
                                line("ItemDescription") = drL("Dscription").ToString()
                            End If

                            lines.Add(line)
                        End While
                        drL.Close()
                    End Using
                End Using

                invoice("DocumentLines") = lines
            End Using

            Return invoice.ToString(Formatting.None)

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

#End Region

#Region "靜態方法 (便於頁面呼叫)"

    ''' <summary>
    ''' 匯入單一 AP 發票 (靜態方法，自動處理連線)
    ''' </summary>
    ''' <param name="docEntry">jOPCH 的 DocEntry</param>
    ''' <returns>成功回傳訊息，失敗回傳錯誤訊息</returns>
    Public Shared Async Function ImportAsync(docEntry As Integer) As Task(Of String)
        Dim importer As New SapAPInvoiceImporterSL()

        Try
            Await importer.LoginAsync()
            Dim sapDocEntry As Integer = Await importer.ImportAPInvoiceAsync(docEntry)
            Return "匯入成功! SAP 單號: " & sapDocEntry

        Catch ex As Exception
            Return "匯入失敗: " & ex.Message
        Finally
            Await importer.LogoutAsync()
        End Try
    End Function

    ''' <summary>
    ''' 批次匯入所有待過帳文件 (靜態方法)
    ''' </summary>
    ''' <returns>匯入結果訊息</returns>
    Public Shared Async Function ImportAllPendingAsync() As Task(Of String)
        Dim importer As New SapAPInvoiceImporterSL()
        Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

        Try
            ' 取得所有待過帳的文件
            Dim pendingList As New List(Of Integer)
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT DocEntry FROM jOPCH " &
                                   "WHERE ApprovalStatus = 'A' AND (B1PostStatus IS NULL OR B1PostStatus <> 'Y')"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            pendingList.Add(Convert.ToInt32(dr("DocEntry")))
                        End While
                    End Using
                End Using
            End Using

            If pendingList.Count = 0 Then
                Return "沒有待過帳的文件"
            End If

            Await importer.LoginAsync()
            Dim result = Await importer.ImportBatchAsync(pendingList)
            Return String.Format("批次匯入完成! 成功: {0} 筆, 失敗: {1} 筆", result.Item1, result.Item2)

        Catch ex As Exception
            Return "批次匯入失敗: " & ex.Message
        Finally
            Await importer.LogoutAsync()
        End Try
    End Function

#End Region

End Class


'=============================================================================
' Web.config 設定範例 (Service Layer)
'=============================================================================
'
' 在 <appSettings> 區段加入以下設定:
'
'   <!-- SAP Service Layer 連線設定 -->
'   <add key="SL_BaseUrl" value="https://192.168.1.219:50001/b1s/v1"/>
'   <add key="SL_CompanyDB" value="JTTST"/>
'   <add key="SL_UserName" value="B1i"/>
'   <add key="SL_Password" value="5587"/>
'
' 注意：
' - HTTP Port 預設為 50000
' - HTTPS Port 預設為 50001
' - 建議使用 HTTPS
'
'=============================================================================
' 使用範例 (VB.NET)
'=============================================================================
'
' 方法一：單筆匯入 (非同步)
'   Dim result As String = Await SapAPInvoiceImporterSL.ImportAsync(docEntry:=123)
'   lblMessage.Text = result
'
' 方法二：批次匯入 (非同步)
'   Dim result As String = Await SapAPInvoiceImporterSL.ImportAllPendingAsync()
'   lblMessage.Text = result
'
' 方法三：同步呼叫 (在非 Async 方法中)
'   Dim result As String = SapAPInvoiceImporterSL.ImportAsync(123).GetAwaiter().GetResult()
'
'=============================================================================
' Service Layer 特有功能：Batch Request
'=============================================================================
'
' Service Layer 支援 OData Batch Request，可以將多個操作合併成一個 HTTP 請求：
'
' 優點：
' 1. 減少網路往返次數
' 2. 整個批次可以在同一個交易中執行
' 3. 提高整體效能
'
' 限制：
' 1. 單一批次最多 100 個操作
' 2. 批次中的操作必須是相互獨立的
'
'=============================================================================
' DI API vs Service Layer 程式碼對照
'=============================================================================
'
' ┌─────────────────────────────────────────────────────────────────────────┐
' │ 功能                │ DI API                      │ Service Layer       │
' ├─────────────────────┼─────────────────────────────┼─────────────────────┤
' │ 連線               │ oCompany.Connect()          │ POST /Login         │
' │ 斷線               │ oCompany.Disconnect()       │ POST /Logout        │
' │ 建立物件           │ oCompany.GetBusinessObject  │ POST /物件名稱      │
' │ 設定欄位           │ oInvoice.CardCode = "xxx"   │ JSON: CardCode="xxx"│
' │ 新增資料           │ oInvoice.Add()              │ POST + JSON Body    │
' │ 取得新 ID          │ oCompany.GetNewObjectKey()  │ 回應 JSON 中的 ID   │
' │ 錯誤處理           │ oCompany.GetLastError()     │ HTTP Status + JSON  │
' └─────────────────────┴─────────────────────────────┴─────────────────────┘
'
'=============================================================================
