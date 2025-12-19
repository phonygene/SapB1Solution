'=============================================================================
' SAP B1 Service Layer - Excel VBA 範例
'
' 用途：透過 Excel VBA 呼叫 SAP B1 Service Layer 建立 AP 發票
' 日期：2025/12/15
' 版本：1.0
'
' 優點：
' - 不需安裝 SAP B1 Client
' - 不需 DI API COM 元件
' - 可在任何有網路連線的電腦使用
' - 適合 End User 自行操作
'
' 需求：
' - Excel 2010 以上版本
' - 參考設定 (VBA 編輯器 -> 工具 -> 設定引用項目):
'   [x] Microsoft Scripting Runtime (字典物件)
'   [x] Microsoft XML, v6.0 (XMLHTTP) - 選用，也可用 Late Binding
'
' 注意事項：
' - Service Layer 預設 Port: 50000 (HTTP) 或 50001 (HTTPS)
' - Session 預設 30 分鐘逾時
' - 建議使用 HTTPS 確保安全性
'=============================================================================

Option Explicit

' ===== 全域變數 =====
Private g_SessionId As String
Private g_BaseUrl As String

' ===== 設定區 (請依實際環境修改) =====
Private Const SL_SERVER As String = "192.168.1.219"
Private Const SL_PORT As String = "50001"           ' 50000=HTTP, 50001=HTTPS
Private Const SL_USE_HTTPS As Boolean = True
Private Const SL_COMPANY As String = "JTTST"
Private Const SL_USER As String = "B1i"
Private Const SL_PASSWORD As String = "5587"

'=============================================================================
' 主要功能函數
'=============================================================================

''' <summary>
''' 登入 Service Layer
''' </summary>
''' <returns>True=成功, False=失敗</returns>
Public Function SL_Login() As Boolean
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim body As String
    Dim response As String

    ' 建立基礎 URL
    If SL_USE_HTTPS Then
        g_BaseUrl = "https://" & SL_SERVER & ":" & SL_PORT & "/b1s/v1"
    Else
        g_BaseUrl = "http://" & SL_SERVER & ":" & SL_PORT & "/b1s/v1"
    End If

    ' 建立 HTTP 物件
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' 忽略 SSL 憑證錯誤 (測試用)
    http.setOption 2, 13056  ' SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

    ' 登入請求
    url = g_BaseUrl & "/Login"
    body = "{""CompanyDB"":""" & SL_COMPANY & """,""UserName"":""" & SL_USER & """,""Password"":""" & SL_PASSWORD & """}"

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.send body

    ' 檢查回應
    If http.Status = 200 Then
        ' 從 Set-Cookie 取得 Session ID
        Dim cookies As String
        cookies = http.getResponseHeader("Set-Cookie")

        If InStr(cookies, "B1SESSION=") > 0 Then
            g_SessionId = ExtractSessionId(cookies)
            SL_Login = True
            Debug.Print "登入成功! Session: " & g_SessionId
        Else
            Debug.Print "無法取得 Session ID"
            SL_Login = False
        End If
    Else
        Debug.Print "登入失敗: " & http.Status & " - " & http.responseText
        SL_Login = False
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "登入錯誤: " & Err.Description
    SL_Login = False
End Function

''' <summary>
''' 登出 Service Layer
''' </summary>
Public Sub SL_Logout()
    On Error Resume Next

    If Len(g_SessionId) = 0 Then Exit Sub

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setOption 2, 13056

    http.Open "POST", g_BaseUrl & "/Logout", False
    http.setRequestHeader "Cookie", "B1SESSION=" & g_SessionId
    http.send

    g_SessionId = ""
    Debug.Print "已登出"

    Set http = Nothing
End Sub

''' <summary>
''' 建立 AP 發票 (費用類)
''' </summary>
''' <param name="cardCode">供應商代碼</param>
''' <param name="docDate">文件日期 (yyyy-mm-dd)</param>
''' <param name="numAtCard">發票號碼</param>
''' <param name="lines">明細陣列 (二維陣列: 科目, 金額, 說明, 稅碼)</param>
''' <returns>成功回傳 DocEntry，失敗回傳 -1</returns>
Public Function SL_CreateAPInvoice(ByVal cardCode As String, _
                                   ByVal docDate As String, _
                                   ByVal numAtCard As String, _
                                   ByRef lines As Variant) As Long
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String
    Dim body As String
    Dim response As String
    Dim i As Long

    ' 檢查登入狀態
    If Len(g_SessionId) = 0 Then
        Debug.Print "請先登入!"
        SL_CreateAPInvoice = -1
        Exit Function
    End If

    ' 建立 JSON 主體
    body = "{"
    body = body & """CardCode"":""" & cardCode & ""","
    body = body & """DocDate"":""" & docDate & ""","
    body = body & """DocDueDate"":""" & docDate & ""","

    If Len(numAtCard) > 0 Then
        body = body & """NumAtCard"":""" & numAtCard & ""","
    End If

    ' 明細行
    body = body & """DocumentLines"":["

    For i = LBound(lines, 1) To UBound(lines, 1)
        If i > LBound(lines, 1) Then body = body & ","

        body = body & "{"
        body = body & """AccountCode"":""" & lines(i, 1) & ""","  ' 科目
        body = body & """LineTotal"":" & lines(i, 2) & ","         ' 金額
        body = body & """ItemDescription"":""" & lines(i, 3) & """" ' 說明

        ' 稅碼 (選用)
        If UBound(lines, 2) >= 4 Then
            If Len(lines(i, 4)) > 0 Then
                body = body & ",""VatGroup"":""" & lines(i, 4) & """"
            End If
        End If

        body = body & "}"
    Next i

    body = body & "]}"

    ' 發送請求
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setOption 2, 13056

    url = g_BaseUrl & "/PurchaseInvoices"

    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Cookie", "B1SESSION=" & g_SessionId
    http.send body

    ' 處理回應
    If http.Status = 201 Then
        ' 成功，解析 DocEntry
        response = http.responseText
        SL_CreateAPInvoice = ExtractDocEntry(response)
        Debug.Print "建立成功! DocEntry: " & SL_CreateAPInvoice
    Else
        ' 失敗
        Debug.Print "建立失敗: " & http.Status & " - " & http.responseText
        SL_CreateAPInvoice = -1
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "錯誤: " & Err.Description
    SL_CreateAPInvoice = -1
End Function

''' <summary>
''' 查詢 AP 發票
''' </summary>
''' <param name="docEntry">文件編號</param>
''' <returns>JSON 字串</returns>
Public Function SL_GetAPInvoice(ByVal docEntry As Long) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String

    If Len(g_SessionId) = 0 Then
        SL_GetAPInvoice = "Error: 請先登入"
        Exit Function
    End If

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setOption 2, 13056

    url = g_BaseUrl & "/PurchaseInvoices(" & docEntry & ")"

    http.Open "GET", url, False
    http.setRequestHeader "Cookie", "B1SESSION=" & g_SessionId
    http.send

    If http.Status = 200 Then
        SL_GetAPInvoice = http.responseText
    Else
        SL_GetAPInvoice = "Error: " & http.Status & " - " & http.responseText
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    SL_GetAPInvoice = "Error: " & Err.Description
End Function

''' <summary>
''' 查詢供應商清單
''' </summary>
''' <param name="top">筆數限制</param>
''' <returns>JSON 字串</returns>
Public Function SL_GetVendors(Optional ByVal top As Long = 20) As String
    On Error GoTo ErrorHandler

    Dim http As Object
    Dim url As String

    If Len(g_SessionId) = 0 Then
        SL_GetVendors = "Error: 請先登入"
        Exit Function
    End If

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setOption 2, 13056

    ' CardType='cSupplier' 表示供應商
    url = g_BaseUrl & "/BusinessPartners?$filter=CardType eq 'cSupplier'&$top=" & top & "&$select=CardCode,CardName"

    http.Open "GET", url, False
    http.setRequestHeader "Cookie", "B1SESSION=" & g_SessionId
    http.send

    If http.Status = 200 Then
        SL_GetVendors = http.responseText
    Else
        SL_GetVendors = "Error: " & http.Status & " - " & http.responseText
    End If

    Set http = Nothing
    Exit Function

ErrorHandler:
    SL_GetVendors = "Error: " & Err.Description
End Function

'=============================================================================
' 輔助函數
'=============================================================================

''' <summary>
''' 從 Set-Cookie 標頭擷取 Session ID
''' </summary>
Private Function ExtractSessionId(ByVal cookies As String) As String
    Dim startPos As Long
    Dim endPos As Long

    startPos = InStr(cookies, "B1SESSION=")
    If startPos > 0 Then
        startPos = startPos + 10
        endPos = InStr(startPos, cookies, ";")
        If endPos > 0 Then
            ExtractSessionId = Mid(cookies, startPos, endPos - startPos)
        Else
            ExtractSessionId = Mid(cookies, startPos)
        End If
    Else
        ExtractSessionId = ""
    End If
End Function

''' <summary>
''' 從 JSON 回應擷取 DocEntry
''' </summary>
Private Function ExtractDocEntry(ByVal json As String) As Long
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String

    startPos = InStr(json, """DocEntry"":")
    If startPos > 0 Then
        startPos = startPos + 11
        endPos = InStr(startPos, json, ",")
        If endPos = 0 Then endPos = InStr(startPos, json, "}")
        value = Trim(Mid(json, startPos, endPos - startPos))
        ExtractDocEntry = CLng(value)
    Else
        ExtractDocEntry = -1
    End If
End Function

'=============================================================================
' 測試程序
'=============================================================================

''' <summary>
''' 測試登入
''' </summary>
Public Sub TestLogin()
    If SL_Login() Then
        MsgBox "登入成功! Session: " & g_SessionId, vbInformation
    Else
        MsgBox "登入失敗!", vbCritical
    End If
End Sub

''' <summary>
''' 測試建立 AP 發票
''' </summary>
Public Sub TestCreateAPInvoice()
    Dim lines(1 To 2, 1 To 4) As Variant
    Dim result As Long

    ' 先登入
    If Not SL_Login() Then
        MsgBox "登入失敗!", vbCritical
        Exit Sub
    End If

    ' 準備明細資料
    ' 格式: 科目, 金額, 說明, 稅碼
    lines(1, 1) = "6001001"     ' 費用科目
    lines(1, 2) = 1000          ' 金額
    lines(1, 3) = "測試費用1"   ' 說明
    lines(1, 4) = "TX"          ' 稅碼

    lines(2, 1) = "6001002"
    lines(2, 2) = 500
    lines(2, 3) = "測試費用2"
    lines(2, 4) = "TX"

    ' 建立 AP 發票
    result = SL_CreateAPInvoice("V10001", Format(Date, "yyyy-mm-dd"), "TEST-001", lines)

    If result > 0 Then
        MsgBox "建立成功! DocEntry: " & result, vbInformation
    Else
        MsgBox "建立失敗!", vbCritical
    End If

    ' 登出
    SL_Logout
End Sub

''' <summary>
''' 從 Excel 工作表建立 AP 發票
''' 假設工作表格式：
''' A1:供應商代碼  B1:日期  C1:發票號碼
''' A3起:科目 B3:金額 C3:說明 D3:稅碼
''' </summary>
Public Sub CreateAPInvoiceFromSheet()
    Dim ws As Worksheet
    Dim cardCode As String
    Dim docDate As String
    Dim numAtCard As String
    Dim lastRow As Long
    Dim lines() As Variant
    Dim i As Long, lineCount As Long
    Dim result As Long

    Set ws = ActiveSheet

    ' 讀取表頭資料
    cardCode = ws.Range("A1").Value
    docDate = Format(ws.Range("B1").Value, "yyyy-mm-dd")
    numAtCard = ws.Range("C1").Value

    ' 計算明細行數
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "沒有明細資料!", vbExclamation
        Exit Sub
    End If

    lineCount = lastRow - 2
    ReDim lines(1 To lineCount, 1 To 4)

    ' 讀取明細資料
    For i = 1 To lineCount
        lines(i, 1) = ws.Cells(i + 2, 1).Value  ' 科目
        lines(i, 2) = ws.Cells(i + 2, 2).Value  ' 金額
        lines(i, 3) = ws.Cells(i + 2, 3).Value  ' 說明
        lines(i, 4) = ws.Cells(i + 2, 4).Value  ' 稅碼
    Next i

    ' 登入並建立
    If Not SL_Login() Then
        MsgBox "登入失敗!", vbCritical
        Exit Sub
    End If

    result = SL_CreateAPInvoice(cardCode, docDate, numAtCard, lines)

    If result > 0 Then
        MsgBox "建立成功! DocEntry: " & result, vbInformation
        ws.Range("E1").Value = result  ' 將結果寫回工作表
    Else
        MsgBox "建立失敗!", vbCritical
    End If

    SL_Logout
End Sub

'=============================================================================
' Service Layer API 常用端點參考
'=============================================================================
'
' 登入/登出：
'   POST /Login           - 登入
'   POST /Logout          - 登出
'
' 業務夥伴：
'   GET  /BusinessPartners                    - 查詢清單
'   GET  /BusinessPartners('V10001')          - 查詢單筆
'   POST /BusinessPartners                    - 新增
'   PATCH /BusinessPartners('V10001')         - 更新
'
' AP 發票：
'   GET  /PurchaseInvoices                    - 查詢清單
'   GET  /PurchaseInvoices(123)               - 查詢單筆
'   POST /PurchaseInvoices                    - 新增
'   POST /PurchaseInvoices(123)/Cancel        - 取消
'
' AR 發票：
'   GET  /Invoices                            - 查詢清單
'   GET  /Invoices(123)                       - 查詢單筆
'   POST /Invoices                            - 新增
'
' 採購單：
'   GET  /PurchaseOrders
'   POST /PurchaseOrders
'
' 銷售單：
'   GET  /Orders
'   POST /Orders
'
' 會計科目：
'   GET  /ChartOfAccounts
'
' OData 查詢參數：
'   $filter   - 篩選 (例如: $filter=CardCode eq 'V10001')
'   $select   - 選擇欄位 (例如: $select=CardCode,CardName)
'   $top      - 取前 N 筆 (例如: $top=10)
'   $skip     - 跳過 N 筆 (例如: $skip=10)
'   $orderby  - 排序 (例如: $orderby=DocEntry desc)
'
'=============================================================================
' 併發控制說明
'=============================================================================
'
' Service Layer 的併發處理機制：
'
' 1. Session 序列化
'    - 同一個 Session 內的請求會序列化執行
'    - 不同 Session 可以並行
'    - 建議每個使用者使用獨立 Session
'
' 2. 樂觀鎖 (Optimistic Locking)
'    - 更新時會檢查版本號
'    - 如果資料已被其他人修改，會回傳 412 Precondition Failed
'
' 3. 沒有寫入佇列
'    - Service Layer 仍是即時寫入資料庫
'    - 如需佇列功能，需自行在應用層實作
'
' 防呆建議：
' - 每次操作後檢查回傳狀態
' - 實作重試機制 (Retry with Exponential Backoff)
' - 記錄所有操作日誌
' - 考慮使用 Batch Request 減少併發問題
'
'=============================================================================
