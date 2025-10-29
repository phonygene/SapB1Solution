Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

''' <summary>
''' B1Transaction 抽象基底類別
''' 用途：模仿 SAP B1 操作介面的交易表單基礎功能
''' 版本：1.1.0
''' 作者：Claude
''' 日期：2025-10-29
''' </summary>
Public MustInherit Class B1TransactionBase
    Inherits System.Web.UI.Page

#Region "屬性與欄位"

    ''' <summary>
    ''' JSON 配置物件
    ''' </summary>
    Protected Config As JObject

    ''' <summary>
    ''' 當前操作模式：create, search, update
    ''' </summary>
    Protected CurrentMode As String = "create"

    ''' <summary>
    ''' 配置檔路徑
    ''' </summary>
    Protected ConfigFilePath As String

    ''' <summary>
    ''' 本機資料庫連線
    ''' </summary>
    Protected conn As SqlConnection

    ''' <summary>
    ''' SAP 資料庫連線
    ''' </summary>
    Protected connSap As SqlConnection

#End Region

#Region "建構子與初始化"

    ''' <summary>
    ''' 建構子
    ''' </summary>
    Protected Sub New()
        conn = New SqlConnection()
        connSap = New SqlConnection()
    End Sub

    ''' <summary>
    ''' 載入 JSON 配置檔
    ''' </summary>
    ''' <param name="configPath">配置檔完整路徑</param>
    Protected Sub LoadConfig(configPath As String)
        Try
            ConfigFilePath = configPath
            Dim jsonText As String = File.ReadAllText(configPath)
            Config = JObject.Parse(jsonText)
        Catch ex As Exception
            Throw New ApplicationException($"載入配置檔失敗: {configPath}" & vbCrLf & ex.Message, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初始化資料庫連線
    ''' </summary>
    Protected Sub InitConnections()
        ' 初始化本機連線
        conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

        ' 初始化 SAP 連線
        connSap.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
    End Sub

#End Region

#Region "模式管理"

    ''' <summary>
    ''' 切換操作模式
    ''' </summary>
    ''' <param name="mode">模式：create, search, update</param>
    Public Sub SwitchMode(mode As String)
        If mode <> "create" AndAlso mode <> "search" AndAlso mode <> "update" Then
            Throw New ArgumentException($"不支援的模式: {mode}")
        End If

        CurrentMode = mode
        ApplyModePolicy(mode)
    End Sub

    ''' <summary>
    ''' 套用模式策略
    ''' </summary>
    ''' <param name="mode">模式名稱</param>
    Protected Sub ApplyModePolicy(mode As String)
        Try
            Dim modePolicies As JObject = CType(Config("modePolicies"), JObject)
            If modePolicies Is Nothing OrElse modePolicies(mode) Is Nothing Then
                Return
            End If

            Dim policy As JObject = CType(modePolicies(mode), JObject)
            Dim fieldDefaults As JObject = CType(policy("fieldDefaults"), JObject)
            Dim stateRules As JArray = CType(policy("stateRules"), JArray)

            ' 套用欄位預設值
            If fieldDefaults IsNot Nothing Then
                ApplyFieldDefaults(fieldDefaults)
            End If

            ' 套用狀態規則
            If stateRules IsNot Nothing Then
                ApplyStateRules(stateRules)
            End If

        Catch ex As Exception
            Throw New ApplicationException($"套用模式策略失敗: {mode}" & vbCrLf & ex.Message, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 套用欄位預設值（由子類別實作）
    ''' </summary>
    ''' <param name="defaults">預設值設定</param>
    Protected MustOverride Sub ApplyFieldDefaults(defaults As JObject)

    ''' <summary>
    ''' 套用狀態規則（由子類別實作）
    ''' </summary>
    ''' <param name="rules">規則陣列</param>
    Protected MustOverride Sub ApplyStateRules(rules As JArray)

#End Region

#Region "欄位屬性查詢"

    ''' <summary>
    ''' 取得表頭欄位屬性
    ''' </summary>
    ''' <param name="fieldName">欄位名稱</param>
    ''' <returns>欄位屬性 JObject，若不存在則傳回 Nothing</returns>
    Protected Function GetMasterFieldAttr(fieldName As String) As JObject
        Try
            Dim master As JObject = CType(Config("ui")("layout")("master")("fields")(fieldName), JObject)
            Return master
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 取得明細欄位屬性
    ''' </summary>
    ''' <param name="detailName">明細名稱（如 invoiceDetail）</param>
    ''' <param name="fieldName">欄位名稱</param>
    ''' <returns>欄位屬性 JObject，若不存在則傳回 Nothing</returns>
    Protected Function GetDetailFieldAttr(detailName As String, fieldName As String) As JObject
        Try
            Dim details As JArray = CType(Config("ui")("layout")("details"), JArray)
            For Each detail As JObject In details
                If detail("name").ToString() = detailName Then
                    Dim field As JObject = CType(detail("fields")(fieldName), JObject)
                    Return field
                End If
            Next
            Return Nothing
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 檢查欄位是否必填
    ''' </summary>
    ''' <param name="fieldAttr">欄位屬性</param>
    ''' <returns>True 表示必填</returns>
    Protected Function IsRequired(fieldAttr As JObject) As Boolean
        If fieldAttr Is Nothing Then Return False
        Dim required As JToken = fieldAttr("required")
        If required Is Nothing Then Return False
        Return CBool(required)
    End Function

    ''' <summary>
    ''' 檢查欄位是否可編輯
    ''' </summary>
    ''' <param name="fieldAttr">欄位屬性</param>
    ''' <returns>True 表示可編輯</returns>
    Protected Function IsEditable(fieldAttr As JObject) As Boolean
        If fieldAttr Is Nothing Then Return False
        Dim editable As JToken = fieldAttr("editable")
        If editable Is Nothing Then Return False
        Return CBool(editable)
    End Function

    ''' <summary>
    ''' 取得欄位預設值
    ''' </summary>
    ''' <param name="fieldAttr">欄位屬性</param>
    ''' <returns>預設值字串，若無則傳回空字串</returns>
    Protected Function GetDefaultValue(fieldAttr As JObject) As String
        If fieldAttr Is Nothing Then Return String.Empty
        Dim defaultValue As JToken = fieldAttr("defaultValue")
        If defaultValue Is Nothing Then Return String.Empty
        Return defaultValue.ToString()
    End Function

#End Region

#Region "欄位驗證"

    ''' <summary>
    ''' 驗證欄位值
    ''' </summary>
    ''' <param name="fieldAttr">欄位屬性</param>
    ''' <param name="value">欄位值</param>
    ''' <param name="errorMessage">錯誤訊息（輸出參數）</param>
    ''' <returns>True 表示驗證通過</returns>
    Protected Function ValidateField(fieldAttr As JObject, value As String, ByRef errorMessage As String) As Boolean
        If fieldAttr Is Nothing Then
            errorMessage = "欄位屬性不存在"
            Return False
        End If

        ' 檢查必填
        If IsRequired(fieldAttr) AndAlso String.IsNullOrWhiteSpace(value) Then
            Dim label As String = fieldAttr("label").ToString()
            errorMessage = $"{label} 為必填欄位"
            Return False
        End If

        ' 檢查正規表達式驗證
        Dim validation As JObject = CType(fieldAttr("validation"), JObject)
        If validation IsNot Nothing Then
            Dim pattern As JToken = validation("pattern")
            If pattern IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(value) Then
                Dim regex As New System.Text.RegularExpressions.Regex(pattern.ToString())
                If Not regex.IsMatch(value) Then
                    errorMessage = validation("message").ToString()
                    Return False
                End If
            End If

            ' 檢查最小值
            Dim minValue As JToken = validation("min")
            If minValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(value) Then
                Dim numValue As Decimal
                If Decimal.TryParse(value, numValue) Then
                    If numValue < CDec(minValue) Then
                        errorMessage = validation("message").ToString()
                        Return False
                    End If
                End If
            End If
        End If

        errorMessage = String.Empty
        Return True
    End Function

    ''' <summary>
    ''' 驗證所有表頭欄位
    ''' </summary>
    ''' <returns>驗證結果清單（欄位名稱 → 錯誤訊息）</returns>
    Protected Function ValidateMasterFields() As Dictionary(Of String, String)
        Dim errors As New Dictionary(Of String, String)()

        Try
            Dim masterFields As JObject = CType(Config("ui")("layout")("master")("fields"), JObject)
            If masterFields Is Nothing Then Return errors

            For Each prop As JProperty In masterFields.Properties()
                Dim fieldName As String = prop.Name
                Dim fieldAttr As JObject = CType(prop.Value, JObject)

                ' 取得欄位值（由子類別提供）
                Dim value As String = GetMasterFieldValue(fieldName)

                ' 驗證
                Dim errorMsg As String = String.Empty
                If Not ValidateField(fieldAttr, value, errorMsg) Then
                    errors(fieldName) = errorMsg
                End If
            Next

        Catch ex As Exception
            errors("_error") = "驗證過程發生錯誤: " & ex.Message
        End Try

        Return errors
    End Function

    ''' <summary>
    ''' 取得表頭欄位值（由子類別實作）
    ''' </summary>
    ''' <param name="fieldName">欄位名稱</param>
    ''' <returns>欄位值</returns>
    Protected MustOverride Function GetMasterFieldValue(fieldName As String) As String

#End Region

#Region "計算欄位"

    ''' <summary>
    ''' 計算表頭總金額
    ''' </summary>
    ''' <param name="details">明細資料表</param>
    ''' <returns>總金額</returns>
    Protected Function CalculateDocTotal(details As DataTable) As Decimal
        Dim total As Decimal = 0

        If details Is Nothing OrElse details.Rows.Count = 0 Then
            Return total
        End If

        For Each row As DataRow In details.Rows
            Dim hwbas As Decimal = If(IsDBNull(row("U_HWBAS")), 0D, CDec(row("U_HWBAS")))
            Dim hwste As Decimal = If(IsDBNull(row("U_HWSTE")), 0D, CDec(row("U_HWSTE")))
            total += hwbas + hwste
        Next

        Return total
    End Function

    ''' <summary>
    ''' 計算表頭稅額總金額
    ''' </summary>
    ''' <param name="details">明細資料表</param>
    ''' <returns>稅額總金額</returns>
    Protected Function CalculateVatSum(details As DataTable) As Decimal
        Dim total As Decimal = 0

        If details Is Nothing OrElse details.Rows.Count = 0 Then
            Return total
        End If

        For Each row As DataRow In details.Rows
            Dim hwste As Decimal = If(IsDBNull(row("U_HWSTE")), 0D, CDec(row("U_HWSTE")))
            total += hwste
        Next

        Return total
    End Function

    ''' <summary>
    ''' 計算明細稅額（根據稅別自動計算）
    ''' </summary>
    ''' <param name="hwbas">未稅金額</param>
    ''' <param name="taxType">稅別（1-應稅 2-零稅 3-免稅）</param>
    ''' <returns>稅額</returns>
    Protected Function CalculateTax(hwbas As Decimal, taxType As String) As Decimal
        Select Case taxType
            Case "1" ' 應稅（5%）
                Return Math.Round(hwbas * 0.05D, 2)
            Case "2", "3" ' 零稅、免稅
                Return 0D
            Case Else
                Return 0D
        End Select
    End Function

#End Region

#Region "資料庫操作"

    ''' <summary>
    ''' 執行本機資料庫查詢
    ''' </summary>
    ''' <param name="sql">SQL 查詢語句</param>
    ''' <param name="parameters">參數陣列</param>
    ''' <returns>DataTable</returns>
    Protected Function ExecuteLocalQuery(sql As String, ParamArray parameters As SqlParameter()) As DataTable
        Dim dt As New DataTable()

        Try
            If conn.State <> ConnectionState.Open Then
                conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
                conn.Open()
            End If

            Using cmd As New SqlCommand(sql, conn)
                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters)
                End If

                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using

        Catch ex As Exception
            Throw New ApplicationException("執行本機查詢失敗: " & ex.Message, ex)
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        Return dt
    End Function

    ''' <summary>
    ''' 執行本機資料庫非查詢命令（INSERT, UPDATE, DELETE）
    ''' </summary>
    ''' <param name="sql">SQL 命令</param>
    ''' <param name="parameters">參數陣列</param>
    ''' <returns>受影響的列數</returns>
    Protected Function ExecuteLocalNonQuery(sql As String, ParamArray parameters As SqlParameter()) As Integer
        Dim rowsAffected As Integer = 0

        Try
            If conn.State <> ConnectionState.Open Then
                conn.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
                conn.Open()
            End If

            Using cmd As New SqlCommand(sql, conn)
                If parameters IsNot Nothing Then
                    cmd.Parameters.AddRange(parameters)
                End If

                rowsAffected = cmd.ExecuteNonQuery()
            End Using

        Catch ex As Exception
            Throw New ApplicationException("執行本機命令失敗: " & ex.Message, ex)
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

        Return rowsAffected
    End Function

#End Region

#Region "輔助方法"

    ''' <summary>
    ''' 取得當前使用者 ID（從 Session）
    ''' </summary>
    ''' <returns>使用者 ID</returns>
    Protected Function GetCurrentUserId() As String
        If Session("s_id") IsNot Nothing Then
            Return Session("s_id").ToString()
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' 顯示訊息
    ''' </summary>
    ''' <param name="message">訊息內容</param>
    Protected Sub ShowMessage(message As String)
        Dim sMessage As String = message.Replace("'", "\'").Replace(vbNewLine, "\n")
        Dim sScript As String = String.Format("alert('{0}');", sMessage)
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alert", sScript, True)
    End Sub

    ''' <summary>
    ''' 顯示確認對話框
    ''' </summary>
    ''' <param name="message">訊息內容</param>
    Protected Sub ShowConfirm(message As String)
        Dim sMessage As String = message.Replace("'", "\'").Replace(vbNewLine, "\n")
        Dim sScript As String = String.Format("confirm('{0}');", sMessage)
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "confirm", sScript, True)
    End Sub

#End Region

End Class
