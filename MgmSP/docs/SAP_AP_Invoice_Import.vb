'=============================================================================
' SAP B1 AP 發票 (A/P Invoice) 匯入功能 - DI API 範例代碼
'
' 用途：將 jOPCH/jPCH1 資料匯入 SAP Business One 成為 AP 發票
' 日期：2025/12/08
' 版本：1.2 - 連線資訊完全集中於 Web.config，使用平台統一帳號 B1i
'
' 注意事項：
' 1. 需要安裝 SAP Business One DI API
' 2. 所有連線資訊統一在 Web.config 的 appSettings 區段設定
' 3. 平台統一使用 B1i 帳號，節省 License
' 4. 此為範例代碼，請檢查後依實際需求修改
'=============================================================================

Imports SAPbobsCOM
Imports System.Data.SqlClient
Imports System.Configuration

''' <summary>
''' SAP AP 發票匯入類別
''' </summary>
Public Class SapAPInvoiceImporter

    Private oCompany As SAPbobsCOM.Company
    Private connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    ' 本幣代碼 (可考慮也放到 Web.config)
    Private Const LOCAL_CURRENCY As String = "TWD"

#Region "Web.config 設定讀取"

    ''' <summary>
    ''' 從 Web.config 讀取 SAP 連線設定
    ''' 所有連線資訊集中在此，代碼中不再硬編碼任何連線資訊
    ''' </summary>
    Private Class SapConfig
        ''' <summary>SAP 資料庫伺服器 IP</summary>
        Public Shared ReadOnly Property ServerIP As String
            Get
                Return ConfigurationManager.AppSettings("SAP_ServerIP")
            End Get
        End Property

        ''' <summary>SAP 公司資料庫名稱</summary>
        Public Shared ReadOnly Property CompanyDB As String
            Get
                Return ConfigurationManager.AppSettings("SAP_CompanyDB")
            End Get
        End Property

        ''' <summary>SQL Server 帳號</summary>
        Public Shared ReadOnly Property DbUser As String
            Get
                Return ConfigurationManager.AppSettings("SAP_DbUser")
            End Get
        End Property

        ''' <summary>SQL Server 密碼</summary>
        Public Shared ReadOnly Property DbPassword As String
            Get
                Return ConfigurationManager.AppSettings("SAP_DbPassword")
            End Get
        End Property

        ''' <summary>SAP DI API 登入帳號 (平台統一使用 B1i)</summary>
        Public Shared ReadOnly Property SapUser As String
            Get
                Return ConfigurationManager.AppSettings("SAP_User")
            End Get
        End Property

        ''' <summary>SAP DI API 登入密碼</summary>
        Public Shared ReadOnly Property SapPassword As String
            Get
                Return ConfigurationManager.AppSettings("SAP_Password")
            End Get
        End Property

        ''' <summary>SQL Server 版本</summary>
        Public Shared ReadOnly Property DbServerType As SAPbobsCOM.BoDataServerTypes
            Get
                Dim serverType As String = ConfigurationManager.AppSettings("SAP_DbServerType")
                Select Case serverType
                    Case "dst_MSSQL2005"
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
                    Case "dst_MSSQL2008"
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                    Case "dst_MSSQL2012"
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                    Case "dst_MSSQL2014"
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                    Case "dst_MSSQL2016"
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
                    Case "dst_MSSQL2017"
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017
                    Case "dst_MSSQL2019"
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019
                    Case Else
                        Return SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
                End Select
            End Get
        End Property
    End Class

#End Region

#Region "SAP 連線"

    ''' <summary>
    ''' 連線到 SAP (使用 Web.config 設定，平台統一帳號)
    ''' </summary>
    ''' <returns>0 = 成功, 其他 = 錯誤碼</returns>
    Public Function ConnectToSAP() As Integer
        Try
            oCompany = New SAPbobsCOM.Company()

            With oCompany
                .Server = SapConfig.ServerIP
                .CompanyDB = SapConfig.CompanyDB
                .UserName = SapConfig.SapUser
                .Password = SapConfig.SapPassword
                .DbUserName = SapConfig.DbUser
                .DbPassword = SapConfig.DbPassword
                .UseTrusted = False
                .language = SAPbobsCOM.BoSuppLangs.ln_Chinese     ' 繁體中文
                .DbServerType = SapConfig.DbServerType
            End With

            Dim retCode As Integer = oCompany.Connect()

            If retCode <> 0 Then
                Dim errMsg As String = oCompany.GetLastErrorDescription()
                Throw New Exception("SAP 連線失敗: " & errMsg)
            End If

            Return 0
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 斷開 SAP 連線
    ''' </summary>
    Public Sub DisconnectFromSAP()
        If oCompany IsNot Nothing AndAlso oCompany.Connected Then
            oCompany.Disconnect()
        End If
    End Sub

    ''' <summary>
    ''' 檢查 SAP 是否已連線
    ''' </summary>
    Public ReadOnly Property IsConnected As Boolean
        Get
            Return oCompany IsNot Nothing AndAlso oCompany.Connected
        End Get
    End Property

#End Region

#Region "AP 發票匯入"

    ''' <summary>
    ''' 匯入單一 AP 發票到 SAP
    ''' </summary>
    ''' <param name="docEntry">jOPCH 的 DocEntry</param>
    ''' <returns>成功回傳 SAP DocEntry，失敗拋出 Exception</returns>
    Public Function ImportAPInvoice(docEntry As Integer) As Integer
        Dim oInvoice As SAPbobsCOM.Documents = Nothing
        Dim sapDocEntry As Integer = -1

        Try
            ' 檢查 SAP 連線
            If Not IsConnected Then
                Throw New Exception("SAP 未連線")
            End If

            ' 建立 AP Invoice 物件
            oInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            ' 文件幣別和匯率 (用於明細計算)
            Dim docCurrency As String = LOCAL_CURRENCY
            Dim docRate As Double = 1.0
            Dim isForeignCurrency As Boolean = False

            ' 從 jtdb 讀取表頭資料
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

                        ' ===== 表頭欄位設定 =====
                        oInvoice.CardCode = drH("CardCode").ToString()
                        oInvoice.CardName = drH("CardName").ToString()

                        ' 供應商參考號 (發票號碼)
                        If Not IsDBNull(drH("NumAtCard")) Then
                            oInvoice.NumAtCard = drH("NumAtCard").ToString()
                        End If

                        ' 文件日期
                        oInvoice.DocDate = Convert.ToDateTime(drH("DocDate"))

                        ' 到期日
                        If Not IsDBNull(drH("DocDueDate")) Then
                            oInvoice.DocDueDate = Convert.ToDateTime(drH("DocDueDate"))
                        End If

                        ' 過帳日期 (Tax Date)
                        If Not IsDBNull(drH("TaxDate")) Then
                            oInvoice.TaxDate = Convert.ToDateTime(drH("TaxDate"))
                        End If

                        ' 幣別
                        If Not IsDBNull(drH("DocCurrency")) Then
                            docCurrency = drH("DocCurrency").ToString()
                            oInvoice.DocCurrency = docCurrency
                            isForeignCurrency = (docCurrency <> LOCAL_CURRENCY)
                        End If

                        ' 匯率 (若非本幣)
                        If Not IsDBNull(drH("DocRate")) Then
                            docRate = Convert.ToDouble(drH("DocRate"))
                            If isForeignCurrency Then
                                oInvoice.DocRate = docRate
                            End If
                        End If

                        ' 付款條件
                        If Not IsDBNull(drH("GroupNum")) Then
                            oInvoice.GroupNumber = Convert.ToInt32(drH("GroupNum"))
                        End If

                        ' 備註
                        If Not IsDBNull(drH("Comments")) Then
                            oInvoice.Comments = drH("Comments").ToString()
                        End If

                        ' 日記帳備註 (JrnlMemo)
                        If Not IsDBNull(drH("JrnlMemo")) Then
                            oInvoice.JournalMemo = drH("JrnlMemo").ToString()
                        End If

                        ' 收貨地址
                        If Not IsDBNull(drH("Address")) Then
                            oInvoice.Address = drH("Address").ToString()
                        End If

                        ' 採購人員 (SlpCode)
                        If Not IsDBNull(drH("SlpCode")) Then
                            oInvoice.SalesPersonCode = Convert.ToInt32(drH("SlpCode"))
                        End If

                        drH.Close()
                    End Using
                End Using

                ' 讀取 Lines (明細)
                Dim sqlL As String = "SELECT * FROM jPCH1 WHERE DocEntry = @DocEntry ORDER BY LineNum"
                Using cmdL As New SqlCommand(sqlL, conn)
                    cmdL.Parameters.AddWithValue("@DocEntry", docEntry)
                    Using drL As SqlDataReader = cmdL.ExecuteReader()
                        Dim lineIndex As Integer = 0

                        While drL.Read()
                            ' 第一行不需要 Add，後續行要 Add
                            If lineIndex > 0 Then
                                oInvoice.Lines.Add()
                            End If

                            ' ===== 明細欄位設定 =====

                            ' 會計科目 (費用類)
                            oInvoice.Lines.AccountCode = drL("AcctCode").ToString()

                            ' 說明
                            If Not IsDBNull(drL("Dscription")) Then
                                oInvoice.Lines.ItemDescription = drL("Dscription").ToString()
                            End If

                            ' ===== 金額處理 (本幣 vs 外幣) =====
                            Dim lineTotal As Double = 0
                            If Not IsDBNull(drL("LineTotal")) Then
                                lineTotal = Convert.ToDouble(drL("LineTotal"))
                            End If

                            If isForeignCurrency Then
                                ' 外幣單據：
                                ' - LineTotalFC = 外幣金額 (原始輸入金額)
                                ' - LineTotal = 外幣金額 * 匯率 (本幣金額)
                                oInvoice.Lines.LineTotalForeignCurrency = lineTotal
                                oInvoice.Lines.LineTotal = Math.Round(lineTotal * docRate, 2)
                            Else
                                ' 本幣單據：
                                ' - LineTotal = 本幣金額
                                ' - LineTotalFC = 0 (不使用)
                                oInvoice.Lines.LineTotal = lineTotal
                            End If

                            ' 稅碼
                            If Not IsDBNull(drL("VatGroup")) Then
                                oInvoice.Lines.VatGroup = drL("VatGroup").ToString()
                            End If

                            ' 成本中心1 (產品/專案)
                            If Not IsDBNull(drL("CostingCode")) AndAlso drL("CostingCode").ToString() <> "" Then
                                oInvoice.Lines.CostingCode = drL("CostingCode").ToString()
                            End If

                            ' 成本中心2 (部門)
                            If Not IsDBNull(drL("CostingCode2")) AndAlso drL("CostingCode2").ToString() <> "" Then
                                oInvoice.Lines.CostingCode2 = drL("CostingCode2").ToString()
                            End If

                            lineIndex += 1
                        End While

                        drL.Close()
                    End Using
                End Using

                conn.Close()
            End Using

            ' ===== 新增到 SAP =====
            Dim addResult As Integer = oInvoice.Add()

            If addResult = 0 Then
                ' 成功，取得 SAP DocEntry
                sapDocEntry = Convert.ToInt32(oCompany.GetNewObjectKey())

                ' 更新 jtdb 的狀態
                UpdatePostStatus(docEntry, sapDocEntry, "Y", "")

                Return sapDocEntry
            Else
                ' 失敗，取得錯誤訊息
                Dim errCode As Integer
                Dim errMsg As String = ""
                oCompany.GetLastError(errCode, errMsg)

                ' 更新 jtdb 的錯誤狀態
                UpdatePostStatus(docEntry, 0, "E", errMsg)

                Throw New Exception("SAP 新增失敗 [" & errCode & "]: " & errMsg)
            End If

        Catch ex As Exception
            ' 記錄錯誤
            UpdatePostStatus(docEntry, 0, "E", ex.Message)
            Throw
        Finally
            ' 釋放 COM 物件
            If oInvoice IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice)
                oInvoice = Nothing
            End If
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

#Region "批次匯入"

    ''' <summary>
    ''' 批次匯入所有待過帳的 AP 發票
    ''' </summary>
    ''' <returns>匯入結果 (成功數, 失敗數)</returns>
    Public Function ImportAllPendingInvoices() As Tuple(Of Integer, Integer)
        Dim successCount As Integer = 0
        Dim failCount As Integer = 0
        Dim pendingList As New List(Of Integer)

        ' 取得所有待過帳的文件 (ApprovalStatus = 'A' 且 B1PostStatus <> 'Y')
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

        ' 逐筆匯入
        For Each docEntry As Integer In pendingList
            Try
                ImportAPInvoice(docEntry)
                successCount += 1
            Catch
                failCount += 1
            End Try
        Next

        Return New Tuple(Of Integer, Integer)(successCount, failCount)
    End Function

#End Region

#Region "靜態方法 (便於頁面呼叫)"

    ''' <summary>
    ''' 匯入單一 AP 發票 (靜態方法，自動處理連線)
    ''' </summary>
    ''' <param name="docEntry">jOPCH 的 DocEntry</param>
    ''' <returns>成功回傳訊息，失敗回傳錯誤訊息</returns>
    Public Shared Function Import(docEntry As Integer) As String
        Dim importer As New SapAPInvoiceImporter()

        Try
            importer.ConnectToSAP()
            Dim sapDocEntry As Integer = importer.ImportAPInvoice(docEntry)
            Return "匯入成功! SAP 單號: " & sapDocEntry

        Catch ex As Exception
            Return "匯入失敗: " & ex.Message
        Finally
            importer.DisconnectFromSAP()
        End Try
    End Function

    ''' <summary>
    ''' 批次匯入所有待過帳文件 (靜態方法，自動處理連線)
    ''' </summary>
    ''' <returns>匯入結果訊息</returns>
    Public Shared Function ImportAllPending() As String
        Dim importer As New SapAPInvoiceImporter()

        Try
            importer.ConnectToSAP()
            Dim result = importer.ImportAllPendingInvoices()
            Return String.Format("批次匯入完成! 成功: {0} 筆, 失敗: {1} 筆", result.Item1, result.Item2)

        Catch ex As Exception
            Return "批次匯入失敗: " & ex.Message
        Finally
            importer.DisconnectFromSAP()
        End Try
    End Function

#End Region

End Class


'=============================================================================
' Web.config 設定範例
'=============================================================================
'
' 在 <appSettings> 區段加入以下設定:
'
'   <!-- SAP DI API 連線設定 -->
'   <add key="SAP_ServerIP" value="192.168.1.219"/>
'   <add key="SAP_CompanyDB" value="JTTST"/>
'   <add key="SAP_DbServerType" value="dst_MSSQL2014"/>
'   <!-- SAP 資料庫帳密 -->
'   <add key="SAP_DbUser" value="sa"/>
'   <add key="SAP_DbPassword" value="sap19690123"/>
'   <!-- SAP DI API 登入帳密 (平台統一使用 B1i 帳號) -->
'   <add key="SAP_User" value="B1i"/>
'   <add key="SAP_Password" value="5587"/>
'
' 支援的 DbServerType 值:
'   dst_MSSQL2005, dst_MSSQL2008, dst_MSSQL2012, dst_MSSQL2014,
'   dst_MSSQL2016, dst_MSSQL2017, dst_MSSQL2019
'
'=============================================================================
' 使用範例
'=============================================================================
'
' 方法一：靜態方法 (簡單，適合頁面呼叫)
'   Dim result As String = SapAPInvoiceImporter.Import(docEntry:=123)
'   lblMessage.Text = result
'
' 方法二：實例方法 (適合批次處理或需要更多控制)
'   Dim importer As New SapAPInvoiceImporter()
'   Try
'       importer.ConnectToSAP()
'       Dim sapDocEntry As Integer = importer.ImportAPInvoice(123)
'       ' 處理成功
'   Catch ex As Exception
'       ' 處理失敗
'   Finally
'       importer.DisconnectFromSAP()
'   End Try
'
'=============================================================================
' 本幣 vs 外幣金額處理
'=============================================================================
'
' [本幣單據] (DocCurrency = "TWD")
'   - LineTotal = 本幣金額
'   - LineTotalFC = 0 (不使用)
'
' [外幣單據] (DocCurrency = "USD", "JPY", etc.)
'   - LineTotalFC = 外幣金額 (使用者輸入的原始金額)
'   - LineTotal = 外幣金額 * 匯率 (換算後的本幣金額)
'   - DocRate = 匯率
'
' 範例：USD 單據，金額 100 USD，匯率 31.5
'   - LineTotalFC = 100
'   - LineTotal = 100 * 31.5 = 3150
'
'=============================================================================
' B1PostStatus 狀態說明
'=============================================================================
' N 或 NULL - 未過帳
' Y         - 已過帳成功
' E         - 過帳失敗 (錯誤訊息存在 B1ErrMsg)
'
'=============================================================================
