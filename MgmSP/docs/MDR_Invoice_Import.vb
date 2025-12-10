'=============================================================================
' MDR 營業稅發票資料匯入功能
'
' 用途：將 jtdb.jMGUIAP/jMGUIAPDetail 資料同步到 MDR.MGUIAP/MGUIAPDetail
' 日期：2025/12/08
' 版本：1.0
'
' 資料流向：
'   jtdb.jMGUIAP       -> MDR.MGUIAP
'   jtdb.jMGUIAPDetail -> MDR.MGUIAPDetail
'
' 注意事項：
' 1. 連線資訊在 Web.config 的 connectionStrings 區段
' 2. jtdb 使用 "jtdbConnectionString"
' 3. MDR 使用 "MDRConnectionString"
'=============================================================================

Imports System.Data.SqlClient
Imports System.Configuration

''' <summary>
''' MDR 營業稅發票資料匯入類別
''' </summary>
Public Class MdrInvoiceImporter

    Private jtdbConnStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private mdrConnStr As String = ConfigurationManager.ConnectionStrings("MDRConnectionString").ConnectionString

#Region "單筆匯入"

    ''' <summary>
    ''' 匯入單一憑證到 MDR
    ''' </summary>
    ''' <param name="docEntry">jOPCH 的 DocEntry (對應 jMGUIAP.DocEntry)</param>
    ''' <returns>成功回傳 MDR ID，失敗拋出 Exception</returns>
    Public Function ImportToMDR(docEntry As Integer) As Integer
        Dim mdrID As Integer = -1

        Try
            ' 從 jtdb 讀取資料
            Dim headerData As DataTable = Nothing
            Dim detailData As DataTable = Nothing

            Using jtdbConn As New SqlConnection(jtdbConnStr)
                jtdbConn.Open()

                ' 讀取表頭
                Dim sqlH As String = "SELECT * FROM jMGUIAP WHERE DocEntry = @DocEntry"
                Using adapter As New SqlDataAdapter(sqlH, jtdbConn)
                    adapter.SelectCommand.Parameters.AddWithValue("@DocEntry", docEntry)
                    headerData = New DataTable()
                    adapter.Fill(headerData)
                End Using

                If headerData.Rows.Count = 0 Then
                    Throw New Exception("找不到 MDR 表頭資料: DocEntry=" & docEntry)
                End If

                ' 讀取明細
                Dim jID As Integer = Convert.ToInt32(headerData.Rows(0)("jID"))
                Dim sqlL As String = "SELECT * FROM jMGUIAPDetail WHERE DocEntry = @DocEntry ORDER BY LineNum"
                Using adapter As New SqlDataAdapter(sqlL, jtdbConn)
                    adapter.SelectCommand.Parameters.AddWithValue("@DocEntry", docEntry)
                    detailData = New DataTable()
                    adapter.Fill(detailData)
                End Using
            End Using

            ' 寫入 MDR 資料庫
            Using mdrConn As New SqlConnection(mdrConnStr)
                mdrConn.Open()
                Using trans As SqlTransaction = mdrConn.BeginTransaction()
                    Try
                        Dim headerRow As DataRow = headerData.Rows(0)

                        ' 檢查是否已存在 (用 DocEntry 判斷)
                        Dim existingID As Integer = 0
                        Dim sqlCheck As String = "SELECT ID FROM MGUIAP WHERE DocEntry = @DocEntry"
                        Using cmdCheck As New SqlCommand(sqlCheck, mdrConn, trans)
                            cmdCheck.Parameters.AddWithValue("@DocEntry", docEntry)
                            Dim result = cmdCheck.ExecuteScalar()
                            If result IsNot Nothing Then
                                existingID = Convert.ToInt32(result)
                            End If
                        End Using

                        If existingID > 0 Then
                            ' 已存在，更新
                            mdrID = existingID
                            UpdateMDRHeader(mdrConn, trans, mdrID, headerRow)
                            DeleteMDRDetails(mdrConn, trans, mdrID)
                        Else
                            ' 不存在，新增
                            mdrID = InsertMDRHeader(mdrConn, trans, headerRow)
                        End If

                        ' 新增明細
                        For Each detailRow As DataRow In detailData.Rows
                            InsertMDRDetail(mdrConn, trans, mdrID, detailRow)
                        Next

                        trans.Commit()

                        ' 更新 jtdb 的同步狀態
                        UpdateSyncStatus(docEntry, "Y", "")

                    Catch ex As Exception
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            Return mdrID

        Catch ex As Exception
            ' 更新錯誤狀態
            UpdateSyncStatus(docEntry, "E", ex.Message)
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 新增 MDR 表頭
    ''' </summary>
    Private Function InsertMDRHeader(conn As SqlConnection, trans As SqlTransaction, row As DataRow) As Integer
        Dim sql As String = "INSERT INTO MGUIAP (DocEntry, DocNum, DocTotal, VatSum, U_OBJTYPE, CreateDate, CreateBy) " &
                           "VALUES (@DocEntry, @DocNum, @DocTotal, @VatSum, @OBJTYPE, GETDATE(), @CreateBy); " &
                           "SELECT SCOPE_IDENTITY();"

        Using cmd As New SqlCommand(sql, conn, trans)
            cmd.Parameters.AddWithValue("@DocEntry", GetDbValue(row, "DocEntry"))
            cmd.Parameters.AddWithValue("@DocNum", GetDbValue(row, "DocNum"))
            cmd.Parameters.AddWithValue("@DocTotal", GetDbValue(row, "DocTotal"))
            cmd.Parameters.AddWithValue("@VatSum", GetDbValue(row, "VatSum"))
            cmd.Parameters.AddWithValue("@OBJTYPE", GetDbValue(row, "U_OBJTYPE", "18")) ' 預設 18 = AP Invoice
            cmd.Parameters.AddWithValue("@CreateBy", GetDbValue(row, "CreateBy", "SYSTEM"))

            Return Convert.ToInt32(cmd.ExecuteScalar())
        End Using
    End Function

    ''' <summary>
    ''' 更新 MDR 表頭
    ''' </summary>
    Private Sub UpdateMDRHeader(conn As SqlConnection, trans As SqlTransaction, mdrID As Integer, row As DataRow)
        Dim sql As String = "UPDATE MGUIAP SET DocNum=@DocNum, DocTotal=@DocTotal, VatSum=@VatSum, " &
                           "U_OBJTYPE=@OBJTYPE, UpdateDate=GETDATE(), UpdateBy=@UpdateBy WHERE ID=@ID"

        Using cmd As New SqlCommand(sql, conn, trans)
            cmd.Parameters.AddWithValue("@ID", mdrID)
            cmd.Parameters.AddWithValue("@DocNum", GetDbValue(row, "DocNum"))
            cmd.Parameters.AddWithValue("@DocTotal", GetDbValue(row, "DocTotal"))
            cmd.Parameters.AddWithValue("@VatSum", GetDbValue(row, "VatSum"))
            cmd.Parameters.AddWithValue("@OBJTYPE", GetDbValue(row, "U_OBJTYPE", "18"))
            cmd.Parameters.AddWithValue("@UpdateBy", GetDbValue(row, "UpdateBy", "SYSTEM"))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' 刪除 MDR 明細 (更新前先清除舊資料)
    ''' </summary>
    Private Sub DeleteMDRDetails(conn As SqlConnection, trans As SqlTransaction, mdrID As Integer)
        Dim sql As String = "DELETE FROM MGUIAPDetail WHERE MGUIAP_ID = @ID"
        Using cmd As New SqlCommand(sql, conn, trans)
            cmd.Parameters.AddWithValue("@ID", mdrID)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' 新增 MDR 明細
    ''' </summary>
    Private Sub InsertMDRDetail(conn As SqlConnection, trans As SqlTransaction, mdrID As Integer, row As DataRow)
        Dim sql As String = "INSERT INTO MGUIAPDetail (MGUIAP_ID, LineNum, U_LIFNR, U_STCEG, U_XBLNR, U_ZFORM_CODE, " &
                           "U_BLDAT, U_VATDATE, U_HWBAS, U_HWSTE, U_TAX_TYPE, U_CUS_TYPE, U_AM_TYPE, " &
                           "U_VATCODE, U_BUKRS, U_MWSKZ, U_BELNR, U_FA_DESC, U_FA_QTY, U_FA_USE, " &
                           "U_GatherMark, U_ConsolidQty, CreateDate, CreateBy) " &
                           "VALUES (@MGUIAP_ID, @LineNum, @LIFNR, @STCEG, @XBLNR, @ZFORM, " &
                           "@BLDAT, @VATDATE, @HWBAS, @HWSTE, @TAX_TYPE, @CUS_TYPE, @AM_TYPE, " &
                           "@VATCODE, @BUKRS, @MWSKZ, @BELNR, @FA_DESC, @FA_QTY, @FA_USE, " &
                           "@GatherMark, @ConsolidQty, GETDATE(), @CreateBy)"

        Using cmd As New SqlCommand(sql, conn, trans)
            cmd.Parameters.AddWithValue("@MGUIAP_ID", mdrID)
            cmd.Parameters.AddWithValue("@LineNum", GetDbValue(row, "LineNum"))
            cmd.Parameters.AddWithValue("@LIFNR", GetDbValue(row, "U_LIFNR"))
            cmd.Parameters.AddWithValue("@STCEG", GetDbValue(row, "U_STCEG"))
            cmd.Parameters.AddWithValue("@XBLNR", GetDbValue(row, "U_XBLNR"))
            cmd.Parameters.AddWithValue("@ZFORM", GetDbValue(row, "U_ZFORM_CODE"))
            cmd.Parameters.AddWithValue("@BLDAT", GetDbValue(row, "U_BLDAT"))
            cmd.Parameters.AddWithValue("@VATDATE", GetDbValue(row, "U_VATDATE"))
            cmd.Parameters.AddWithValue("@HWBAS", GetDbValue(row, "U_HWBAS"))
            cmd.Parameters.AddWithValue("@HWSTE", GetDbValue(row, "U_HWSTE"))
            cmd.Parameters.AddWithValue("@TAX_TYPE", GetDbValue(row, "U_TAX_TYPE"))
            cmd.Parameters.AddWithValue("@CUS_TYPE", GetDbValue(row, "U_CUS_TYPE"))
            cmd.Parameters.AddWithValue("@AM_TYPE", GetDbValue(row, "U_AM_TYPE"))
            cmd.Parameters.AddWithValue("@VATCODE", GetDbValue(row, "U_VATCODE"))
            cmd.Parameters.AddWithValue("@BUKRS", GetDbValue(row, "U_BUKRS"))
            cmd.Parameters.AddWithValue("@MWSKZ", GetDbValue(row, "U_MWSKZ"))
            cmd.Parameters.AddWithValue("@BELNR", GetDbValue(row, "U_BELNR"))
            cmd.Parameters.AddWithValue("@FA_DESC", GetDbValue(row, "U_FA_DESC"))
            cmd.Parameters.AddWithValue("@FA_QTY", GetDbValue(row, "U_FA_QTY"))
            cmd.Parameters.AddWithValue("@FA_USE", GetDbValue(row, "U_FA_USE"))
            cmd.Parameters.AddWithValue("@GatherMark", GetDbValue(row, "U_GatherMark"))
            cmd.Parameters.AddWithValue("@ConsolidQty", GetDbValue(row, "U_ConsolidQty"))
            cmd.Parameters.AddWithValue("@CreateBy", GetDbValue(row, "CreateBy", "SYSTEM"))
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' 更新 jtdb 的同步狀態
    ''' </summary>
    Private Sub UpdateSyncStatus(docEntry As Integer, status As String, errMsg As String)
        Try
            Using conn As New SqlConnection(jtdbConnStr)
                conn.Open()
                Dim sql As String = "UPDATE jMGUIAP SET MDRPostStatus=@Status, MDRPostDate=GETDATE(), " &
                                   "MDRErrMsg=@ErrMsg WHERE DocEntry=@DocEntry"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@ErrMsg", If(String.IsNullOrEmpty(errMsg), DBNull.Value, errMsg))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
            ' 忽略狀態更新錯誤
        End Try
    End Sub

    ''' <summary>
    ''' 取得 DataRow 欄位值，處理 DBNull
    ''' </summary>
    Private Function GetDbValue(row As DataRow, columnName As String, Optional defaultValue As Object = Nothing) As Object
        If row.Table.Columns.Contains(columnName) AndAlso Not IsDBNull(row(columnName)) Then
            Return row(columnName)
        Else
            Return If(defaultValue, DBNull.Value)
        End If
    End Function

#End Region

#Region "批次匯入"

    ''' <summary>
    ''' 批次匯入所有待同步的 MDR 資料
    ''' </summary>
    ''' <returns>(成功數, 失敗數)</returns>
    Public Function ImportAllPending() As Tuple(Of Integer, Integer)
        Dim successCount As Integer = 0
        Dim failCount As Integer = 0
        Dim pendingList As New List(Of Integer)

        ' 取得所有待同步的文件 (MDRPostStatus <> 'Y')
        Using conn As New SqlConnection(jtdbConnStr)
            conn.Open()
            Dim sql As String = "SELECT DISTINCT DocEntry FROM jMGUIAP " &
                               "WHERE DocEntry IS NOT NULL AND (MDRPostStatus IS NULL OR MDRPostStatus <> 'Y')"
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
                ImportToMDR(docEntry)
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
    ''' 匯入單筆 MDR 資料 (靜態方法)
    ''' </summary>
    ''' <param name="docEntry">jOPCH 的 DocEntry</param>
    ''' <returns>結果訊息</returns>
    Public Shared Function Import(docEntry As Integer) As String
        Try
            Dim importer As New MdrInvoiceImporter()
            Dim mdrID As Integer = importer.ImportToMDR(docEntry)
            Return "MDR 同步成功! ID: " & mdrID
        Catch ex As Exception
            Return "MDR 同步失敗: " & ex.Message
        End Try
    End Function

    ''' <summary>
    ''' 批次匯入所有待同步 MDR 資料 (靜態方法)
    ''' </summary>
    ''' <returns>結果訊息</returns>
    Public Shared Function ImportAll() As String
        Try
            Dim importer As New MdrInvoiceImporter()
            Dim result = importer.ImportAllPending()
            Return String.Format("MDR 批次同步完成! 成功: {0} 筆, 失敗: {1} 筆", result.Item1, result.Item2)
        Catch ex As Exception
            Return "MDR 批次同步失敗: " & ex.Message
        End Try
    End Function

#End Region

End Class


'=============================================================================
' 使用範例
'=============================================================================
'
' 方法一：單筆匯入
'   Dim result As String = MdrInvoiceImporter.Import(docEntry:=123)
'   lblMessage.Text = result
'   ' 回傳: "MDR 同步成功! ID: 456" 或 "MDR 同步失敗: xxx"
'
' 方法二：批次匯入
'   Dim result As String = MdrInvoiceImporter.ImportAll()
'   lblMessage.Text = result
'   ' 回傳: "MDR 批次同步完成! 成功: 5 筆, 失敗: 1 筆"
'
'=============================================================================
' 完整匯入流程 (AP + MDR)
'=============================================================================
'
' 當單據核准後，執行以下兩個匯入：
'
'   ' 1. 匯入 SAP AP Invoice
'   Dim apResult As String = SapAPInvoiceImporter.Import(docEntry)
'
'   ' 2. 同步 MDR 營業稅資料
'   Dim mdrResult As String = MdrInvoiceImporter.Import(docEntry)
'
'   lblMessage.Text = apResult & " | " & mdrResult
'
'=============================================================================
' MDRPostStatus 狀態說明
'=============================================================================
' NULL 或空 - 未同步
' Y         - 同步成功
' E         - 同步失敗 (錯誤訊息存在 MDRErrMsg)
'
'=============================================================================
