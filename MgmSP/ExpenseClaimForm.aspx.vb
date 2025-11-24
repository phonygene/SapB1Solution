Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Configuration

''' <summary>
''' 費用申請單 - 第一階段：表頭欄位與基本架構
''' 建立日期: 2025-11-05
''' 說明: 實作表頭欄位、檔案上傳、審核功能
''' </summary>
Partial Public Class ExpenseClaimForm
    Inherits System.Web.UI.Page

#Region "變數宣告"
    ' 資料庫連線字串
    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private ReadOnly sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString

    ' 當前使用者資訊
    Private currentUserId As String = ""
    Private canApprove As Boolean = False

    ' 當前單據編號（編輯模式）
    Private currentDocEntry As Integer = 0
#End Region

#Region "頁面載入"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ' 檢查使用者登入
            If Session("s_id") Is Nothing Then
                Response.Redirect("~/usermgm/login.aspx")
                Return
            End If

            currentUserId = Session("s_id").ToString()

            ' 檢查審核權限
            CheckApprovalPermission()

            If Not IsPostBack Then
                ' 初始化下拉選單
                InitializeDropDowns()

                ' 檢查是否為編輯模式
                If Request.QueryString("DocEntry") IsNot Nothing Then
                    currentDocEntry = Convert.ToInt32(Request.QueryString("DocEntry"))
                    LoadDocument(currentDocEntry)
                Else
                    ' 新增模式：設定預設值
                    SetDefaultValues()
                    InitializeGridView()
                    InitializeMDRGridView()
                End If
            End If

        Catch ex As Exception
            lblMessage.Text = "頁面載入錯誤: " & ex.Message
            LogError("Page_Load", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 檢查使用者是否有審核權限
    ''' </summary>
    Private Sub CheckApprovalPermission()
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT Approver FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", currentUserId)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        canApprove = (Convert.ToInt32(result) = 1)
                    End If
                End Using
            End Using

            ' 根據權限顯示/隱藏審核區塊
            pnlApproval.Visible = canApprove

        Catch ex As Exception
            LogError("CheckApprovalPermission", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初始化所有下拉選單
    ''' </summary>
    Private Sub InitializeDropDowns()
        LoadVendors()           ' 供應商
        LoadDeliveryAddress()   ' 收貨地址
        LoadDepartments()       ' 產品/部門
        LoadCurrencies()        ' 幣別
    End Sub

    ''' <summary>
    ''' 載入供應商清單（從 SAP B1）
    ''' </summary>
    Private Sub LoadVendors()
        Try
            ddlCardCode.Items.Clear()
            ddlCardCode.Items.Add(New ListItem("-- 請選擇供應商 --", ""))

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT CardCode, CardName FROM OCRD WHERE CardType = 'S' AND frozenFor = 'N' ORDER BY CardName"
                Using cmd As New SqlCommand(sql, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim code As String = reader("CardCode").ToString()
                            Dim name As String = reader("CardName").ToString()
                            ddlCardCode.Items.Add(New ListItem($"{code} - {name}", code))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            lblMessage.Text = "載入供應商失敗: " & ex.Message
            LogError("LoadVendors", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 載入收貨地址清單
    ''' </summary>
    Private Sub LoadDeliveryAddress()
        Try
            ddlDeliveryAddr.Items.Clear()
            ddlDeliveryAddr.Items.Add(New ListItem("-- 請選擇收貨地址 --", ""))

            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT ID, addrName, address FROM addr WHERE addrType = 'R' AND active = 'Y' ORDER BY addrName"
                Using cmd As New SqlCommand(sql, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim id As String = reader("ID").ToString()
                            Dim name As String = reader("addrName").ToString()
                            Dim addr As String = reader("address").ToString()
                            ddlDeliveryAddr.Items.Add(New ListItem($"{name} - {addr}", id))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            lblMessage.Text = "載入收貨地址失敗: " & ex.Message
            LogError("LoadDeliveryAddress", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 載入產品/部門清單（從 SAP B1 OOCR）
    ''' </summary>
    Private Sub LoadDepartments()
        Try
            ddlOcrCode.Items.Clear()
            ddlOcrCode.Items.Add(New ListItem("-- 請選擇產品/部門 --", ""))

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                ' 根據實作計畫書 v1.1，使用 OOCR 表並過濾 DimCode
                Dim sql As String = "SELECT OcrCode, OcrName FROM OOCR WHERE DimCode = 1 AND Active = 'Y' ORDER BY OcrName"
                Using cmd As New SqlCommand(sql, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim code As String = reader("OcrCode").ToString()
                            Dim name As String = reader("OcrName").ToString()
                            ddlOcrCode.Items.Add(New ListItem($"{code} - {name}", code))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            lblMessage.Text = "載入產品/部門失敗: " & ex.Message
            LogError("LoadDepartments", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 載入幣別清單（從 SAP B1 OCRN）
    ''' </summary>
    Private Sub LoadCurrencies()
        Try
            ddlDocCurrency.Items.Clear()

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT CurrCode, CurrName FROM OCRN ORDER BY CurrCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim code As String = reader("CurrCode").ToString()
                            Dim name As String = reader("CurrName").ToString()
                            ddlDocCurrency.Items.Add(New ListItem($"{code} - {name}", code))
                        End While
                    End Using
                End Using
            End Using

            ' 預設選擇本地幣別（TWD）
            If ddlDocCurrency.Items.FindByValue("TWD") IsNot Nothing Then
                ddlDocCurrency.SelectedValue = "TWD"
            End If

        Catch ex As Exception
            lblMessage.Text = "載入幣別失敗: " & ex.Message
            LogError("LoadCurrencies", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 設定新增模式的預設值
    ''' </summary>
    Private Sub SetDefaultValues()
        lblDocNum.Text = "[自動產生]"
        lblDocStatus.Text = "草稿"
        lblCreateBy.Text = currentUserId
        lblCreateDate.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        txtDocDate.Text = DateTime.Now.ToString("yyyy-MM-dd")
        txtDocDueDate.Text = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd")
        txtDocRate.Text = "1.0"

        lblAttachment.Text = "無"
        btnDownload.Visible = False
    End Sub

#End Region

#Region "供應商選擇事件"
    ''' <summary>
    ''' 供應商選擇變更：自動填入供應商名稱並載入聯絡人
    ''' </summary>
    Protected Sub ddlCardCode_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            If String.IsNullOrEmpty(ddlCardCode.SelectedValue) Then
                txtCardName.Text = ""
                ddlContactPerson.Items.Clear()
                Return
            End If

            Dim cardCode As String = ddlCardCode.SelectedValue

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()

                ' 取得供應商名稱
                Dim sql As String = "SELECT CardName FROM OCRD WHERE CardCode = @CardCode"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@CardCode", cardCode)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        txtCardName.Text = result.ToString()
                    End If
                End Using

                ' 載入聯絡人清單
                LoadContactPersons(cardCode, conn)
            End Using

            ' 同步 MDR Tab 表頭資訊
            SyncMDRHeaderInfo()

        Catch ex As Exception
            lblMessage.Text = "載入供應商資料失敗: " & ex.Message
            LogError("ddlCardCode_SelectedIndexChanged", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 載入聯絡人清單
    ''' </summary>
    Private Sub LoadContactPersons(cardCode As String, conn As SqlConnection)
        Try
            ddlContactPerson.Items.Clear()
            ddlContactPerson.Items.Add(New ListItem("-- 請選擇聯絡人 --", ""))

            Dim sql As String = "SELECT CntctCode, Name FROM OCPR WHERE CardCode = @CardCode AND Active = 'Y' ORDER BY Name"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@CardCode", cardCode)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim code As String = reader("CntctCode").ToString()
                        Dim name As String = reader("Name").ToString()
                        ddlContactPerson.Items.Add(New ListItem(name, code))
                    End While
                End Using
            End Using

        Catch ex As Exception
            LogError("LoadContactPersons", ex)
        End Try
    End Sub
#End Region

#Region "幣別與匯率處理"
    ''' <summary>
    ''' 幣別變更：自動更新匯率
    ''' </summary>
    Protected Sub ddlDocCurrency_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            UpdateExchangeRate()
            ' 同步 MDR Tab 表頭資訊
            SyncMDRHeaderInfo()
        Catch ex As Exception
            lblMessage.Text = "更新匯率失敗: " & ex.Message
            LogError("ddlDocCurrency_SelectedIndexChanged", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 手動更新匯率按鈕
    ''' </summary>
    Protected Sub btnRefreshRate_Click(sender As Object, e As EventArgs)
        Try
            UpdateExchangeRate()
            ' 同步 MDR Tab 表頭資訊
            SyncMDRHeaderInfo()
            lblMessage.Text = "匯率已更新"
        Catch ex As Exception
            lblMessage.Text = "更新匯率失敗: " & ex.Message
            LogError("btnRefreshRate_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 從 SAP B1 ORTT 取得匯率
    ''' </summary>
    Private Sub UpdateExchangeRate()
        Dim currency As String = ddlDocCurrency.SelectedValue

        ' 本地幣別匯率固定為 1
        If currency = "TWD" Then
            txtDocRate.Text = "1.0"
            txtDocRate.ReadOnly = True
            Return
        End If

        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                ' 取得最新匯率（依照日期降序）
                Dim sql As String = "SELECT TOP 1 Rate FROM ORTT WHERE Currency = @Currency ORDER BY RateDate DESC"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Currency", currency)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso IsNumeric(result) Then
                        txtDocRate.Text = Convert.ToDecimal(result).ToString("F6")
                    Else
                        txtDocRate.Text = "1.0"
                        lblMessage.Text = $"查無 {currency} 匯率，請手動輸入"
                    End If
                End Using
            End Using

            ' 外幣允許手動編輯
            txtDocRate.ReadOnly = False

        Catch ex As Exception
            txtDocRate.Text = "1.0"
            txtDocRate.ReadOnly = False
            LogError("UpdateExchangeRate", ex)
        End Try
    End Sub
#End Region

#Region "檔案上傳處理"
    ''' <summary>
    ''' 檔案上傳按鈕
    ''' </summary>
    Protected Sub btnUpload_Click(sender As Object, e As EventArgs)
        Try
            If Not fileUpload.HasFile Then
                lblMessage.Text = "請選擇要上傳的檔案"
                Return
            End If

            ' 檔案大小限制（10MB）
            If fileUpload.PostedFile.ContentLength > 10485760 Then
                lblMessage.Text = "檔案大小不可超過 10MB"
                Return
            End If

            ' 取得原始檔名並清理
            Dim originalFileName As String = Path.GetFileName(fileUpload.FileName)
            Dim cleanedFileName As String = CleanFileName(originalFileName)

            ' 產生唯一檔名：清理後檔名 + 時間戳記
            Dim timestamp As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim extension As String = Path.GetExtension(cleanedFileName)
            Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(cleanedFileName)
            Dim uniqueFileName As String = $"{nameWithoutExt}_{timestamp}{extension}"

            ' 儲存路徑
            Dim uploadFolder As String = Server.MapPath("~/Uploads/ExpenseClaims/")
            If Not Directory.Exists(uploadFolder) Then
                Directory.CreateDirectory(uploadFolder)
            End If

            Dim filePath As String = Path.Combine(uploadFolder, uniqueFileName)
            fileUpload.SaveAs(filePath)

            ' 更新顯示
            lblAttachment.Text = uniqueFileName
            btnDownload.Visible = True
            lblMessage.Text = "檔案上傳成功"

            ' 若為編輯模式，更新資料庫
            If currentDocEntry > 0 Then
                UpdateAttachmentInDB(uniqueFileName)
            End If

        Catch ex As Exception
            lblMessage.Text = "檔案上傳失敗: " & ex.Message
            LogError("btnUpload_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 檔案下載按鈕
    ''' </summary>
    Protected Sub btnDownload_Click(sender As Object, e As EventArgs)
        Try
            Dim fileName As String = lblAttachment.Text
            If fileName = "無" OrElse String.IsNullOrEmpty(fileName) Then
                lblMessage.Text = "無可下載的檔案"
                Return
            End If

            Dim filePath As String = Server.MapPath($"~/Uploads/ExpenseClaims/{fileName}")
            If File.Exists(filePath) Then
                Response.ContentType = "application/octet-stream"
                Response.AppendHeader("Content-Disposition", $"attachment; filename=""{fileName}""")
                Response.TransmitFile(filePath)
                Response.End()
            Else
                lblMessage.Text = "檔案不存在"
            End If

        Catch ex As Exception
            lblMessage.Text = "檔案下載失敗: " & ex.Message
            LogError("btnDownload_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 清理檔名：移除不安全字元
    ''' </summary>
    Private Function CleanFileName(fileName As String) As String
        ' 移除路徑字元和特殊字元
        Dim invalidChars As Char() = Path.GetInvalidFileNameChars()
        For Each c As Char In invalidChars
            fileName = fileName.Replace(c, "_"c)
        Next

        ' 移除額外的不安全字元
        fileName = fileName.Replace(" ", "_")
        fileName = fileName.Replace("&", "and")

        Return fileName
    End Function

    ''' <summary>
    ''' 更新資料庫中的附件欄位
    ''' </summary>
    Private Sub UpdateAttachmentInDB(fileName As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE jOPCH SET Attachment = @Attachment WHERE DocEntry = @DocEntry"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Attachment", fileName)
                    cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            LogError("UpdateAttachmentInDB", ex)
        End Try
    End Sub
#End Region

#Region "審核功能"
    ''' <summary>
    ''' 放行按鈕
    ''' </summary>
    Protected Sub btnApprove_Click(sender As Object, e As EventArgs)
        Try
            If currentDocEntry = 0 Then
                lblMessage.Text = "請先儲存單據"
                Return
            End If

            ' 更新審核狀態
            UpdateApprovalStatus("A", "已核准")

            lblMessage.Text = "費用申請單已放行"
            lblApprovalStatus.Text = "已核准"
            lblApprovalStatus.CssClass = "status-approved"

        Catch ex As Exception
            lblMessage.Text = "放行失敗: " & ex.Message
            LogError("btnApprove_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 駁回按鈕
    ''' </summary>
    Protected Sub btnReject_Click(sender As Object, e As EventArgs)
        Try
            If currentDocEntry = 0 Then
                lblMessage.Text = "請先儲存單據"
                Return
            End If

            ' 更新審核狀態
            UpdateApprovalStatus("R", "已駁回")

            lblMessage.Text = "費用申請單已駁回"
            lblApprovalStatus.Text = "已駁回"
            lblApprovalStatus.CssClass = "status-rejected"

        Catch ex As Exception
            lblMessage.Text = "駁回失敗: " & ex.Message
            LogError("btnReject_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 發送通知按鈕
    ''' </summary>
    Protected Sub btnSendNotification_Click(sender As Object, e As EventArgs)
        Try
            If currentDocEntry = 0 Then
                lblMessage.Text = "請先儲存單據"
                Return
            End If

            ' 取得建立人員的 Email
            Dim creatorEmail As String = GetUserEmail(lblCreateBy.Text)
            If String.IsNullOrEmpty(creatorEmail) Then
                lblMessage.Text = "無法取得建立人員的 Email"
                Return
            End If

            ' 發送通知郵件
            Dim subject As String = $"費用申請單 {lblDocNum.Text} 審核通知"
            Dim body As String = BuildNotificationEmail()

            ' 使用 CommUtil.SendMail 發送郵件
            Dim mailUtil As New CommUtil()
            mailUtil.SendMail(creatorEmail, subject, body)
            
            lblMessage.Text = "通知郵件已發送"

        Catch ex As Exception
            lblMessage.Text = "發送通知失敗: " & ex.Message
            LogError("btnSendNotification_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 更新審核狀態
    ''' </summary>
    Private Sub UpdateApprovalStatus(status As String, statusText As String)
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Dim sql As String = "UPDATE jOPCH SET ApprovalStatus = @Status, ApprovedBy = @ApprovedBy, " &
                              "ApprovalDate = CONVERT(DATE, GETDATE()), ApprovalTime = CONVERT(TIME, GETDATE()), " &
                              "ApprovalComments = @Comments WHERE DocEntry = @DocEntry"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Status", status)
                cmd.Parameters.AddWithValue("@ApprovedBy", currentUserId)
                cmd.Parameters.AddWithValue("@Comments", txtApprovalComments.Text)
                cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 建立通知郵件內容
    ''' </summary>
    Private Function BuildNotificationEmail() As String
        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("<html><body>")
        sb.AppendLine("<h2>費用申請單審核通知</h2>")
        sb.AppendLine($"<p><strong>單據編號:</strong> {lblDocNum.Text}</p>")
        sb.AppendLine($"<p><strong>供應商:</strong> {txtCardName.Text}</p>")
        sb.AppendLine($"<p><strong>審核狀態:</strong> {lblApprovalStatus.Text}</p>")
        sb.AppendLine($"<p><strong>審核人員:</strong> {currentUserId}</p>")
        sb.AppendLine($"<p><strong>審核日期:</strong> {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>")

        If Not String.IsNullOrEmpty(txtApprovalComments.Text) Then
            sb.AppendLine($"<p><strong>審核意見:</strong></p>")
            sb.AppendLine($"<p>{txtApprovalComments.Text.Replace(vbCrLf, "<br/>")}</p>")
        End If

        sb.AppendLine("<hr/>")
        sb.AppendLine("<p style='color: gray; font-size: 12px;'>此郵件由系統自動發送，請勿直接回覆。</p>")
        sb.AppendLine("</body></html>")

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 取得使用者 Email
    ''' </summary>
    Private Function GetUserEmail(userId As String) As String
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT email FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        Return result.ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            LogError("GetUserEmail", ex)
        End Try
        Return ""
    End Function
#End Region

#Region "按鈕事件 - 儲存/送出/刪除/取消"
    ''' <summary>
    ''' 儲存按鈕：儲存草稿（狀態保持 P=Pending）
    ''' </summary>
    Protected Sub btnSave_Click(sender As Object, e As EventArgs)
        Try
            ' 基本驗證
            If Not ValidateBasicFields() Then
                Return
            End If

            ' 收集 GridView 資料
            CollectGridViewData()
            If Not ValidateDetailLines() Then
                Return
            End If

            ' 收集 MDR GridView 資料
            CollectMDRGridViewData()
            If Not ValidateMDRLines() Then
                Return
            End If

            If currentDocEntry = 0 Then
                ' 新增模式
                currentDocEntry = InsertNewDocument("P")
                lblDocNum.Text = currentDocEntry.ToString()
                lblMessage.Text = "儲存成功"
            Else
                ' 更新模式
                UpdateDocument()
                lblMessage.Text = "更新成功"
            End If

        Catch ex As Exception
            lblMessage.Text = "儲存失敗: " & ex.Message
            LogError("btnSave_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 送出按鈕：送出至待審核狀態（ApprovalStatus = W）
    ''' </summary>
    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs)
        Try
            ' 完整驗證
            If Not ValidateBasicFields() Then
                Return
            End If

            ' 收集 GridView 資料
            CollectGridViewData()
            If Not ValidateDetailLines() Then
                Return
            End If

            ' 收集 MDR GridView 資料
            CollectMDRGridViewData()
            If Not ValidateMDRLines() Then
                Return
            End If

            If currentDocEntry = 0 Then
                ' 新增模式：儲存並設為待審核
                currentDocEntry = InsertNewDocument("W")
                lblDocNum.Text = currentDocEntry.ToString()
            Else
                ' 更新模式：更新並設為待審核
                UpdateDocument()
                UpdateDocumentStatus("W")
            End If

            lblDocStatus.Text = "待審核"
            lblDocStatus.CssClass = "status-pending"
            lblMessage.Text = "送出成功，等待審核"

        Catch ex As Exception
            lblMessage.Text = "送出失敗: " & ex.Message
            LogError("btnSubmit_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 刪除按鈕
    ''' </summary>
    Protected Sub btnDelete_Click(sender As Object, e As EventArgs)
        Try
            If currentDocEntry = 0 Then
                lblMessage.Text = "無可刪除的單據"
                Return
            End If

            ' 刪除單據（實際應改為邏輯刪除）
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using trans As SqlTransaction = conn.BeginTransaction()
                    Try
                        ' 刪除明細
                        Dim sql1 As String = "DELETE FROM jOPC1 WHERE DocEntry = @DocEntry"
                        Using cmd As New SqlCommand(sql1, conn, trans)
                            cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
                            cmd.ExecuteNonQuery()
                        End Using

                        ' 刪除表頭
                        Dim sql2 As String = "DELETE FROM jOPCH WHERE DocEntry = @DocEntry"
                        Using cmd As New SqlCommand(sql2, conn, trans)
                            cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
                            cmd.ExecuteNonQuery()
                        End Using

                        trans.Commit()
                        Response.Redirect("Index.aspx")

                    Catch ex As Exception
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As Exception
            lblMessage.Text = "刪除失敗: " & ex.Message
            LogError("btnDelete_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取消按鈕：返回列表頁
    ''' </summary>
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs)
        Response.Redirect("Index.aspx")
    End Sub
#End Region

#Region "資料庫操作"
    ''' <summary>
    ''' 新增單據（返回 DocEntry）
    ''' </summary>
    Private Function InsertNewDocument(approvalStatus As String) As Integer
        Dim newDocEntry As Integer = 0

        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using trans As SqlTransaction = conn.BeginTransaction()
                Try
                    ' 插入表頭
                    Dim sql As String = "INSERT INTO jOPCH (CardCode, CardName, ContactPerson, DeliveryAddrID, " &
                                      "DocDate, DocDueDate, OcrCode, DocCurrency, DocRate, Remarks, Attachment, " &
                                      "ApprovalStatus, CreateBy, CreateDate) " &
                                      "VALUES (@CardCode, @CardName, @ContactPerson, @DeliveryAddrID, " &
                                      "@DocDate, @DocDueDate, @OcrCode, @DocCurrency, @DocRate, @Remarks, @Attachment, " &
                                      "@ApprovalStatus, @CreateBy, GETDATE()); SELECT SCOPE_IDENTITY()"

                    Using cmd As New SqlCommand(sql, conn, trans)
                        cmd.Parameters.AddWithValue("@CardCode", GetSelectedValue(ddlCardCode.SelectedValue))
                        cmd.Parameters.AddWithValue("@CardName", txtCardName.Text)
                        cmd.Parameters.AddWithValue("@ContactPerson", GetSelectedValue(ddlContactPerson.SelectedValue))
                        cmd.Parameters.AddWithValue("@DeliveryAddrID", GetSelectedValue(ddlDeliveryAddr.SelectedValue))
                        cmd.Parameters.AddWithValue("@DocDate", DateTime.Parse(txtDocDate.Text))
                        cmd.Parameters.AddWithValue("@DocDueDate", DateTime.Parse(txtDocDueDate.Text))
                        cmd.Parameters.AddWithValue("@OcrCode", GetSelectedValue(ddlOcrCode.SelectedValue))
                        cmd.Parameters.AddWithValue("@DocCurrency", ddlDocCurrency.SelectedValue)
                        cmd.Parameters.AddWithValue("@DocRate", Decimal.Parse(txtDocRate.Text))
                        cmd.Parameters.AddWithValue("@Remarks", txtRemarks.Text)
                        cmd.Parameters.AddWithValue("@Attachment", If(lblAttachment.Text = "無", DBNull.Value, lblAttachment.Text))
                        cmd.Parameters.AddWithValue("@ApprovalStatus", approvalStatus)
                        cmd.Parameters.AddWithValue("@CreateBy", currentUserId)

                        newDocEntry = Convert.ToInt32(cmd.ExecuteScalar())
                    End Using

                    ' 儲存明細
                    SaveDetailLinesToDB(newDocEntry, conn, trans)
                    SaveMDRLinesToDB(newDocEntry, conn, trans)

                    trans.Commit()

                Catch ex As Exception
                    trans.Rollback()
                    Throw
                End Try
            End Using
        End Using

        Return newDocEntry
    End Function

    ''' <summary>
    ''' 更新單據
    ''' </summary>
    Private Sub UpdateDocument()
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using trans As SqlTransaction = conn.BeginTransaction()
                Try
                    ' 更新表頭
                    Dim sql As String = "UPDATE jOPCH SET CardCode = @CardCode, CardName = @CardName, " &
                                      "ContactPerson = @ContactPerson, DeliveryAddrID = @DeliveryAddrID, " &
                                      "DocDate = @DocDate, DocDueDate = @DocDueDate, OcrCode = @OcrCode, " &
                                      "DocCurrency = @DocCurrency, DocRate = @DocRate, Remarks = @Remarks, " &
                                      "Attachment = @Attachment, UpdateDate = GETDATE() " &
                                      "WHERE DocEntry = @DocEntry"

                    Using cmd As New SqlCommand(sql, conn, trans)
                        cmd.Parameters.AddWithValue("@CardCode", GetSelectedValue(ddlCardCode.SelectedValue))
                        cmd.Parameters.AddWithValue("@CardName", txtCardName.Text)
                        cmd.Parameters.AddWithValue("@ContactPerson", GetSelectedValue(ddlContactPerson.SelectedValue))
                        cmd.Parameters.AddWithValue("@DeliveryAddrID", GetSelectedValue(ddlDeliveryAddr.SelectedValue))
                        cmd.Parameters.AddWithValue("@DocDate", DateTime.Parse(txtDocDate.Text))
                        cmd.Parameters.AddWithValue("@DocDueDate", DateTime.Parse(txtDocDueDate.Text))
                        cmd.Parameters.AddWithValue("@OcrCode", GetSelectedValue(ddlOcrCode.SelectedValue))
                        cmd.Parameters.AddWithValue("@DocCurrency", ddlDocCurrency.SelectedValue)
                        cmd.Parameters.AddWithValue("@DocRate", Decimal.Parse(txtDocRate.Text))
                        cmd.Parameters.AddWithValue("@Remarks", txtRemarks.Text)
                        cmd.Parameters.AddWithValue("@Attachment", If(lblAttachment.Text = "無", DBNull.Value, lblAttachment.Text))
                        cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)

                        cmd.ExecuteNonQuery()
                    End Using

                    ' 更新明細
                    SaveDetailLinesToDB(currentDocEntry, conn, trans)
                    SaveMDRLinesToDB(currentDocEntry, conn, trans)

                    trans.Commit()

                Catch ex As Exception
                    trans.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 更新單據狀態
    ''' </summary>
    Private Sub UpdateDocumentStatus(status As String)
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Dim sql As String = "UPDATE jOPCH SET ApprovalStatus = @Status WHERE DocEntry = @DocEntry"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Status", status)
                cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 載入單據（編輯模式）
    ''' </summary>
    Private Sub LoadDocument(docEntry As Integer)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 載入表頭
                Dim sql As String = "SELECT * FROM jOPCH WHERE DocEntry = @DocEntry"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            ' 基本資訊
                            lblDocNum.Text = reader("DocEntry").ToString()
                            lblCreateBy.Text = If(IsDBNull(reader("CreateBy")), "", reader("CreateBy").ToString())
                            lblCreateDate.Text = If(IsDBNull(reader("CreateDate")), "", Convert.ToDateTime(reader("CreateDate")).ToString("yyyy-MM-dd HH:mm:ss"))

                            ' 狀態
                            Dim approvalStatus As String = If(IsDBNull(reader("ApprovalStatus")), "P", reader("ApprovalStatus").ToString())
                            Select Case approvalStatus
                                Case "P"
                                    lblDocStatus.Text = "草稿"
                                    lblDocStatus.CssClass = "status-pending"
                                Case "W"
                                    lblDocStatus.Text = "待審核"
                                    lblDocStatus.CssClass = "status-pending"
                                Case "A"
                                    lblDocStatus.Text = "已核准"
                                    lblDocStatus.CssClass = "status-approved"
                                Case "R"
                                    lblDocStatus.Text = "已駁回"
                                    lblDocStatus.CssClass = "status-rejected"
                            End Select

                            ' 供應商資訊
                            ddlCardCode.SelectedValue = If(IsDBNull(reader("CardCode")), "", reader("CardCode").ToString())
                            txtCardName.Text = If(IsDBNull(reader("CardName")), "", reader("CardName").ToString())

                            ' 載入聯絡人並設定選中值
                            If Not String.IsNullOrEmpty(ddlCardCode.SelectedValue) Then
                                LoadContactPersons(ddlCardCode.SelectedValue, conn)
                                If Not IsDBNull(reader("ContactPerson")) Then
                                    Dim contactPerson As String = reader("ContactPerson").ToString()
                                    If ddlContactPerson.Items.FindByValue(contactPerson) IsNot Nothing Then
                                        ddlContactPerson.SelectedValue = contactPerson
                                    End If
                                End If
                            End If

                            ' 單據資訊
                            If Not IsDBNull(reader("DeliveryAddrID")) Then
                                ddlDeliveryAddr.SelectedValue = reader("DeliveryAddrID").ToString()
                            End If

                            If Not IsDBNull(reader("DocDate")) Then
                                txtDocDate.Text = Convert.ToDateTime(reader("DocDate")).ToString("yyyy-MM-dd")
                            End If

                            If Not IsDBNull(reader("DocDueDate")) Then
                                txtDocDueDate.Text = Convert.ToDateTime(reader("DocDueDate")).ToString("yyyy-MM-dd")
                            End If

                            If Not IsDBNull(reader("OcrCode")) Then
                                ddlOcrCode.SelectedValue = reader("OcrCode").ToString()
                            End If

                            If Not IsDBNull(reader("DocCurrency")) Then
                                ddlDocCurrency.SelectedValue = reader("DocCurrency").ToString()
                            End If

                            txtDocRate.Text = If(IsDBNull(reader("DocRate")), "1.0", Convert.ToDecimal(reader("DocRate")).ToString("F6"))
                            txtRemarks.Text = If(IsDBNull(reader("Remarks")), "", reader("Remarks").ToString())

                            ' 附件
                            If Not IsDBNull(reader("Attachment")) Then
                                lblAttachment.Text = reader("Attachment").ToString()
                                btnDownload.Visible = True
                            Else
                                lblAttachment.Text = "無"
                                btnDownload.Visible = False
                            End If

                            ' 審核資訊
                            If Not IsDBNull(reader("ApprovedBy")) Then
                                lblApprovedBy.Text = reader("ApprovedBy").ToString()
                            End If

                            ' 合併 ApprovalDate 和 ApprovalTime 顯示
                            If Not IsDBNull(reader("ApprovalDate")) AndAlso Not IsDBNull(reader("ApprovalTime")) Then
                                Dim approvalDate As Date = Convert.ToDateTime(reader("ApprovalDate"))
                                Dim approvalTime As TimeSpan = CType(reader("ApprovalTime"), TimeSpan)
                                lblApprovedDate.Text = approvalDate.ToString("yyyy-MM-dd") & " " & approvalTime.ToString("hh\:mm\:ss")
                            ElseIf Not IsDBNull(reader("ApprovalDate")) Then
                                lblApprovedDate.Text = Convert.ToDateTime(reader("ApprovalDate")).ToString("yyyy-MM-dd")
                            End If

                            If Not IsDBNull(reader("ApprovalComments")) Then
                                txtApprovalComments.Text = reader("ApprovalComments").ToString()
                            End If

                            ' 審核意見：非審核人員唯讀
                            If Not canApprove Then
                                txtApprovalComments.ReadOnly = True
                            End If
                        End If
                    End Using
                End Using

                ' 載入明細
                LoadDetailLinesFromDB(docEntry, conn)
                LoadMDRLinesFromDB(docEntry, conn)
            End Using

        Catch ex As Exception
            lblMessage.Text = "載入單據失敗: " & ex.Message
            LogError("LoadDocument", ex)
        End Try
    End Sub
#End Region

#Region "驗證邏輯"
    ''' <summary>
    ''' 驗證基本欄位
    ''' </summary>
    Private Function ValidateBasicFields() As Boolean
        If String.IsNullOrEmpty(ddlCardCode.SelectedValue) Then
            lblMessage.Text = "請選擇供應商"
            Return False
        End If

        If String.IsNullOrEmpty(ddlDeliveryAddr.SelectedValue) Then
            lblMessage.Text = "請選擇收貨地址"
            Return False
        End If

        If String.IsNullOrEmpty(txtDocDate.Text) Then
            lblMessage.Text = "請選擇請款日期"
            Return False
        End If

        If String.IsNullOrEmpty(txtDocDueDate.Text) Then
            lblMessage.Text = "請選擇到期日"
            Return False
        End If

        ' 驗證匯率
        Dim rate As Decimal
        If Not Decimal.TryParse(txtDocRate.Text, rate) OrElse rate <= 0 Then
            lblMessage.Text = "匯率必須為正數"
            Return False
        End If

        Return True
    End Function
#End Region

#Region "輔助方法"
    ''' <summary>
    ''' 取得下拉選單值（避免 DBNull）
    ''' </summary>
    Private Function GetSelectedValue(value As String) As Object
        If String.IsNullOrEmpty(value) Then
            Return DBNull.Value
        Else
            Return value
        End If
    End Function

    ''' <summary>
    ''' 記錄錯誤
    ''' </summary>
    Private Sub LogError(methodName As String, ex As Exception)
        ' TODO: 實作錯誤日誌記錄
        System.Diagnostics.Debug.WriteLine($"[{methodName}] Error: {ex.Message}")
    End Sub
#End Region

#Region "費用明細 GridView"
    ''' <summary>
    ''' GridView 資料來源類別
    ''' </summary>
    <Serializable()>
    Public Class ExpenseDetailLine
        Public Property LineNum As Integer
        Public Property CategoryCode As String
        Public Property CategoryName As String
        Public Property AcctCode As String
        Public Property Description As String
        Public Property Quantity As Decimal
        Public Property Price As Decimal
        Public Property LineTotal As Decimal
        Public Property LineRate As Decimal
        Public Property LineTotalTWD As Decimal
    End Class

    ' 明細資料暫存（使用 ViewState）
    Private Property DetailLines As List(Of ExpenseDetailLine)
        Get
            If ViewState("DetailLines") Is Nothing Then
                ViewState("DetailLines") = New List(Of ExpenseDetailLine)()
            End If
            Return CType(ViewState("DetailLines"), List(Of ExpenseDetailLine))
        End Get
        Set(value As List(Of ExpenseDetailLine))
            ViewState("DetailLines") = value
        End Set
    End Property

    ''' <summary>
    ''' 初始化 GridView（在 Page_Load 的 If Not IsPostBack 中呼叫）
    ''' </summary>
    Private Sub InitializeGridView()
        If currentDocEntry = 0 Then
            ' 新增模式：空白 GridView
            DetailLines = New List(Of ExpenseDetailLine)()
        End If
        BindGridView()
    End Sub

    ''' <summary>
    ''' 綁定 GridView
    ''' </summary>
    Private Sub BindGridView()
        gvExpenseDetail.DataSource = DetailLines
        gvExpenseDetail.DataBind()
        CalculateTotals()
    End Sub

    ''' <summary>
    ''' GridView RowDataBound 事件
    ''' </summary>
    Protected Sub gvExpenseDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim line As ExpenseDetailLine = CType(e.Row.DataItem, ExpenseDetailLine)

            ' 費用類別下拉選單
            Dim ddlCategory As DropDownList = CType(e.Row.FindControl("ddlExpenseCategory"), DropDownList)
            LoadExpenseCategories(ddlCategory)
            If Not String.IsNullOrEmpty(line.CategoryCode) Then
                If ddlCategory.Items.FindByValue(line.CategoryCode) IsNot Nothing Then
                    ddlCategory.SelectedValue = line.CategoryCode
                End If
            End If

            ' 總帳科目
            Dim txtAcctCode As TextBox = CType(e.Row.FindControl("txtAcctCode"), TextBox)
            txtAcctCode.Text = line.AcctCode

            ' 費用說明
            Dim txtDescription As TextBox = CType(e.Row.FindControl("txtDescription"), TextBox)
            txtDescription.Text = line.Description

            ' 數量
            Dim txtQuantity As TextBox = CType(e.Row.FindControl("txtQuantity"), TextBox)
            txtQuantity.Text = line.Quantity.ToString("F2")

            ' 單價
            Dim txtPrice As TextBox = CType(e.Row.FindControl("txtPrice"), TextBox)
            txtPrice.Text = line.Price.ToString("F2")

            ' 金額（外幣）
            Dim lblLineTotal As Label = CType(e.Row.FindControl("lblLineTotal"), Label)
            lblLineTotal.Text = line.LineTotal.ToString("N2")

            ' 匯率
            Dim txtLineRate As TextBox = CType(e.Row.FindControl("txtLineRate"), TextBox)
            txtLineRate.Text = line.LineRate.ToString("F6")

            ' 本幣金額
            Dim lblLineTotalTWD As Label = CType(e.Row.FindControl("lblLineTotalTWD"), Label)
            lblLineTotalTWD.Text = line.LineTotalTWD.ToString("N2")
        End If
    End Sub

    ''' <summary>
    ''' 載入費用類別下拉選單
    ''' </summary>
    Private Sub LoadExpenseCategories(ddl As DropDownList)
        Try
            ddl.Items.Clear()
            ddl.Items.Add(New ListItem("-- 請選擇費用類別 --", ""))

            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT CategoryCode, CategoryName, AcctCode FROM expense_category WHERE Active = 'Y' ORDER BY CategoryCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim code As String = reader("CategoryCode").ToString()
                            Dim name As String = reader("CategoryName").ToString()
                            ddl.Items.Add(New ListItem($"{code} - {name}", code))
                        End While
                    End Using
                End Using
            End Using

        Catch ex As Exception
            LogError("LoadExpenseCategories", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 費用類別選擇變更：自動填入總帳科目
    ''' </summary>
    Protected Sub ddlExpenseCategory_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Dim ddl As DropDownList = CType(sender, DropDownList)
            Dim row As GridViewRow = CType(ddl.NamingContainer, GridViewRow)
            Dim rowIndex As Integer = row.RowIndex

            If String.IsNullOrEmpty(ddl.SelectedValue) Then
                Return
            End If

            ' 查詢對應的總帳科目
            Dim acctCode As String = ""
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT AcctCode FROM expense_category WHERE CategoryCode = @CategoryCode"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@CategoryCode", ddl.SelectedValue)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        acctCode = result.ToString()
                    End If
                End Using
            End Using

            ' 更新 ViewState 和介面
            If rowIndex < DetailLines.Count Then
                DetailLines(rowIndex).CategoryCode = ddl.SelectedValue
                DetailLines(rowIndex).CategoryName = ddl.SelectedItem.Text
                DetailLines(rowIndex).AcctCode = acctCode
            End If

            ' 更新總帳科目欄位
            Dim txtAcctCode As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
            txtAcctCode.Text = acctCode

        Catch ex As Exception
            lblMessage.Text = "更新總帳科目失敗: " & ex.Message
            LogError("ddlExpenseCategory_SelectedIndexChanged", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 數量變更：重新計算金額
    ''' </summary>
    Protected Sub txtQuantity_TextChanged(sender As Object, e As EventArgs)
        RecalculateLineTotals(sender)
    End Sub

    ''' <summary>
    ''' 單價變更：重新計算金額
    ''' </summary>
    Protected Sub txtPrice_TextChanged(sender As Object, e As EventArgs)
        RecalculateLineTotals(sender)
    End Sub

    ''' <summary>
    ''' 重新計算行金額
    ''' </summary>
    Private Sub RecalculateLineTotals(sender As Object)
        Try
            Dim txt As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
            Dim rowIndex As Integer = row.RowIndex

            If rowIndex >= DetailLines.Count Then Return

            ' 取得數量和單價
            Dim txtQuantity As TextBox = CType(row.FindControl("txtQuantity"), TextBox)
            Dim txtPrice As TextBox = CType(row.FindControl("txtPrice"), TextBox)
            Dim txtLineRate As TextBox = CType(row.FindControl("txtLineRate"), TextBox)

            Dim quantity As Decimal = 0
            Dim price As Decimal = 0
            Dim lineRate As Decimal = 1

            Decimal.TryParse(txtQuantity.Text, quantity)
            Decimal.TryParse(txtPrice.Text, price)
            Decimal.TryParse(txtLineRate.Text, lineRate)

            ' 計算金額
            Dim lineTotal As Decimal = quantity * price
            Dim lineTotalTWD As Decimal = lineTotal * lineRate

            ' 更新 ViewState
            DetailLines(rowIndex).Quantity = quantity
            DetailLines(rowIndex).Price = price
            DetailLines(rowIndex).LineTotal = lineTotal
            DetailLines(rowIndex).LineRate = lineRate
            DetailLines(rowIndex).LineTotalTWD = lineTotalTWD

            ' 更新顯示
            Dim lblLineTotal As Label = CType(row.FindControl("lblLineTotal"), Label)
            lblLineTotal.Text = lineTotal.ToString("N2")

            Dim lblLineTotalTWD As Label = CType(row.FindControl("lblLineTotalTWD"), Label)
            lblLineTotalTWD.Text = lineTotalTWD.ToString("N2")

            ' 重新計算總計
            CalculateTotals()

        Catch ex As Exception
            lblMessage.Text = "計算金額失敗: " & ex.Message
            LogError("RecalculateLineTotals", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 計算合計
    ''' </summary>
    Private Sub CalculateTotals()
        Dim totalFC As Decimal = 0
        Dim totalLC As Decimal = 0

        For Each line As ExpenseDetailLine In DetailLines
            totalFC += line.LineTotal
            totalLC += line.LineTotalTWD
        Next

        lblTotalFC.Text = totalFC.ToString("N2")
        lblTotalLC.Text = totalLC.ToString("N2")
        lblCurrency.Text = ddlDocCurrency.SelectedValue
    End Sub

    ''' <summary>
    ''' 新增明細行按鈕
    ''' </summary>
    Protected Sub btnAddLine_Click(sender As Object, e As EventArgs)
        Try
            ' 取得表頭匯率
            Dim docRate As Decimal = 1
            Decimal.TryParse(txtDocRate.Text, docRate)

            ' 新增空白行
            Dim newLine As New ExpenseDetailLine() With {
                .LineNum = DetailLines.Count + 1,
                .CategoryCode = "",
                .CategoryName = "",
                .AcctCode = "",
                .Description = "",
                .Quantity = 1,
                .Price = 0,
                .LineTotal = 0,
                .LineRate = docRate,
                .LineTotalTWD = 0
            }

            DetailLines.Add(newLine)
            BindGridView()

            lblMessage.Text = "已新增明細行"

        Catch ex As Exception
            lblMessage.Text = "新增明細行失敗: " & ex.Message
            LogError("btnAddLine_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 刪除選中行按鈕
    ''' </summary>
    Protected Sub btnDeleteLine_Click(sender As Object, e As EventArgs)
        Try
            Dim deleteCount As Integer = 0

            ' 從後往前刪除（避免索引問題）
            For i As Integer = gvExpenseDetail.Rows.Count - 1 To 0 Step -1
                Dim chkSelect As CheckBox = CType(gvExpenseDetail.Rows(i).FindControl("chkSelect"), CheckBox)
                If chkSelect IsNot Nothing AndAlso chkSelect.Checked Then
                    If i < DetailLines.Count Then
                        DetailLines.RemoveAt(i)
                        deleteCount += 1
                    End If
                End If
            Next

            If deleteCount > 0 Then
                ' 重新編號
                For i As Integer = 0 To DetailLines.Count - 1
                    DetailLines(i).LineNum = i + 1
                Next

                BindGridView()
                lblMessage.Text = $"已刪除 {deleteCount} 行明細"
            Else
                lblMessage.Text = "請先勾選要刪除的明細行"
            End If

        Catch ex As Exception
            lblMessage.Text = "刪除明細行失敗: " & ex.Message
            LogError("btnDeleteLine_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' GridView RowDeleting 事件（單行刪除）
    ''' </summary>
    Protected Sub gvExpenseDetail_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)
        Try
            Dim rowIndex As Integer = e.RowIndex

            If rowIndex < DetailLines.Count Then
                DetailLines.RemoveAt(rowIndex)

                ' 重新編號
                For i As Integer = 0 To DetailLines.Count - 1
                    DetailLines(i).LineNum = i + 1
                Next

                BindGridView()
                lblMessage.Text = "已刪除明細行"
            End If

        Catch ex As Exception
            lblMessage.Text = "刪除明細行失敗: " & ex.Message
            LogError("gvExpenseDetail_RowDeleting", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 儲存前收集 GridView 資料
    ''' </summary>
    Private Sub CollectGridViewData()
        Try
            DetailLines.Clear()

            For Each row As GridViewRow In gvExpenseDetail.Rows
                If row.RowType = DataControlRowType.DataRow Then
                    Dim ddlCategory As DropDownList = CType(row.FindControl("ddlExpenseCategory"), DropDownList)
                    Dim txtAcctCode As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
                    Dim txtDescription As TextBox = CType(row.FindControl("txtDescription"), TextBox)
                    Dim txtQuantity As TextBox = CType(row.FindControl("txtQuantity"), TextBox)
                    Dim txtPrice As TextBox = CType(row.FindControl("txtPrice"), TextBox)
                    Dim txtLineRate As TextBox = CType(row.FindControl("txtLineRate"), TextBox)
                    Dim lblLineTotal As Label = CType(row.FindControl("lblLineTotal"), Label)
                    Dim lblLineTotalTWD As Label = CType(row.FindControl("lblLineTotalTWD"), Label)

                    Dim line As New ExpenseDetailLine() With {
                        .LineNum = row.RowIndex + 1,
                        .CategoryCode = ddlCategory.SelectedValue,
                        .CategoryName = ddlCategory.SelectedItem.Text,
                        .AcctCode = txtAcctCode.Text,
                        .Description = txtDescription.Text,
                        .Quantity = If(Decimal.TryParse(txtQuantity.Text, Nothing), Convert.ToDecimal(txtQuantity.Text), 0),
                        .Price = If(Decimal.TryParse(txtPrice.Text, Nothing), Convert.ToDecimal(txtPrice.Text), 0),
                        .LineTotal = If(Decimal.TryParse(lblLineTotal.Text.Replace(",", ""), Nothing), Convert.ToDecimal(lblLineTotal.Text.Replace(",", "")), 0),
                        .LineRate = If(Decimal.TryParse(txtLineRate.Text, Nothing), Convert.ToDecimal(txtLineRate.Text), 1),
                        .LineTotalTWD = If(Decimal.TryParse(lblLineTotalTWD.Text.Replace(",", ""), Nothing), Convert.ToDecimal(lblLineTotalTWD.Text.Replace(",", "")), 0)
                    }

                    DetailLines.Add(line)
                End If
            Next

        Catch ex As Exception
            LogError("CollectGridViewData", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 驗證明細資料
    ''' </summary>
    Private Function ValidateDetailLines() As Boolean
        If DetailLines.Count = 0 Then
            lblMessage.Text = "請至少新增一筆費用明細"
            Return False
        End If

        For i As Integer = 0 To DetailLines.Count - 1
            Dim line As ExpenseDetailLine = DetailLines(i)

            If String.IsNullOrEmpty(line.CategoryCode) Then
                lblMessage.Text = $"第 {i + 1} 行：請選擇費用類別"
                Return False
            End If

            If String.IsNullOrEmpty(line.Description) Then
                lblMessage.Text = $"第 {i + 1} 行：請輸入費用說明"
                Return False
            End If

            If line.Quantity <= 0 Then
                lblMessage.Text = $"第 {i + 1} 行：數量必須大於 0"
                Return False
            End If

            If line.Price < 0 Then
                lblMessage.Text = $"第 {i + 1} 行：單價不可為負數"
                Return False
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' 儲存明細到資料庫（在 InsertNewDocument 和 UpdateDocument 中呼叫）
    ''' </summary>
    Private Sub SaveDetailLinesToDB(docEntry As Integer, conn As SqlConnection, trans As SqlTransaction)
        Try
            ' 先刪除舊明細（更新模式）
            Dim sqlDelete As String = "DELETE FROM jOPC1 WHERE DocEntry = @DocEntry"
            Using cmd As New SqlCommand(sqlDelete, conn, trans)
                cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                cmd.ExecuteNonQuery()
            End Using

            ' 插入新明細
            Dim sqlInsert As String = "INSERT INTO jOPC1 (DocEntry, LineNum, CategoryCode, CategoryName, AcctCode, " &
                                     "LineDescription, Quantity, Price, LineTotal, LineRate, LineTotalTWD) " &
                                     "VALUES (@DocEntry, @LineNum, @CategoryCode, @CategoryName, @AcctCode, " &
                                     "@LineDescription, @Quantity, @Price, @LineTotal, @LineRate, @LineTotalTWD)"

            For Each line As ExpenseDetailLine In DetailLines
                Using cmd As New SqlCommand(sqlInsert, conn, trans)
                    cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                    cmd.Parameters.AddWithValue("@LineNum", line.LineNum)
                    cmd.Parameters.AddWithValue("@CategoryCode", line.CategoryCode)
                    cmd.Parameters.AddWithValue("@CategoryName", line.CategoryName)
                    cmd.Parameters.AddWithValue("@AcctCode", line.AcctCode)
                    cmd.Parameters.AddWithValue("@LineDescription", line.Description)
                    cmd.Parameters.AddWithValue("@Quantity", line.Quantity)
                    cmd.Parameters.AddWithValue("@Price", line.Price)
                    cmd.Parameters.AddWithValue("@LineTotal", line.LineTotal)
                    cmd.Parameters.AddWithValue("@LineRate", line.LineRate)
                    cmd.Parameters.AddWithValue("@LineTotalTWD", line.LineTotalTWD)
                    cmd.ExecuteNonQuery()
                End Using
            Next

        Catch ex As Exception
            LogError("SaveDetailLinesToDB", ex)
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 從資料庫載入明細（在 LoadDocument 中呼叫）
    ''' </summary>
    Private Sub LoadDetailLinesFromDB(docEntry As Integer, conn As SqlConnection)
        Try
            DetailLines.Clear()

            Dim sql As String = "SELECT * FROM jOPC1 WHERE DocEntry = @DocEntry ORDER BY LineNum"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim line As New ExpenseDetailLine() With {
                            .LineNum = Convert.ToInt32(reader("LineNum")),
                            .CategoryCode = If(IsDBNull(reader("CategoryCode")), "", reader("CategoryCode").ToString()),
                            .CategoryName = If(IsDBNull(reader("CategoryName")), "", reader("CategoryName").ToString()),
                            .AcctCode = If(IsDBNull(reader("AcctCode")), "", reader("AcctCode").ToString()),
                            .Description = If(IsDBNull(reader("LineDescription")), "", reader("LineDescription").ToString()),
                            .Quantity = If(IsDBNull(reader("Quantity")), 0, Convert.ToDecimal(reader("Quantity"))),
                            .Price = If(IsDBNull(reader("Price")), 0, Convert.ToDecimal(reader("Price"))),
                            .LineTotal = If(IsDBNull(reader("LineTotal")), 0, Convert.ToDecimal(reader("LineTotal"))),
                            .LineRate = If(IsDBNull(reader("LineRate")), 1, Convert.ToDecimal(reader("LineRate"))),
                            .LineTotalTWD = If(IsDBNull(reader("LineTotalTWD")), 0, Convert.ToDecimal(reader("LineTotalTWD")))
                        }
                        DetailLines.Add(line)
                    End While
                End Using
            End Using

            BindGridView()

        Catch ex As Exception
            LogError("LoadDetailLinesFromDB", ex)
        End Try
    End Sub
#End Region

#Region "MDR 發票明細 Tab"
    ''' <summary>
    ''' MDR 明細資料類別
    ''' </summary>
    <Serializable()>
    Public Class MDRDetailLine
        Public Property LineNum As Integer
        Public Property InvoiceNum As String
        Public Property InvoiceDate As DateTime
        Public Property InvoiceAmount As Decimal      ' 未稅
        Public Property TaxAmount As Decimal
        Public Property InvoiceTotal As Decimal       ' 含稅
        Public Property Remarks As String
    End Class

    ' MDR 明細資料暫存（使用 ViewState）
    Private Property MDRLines As List(Of MDRDetailLine)
        Get
            If ViewState("MDRLines") Is Nothing Then
                ViewState("MDRLines") = New List(Of MDRDetailLine)()
            End If
            Return CType(ViewState("MDRLines"), List(Of MDRDetailLine))
        End Get
        Set(value As List(Of MDRDetailLine))
            ViewState("MDRLines") = value
        End Set
    End Property

    ''' <summary>
    ''' 初始化 MDR GridView（在 Page_Load 的 If Not IsPostBack 中呼叫）
    ''' </summary>
    Private Sub InitializeMDRGridView()
        If currentDocEntry = 0 Then
            ' 新增模式：空白 GridView
            MDRLines = New List(Of MDRDetailLine)()
        End If
        BindMDRGridView()
        SyncMDRHeaderInfo()
    End Sub

    ''' <summary>
    ''' 綁定 MDR GridView
    ''' </summary>
    Private Sub BindMDRGridView()
        gvMDRDetail.DataSource = MDRLines
        gvMDRDetail.DataBind()
        CalculateMDRTotals()
    End Sub

    ''' <summary>
    ''' 同步 MDR 表頭資訊（從費用申請 Tab）
    ''' </summary>
    Private Sub SyncMDRHeaderInfo()
        lblMDR_CardCode.Text = ddlCardCode.SelectedValue
        lblMDR_CardName.Text = txtCardName.Text
        lblMDR_DocDate.Text = txtDocDate.Text
        lblMDR_DocDueDate.Text = txtDocDueDate.Text
        lblMDR_DocCurrency.Text = ddlDocCurrency.SelectedValue
        lblMDR_DocRate.Text = txtDocRate.Text
    End Sub

    ''' <summary>
    ''' MDR GridView RowDataBound 事件
    ''' </summary>
    Protected Sub gvMDRDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim line As MDRDetailLine = CType(e.Row.DataItem, MDRDetailLine)

            ' 發票號碼
            Dim txtInvoiceNum As TextBox = CType(e.Row.FindControl("txtInvoiceNum"), TextBox)
            txtInvoiceNum.Text = line.InvoiceNum

            ' 發票日期
            Dim txtInvoiceDate As TextBox = CType(e.Row.FindControl("txtInvoiceDate"), TextBox)
            If line.InvoiceDate > DateTime.MinValue Then
                txtInvoiceDate.Text = line.InvoiceDate.ToString("yyyy-MM-dd")
            End If

            ' 發票金額（未稅）
            Dim txtInvoiceAmount As TextBox = CType(e.Row.FindControl("txtInvoiceAmount"), TextBox)
            txtInvoiceAmount.Text = line.InvoiceAmount.ToString("F2")

            ' 稅額
            Dim txtTaxAmount As TextBox = CType(e.Row.FindControl("txtTaxAmount"), TextBox)
            txtTaxAmount.Text = line.TaxAmount.ToString("F2")

            ' 發票總額（含稅）
            Dim lblInvoiceTotal As Label = CType(e.Row.FindControl("lblInvoiceTotal"), Label)
            lblInvoiceTotal.Text = line.InvoiceTotal.ToString("N2")

            ' 備註
            Dim txtMDRRemarks As TextBox = CType(e.Row.FindControl("txtMDRRemarks"), TextBox)
            txtMDRRemarks.Text = line.Remarks
        End If
    End Sub

    ''' <summary>
    ''' 發票金額變更：重新計算發票總額
    ''' </summary>
    Protected Sub txtInvoiceAmount_TextChanged(sender As Object, e As EventArgs)
        RecalculateMDRLineTotal(sender)
    End Sub

    ''' <summary>
    ''' 稅額變更：重新計算發票總額
    ''' </summary>
    Protected Sub txtTaxAmount_TextChanged(sender As Object, e As EventArgs)
        RecalculateMDRLineTotal(sender)
    End Sub

    ''' <summary>
    ''' 重新計算 MDR 行總額
    ''' </summary>
    Private Sub RecalculateMDRLineTotal(sender As Object)
        Try
            Dim txt As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
            Dim rowIndex As Integer = row.RowIndex

            If rowIndex >= MDRLines.Count Then Return

            ' 取得發票金額和稅額
            Dim txtInvoiceAmount As TextBox = CType(row.FindControl("txtInvoiceAmount"), TextBox)
            Dim txtTaxAmount As TextBox = CType(row.FindControl("txtTaxAmount"), TextBox)

            Dim invoiceAmount As Decimal = 0
            Dim taxAmount As Decimal = 0

            Decimal.TryParse(txtInvoiceAmount.Text, invoiceAmount)
            Decimal.TryParse(txtTaxAmount.Text, taxAmount)

            ' 計算發票總額（含稅）
            Dim invoiceTotal As Decimal = invoiceAmount + taxAmount

            ' 更新 ViewState
            MDRLines(rowIndex).InvoiceAmount = invoiceAmount
            MDRLines(rowIndex).TaxAmount = taxAmount
            MDRLines(rowIndex).InvoiceTotal = invoiceTotal

            ' 更新顯示
            Dim lblInvoiceTotal As Label = CType(row.FindControl("lblInvoiceTotal"), Label)
            lblInvoiceTotal.Text = invoiceTotal.ToString("N2")

            ' 重新計算總計
            CalculateMDRTotals()

        Catch ex As Exception
            lblMessage.Text = "計算發票金額失敗: " & ex.Message
            LogError("RecalculateMDRLineTotal", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 計算 MDR 合計
    ''' </summary>
    Private Sub CalculateMDRTotals()
        Dim totalAmount As Decimal = 0
        Dim totalTax As Decimal = 0
        Dim grandTotal As Decimal = 0

        For Each line As MDRDetailLine In MDRLines
            totalAmount += line.InvoiceAmount
            totalTax += line.TaxAmount
            grandTotal += line.InvoiceTotal
        Next

        lblMDR_TotalAmount.Text = totalAmount.ToString("N2")
        lblMDR_TotalTax.Text = totalTax.ToString("N2")
        lblMDR_GrandTotal.Text = grandTotal.ToString("N2")
    End Sub

    ''' <summary>
    ''' 新增 MDR 明細行按鈕
    ''' </summary>
    Protected Sub btnMDR_AddLine_Click(sender As Object, e As EventArgs)
        Try
            ' 新增空白發票明細行
            Dim newLine As New MDRDetailLine() With {
                .LineNum = MDRLines.Count + 1,
                .InvoiceNum = "",
                .InvoiceDate = DateTime.Now,
                .InvoiceAmount = 0,
                .TaxAmount = 0,
                .InvoiceTotal = 0,
                .Remarks = ""
            }

            MDRLines.Add(newLine)
            BindMDRGridView()

            lblMessage.Text = "已新增發票明細行"

        Catch ex As Exception
            lblMessage.Text = "新增發票明細失敗: " & ex.Message
            LogError("btnMDR_AddLine_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 刪除選中 MDR 明細行按鈕
    ''' </summary>
    Protected Sub btnMDR_DeleteLine_Click(sender As Object, e As EventArgs)
        Try
            Dim deleteCount As Integer = 0

            ' 從後往前刪除（避免索引問題）
            For i As Integer = gvMDRDetail.Rows.Count - 1 To 0 Step -1
                Dim chkSelect As CheckBox = CType(gvMDRDetail.Rows(i).FindControl("chkMDRSelect"), CheckBox)
                If chkSelect IsNot Nothing AndAlso chkSelect.Checked Then
                    If i < MDRLines.Count Then
                        MDRLines.RemoveAt(i)
                        deleteCount += 1
                    End If
                End If
            Next

            If deleteCount > 0 Then
                ' 重新編號
                For i As Integer = 0 To MDRLines.Count - 1
                    MDRLines(i).LineNum = i + 1
                Next

                BindMDRGridView()
                lblMessage.Text = $"已刪除 {deleteCount} 行發票明細"
            Else
                lblMessage.Text = "請先勾選要刪除的發票明細"
            End If

        Catch ex As Exception
            lblMessage.Text = "刪除發票明細失敗: " & ex.Message
            LogError("btnMDR_DeleteLine_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' MDR GridView RowDeleting 事件（單行刪除）
    ''' </summary>
    Protected Sub gvMDRDetail_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)
        Try
            Dim rowIndex As Integer = e.RowIndex

            If rowIndex < MDRLines.Count Then
                MDRLines.RemoveAt(rowIndex)

                ' 重新編號
                For i As Integer = 0 To MDRLines.Count - 1
                    MDRLines(i).LineNum = i + 1
                Next

                BindMDRGridView()
                lblMessage.Text = "已刪除發票明細"
            End If

        Catch ex As Exception
            lblMessage.Text = "刪除發票明細失敗: " & ex.Message
            LogError("gvMDRDetail_RowDeleting", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 驗證金額總和按鈕
    ''' </summary>
    Protected Sub btnMDR_ValidateSum_Click(sender As Object, e As EventArgs)
        Try
            ' 收集 MDR GridView 最新資料
            CollectMDRGridViewData()

            ' 計算 AP 發票總額（來自費用明細）
            Dim apTotal As Decimal = 0
            If Decimal.TryParse(lblTotalLC.Text.Replace(",", ""), apTotal) Then
                ' AP 總額已取得
            Else
                apTotal = 0
            End If

            ' 計算 MDR 發票總額（含稅）
            Dim mdrTotal As Decimal = 0
            If Decimal.TryParse(lblMDR_GrandTotal.Text.Replace(",", ""), mdrTotal) Then
                ' MDR 總額已取得
            Else
                mdrTotal = 0
            End If

            ' 驗證金額是否相等
            Dim difference As Decimal = Math.Abs(apTotal - mdrTotal)

            pnlMDR_ValidationResult.Visible = True

            If difference < 0.01D Then
                ' 驗證成功
                pnlMDR_ValidationResult.CssClass = "validation-success"
                lblMDR_ValidationMessage.Text = $"✓ 驗證通過！AP 發票總額（{apTotal:N2}）與 MDR 發票總額（{mdrTotal:N2}）相符。"
            Else
                ' 驗證失敗
                pnlMDR_ValidationResult.CssClass = "validation-error"
                lblMDR_ValidationMessage.Text = $"✗ 驗證失敗！AP 發票總額（{apTotal:N2}）與 MDR 發票總額（{mdrTotal:N2}）不符，差異：{difference:N2}"
            End If

        Catch ex As Exception
            lblMessage.Text = "驗證金額失敗: " & ex.Message
            LogError("btnMDR_ValidateSum_Click", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 儲存前收集 MDR GridView 資料
    ''' </summary>
    Private Sub CollectMDRGridViewData()
        Try
            MDRLines.Clear()

            For Each row As GridViewRow In gvMDRDetail.Rows
                If row.RowType = DataControlRowType.DataRow Then
                    Dim txtInvoiceNum As TextBox = CType(row.FindControl("txtInvoiceNum"), TextBox)
                    Dim txtInvoiceDate As TextBox = CType(row.FindControl("txtInvoiceDate"), TextBox)
                    Dim txtInvoiceAmount As TextBox = CType(row.FindControl("txtInvoiceAmount"), TextBox)
                    Dim txtTaxAmount As TextBox = CType(row.FindControl("txtTaxAmount"), TextBox)
                    Dim lblInvoiceTotal As Label = CType(row.FindControl("lblInvoiceTotal"), Label)
                    Dim txtMDRRemarks As TextBox = CType(row.FindControl("txtMDRRemarks"), TextBox)

                    Dim invoiceDate As DateTime = DateTime.Now
                    If Not String.IsNullOrEmpty(txtInvoiceDate.Text) Then
                        DateTime.TryParse(txtInvoiceDate.Text, invoiceDate)
                    End If

                    Dim line As New MDRDetailLine() With {
                        .LineNum = row.RowIndex + 1,
                        .InvoiceNum = txtInvoiceNum.Text,
                        .InvoiceDate = invoiceDate,
                        .InvoiceAmount = If(Decimal.TryParse(txtInvoiceAmount.Text, Nothing), Convert.ToDecimal(txtInvoiceAmount.Text), 0),
                        .TaxAmount = If(Decimal.TryParse(txtTaxAmount.Text, Nothing), Convert.ToDecimal(txtTaxAmount.Text), 0),
                        .InvoiceTotal = If(Decimal.TryParse(lblInvoiceTotal.Text.Replace(",", ""), Nothing), Convert.ToDecimal(lblInvoiceTotal.Text.Replace(",", "")), 0),
                        .Remarks = txtMDRRemarks.Text
                    }

                    MDRLines.Add(line)
                End If
            Next

        Catch ex As Exception
            LogError("CollectMDRGridViewData", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 驗證 MDR 明細資料
    ''' </summary>
    Private Function ValidateMDRLines() As Boolean
        ' MDR 明細可以為空（視業務需求）
        If MDRLines.Count = 0 Then
            Return True  ' 允許不輸入 MDR 明細
        End If

        For i As Integer = 0 To MDRLines.Count - 1
            Dim line As MDRDetailLine = MDRLines(i)

            If String.IsNullOrEmpty(line.InvoiceNum) Then
                lblMessage.Text = $"MDR 第 {i + 1} 行：請輸入發票號碼"
                Return False
            End If

            If line.InvoiceDate = DateTime.MinValue Then
                lblMessage.Text = $"MDR 第 {i + 1} 行：請輸入發票日期"
                Return False
            End If

            If line.InvoiceAmount < 0 Then
                lblMessage.Text = $"MDR 第 {i + 1} 行：發票金額不可為負數"
                Return False
            End If

            If line.TaxAmount < 0 Then
                lblMessage.Text = $"MDR 第 {i + 1} 行：稅額不可為負數"
                Return False
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' 儲存 MDR 明細到資料庫（在 InsertNewDocument 和 UpdateDocument 中呼叫）
    ''' </summary>
    Private Sub SaveMDRLinesToDB(docEntry As Integer, conn As SqlConnection, trans As SqlTransaction)
        Try
            ' 先刪除舊 MDR 明細（更新模式）
            Dim sqlDelete As String = "DELETE FROM jMDR1 WHERE DocEntry = @DocEntry"
            Using cmd As New SqlCommand(sqlDelete, conn, trans)
                cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                cmd.ExecuteNonQuery()
            End Using

            ' 插入新 MDR 明細
            If MDRLines.Count > 0 Then
                Dim sqlInsert As String = "INSERT INTO jMDR1 (DocEntry, LineNum, InvoiceNum, InvoiceDate, " &
                                         "InvoiceAmount, TaxAmount, InvoiceTotal, Remarks) " &
                                         "VALUES (@DocEntry, @LineNum, @InvoiceNum, @InvoiceDate, " &
                                         "@InvoiceAmount, @TaxAmount, @InvoiceTotal, @Remarks)"

                For Each line As MDRDetailLine In MDRLines
                    Using cmd As New SqlCommand(sqlInsert, conn, trans)
                        cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                        cmd.Parameters.AddWithValue("@LineNum", line.LineNum)
                        cmd.Parameters.AddWithValue("@InvoiceNum", line.InvoiceNum)
                        cmd.Parameters.AddWithValue("@InvoiceDate", line.InvoiceDate)
                        cmd.Parameters.AddWithValue("@InvoiceAmount", line.InvoiceAmount)
                        cmd.Parameters.AddWithValue("@TaxAmount", line.TaxAmount)
                        cmd.Parameters.AddWithValue("@InvoiceTotal", line.InvoiceTotal)
                        cmd.Parameters.AddWithValue("@Remarks", If(String.IsNullOrEmpty(line.Remarks), DBNull.Value, line.Remarks))
                        cmd.ExecuteNonQuery()
                    End Using
                Next
            End If

        Catch ex As Exception
            LogError("SaveMDRLinesToDB", ex)
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 從資料庫載入 MDR 明細（在 LoadDocument 中呼叫）
    ''' </summary>
    Private Sub LoadMDRLinesFromDB(docEntry As Integer, conn As SqlConnection)
        Try
            MDRLines.Clear()

            Dim sql As String = "SELECT * FROM jMDR1 WHERE DocEntry = @DocEntry ORDER BY LineNum"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim line As New MDRDetailLine() With {
                            .LineNum = Convert.ToInt32(reader("LineNum")),
                            .InvoiceNum = If(IsDBNull(reader("InvoiceNum")), "", reader("InvoiceNum").ToString()),
                            .InvoiceDate = If(IsDBNull(reader("InvoiceDate")), DateTime.Now, Convert.ToDateTime(reader("InvoiceDate"))),
                            .InvoiceAmount = If(IsDBNull(reader("InvoiceAmount")), 0, Convert.ToDecimal(reader("InvoiceAmount"))),
                            .TaxAmount = If(IsDBNull(reader("TaxAmount")), 0, Convert.ToDecimal(reader("TaxAmount"))),
                            .InvoiceTotal = If(IsDBNull(reader("InvoiceTotal")), 0, Convert.ToDecimal(reader("InvoiceTotal"))),
                            .Remarks = If(IsDBNull(reader("Remarks")), "", reader("Remarks").ToString())
                        }
                        MDRLines.Add(line)
                    End While
                End Using
            End Using

            BindMDRGridView()
            SyncMDRHeaderInfo()

        Catch ex As Exception
            LogError("LoadMDRLinesFromDB", ex)
        End Try
    End Sub
#End Region

End Class
