Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Configuration
Imports System.Web.UI.HtmlControls
Imports SAPbobsCOM

''' <summary>
''' 費用申請單 (Expense Claim Form)
''' 重構版本: 2025-11-26 (v2.0)
''' 對應規格: shopfloor/Claude_TMP/ExpenseClaimForm_Spec_Integration_20251126.md
''' </summary>
Partial Public Class ExpenseClaimForm
    Inherits System.Web.UI.Page

#Region "類別定義"
    <Serializable()>
    Public Class ExpenseLine
        Public Property LineNum As Integer
        Public Property CategoryCode As String
        Public Property Description As String
        Public Property AcctCode As String
        Public Property LineTotal As Decimal ' 未稅金額
        Public Property VatGroup As String
        Public Property VatRate As Decimal
        Public Property VatSum As Decimal ' 稅額
        Public Property PriceAfterVat As Decimal ' 含稅金額
        Public Property CostingCode As String
        Public Property CostingCode2 As String
        Public Property UserSelectedAcct As Boolean
        ' 顯示用
        Public Property Currency As String
        Public Property Rate As Decimal
    End Class

    <Serializable()>
    Public Class MDRLine
        Public Property LineNum As Integer
        Public Property U_LIFNR As String
        Public Property U_STCEG As String
        Public Property U_XBLNR As String
        Public Property U_ZFORM_CODE As String
        Public Property U_BLDAT As DateTime?
        Public Property U_VATDATE As DateTime?
        Public Property U_HWBAS As Decimal ' 未稅金額
        Public Property U_HWSTE As Decimal ' 稅額
        Public Property U_TAX_TYPE As String
    End Class

    <Serializable()>
    Public Class AttachmentItem
        Public Property ID As Integer
        Public Property FileName As String
        Public Property FilePath As String
        Public Property UploadDate As DateTime
        Public Property UploadTime As String
        Public Property Uploader As String
    End Class
#End Region

#Region "變數宣告"
    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private ReadOnly sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
    Public CommUtil As New CommUtil

    ' SAP DI API Company 物件 (Page 層級，參考廠長的做法)
    Public oCompany As New SAPbobsCOM.Company

    Private currentUserId As String = ""
    Private currentJID As Integer = 0
    Private canApprove As Boolean = False
    Private isApUser As Boolean = False ' AP_App 權限
#End Region

#Region "屬性 (ViewState)"
    Private Property CurrentLines As List(Of ExpenseLine)
        Get
            If ViewState("CurrentLines") Is Nothing Then
                ViewState("CurrentLines") = New List(Of ExpenseLine)()
            End If
            Return CType(ViewState("CurrentLines"), List(Of ExpenseLine))
        End Get
        Set(value As List(Of ExpenseLine))
            ViewState("CurrentLines") = value
        End Set
    End Property

    Private Property CurrentMDRLines As List(Of MDRLine)
        Get
            If ViewState("CurrentMDRLines") Is Nothing Then
                ViewState("CurrentMDRLines") = New List(Of MDRLine)()
            End If
            Return CType(ViewState("CurrentMDRLines"), List(Of MDRLine))
        End Get
        Set(value As List(Of MDRLine))
            ViewState("CurrentMDRLines") = value
        End Set
    End Property

    Private Property CurrentAttachments As List(Of AttachmentItem)
        Get
            If ViewState("CurrentAttachments") Is Nothing Then
                ViewState("CurrentAttachments") = New List(Of AttachmentItem)()
            End If
            Return CType(ViewState("CurrentAttachments"), List(Of AttachmentItem))
        End Get
        Set(value As List(Of AttachmentItem))
            ViewState("CurrentAttachments") = value
        End Set
    End Property

    Private Property AcctSearchData As DataTable
        Get
            Return CType(ViewState("AcctSearchData"), DataTable)
        End Get
        Set(value As DataTable)
            ViewState("AcctSearchData") = value
        End Set
    End Property

    Private Property AcctSearchSortExpression As String
        Get
            Return If(TryCast(ViewState("AcctSearchSortExpression"), String), "")
        End Get
        Set(value As String)
            ViewState("AcctSearchSortExpression") = value
        End Set
    End Property

    Private Property AcctSearchSortDirection As String
        Get
            Return If(TryCast(ViewState("AcctSearchSortDirection"), String), "ASC")
        End Get
        Set(value As String)
            ViewState("AcctSearchSortDirection") = value
        End Set
    End Property

    Private Property CopyFromJID As Integer
        Get
            If ViewState("CopyFromJID") Is Nothing Then Return 0
            Return Convert.ToInt32(ViewState("CopyFromJID"))
        End Get
        Set(value As Integer)
            ViewState("CopyFromJID") = value
        End Set
    End Property

    Private Property CopyAttachments As Boolean
        Get
            If ViewState("CopyAttachments") Is Nothing Then Return False
            Return Convert.ToBoolean(ViewState("CopyAttachments"))
        End Get
        Set(value As Boolean)
            ViewState("CopyAttachments") = value
        End Set
    End Property

    Private Property CopyMDR As Boolean
        Get
            If ViewState("CopyMDR") Is Nothing Then Return False
            Return Convert.ToBoolean(ViewState("CopyMDR"))
        End Get
        Set(value As Boolean)
            ViewState("CopyMDR") = value
        End Set
    End Property
#End Region

#Region "頁面載入"
    Protected Sub Page_PreRender(sender As Object, e As EventArgs) Handles Me.PreRender
        ' 根據 hfActiveTab 設定頁籤狀態，確保 Postback (如新增明細) 後 HTML 狀態正確
        ' 這能避免 UpdatePanel 更新後頁籤跳回預設值的問題
        Dim activeTab As String = hfActiveTab.Value

        ' Reset CSS
        btnTabExpense.Attributes("class") = "tab-button"
        divContentExpense.Attributes("class") = "tab-content"
        divContentExpense.Style("display") = "none"
        
        btnTabMDR.Attributes("class") = "tab-button"
        divContentMDR.Attributes("class") = "tab-content"
        divContentMDR.Style("display") = "none"

        If activeTab = "mdr" Then
            btnTabMDR.Attributes("class") += " active"
            divContentMDR.Attributes("class") += " active"
            divContentMDR.Style("display") = "block"
        Else
            ' Default to expense
            btnTabExpense.Attributes("class") += " active"
            divContentExpense.Attributes("class") += " active"
            divContentExpense.Style("display") = "block"
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Session("s_id") Is Nothing Then
                Response.Redirect("~/usermgm/login.aspx")
                Return
            End If
            currentUserId = Session("s_id").ToString()
            Dim headerUser As Label = TryCast(FindControlRecursive(Me, "lblCurrentUser"), Label)
            If headerUser IsNot Nothing Then
                headerUser.Text = currentUserId
            End If

            CheckApprovalPermission()

            ' 支援 jID 和 DocEntry 兩種參數名稱（向下相容）
            If Request.QueryString("jID") IsNot Nothing Then
                Integer.TryParse(Request.QueryString("jID"), currentJID)
            ElseIf Request.QueryString("DocEntry") IsNot Nothing Then
                Integer.TryParse(Request.QueryString("DocEntry"), currentJID)
            End If
            Dim copyFromId As Integer = 0
            If Request.QueryString("CopyFrom") IsNot Nothing Then
                Integer.TryParse(Request.QueryString("CopyFrom"), copyFromId)
                currentJID = 0
            End If

            If Not IsPostBack Then
                InitializeDropDowns()

                ' 檢查使用者費用部門是否已設定
                CheckUserExpDept()

                If copyFromId > 0 Then
                    LoadDocument(copyFromId)
                    Dim copyAttach As Boolean = (Request.QueryString("CopyAttach") = "1")
                    Dim copyMdr As Boolean = (Request.QueryString("CopyMDR") = "1")
                    ApplyCopyOverrides(copyFromId, copyAttach, copyMdr)
                ElseIf currentJID > 0 Then
                    LoadDocument(currentJID)
                Else
                    SetDefaultValues()
                    InitializeGridViews()
                End If
            End If

        Catch ex As Exception
            ShowError("頁面載入錯誤: " & ex.Message)
        End Try
    End Sub

    Protected Sub lnkLogout_Click(sender As Object, e As EventArgs)
        Response.Redirect("~/usermgm/logout.aspx")
    End Sub

    Private Sub CheckApprovalPermission()
        Using conn As New SqlConnection(connStr)
            conn.Open()
            ' 檢查使用者是否有審核權限 (Approver) 與 AP作業權限 (AP_App)
            Dim sql As String = "SELECT Approver, AP_App FROM [User] WHERE id = @UserId"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@UserId", currentUserId)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        canApprove = (Convert.ToInt32(dr("Approver")) = 1)
                        isApUser = (Convert.ToInt32(dr("AP_App")) = 1)
                    End If
                End Using
            End Using
        End Using

        ' 預設不隱藏，由 LoadDocument 或 Page_Load 決定顯示狀態
    End Sub

    ''' <summary>
    ''' 檢查使用者費用部門是否已設定，若未設定則彈出選擇視窗
    ''' </summary>
    Private Sub CheckUserExpDept()
        Dim userExpDept As String = ""

        ' 取得使用者目前的 expDept
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Dim sql As String = "SELECT expDEPT FROM [User] WHERE id = @UserId"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@UserId", currentUserId)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    userExpDept = result.ToString().Trim()
                End If
            End Using
        End Using

        ' 檢查 expDept 是否存在於 jDEPT 中
        Dim isValidDept As Boolean = False
        If Not String.IsNullOrEmpty(userExpDept) Then
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT COUNT(*) FROM jDEPT WHERE EDeptID = @EDeptID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@EDeptID", userExpDept)
                    isValidDept = (Convert.ToInt32(cmd.ExecuteScalar()) > 0)
                End Using
            End Using
        End If

        ' 若未設定或不存在，載入部門選項並顯示彈窗
        If Not isValidDept Then
            LoadExpDeptDropDown()
            mpeExpDept.Show()
        End If
    End Sub

    ''' <summary>
    ''' 載入費用部門下拉選單
    ''' </summary>
    Private Sub LoadExpDeptDropDown()
        ddlExpDeptSelect.Items.Clear()
        ddlExpDeptSelect.Items.Add(New ListItem("- 請選擇 -", ""))
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT EDeptID, EDeptName FROM jDEPT ORDER BY EDeptID"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddlExpDeptSelect.Items.Add(New ListItem(dr("EDeptName").ToString(), dr("EDeptID").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入費用部門失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 費用部門選擇確認按鈕事件
    ''' </summary>
    Protected Sub btnExpDeptConfirm_Click(sender As Object, e As EventArgs)
        If String.IsNullOrEmpty(ddlExpDeptSelect.SelectedValue) Then
            ' 若未選擇，重新顯示彈窗
            mpeExpDept.Show()
            Return
        End If

        ' 更新使用者的 expDept
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE [User] SET expDEPT = @ExpDept WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ExpDept", ddlExpDeptSelect.SelectedValue)
                    cmd.Parameters.AddWithValue("@UserId", currentUserId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            mpeExpDept.Hide()
            lblMessage.Text = "費用部門設定成功！"
            lblMessage.ForeColor = System.Drawing.Color.Green
        Catch ex As Exception
            ShowError("更新費用部門失敗: " & ex.Message)
            mpeExpDept.Show()
        End Try
    End Sub

    Private Sub SetDefaultValues()
        lblDocNum.Text = "[新單據]"
        txtOwner.Text = currentUserId
        ' If ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedValue = currentUserId
        If ddlPurchaser.SelectedIndex = -1 AndAlso ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedIndex = 0
        ' [B] 新單據直接顯示為「新增中」，送出後即為待審核
        lblDocStatus.Text = "新增中"
        lblDocStatus.CssClass = "badge status-W"
        txtStatusDisplay.Text = "新增中"

        Dim today As String = DateTime.Now.ToString("yyyy-MM-dd")
        txtDocDate.Text = today      ' 過帳日期預設今天
        txtTaxDate.Text = today      ' 文件日期預設今天
        txtDocDueDate.Text = ""      ' 到期日預設空白

        If ddlDocCurrency.Items.FindByValue("NTD") IsNot Nothing Then
            ddlDocCurrency.SelectedValue = "NTD"
        End If
        txtDocRate.Text = "1.0"

        ' 新增單據時，審核區塊預設唯讀且停用按鈕 (但保持顯示以確認版面)
        txtApprovalComments.ReadOnly = True

        btnApprove.Visible = True
        btnApprove.Enabled = False

        btnUpdateComment.Visible = True
        btnUpdateComment.Enabled = False

        btnReject.Visible = True
        btnReject.Enabled = False

        ' [B] 按鈕狀態 (新增模式) - 移除暫存草稿按鈕
        btnSave.Visible = False
        btnDelete.Visible = False
    End Sub
#End Region

#Region "初始化資料"
    Private Sub InitializeDropDowns()
        LoadDeliveryAddress()
        LoadCurrencies()
        LoadPaymentGroups()
        LoadPurchasers()
    End Sub

    Private Sub InitializeGridViews()
        If CurrentLines.Count = 0 Then
            ' 預設加入一筆空行
            AddNewEmptyLine()
        End If
        BindGrid()

        If CurrentMDRLines.Count = 0 Then
            ' MDR 預設不加空行，由使用者手動新增
        End If
        BindMDRGrid()
        
        BindAttachmentGrid()
    End Sub

    Private Sub LoadDeliveryAddress()
        ddlDeliveryAddr.Items.Clear()
        ddlDeliveryAddr.Items.Add(New ListItem("- 請選擇 -", ""))
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT ID, addrName, address FROM addr WHERE addrType='R' AND active='Y'"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            Dim item As New ListItem(dr("addrName").ToString(), dr("ID").ToString())
                            item.Attributes.Add("data-address", dr("address").ToString())
                            ddlDeliveryAddr.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入地址失敗: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadCurrencies()
        ddlDocCurrency.Items.Clear()
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT CurrCode FROM OCRN"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddlDocCurrency.Items.Add(New ListItem(dr("CurrCode").ToString(), dr("CurrCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入幣別失敗: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadPaymentGroups()
        ddlGroupNum.Items.Clear()
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT GroupNum, PymntGroup FROM OCTG"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddlGroupNum.Items.Add(New ListItem(dr("PymntGroup").ToString(), dr("GroupNum").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入付款條件失敗: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadPurchasers()
        ddlPurchaser.Items.Clear()
        ddlPurchaser.Items.Add(New ListItem("-- 請選擇 --", ""))  ' 加入空白選項
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                ' 只載入有效的銷售人員 (Active = 'Y')
                Dim sql As String = "SELECT SlpCode, SlpName FROM OSLP WHERE Active = 'Y' ORDER BY SlpName"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddlPurchaser.Items.Add(New ListItem(dr("SlpName").ToString(), dr("SlpCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入採購人員失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 付款條件變更事件 - 自動計算到期日
    ''' </summary>
    Protected Sub ddlGroupNum_SelectedIndexChanged(sender As Object, e As EventArgs)
        CalculateDueDate()
    End Sub

    ''' <summary>
    ''' 根據付款條件計算到期日
    ''' SAP B1 OCTG 表：ExtraMonth=加月數, ExtraDays=加天數
    ''' 到期日 = 過帳日期 + ExtraMonth 月 + ExtraDays 天
    ''' </summary>
    Private Sub CalculateDueDate()
        Try
            If String.IsNullOrEmpty(ddlGroupNum.SelectedValue) Then Return
            If String.IsNullOrEmpty(txtDocDate.Text) Then Return

            Dim groupNum As Integer = Integer.Parse(ddlGroupNum.SelectedValue)
            Dim docDate As DateTime = DateTime.Parse(txtDocDate.Text)

            ' 從 SAP OCTG 取得付款條件的 ExtraMonth 和 ExtraDays
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT ExtraMonth, ExtraDays FROM OCTG WHERE GroupNum = @GroupNum"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@GroupNum", groupNum)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            Dim extraMonth As Integer = If(IsDBNull(dr("ExtraMonth")), 0, Convert.ToInt32(dr("ExtraMonth")))
                            Dim extraDays As Integer = If(IsDBNull(dr("ExtraDays")), 0, Convert.ToInt32(dr("ExtraDays")))

                            ' 計算到期日：過帳日期 + 月數 + 天數
                            Dim dueDate As DateTime = docDate.AddMonths(extraMonth).AddDays(extraDays)
                            txtDocDueDate.Text = dueDate.ToString("yyyy-MM-dd")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 計算失敗不阻斷流程，靜默處理
        End Try
    End Sub

    ''' <summary>
    ''' 載入費用項目 (從 OEPI 取值)
    ''' 顯示格式: {ExpItemName}
    ''' ToolTip 顯示: {ExpItemDescription}
    ''' </summary>
    Private Sub LoadExpenseCategories(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("-選擇-", ""))
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT ExpItemCode, ExpItemName, ExpItemDescription FROM OEPI ORDER BY ExpItemCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            Dim displayText As String = dr("ExpItemName").ToString()
                            Dim description As String = If(dr("ExpItemDescription") IsNot DBNull.Value, dr("ExpItemDescription").ToString(), "")
                            Dim item As New ListItem(displayText, dr("ExpItemCode").ToString())
                            ' 將描述存入 data-desc 屬性，供 RowDataBound 設定 ToolTip
                            item.Attributes.Add("data-desc", description)
                            ' 為選項加入 title 屬性 (滑鼠懸停顯示)
                            item.Attributes.Add("title", description)
                            ddl.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入費用項目失敗")
        End Try
    End Sub

    Private Sub LoadExpenseCategoriesFiltered(ddl As DropDownList, allowedCodes As List(Of String))
        If allowedCodes Is Nothing OrElse allowedCodes.Count = 0 Then
            LoadExpenseCategories(ddl)
            Return
        End If

        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("-選擇-", ""))

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim paramNames As New List(Of String)()
                For i As Integer = 0 To allowedCodes.Count - 1
                    paramNames.Add("@Code" & i.ToString())
                Next
                Dim sql As String = "SELECT ExpItemCode, ExpItemName, ExpItemDescription FROM OEPI " &
                                    "WHERE ExpItemCode IN (" & String.Join(",", paramNames) & ") " &
                                    "ORDER BY ExpItemCode"
                Using cmd As New SqlCommand(sql, conn)
                    For i As Integer = 0 To allowedCodes.Count - 1
                        cmd.Parameters.AddWithValue(paramNames(i), allowedCodes(i))
                    Next
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            Dim displayText As String = dr("ExpItemName").ToString()
                            Dim description As String = If(dr("ExpItemDescription") IsNot DBNull.Value, dr("ExpItemDescription").ToString(), "")
                            Dim item As New ListItem(displayText, dr("ExpItemCode").ToString())
                            item.Attributes.Add("data-desc", description)
                            item.Attributes.Add("title", description)
                            ddl.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
            LoadExpenseCategories(ddl)
        End Try
    End Sub

    Private Sub LoadVatGroups(ddl As DropDownList)
        ddl.Items.Clear()
        ' 1-應稅 (5%), 2-零稅 (0), 3-免稅 (0)
        Dim item1 As New ListItem("1-應稅 (5%)", "1")
        item1.Attributes.Add("data-rate", "5")
        ddl.Items.Add(item1)

        Dim item2 As New ListItem("2-零稅 (0%)", "2")
        item2.Attributes.Add("data-rate", "0")
        ddl.Items.Add(item2)

        Dim item3 As New ListItem("3-免稅 (0%)", "3")
        item3.Attributes.Add("data-rate", "0")
        ddl.Items.Add(item3)
    End Sub

    ''' <summary>
    ''' 載入產品下拉選單 (CostingCode - 成本中心維度1)
    ''' 來源：OPRC 中 DimCode=1 且排除一般中心
    ''' </summary>
    Private Sub LoadProducts(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                ' 產品：DimCode=1 (1030-AOI, 1040-ICT)，排除一般中心 Centr_z
                Dim sql As String = "SELECT PrcCode, PrcName FROM OPRC " &
                                    "WHERE DimCode = 1 AND PrcCode NOT LIKE 'Centr%' " &
                                    "ORDER BY PrcCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddl.Items.Add(New ListItem(dr("PrcName").ToString() & " (" & dr("PrcCode").ToString() & ")", dr("PrcCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入產品失敗")
        End Try
    End Sub

    ''' <summary>
    ''' 載入部門下拉選單 (CostingCode2 - 成本中心維度2)
    ''' 來源：OPRC 中 DimCode=2 且排除一般中心
    ''' </summary>
    Private Sub LoadDepartments(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                ' 部門：DimCode=2 (1050~1300)，排除一般中心 Centr_z2
                Dim sql As String = "SELECT PrcCode, PrcName FROM OPRC " &
                                    "WHERE DimCode = 2 AND PrcCode NOT LIKE 'Centr%' " &
                                    "ORDER BY PrcCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddl.Items.Add(New ListItem(dr("PrcName").ToString() & " (" & dr("PrcCode").ToString() & ")", dr("PrcCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入部門失敗")
        End Try
    End Sub

    Private Sub LoadDepartmentsFiltered(ddl As DropDownList, allowedCodes As List(Of String))
        If allowedCodes Is Nothing OrElse allowedCodes.Count = 0 Then
            LoadDepartments(ddl)
            Return
        End If

        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))

        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim paramNames As New List(Of String)()
                For i As Integer = 0 To allowedCodes.Count - 1
                    paramNames.Add("@Code" & i.ToString())
                Next
                Dim sql As String = "SELECT PrcCode, PrcName FROM OPRC " &
                                    "WHERE DimCode = 2 AND PrcCode NOT LIKE 'Centr%' " &
                                    "AND PrcCode IN (" & String.Join(",", paramNames) & ") " &
                                    "ORDER BY PrcCode"
                Using cmd As New SqlCommand(sql, conn)
                    For i As Integer = 0 To allowedCodes.Count - 1
                        cmd.Parameters.AddWithValue(paramNames(i), allowedCodes(i))
                    Next
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddl.Items.Add(New ListItem(dr("PrcName").ToString() & " (" & dr("PrcCode").ToString() & ")", dr("PrcCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
            LoadDepartments(ddl)
        End Try
    End Sub
#End Region

#Region "供應商搜尋 (Search Modal)"
    Protected Sub btnSearchCardCode_Click(sender As Object, e As EventArgs)
        PerformVendorSearch("Code", txtCardCode.Text.Trim())
    End Sub

    Protected Sub btnSearchCardName_Click(sender As Object, e As EventArgs)
        PerformVendorSearch("Name", txtCardName.Text.Trim())
    End Sub

    Private Sub PerformVendorSearch(source As String, keyword As String)
        hfSearchSource.Value = source
        txtVendorSearchKeyword.Text = keyword
        BindVendorSearchGrid(keyword)
        mpeVendor.Show()
        pnlVendorSearch.Style("display") = "block"
    End Sub

    Protected Sub btnDoSearchVendor_Click(sender As Object, e As EventArgs)
        BindVendorSearchGrid(txtVendorSearchKeyword.Text.Trim())
        mpeVendor.Show()
    End Sub

    Protected Sub gvVendorSearch_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
        gvVendorSearch.PageIndex = e.NewPageIndex
        BindVendorSearchGrid(txtVendorSearchKeyword.Text.Trim())
        mpeVendor.Show()
    End Sub

    Private Sub BindVendorSearchGrid(keyword As String)
        Try
            Dim sqlWhere As String = "WHERE CardType='S' AND FrozenFor='N' "

            Dim searchSource As String = hfSearchSource.Value ' 判斷搜尋來源
            Dim isExact As Boolean = (rblSearchMode.SelectedValue = "Exact")

            If Not String.IsNullOrEmpty(keyword) Then
                keyword = keyword.Replace("*", "").Replace("%", "")

                If isExact Then
                    ' 開頭比對
                    If searchSource = "Code" Then
                        sqlWhere &= " AND CardCode LIKE @Kw"
                    Else ' Name
                        sqlWhere &= " AND CardName LIKE @Kw"
                    End If
                    keyword = keyword & "%"
                Else
                    ' 模糊比對
                    If searchSource = "Code" Then
                        sqlWhere &= " AND CardCode LIKE @Kw"
                    Else ' Name
                        sqlWhere &= " AND CardName LIKE @Kw"
                    End If
                    keyword = "%" & keyword & "%"
                End If
            End If

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = $"SELECT CardCode, CardName FROM OCRD {sqlWhere} ORDER BY CardCode"
                Using cmd As New SqlCommand(sql, conn)
                    If Not String.IsNullOrEmpty(keyword) Then
                        cmd.Parameters.AddWithValue("@Kw", keyword)
                    End If

                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)
                        gvVendorSearch.DataSource = dt
                        gvVendorSearch.DataBind()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("搜尋供應商錯誤: " & ex.Message)
        End Try
    End Sub

    Protected Sub gvVendorSearch_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "SelectVendor" Then
            Dim args As String() = e.CommandArgument.ToString().Split("|"c)
            If args.Length >= 2 Then
                txtCardCode.Text = args(0)
                txtCardName.Text = args(1)

                ' 連動資料 (幣別, 付款條件)
                LoadVendorLinkedData(args(0))

                ' 清除錯誤訊息
                lblErrCardCode.Visible = False
                lblErrCardName.Visible = False
            End If
            mpeVendor.Hide()
            pnlVendorSearch.Style("display") = "none"
        End If
    End Sub

    Private Sub LoadVendorLinkedData(cardCode As String)
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                ' 幣別 (包含 U_LastCur 檢查)
                Dim sqlCurr As String = "SELECT Currency, U_LastCur FROM OCRD WHERE CardCode = @CardCode"
                Dim currencyAll As Boolean = False
                Using cmd As New SqlCommand(sqlCurr, conn)
                    cmd.Parameters.AddWithValue("@CardCode", cardCode)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            Dim dbCurr As String = dr("Currency").ToString()
                            Dim lastCur As String = If(IsDBNull(dr("U_LastCur")), "", dr("U_LastCur").ToString())

                            If dbCurr = "##" Then
                                currencyAll = True
                                ddlDocCurrency.Enabled = True
                                ' 優先帶入 U_LastCur，若無則帶入 NTD 或 TWD
                                If Not String.IsNullOrEmpty(lastCur) AndAlso ddlDocCurrency.Items.FindByValue(lastCur) IsNot Nothing Then
                                    ddlDocCurrency.SelectedValue = lastCur
                                Else
                                    If ddlDocCurrency.Items.FindByValue("NTD") IsNot Nothing Then
                                        ddlDocCurrency.SelectedValue = "NTD"
                                    End If
                                End If
                            Else
                                currencyAll = False
                                ddlDocCurrency.Enabled = False ' 鎖定幣別
                                If ddlDocCurrency.Items.FindByValue(dbCurr) IsNot Nothing Then
                                    ddlDocCurrency.SelectedValue = dbCurr
                                End If
                            End If
                        End If
                    End Using
                End Using

                ' 觸發匯率更新與明細連動
                ddlDocCurrency_SelectedIndexChanged(Nothing, Nothing)

                ' 付款條件
                Dim sqlPymnt As String = "SELECT GroupNum FROM OCRD WHERE CardCode = @CardCode"
                Using cmd As New SqlCommand(sqlPymnt, conn)
                    cmd.Parameters.AddWithValue("@CardCode", cardCode)
                    Dim grp As Object = cmd.ExecuteScalar()
                    If grp IsNot Nothing AndAlso Not IsDBNull(grp) Then
                        Dim grpNum As String = grp.ToString()
                        If ddlGroupNum.Items.FindByValue(grpNum) IsNot Nothing Then
                            ddlGroupNum.SelectedValue = grpNum
                            ' 自動計算到期日
                            CalculateDueDate()
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入供應商關聯資料失敗: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "會計科目搜尋 (Search Modal)"
    Protected Sub gvExpenseDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "SearchAcct" Then
            Dim rowIndex As Integer
            If Integer.TryParse(e.CommandArgument.ToString(), rowIndex) Then
                Dim keyword As String = ""
                Try
                    Dim row As GridViewRow = gvExpenseDetail.Rows(rowIndex)
                    Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
                    If txtAcct IsNot Nothing Then keyword = txtAcct.Text.Trim()
                Catch ex As Exception
                    keyword = ""
                End Try
                hfActiveTab.Value = "expense"
                OpenAcctSearch(rowIndex, keyword)
            End If
        End If
    End Sub

    Private Sub OpenAcctSearch(rowIndex As Integer, keyword As String)
        Try
            hfAcctSearchRowIndex.Value = rowIndex.ToString()
            txtAcctSearchKeyword.Text = keyword
            AcctSearchSortExpression = ""
            AcctSearchSortDirection = "ASC"
            BindAcctSearchGrid(keyword)
            mpeAcct.Show()
            pnlAcctSearch.Style("display") = "block"
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
        End Try
    End Sub

    Protected Sub btnDoSearchAcct_Click(sender As Object, e As EventArgs)
        BindAcctSearchGrid(txtAcctSearchKeyword.Text.Trim())
        mpeAcct.Show()
    End Sub

    Protected Sub gvAcctSearch_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
        gvAcctSearch.PageIndex = e.NewPageIndex
        BindAcctSearchGrid(txtAcctSearchKeyword.Text.Trim())
        mpeAcct.Show()
    End Sub

    Protected Sub gvAcctSearch_Sorting(sender As Object, e As GridViewSortEventArgs)
        Try
            If AcctSearchSortExpression = e.SortExpression Then
                AcctSearchSortDirection = If(AcctSearchSortDirection = "ASC", "DESC", "ASC")
            Else
                AcctSearchSortExpression = e.SortExpression
                AcctSearchSortDirection = "ASC"
            End If

            Dim dt As DataTable = AcctSearchData
            If dt Is Nothing Then
                BindAcctSearchGrid(txtAcctSearchKeyword.Text.Trim())
            Else
                Dim view As DataView = dt.DefaultView
                If Not String.IsNullOrEmpty(AcctSearchSortExpression) Then
                    view.Sort = AcctSearchSortExpression & " " & AcctSearchSortDirection
                End If
                gvAcctSearch.DataSource = view
                gvAcctSearch.DataBind()
            End If
            mpeAcct.Show()
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
        End Try
    End Sub

    Protected Sub gvAcctSearch_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "SelectAcct" Then
            Dim args As String() = e.CommandArgument.ToString().Split("|"c)
            If args.Length >= 1 Then
                Dim acctCode As String = args(0)
                Dim rowIndex As Integer
                If Integer.TryParse(hfAcctSearchRowIndex.Value, rowIndex) Then
                    Try
                        SyncGridDataToModel()
                        Dim lines = CurrentLines
                        If rowIndex < lines.Count Then
                            Dim line = lines(rowIndex)
                            line.AcctCode = acctCode
                            line.UserSelectedAcct = True

                            Dim itemCode As String = ""
                            Dim deptCode As String = ""
                            If TryGetUedlMapping(currentUserId, acctCode, itemCode, deptCode) Then
                                line.CategoryCode = itemCode
                                line.CostingCode2 = deptCode
                            Else
                                ApplyReverseMapToLine(line, acctCode)
                            End If

                            lines(rowIndex) = line
                            CurrentLines = lines
                            hfAcctPendingRowIndex.Value = rowIndex.ToString()
                            BindGrid()
                        End If
                    Catch ex As Exception
                        ' UI/UX 輔助失敗時靜默
                    End Try
                End If
            End If

            mpeAcct.Hide()
            pnlAcctSearch.Style("display") = "none"
        End If
    End Sub

    Protected Sub btnAcctRowLeave_Click(sender As Object, e As EventArgs)
        Try
            Dim rowIndex As Integer
            If Not Integer.TryParse(hfAcctPendingRowIndex.Value, rowIndex) Then
                Return
            End If

            SyncGridDataToModel()
            Dim lines = CurrentLines
            If rowIndex >= 0 AndAlso rowIndex < lines.Count Then
                TryUpsertUedlLog(lines(rowIndex))
            End If
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
        Finally
            hfAcctPendingRowIndex.Value = ""
        End Try
    End Sub

    Private Sub BindAcctSearchGrid(keyword As String)
        Try
            Dim dt As DataTable = GetAcctSearchData(keyword)
            AcctSearchData = dt

            Dim view As DataView = dt.DefaultView
            If Not String.IsNullOrEmpty(AcctSearchSortExpression) Then
                view.Sort = AcctSearchSortExpression & " " & AcctSearchSortDirection
            End If

            gvAcctSearch.DataSource = view
            gvAcctSearch.DataBind()
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
        End Try
    End Sub

    Private Function GetAcctSearchData(keyword As String) As DataTable
        Dim dt As New DataTable()
        dt.Columns.Add("AcctCode")
        dt.Columns.Add("AcctName")

        Try
            Dim searchMode As String = If(rblAcctSearchMode.SelectedValue, "Exact")
            Dim hasKeyword As Boolean = Not String.IsNullOrEmpty(keyword)
            Dim kw As String = keyword.Replace("*", "").Replace("%", "")

            If hasKeyword Then
                If searchMode = "Exact" Then
                    kw = kw & "%"
                Else
                    kw = "%" & kw & "%"
                End If
            End If

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 500 AcctCode, AcctName FROM OACT WHERE Postable='Y'"
                If hasKeyword Then
                    sql &= " AND (AcctCode LIKE @Kw OR AcctName LIKE @Kw)"
                End If
                sql &= " ORDER BY AcctCode"

                Using cmd As New SqlCommand(sql, conn)
                    If hasKeyword Then
                        cmd.Parameters.AddWithValue("@Kw", kw)
                    End If
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            Dim code As String = dr("AcctCode").ToString().Trim()
                            Dim name As String = dr("AcctName").ToString()
                            dt.Rows.Add(code, name)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
        End Try

        Return dt
    End Function

#End Region

#Region "表頭欄位連動"
    Protected Sub ddlDeliveryAddr_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim id As String = ddlDeliveryAddr.SelectedValue
        If String.IsNullOrEmpty(id) Then
            txtAddress.Text = ""
            Return
        End If

        Using conn As New SqlConnection(connStr)
            conn.Open()
            Dim sql As String = "SELECT address FROM addr WHERE ID=@ID"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ID", id)
                Dim res = cmd.ExecuteScalar()
                If res IsNot Nothing Then
                    txtAddress.Text = res.ToString()
                End If
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' [F] 匯率取得邏輯，加入日期範圍限制和警告
    ''' </summary>
    Protected Sub ddlDocCurrency_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim curr As String = ddlDocCurrency.SelectedValue
        Dim rate As Decimal = 1D
        Dim rateDate As DateTime? = Nothing

        If curr = "TWD" OrElse curr = "NTD" Then
            txtDocRate.Text = "1.0"
            hfRateDate.Value = DateTime.Today.ToString("yyyy-MM-dd")
        Else
            ' 取得 DocDate，若無值則預設為今天
            Dim docDate As DateTime = DateTime.Today
            If Not String.IsNullOrEmpty(txtDocDate.Text) Then
                DateTime.TryParse(txtDocDate.Text, docDate)
            End If

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                ' [F] 取得最近的匯率資料，同時取得匯率日期
                Dim sql As String = "SELECT TOP 1 Rate, RateDate FROM ORTT WHERE Currency=@Curr AND RateDate <= @DocDate ORDER BY RateDate DESC"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Curr", curr)
                    cmd.Parameters.AddWithValue("@DocDate", docDate)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            rate = Convert.ToDecimal(dr("Rate"))
                            rateDate = Convert.ToDateTime(dr("RateDate"))
                            txtDocRate.Text = rate.ToString("F4")
                            hfRateDate.Value = rateDate.Value.ToString("yyyy-MM-dd")
                        Else
                            txtDocRate.Text = "1.0"
                            hfRateDate.Value = ""
                        End If
                    End Using
                End Using
            End Using

            ' [F] 檢查匯率日期是否為當天
            If rateDate.HasValue Then
                Dim daysDiff As Integer = CInt((DateTime.Today - rateDate.Value).TotalDays)
                If daysDiff > 0 Then
                    ShowWarning(String.Format("提醒：匯率資料為 {0:yyyy-MM-dd}，距今已 {1} 天，請確認匯率是否正確。", rateDate.Value, daysDiff))
                End If
            End If
        End If

        ' 同步更新明細列的幣別與匯率
        UpdateLinesCurrency(curr, rate)
    End Sub

    Protected Sub txtDocRate_TextChanged(sender As Object, e As EventArgs)
        Dim curr As String = ddlDocCurrency.SelectedValue
        Dim rate As Decimal = 1D
        Decimal.TryParse(txtDocRate.Text, rate)

        ' 同步更新明細列的幣別與匯率
        UpdateLinesCurrency(curr, rate)
    End Sub

    Private Sub UpdateLinesCurrency(curr As String, rate As Decimal)
        Dim lines = CurrentLines
        For Each line In lines
            line.Currency = curr
            line.Rate = rate
        Next
        CurrentLines = lines
        BindGrid()
    End Sub


    Protected Sub btnRefreshRate_Click(sender As Object, e As EventArgs)
        ddlDocCurrency_SelectedIndexChanged(Nothing, Nothing)
        ' 過帳日期變更時，重新計算到期日
        CalculateDueDate()
    End Sub
#End Region

#Region "Expense GridView"
    Protected Sub btnAddLine_Click(sender As Object, e As EventArgs)
        SyncGridDataToModel()
        AddNewEmptyLine()
        BindGrid()
    End Sub

    Private Sub AddNewEmptyLine()
        Dim lines = CurrentLines
        Dim newNum As Integer = 1
        If lines.Count > 0 Then newNum = lines.Max(Function(x) x.LineNum) + 1

        Dim curr As String = ddlDocCurrency.SelectedValue
        Dim rate As Decimal = 1D
        Decimal.TryParse(txtDocRate.Text, rate)

        lines.Add(New ExpenseLine With {
            .LineNum = newNum,
            .LineTotal = 0,
            .VatRate = 0,
            .PriceAfterVat = 0,
            .Currency = curr,
            .Rate = rate
        })
        CurrentLines = lines
    End Sub

    Protected Sub btnDeleteLine_Click(sender As Object, e As EventArgs)
        SyncGridDataToModel()
        Dim lines = CurrentLines
        Dim newLines As New List(Of ExpenseLine)

        For Each row As GridViewRow In gvExpenseDetail.Rows
            If row.RowType = DataControlRowType.DataRow Then
                Dim chk As CheckBox = CType(row.FindControl("chkSelect"), CheckBox)
                If chk IsNot Nothing AndAlso Not chk.Checked Then
                    Dim idx As Integer = row.DataItemIndex
                    If idx < lines.Count Then newLines.Add(lines(idx))
                End If
            End If
        Next

        ' 重算 LineNum
        For i As Integer = 0 To newLines.Count - 1
            newLines(i).LineNum = i + 1
        Next

        CurrentLines = newLines
        BindGrid()
    End Sub

    Private Sub BindGrid()
        gvExpenseDetail.DataSource = CurrentLines
        gvExpenseDetail.DataBind()
        CalculateFooterTotal()
    End Sub

    Protected Sub gvExpenseDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim line As ExpenseLine = CType(e.Row.DataItem, ExpenseLine)
            e.Row.Attributes("data-rowindex") = e.Row.DataItemIndex.ToString()

            Dim acctFilter As AcctReverseMap = Nothing
            If line.UserSelectedAcct AndAlso Not String.IsNullOrEmpty(line.AcctCode) Then
                acctFilter = GetReverseMapByAcctCode(line.AcctCode)
            End If

            ' Expense Category
            Dim ddlCat As DropDownList = CType(e.Row.FindControl("ddlExpCategory"), DropDownList)
            If acctFilter IsNot Nothing AndAlso acctFilter.ExpItemCodes.Count > 0 Then
                LoadExpenseCategoriesFiltered(ddlCat, acctFilter.ExpItemCodes)
            Else
                LoadExpenseCategories(ddlCat)
            End If
            If ddlCat.Items.FindByValue(line.CategoryCode) IsNot Nothing Then
                ddlCat.SelectedValue = line.CategoryCode
                ' 設定下拉選單 ToolTip 為目前選中項目的描述
                Dim selectedItem = ddlCat.SelectedItem
                If selectedItem IsNot Nothing AndAlso selectedItem.Attributes("data-desc") IsNot Nothing Then
                    ddlCat.ToolTip = selectedItem.Attributes("data-desc")
                End If
            End If

            ' Vat Group
            Dim ddlVat As DropDownList = CType(e.Row.FindControl("ddlVatGroup"), DropDownList)
            LoadVatGroups(ddlVat)
            If ddlVat.Items.FindByValue(line.VatGroup) IsNot Nothing Then ddlVat.SelectedValue = line.VatGroup

            ' Products (CostingCode)
            Dim ddlCost As DropDownList = CType(e.Row.FindControl("ddlCostingCode"), DropDownList)
            LoadProducts(ddlCost)
            If ddlCost.Items.FindByValue(line.CostingCode) IsNot Nothing Then ddlCost.SelectedValue = line.CostingCode

            ' Departments (CostingCode2)
            Dim ddlCost2 As DropDownList = CType(e.Row.FindControl("ddlCostingCode2"), DropDownList)
            If acctFilter IsNot Nothing AndAlso line.UserSelectedAcct Then
                If acctFilter.AllowAllDepartments Then
                    LoadDepartments(ddlCost2)
                ElseIf acctFilter.DeptCodes.Count > 0 Then
                    LoadDepartmentsFiltered(ddlCost2, acctFilter.DeptCodes)
                Else
                    LoadDepartments(ddlCost2)
                End If
            Else
                LoadDepartments(ddlCost2)
            End If
            If ddlCost2.Items.FindByValue(line.CostingCode2) IsNot Nothing Then ddlCost2.SelectedValue = line.CostingCode2

            ' Values
            CType(e.Row.FindControl("txtDescription"), TextBox).Text = line.Description

            ' 會計科目：可搜尋或手動輸入
            Dim txtAcct As TextBox = CType(e.Row.FindControl("txtAcctCode"), TextBox)
            txtAcct.Text = line.AcctCode
            ' 設定 ToolTip 顯示會計科目名稱
            If Not String.IsNullOrEmpty(line.AcctCode) Then
                Dim acctName As String = GetAcctNameByCode(line.AcctCode)
                txtAcct.ToolTip = acctName
            End If
            Dim btnSearchAcct As Button = CType(e.Row.FindControl("btnSearchAcct"), Button)
            If btnSearchAcct IsNot Nothing Then
                btnSearchAcct.Enabled = isApUser
            End If
            If isApUser Then
                txtAcct.ReadOnly = False
                txtAcct.CssClass = ""
            End If

            CType(e.Row.FindControl("txtLineTotal"), TextBox).Text = line.LineTotal.ToString("0.##")
            CType(e.Row.FindControl("txtVatSum"), TextBox).Text = line.VatSum.ToString("0.##")
            CType(e.Row.FindControl("txtPriceAfterVat"), TextBox).Text = line.PriceAfterVat.ToString("0.##")

            ' Currency & Rate - 已移至單頭，不再於明細列顯示
        End If
    End Sub

    Protected Sub ddlExpCategory_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ddl As DropDownList = CType(sender, DropDownList)
        Dim row As GridViewRow = CType(ddl.NamingContainer, GridViewRow)
        Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
        Dim ddlDept As DropDownList = CType(row.FindControl("ddlCostingCode2"), DropDownList)

        If ddl.SelectedIndex > 0 AndAlso Not String.IsNullOrEmpty(ddl.SelectedValue) Then
            ' 根據費用項目和部門查詢對應的會計科目
            Dim deptCode As String = If(ddlDept IsNot Nothing, ddlDept.SelectedValue, "")
            Dim acctInfo As AcctInfo
            If Not String.IsNullOrEmpty(deptCode) Then
                ' 有選擇部門，使用新的函數
                acctInfo = GetAcctCodeByExpItemAndDept(ddl.SelectedValue, deptCode)
            Else
                ' 沒有選擇部門，使用原本的函數（依使用者預設部門）
                acctInfo = GetAcctCodeByExpItem(ddl.SelectedValue)
            End If
            txtAcct.Text = acctInfo.AcctCode
            txtAcct.ToolTip = acctInfo.AcctName ' 設定 ToolTip 顯示科目名稱

            ' 更新費用類別下拉選單的 ToolTip 為選中項目的描述
            Dim selectedItem = ddl.SelectedItem
            If selectedItem IsNot Nothing Then
                ddl.ToolTip = GetExpItemDescription(ddl.SelectedValue)
            End If

            ' 進出口報關費用-其他 (E030) - 提示改選 E030A
            If ddl.SelectedValue = "E030" AndAlso Not E030OtherWarningShown Then
                Dim script As String = "alert('提示：若要登打海關代徵營業稅\n\n" &
                                      "請選擇「進出口報關費用-28-海關代徵營業稅」項目 (E030A)。\n\n" &
                                      "本項目「進出口報關費用-其他」適用於一般進出口報關費。');"
                ScriptManager.RegisterStartupScript(Me.Page, Me.GetType(), "expE030Warning", script, True)
                E030OtherWarningShown = True
            End If

            ' 進出口報關費用-28-海關代徵營業稅 (E030A) - 格式28 特殊警告
            ' 同一次進入單據期間，費用頁籤最多警告一次
            If ddl.SelectedValue = "E030A" AndAlso Not Format28ExpenseWarningShown Then
                Dim script As String = "alert('注意：海關代徵營業稅 (格式28)\n\n" &
                                      "此項目寫入SAP時只會過帳「稅額」金額。\n\n" &
                                      "請在營業稅區填入：\n" &
                                      "- 未稅金額 = 營業稅基\n" &
                                      "- 稅額 = 實際應繳稅額');"
                ScriptManager.RegisterStartupScript(Me.Page, Me.GetType(), "expE030AWarning", script, True)
                Format28ExpenseWarningShown = True
            End If
        Else
            txtAcct.Text = ""
            txtAcct.ToolTip = ""
            ddl.ToolTip = ""
        End If

        SyncGridDataToModel()
    End Sub

    ''' <summary>
    ''' 部門 (CostingCode2) 變更事件 - 根據部門重新查詢會計科目
    ''' </summary>
    Protected Sub ddlCostingCode2_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ddlDept As DropDownList = CType(sender, DropDownList)
        Dim row As GridViewRow = CType(ddlDept.NamingContainer, GridViewRow)
        Dim ddlExpItem As DropDownList = CType(row.FindControl("ddlExpCategory"), DropDownList)
        Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)

        ' 只有當費用項目已選擇時才更新會計科目
        If ddlExpItem IsNot Nothing AndAlso ddlExpItem.SelectedIndex > 0 AndAlso Not String.IsNullOrEmpty(ddlExpItem.SelectedValue) Then
            Dim deptCode As String = ddlDept.SelectedValue
            Dim acctInfo As AcctInfo
            If Not String.IsNullOrEmpty(deptCode) Then
                ' 有選擇部門，根據部門的 CCTypeCode 查詢會計科目
                acctInfo = GetAcctCodeByExpItemAndDept(ddlExpItem.SelectedValue, deptCode)
            Else
                ' 沒有選擇部門，使用原本的函數（依使用者預設部門）
                acctInfo = GetAcctCodeByExpItem(ddlExpItem.SelectedValue)
            End If
            txtAcct.Text = acctInfo.AcctCode
            txtAcct.ToolTip = acctInfo.AcctName
        End If

        SyncGridDataToModel()
    End Sub

    ''' <summary>
    ''' 會計科目資訊結構
    ''' </summary>
    Private Structure AcctInfo
        Public AcctCode As String
        Public AcctName As String
    End Structure

    Private Class AcctReverseMap
        Public Property ExpItemCodes As List(Of String)
        Public Property ExpClasses As List(Of String)
        Public Property DeptCodes As List(Of String)
        Public Property AllowAllDepartments As Boolean
    End Class

    Private Function GetAcctNameByCode(acctCode As String) As String
        If String.IsNullOrEmpty(acctCode) Then Return ""
        Try
            Dim cache = TryCast(Session("AcctNameCache"), Dictionary(Of String, String))
            If cache Is Nothing Then cache = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            If cache.ContainsKey(acctCode) Then Return cache(acctCode)

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT AcctName FROM OACT WHERE AcctCode = @AcctCode"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@AcctCode", acctCode)
                    Dim result = cmd.ExecuteScalar()
                    Dim acctName As String = If(result IsNot Nothing AndAlso Not IsDBNull(result), result.ToString(), "")
                    cache(acctCode) = acctName
                    Session("AcctNameCache") = cache
                    Return acctName
                End Using
            End Using
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 根據費用項目代碼和使用者費用部門，從 EPI1 查詢對應的會計科目
    ''' 邏輯:
    ''' 1. 先檢查該費用項目是否有 ExpClass='公' 的科目（公共費用，不分部門）
    ''' 2. 若有，直接使用「公」的科目
    ''' 3. 若無，則依 User.expDEPT -> jDEPT.ExpClass -> EPI1 對照
    ''' </summary>
    Private Function GetAcctCodeByExpItem(expItemCode As String) As AcctInfo
        Dim result As New AcctInfo()
        result.AcctCode = ""
        result.AcctName = ""

        If String.IsNullOrEmpty(expItemCode) Then Return result

        Try
            ' 步驟1: 先檢查是否為公共費用項目（ExpClass='公'）
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT AcctCode, AcctName FROM EPI1 WHERE ExpItemCode = @ExpItemCode AND ExpClass = N'公'"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ExpItemCode", expItemCode)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            ' 找到「公」的科目，直接返回
                            result.AcctCode = If(dr("AcctCode") IsNot DBNull.Value, dr("AcctCode").ToString(), "")
                            result.AcctName = If(dr("AcctName") IsNot DBNull.Value, dr("AcctName").ToString(), "")
                            Return result
                        End If
                    End Using
                End Using
            End Using

            ' 步驟2: 非公共費用，取得使用者的 expDEPT
            Dim userExpDept As String = ""
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT expDEPT FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", currentUserId)
                    Dim dbResult = cmd.ExecuteScalar()
                    If dbResult IsNot Nothing AndAlso Not IsDBNull(dbResult) Then
                        userExpDept = dbResult.ToString().Trim()
                    End If
                End Using
            End Using

            If String.IsNullOrEmpty(userExpDept) Then Return result

            ' 步驟3: 從 jDEPT 取得 ExpClass
            Dim expClass As String = ""
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT ExpClass FROM jDEPT WHERE EDeptID = @EDeptID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@EDeptID", userExpDept)
                    Dim dbResult = cmd.ExecuteScalar()
                    If dbResult IsNot Nothing AndAlso Not IsDBNull(dbResult) Then
                        expClass = dbResult.ToString().Trim()
                    End If
                End Using
            End Using

            If String.IsNullOrEmpty(expClass) Then Return result

            ' 步驟4: 從 EPI1 查詢會計科目
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT AcctCode, AcctName FROM EPI1 WHERE ExpItemCode = @ExpItemCode AND ExpClass = @ExpClass"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ExpItemCode", expItemCode)
                    cmd.Parameters.AddWithValue("@ExpClass", expClass)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            result.AcctCode = If(dr("AcctCode") IsNot DBNull.Value, dr("AcctCode").ToString(), "")
                            result.AcctName = If(dr("AcctName") IsNot DBNull.Value, dr("AcctName").ToString(), "")
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' 查詢失敗時返回空結構
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 根據 SAP 部門代碼 (PrcCode) 查詢費用分類 (CCTypeCode)
    ''' CCTypeCode 對應：製、銷、管、研、製-CNC
    ''' </summary>
    Private Function GetExpClassByDept(deptCode As String) As String
        If String.IsNullOrEmpty(deptCode) Then Return ""
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT CCTypeCode FROM OPRC WHERE PrcCode = @PrcCode"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@PrcCode", deptCode)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString().Trim()
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' 查詢失敗時返回空字串
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' 根據費用項目代碼和部門代碼，從 EPI1 查詢對應的會計科目
    ''' 邏輯:
    ''' 1. 先檢查該費用項目是否有 ExpClass='公' 的科目（公共費用，不分部門）
    ''' 2. 若有，直接使用「公」的科目
    ''' 3. 若無，則根據部門的 CCTypeCode (製/銷/管/研/製-CNC) 查詢 EPI1
    ''' </summary>
    Private Function GetAcctCodeByExpItemAndDept(expItemCode As String, deptCode As String) As AcctInfo
        Dim result As New AcctInfo()
        result.AcctCode = ""
        result.AcctName = ""

        If String.IsNullOrEmpty(expItemCode) Then Return result

        Try
            ' 步驟1: 先檢查是否為公共費用項目（ExpClass='公'）
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT AcctCode, AcctName FROM EPI1 WHERE ExpItemCode = @ExpItemCode AND ExpClass = N'公'"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ExpItemCode", expItemCode)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            ' 找到「公」的科目，直接返回
                            result.AcctCode = If(dr("AcctCode") IsNot DBNull.Value, dr("AcctCode").ToString(), "")
                            result.AcctName = If(dr("AcctName") IsNot DBNull.Value, dr("AcctName").ToString(), "")
                            Return result
                        End If
                    End Using
                End Using
            End Using

            ' 步驟2: 非公共費用，從 SAP OPRC 取得部門的 ExpClass (CCTypeCode)
            Dim expClass As String = GetExpClassByDept(deptCode)
            If String.IsNullOrEmpty(expClass) Then Return result

            ' 步驟3: 從 EPI1 查詢會計科目
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT AcctCode, AcctName FROM EPI1 WHERE ExpItemCode = @ExpItemCode AND ExpClass = @ExpClass"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ExpItemCode", expItemCode)
                    cmd.Parameters.AddWithValue("@ExpClass", expClass)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            result.AcctCode = If(dr("AcctCode") IsNot DBNull.Value, dr("AcctCode").ToString(), "")
                            result.AcctName = If(dr("AcctName") IsNot DBNull.Value, dr("AcctName").ToString(), "")
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' 查詢失敗時返回空結構
        End Try

        Return result
    End Function

    Private Function GetReverseMapByAcctCode(acctCode As String) As AcctReverseMap
        If String.IsNullOrEmpty(acctCode) Then Return Nothing
        Try
            Dim cache = TryCast(Session("AcctReverseMapCache"), Dictionary(Of String, AcctReverseMap))
            If cache Is Nothing Then cache = New Dictionary(Of String, AcctReverseMap)(StringComparer.OrdinalIgnoreCase)
            If cache.ContainsKey(acctCode) Then Return cache(acctCode)

            Dim map As New AcctReverseMap() With {
                .ExpItemCodes = New List(Of String)(),
                .ExpClasses = New List(Of String)(),
                .DeptCodes = New List(Of String)(),
                .AllowAllDepartments = False
            }

            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT ExpItemCode, ExpClass FROM EPI1 WHERE AcctCode = @AcctCode"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@AcctCode", acctCode)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            Dim expItemCode As String = dr("ExpItemCode").ToString().Trim()
                            Dim expClass As String = dr("ExpClass").ToString().Trim()
                            If expItemCode <> "" AndAlso Not map.ExpItemCodes.Contains(expItemCode) Then map.ExpItemCodes.Add(expItemCode)
                            If expClass <> "" AndAlso Not map.ExpClasses.Contains(expClass) Then map.ExpClasses.Add(expClass)
                        End While
                    End Using
                End Using
            End Using

            If map.ExpClasses.Contains("公") Then
                map.AllowAllDepartments = True
            ElseIf map.ExpClasses.Count > 0 Then
                Using conn As New SqlConnection(sapConnStr)
                    conn.Open()
                    Dim paramNames As New List(Of String)()
                    For i As Integer = 0 To map.ExpClasses.Count - 1
                        paramNames.Add("@Class" & i.ToString())
                    Next
                    Dim sql As String = "SELECT PrcCode FROM OPRC WHERE DimCode = 2 AND PrcCode NOT LIKE 'Centr%' " &
                                        "AND CCTypeCode IN (" & String.Join(",", paramNames) & ")"
                    Using cmd As New SqlCommand(sql, conn)
                        For i As Integer = 0 To map.ExpClasses.Count - 1
                            cmd.Parameters.AddWithValue(paramNames(i), map.ExpClasses(i))
                        Next
                        Using dr As SqlDataReader = cmd.ExecuteReader()
                            While dr.Read()
                                Dim deptCode As String = dr("PrcCode").ToString().Trim()
                                If deptCode <> "" AndAlso Not map.DeptCodes.Contains(deptCode) Then map.DeptCodes.Add(deptCode)
                            End While
                        End Using
                    End Using
                End Using
            End If

            cache(acctCode) = map
            Session("AcctReverseMapCache") = cache
            Return map
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Sub ApplyReverseMapToLine(ByRef line As ExpenseLine, acctCode As String)
        Try
            Dim map = GetReverseMapByAcctCode(acctCode)
            If map Is Nothing Then Return

            If map.ExpItemCodes.Count = 1 Then
                line.CategoryCode = map.ExpItemCodes(0)
            ElseIf map.ExpItemCodes.Count > 1 Then
                If Not map.ExpItemCodes.Contains(line.CategoryCode) Then
                    line.CategoryCode = ""
                End If
            End If

            If map.AllowAllDepartments Then
                ' 允許全部部門，保留原選擇
            ElseIf map.DeptCodes.Count = 1 Then
                line.CostingCode2 = map.DeptCodes(0)
            ElseIf map.DeptCodes.Count > 1 Then
                If Not map.DeptCodes.Contains(line.CostingCode2) Then
                    line.CostingCode2 = ""
                End If
            End If
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
        End Try
    End Sub

    Private Function TryGetUedlMapping(userId As String, acctCode As String, ByRef itemCode As String, ByRef deptCode As String) As Boolean
        itemCode = ""
        deptCode = ""

        If String.IsNullOrEmpty(userId) OrElse String.IsNullOrEmpty(acctCode) Then Return False

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 1 ItemCode, CostingCode2 FROM UEDL " &
                                    "WHERE UserId = @UserId AND AcctCode = @AcctCode " &
                                    "ORDER BY expDate DESC"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    cmd.Parameters.AddWithValue("@AcctCode", acctCode)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            itemCode = If(IsDBNull(dr("ItemCode")), "", dr("ItemCode").ToString())
                            deptCode = If(IsDBNull(dr("CostingCode2")), "", dr("CostingCode2").ToString())
                            Return True
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try

        Return False
    End Function

    Private Sub TryUpsertUedlLog(line As ExpenseLine)
        If line Is Nothing Then Return
        If Not line.UserSelectedAcct Then Return
        If String.IsNullOrEmpty(line.AcctCode) Then Return

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String =
                    "IF EXISTS (SELECT 1 FROM UEDL WHERE UserId = @UserId AND AcctCode = @AcctCode) " &
                    "BEGIN " &
                    "UPDATE UEDL SET ItemCode = @ItemCode, CostingCode2 = @CostingCode2, jID = @jID, LineNum = @LineNum, expDate = GETDATE() " &
                    "WHERE UserId = @UserId AND AcctCode = @AcctCode " &
                    "END " &
                    "ELSE " &
                    "BEGIN " &
                    "INSERT INTO UEDL (UserId, AcctCode, ItemCode, CostingCode2, jID, LineNum, expDate) " &
                    "VALUES (@UserId, @AcctCode, @ItemCode, @CostingCode2, @jID, @LineNum, GETDATE()) " &
                    "END"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", currentUserId)
                    cmd.Parameters.AddWithValue("@AcctCode", line.AcctCode)
                    cmd.Parameters.AddWithValue("@ItemCode", line.CategoryCode)
                    cmd.Parameters.AddWithValue("@CostingCode2", line.CostingCode2)
                    cmd.Parameters.AddWithValue("@jID", currentJID)
                    cmd.Parameters.AddWithValue("@LineNum", line.LineNum)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            ' UI/UX 輔助失敗時靜默
        End Try
    End Sub

    ''' <summary>
    ''' 根據費用項目代碼取得描述 (ExpItemDescription)
    ''' </summary>
    Private Function GetExpItemDescription(expItemCode As String) As String
        If String.IsNullOrEmpty(expItemCode) Then Return ""
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT ExpItemDescription FROM OEPI WHERE ExpItemCode = @ExpItemCode"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ExpItemCode", expItemCode)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' 查詢失敗時返回空字串
        End Try
        Return ""
    End Function

    Protected Sub txtDescription_TextChanged(sender As Object, e As EventArgs)
        SyncGridDataToModel()
    End Sub

    Protected Sub CalculateLineTotal(sender As Object, e As EventArgs)
        ' 當未稅金額、稅別改變時觸發
        SyncGridDataToModel(False)

        Dim ddl As DropDownList = TryCast(sender, DropDownList)
        Dim txt As TextBox = TryCast(sender, TextBox)
        Dim row As GridViewRow = Nothing
        If ddl IsNot Nothing Then row = CType(ddl.NamingContainer, GridViewRow)
        If txt IsNot Nothing Then row = CType(txt.NamingContainer, GridViewRow)

        If row IsNot Nothing Then
            Dim rowIndex As Integer = row.DataItemIndex
            Dim lines = CurrentLines
            If rowIndex < lines.Count Then
                Dim line = lines(rowIndex)

                ' 設定稅率
                If line.VatGroup = "1" Then ' 1-應稅 (5%)
                    line.VatRate = 5
                Else
                    line.VatRate = 0
                End If

                ' 讀取 UI 上當前的稅額（用戶可能已手動修改）
                Dim txtVatSum As TextBox = CType(row.FindControl("txtVatSum"), TextBox)
                Dim currentVatSum As Decimal = 0
                If txtVatSum IsNot Nothing Then
                    Decimal.TryParse(txtVatSum.Text, currentVatSum)
                End If

                ' 稅額處理邏輯：
                ' - 非應稅項目：稅額強制為 0
                ' - 應稅項目：
                '   * 如果稅額為 0（新增明細），自動計算
                '   * 如果稅額不為 0（已有值），保留用戶輸入的稅額
                '     用戶可能根據實際憑證調整稅額（進位/捨去差異）
                If line.VatGroup <> "1" Then
                    ' 非應稅項目，稅額必須為 0
                    line.VatSum = 0
                ElseIf currentVatSum = 0 AndAlso line.LineTotal > 0 Then
                    ' 新增明細或清空稅額，自動計算
                    ' E030A (海關代徵營業稅) 使用無條件捨去，其他項目使用四捨五入
                    If line.CategoryCode = "E030A" Then
                        line.VatSum = Math.Floor(line.LineTotal * 0.05D)
                    Else
                        line.VatSum = Math.Round(line.LineTotal * 0.05D, 0, MidpointRounding.AwayFromZero)
                    End If
                Else
                    ' 已有稅額，保留用戶輸入的值
                    line.VatSum = currentVatSum
                End If

                ' 重新計算含稅金額
                line.PriceAfterVat = line.LineTotal + line.VatSum
            End If
            CurrentLines = lines
        End If

        BindGrid()
    End Sub

    Protected Sub CalculateVatSum(sender As Object, e As EventArgs)
        ' 當稅額被手動修改時
        SyncGridDataToModel(True) ' Update Model with screen VatSum

        Dim lines = CurrentLines
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
        Dim rowIndex As Integer = row.DataItemIndex

        If rowIndex < lines.Count Then
            Dim line = lines(rowIndex)

            ' 檢查用戶輸入的稅額是否與系統計算值不同
            If line.VatGroup = "1" AndAlso line.LineTotal > 0 Then
                ' E030A (海關代徵營業稅) 使用無條件捨去，其他項目使用四捨五入
                Dim calculatedVat As Decimal
                If line.CategoryCode = "E030A" Then
                    calculatedVat = Math.Floor(line.LineTotal * 0.05D)
                Else
                    calculatedVat = Math.Round(line.LineTotal * 0.05D, 0, MidpointRounding.AwayFromZero)
                End If

                If line.VatSum <> calculatedVat Then
                    ' 顯示警示：稅額與系統計算不同
                    Dim script As String = String.Format(
                        "alert('注意：您輸入的稅額 ({0:N0}) 與系統計算值 ({1:N0}) 不同。\n\n" &
                        "若憑證上的稅額確實為 {0:N0} 元，請忽略此訊息。\n\n" &
                        "系統將保留您輸入的稅額。');",
                        line.VatSum, calculatedVat)
                    ScriptManager.RegisterStartupScript(Me.Page, Me.GetType(), "vatDiffWarning" & rowIndex, script, True)
                End If
            End If

            ' 手動修改稅額後，重新計算含稅金額
            line.PriceAfterVat = line.LineTotal + line.VatSum
        End If
        CurrentLines = lines
        BindGrid()
    End Sub

    Protected Sub CalculatePriceAfterVat(sender As Object, e As EventArgs)
        ' 當含稅金額改變時觸發 -> 反算未稅金額
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
        Dim rowIndex As Integer = row.DataItemIndex

        Dim priceAfterVat As Decimal = 0
        Decimal.TryParse(txt.Text, priceAfterVat)

        SyncGridDataToModel(False)

        Dim lines = CurrentLines
        If rowIndex < lines.Count Then
            Dim line = lines(rowIndex)
            line.PriceAfterVat = priceAfterVat

            ' 反算邏輯
            ' 1-應稅 (5%)
            If line.VatGroup = "1" Then
                line.VatRate = 5
                line.LineTotal = Math.Round(priceAfterVat / 1.05D, 0, MidpointRounding.AwayFromZero)
                line.VatSum = priceAfterVat - line.LineTotal
            Else
                line.VatRate = 0
                line.LineTotal = priceAfterVat
                line.VatSum = 0
            End If
        End If
        CurrentLines = lines
        BindGrid()
    End Sub

    Private Sub SyncGridDataToModel(Optional readPriceAfterVat As Boolean = False)
        Dim lines = CurrentLines

        ' 從頁面取得幣別和匯率
        Dim docCurrency As String = If(ddlDocCurrency.SelectedValue, "NTD")
        Dim docRate As Decimal = 1
        Decimal.TryParse(txtDocRate.Text, docRate)

        For i As Integer = 0 To gvExpenseDetail.Rows.Count - 1
            Dim row As GridViewRow = gvExpenseDetail.Rows(i)
            If i < lines.Count Then
                Dim ddlExp As DropDownList = CType(row.FindControl("ddlExpCategory"), DropDownList)
                Dim txtDesc As TextBox = CType(row.FindControl("txtDescription"), TextBox)
                Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
                Dim txtTotal As TextBox = CType(row.FindControl("txtLineTotal"), TextBox)
                Dim ddlVat As DropDownList = CType(row.FindControl("ddlVatGroup"), DropDownList)
                Dim txtVatSum As TextBox = CType(row.FindControl("txtVatSum"), TextBox)
                Dim txtPrice As TextBox = CType(row.FindControl("txtPriceAfterVat"), TextBox)
                Dim ddlCost As DropDownList = CType(row.FindControl("ddlCostingCode"), DropDownList)
                Dim ddlCost2 As DropDownList = CType(row.FindControl("ddlCostingCode2"), DropDownList)

                Dim line = lines(i)
                If ddlExp IsNot Nothing Then line.CategoryCode = ddlExp.SelectedValue
                If txtDesc IsNot Nothing Then line.Description = txtDesc.Text
                If txtAcct IsNot Nothing Then line.AcctCode = txtAcct.Text
                If txtTotal IsNot Nothing Then Decimal.TryParse(txtTotal.Text, line.LineTotal)
                If ddlVat IsNot Nothing Then line.VatGroup = ddlVat.SelectedValue
                If txtVatSum IsNot Nothing Then Decimal.TryParse(txtVatSum.Text, line.VatSum)
                If ddlCost IsNot Nothing Then line.CostingCode = ddlCost.SelectedValue
                If ddlCost2 IsNot Nothing Then line.CostingCode2 = ddlCost2.SelectedValue

                ' 同步幣別和匯率
                line.Currency = docCurrency
                line.Rate = docRate

                ' 注意：這裡不再自動重算 VatSum，避免覆蓋使用者手動輸入的稅額。
                ' 重算邏輯移至 CalculateLineTotal (當 LineTotal 或 VatGroup 改變時)
                ' 或 CalculatePriceAfterVat (當 PriceAfterVat 改變時)

                ' 若是由 CalculatePriceAfterVat 呼叫 (readPriceAfterVat=False)，這時 UI 上的 PriceAfterVat 可能還沒寫入 Model (因為還沒 Assign)，
                ' 但這裡是讀取 UI 上的 VatSum。
                ' 基本上 SyncGridDataToModel 負責將畫面上的值(包含使用者手動改的)同步回 Model。
            End If
        Next
        CurrentLines = lines
    End Sub
#End Region

#Region "MDR GridView"
    Protected Sub btnAddMDRRow_Click(sender As Object, e As EventArgs)
        hfActiveTab.Value = "mdr" ' 保持在 MDR 頁籤
        SyncMDRGridToModel()
        Dim lines = CurrentMDRLines
        Dim newNum As Integer = 1
        If lines.Count > 0 Then newNum = lines.Max(Function(x) x.LineNum) + 1

        lines.Add(New MDRLine With {
            .LineNum = newNum,
            .U_LIFNR = txtCardCode.Text, ' 預設帶入表頭供應商
            .U_BLDAT = DateTime.Now,
            .U_VATDATE = DateTime.Now,
            .U_HWBAS = 0,
            .U_HWSTE = 0,
            .U_TAX_TYPE = "1", ' 預設應稅
            .U_ZFORM_CODE = "21" ' 預設發票類型
        })
        CurrentMDRLines = lines
        BindMDRGrid()
    End Sub

    Protected Sub btnDeleteMDRRow_Click(sender As Object, e As EventArgs)
        hfActiveTab.Value = "mdr" ' 保持在 MDR 頁籤
        SyncMDRGridToModel()
        Dim lines = CurrentMDRLines
        Dim newLines As New List(Of MDRLine)

        For Each row As GridViewRow In gvMDRDetail.Rows
            If row.RowType = DataControlRowType.DataRow Then
                Dim chk As CheckBox = CType(row.FindControl("chkSelectMDR"), CheckBox)
                If chk IsNot Nothing AndAlso Not chk.Checked Then
                    Dim idx As Integer = row.DataItemIndex
                    If idx < lines.Count Then newLines.Add(lines(idx))
                End If
            End If
        Next

        ' 重算 LineNum
        For i As Integer = 0 To newLines.Count - 1
            newLines(i).LineNum = i + 1
        Next

        CurrentMDRLines = newLines
        BindMDRGrid()
    End Sub

    Protected Sub btnGenerateMDR_Click(sender As Object, e As EventArgs)
        ' 產生憑證明細: 依據費用申請明細自動產生對應的憑證明細
        hfActiveTab.Value = "mdr" ' 切換到憑證明細頁籤
        SyncGridDataToModel()
        SyncMDRGridToModel()

        Dim mdrLines = CurrentMDRLines
        Dim expenseLines = CurrentLines

        If expenseLines.Count = 0 Then
            ShowError("請先新增費用申請明細")
            Return
        End If

        Dim startNum As Integer = 1
        If mdrLines.Count > 0 Then startNum = mdrLines.Max(Function(x) x.LineNum) + 1

        ' 依據費用明細產生對應的憑證明細
        For Each exp As ExpenseLine In expenseLines
            ' E030A (海關代徵營業稅) 預設格式28，其他項目預設格式21
            Dim defaultZFormCode As String = If(exp.CategoryCode = "E030A", "28", "21")

            mdrLines.Add(New MDRLine With {
                .LineNum = startNum,
                .U_LIFNR = txtCardCode.Text, ' 從單頭取得供應商
                .U_STCEG = "", ' 統一編號空白，需使用者填寫
                .U_XBLNR = "", ' 憑證號碼空白，需使用者填寫
                .U_BLDAT = DateTime.Now,
                .U_VATDATE = DateTime.Now,
                .U_HWBAS = exp.LineTotal, ' 帶入未稅金額
                .U_HWSTE = exp.VatSum, ' 帶入稅額
                .U_TAX_TYPE = If(exp.VatGroup = "1", "1", "2"), ' 1-應稅, 2-零稅
                .U_ZFORM_CODE = defaultZFormCode
            })
            startNum += 1
        Next

        CurrentMDRLines = mdrLines
        BindMDRGrid()
        ShowInfo("已產生 " & expenseLines.Count.ToString() & " 筆憑證明細，請填寫統一編號與憑證號碼")
    End Sub

    Private Sub BindMDRGrid()
        gvMDRDetail.DataSource = CurrentMDRLines
        gvMDRDetail.DataBind()
    End Sub

    Protected Sub gvMDRDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        ' 簡單綁定，大部分是 TextBox 直接在前端顯示
    End Sub

    Protected Sub CalculateMDRTotal(sender As Object, e As EventArgs)
        SyncMDRGridToModel(False)
        BindMDRGrid()
    End Sub

    Protected Sub CalculateMDRTaxManual(sender As Object, e As EventArgs)
        ' 當稅額被手動修改時
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
        Dim rowIndex As Integer = row.DataItemIndex

        SyncMDRGridToModel(True) ' 手動修改稅額

        ' 檢查用戶輸入的稅額是否與系統計算值不同
        Dim lines = CurrentMDRLines
        If rowIndex < lines.Count Then
            Dim line = lines(rowIndex)
            If line.U_TAX_TYPE = "1" AndAlso line.U_HWBAS > 0 Then
                ' 格式28 (海關代徵營業稅) 使用無條件捨去，其他格式使用四捨五入
                Dim calculatedVat As Decimal
                If line.U_ZFORM_CODE = "28" Then
                    calculatedVat = Math.Floor(line.U_HWBAS * 0.05D)
                Else
                    calculatedVat = Math.Round(line.U_HWBAS * 0.05D, 0, MidpointRounding.AwayFromZero)
                End If

                If line.U_HWSTE <> calculatedVat Then
                    ' 顯示警示：稅額與系統計算不同
                    Dim script As String = String.Format(
                        "alert('注意：您輸入的稅額 ({0:N0}) 與系統計算值 ({1:N0}) 不同。\n\n" &
                        "若憑證上的稅額確實為 {0:N0} 元，請忽略此訊息。\n\n" &
                        "系統將保留您輸入的稅額。');",
                        line.U_HWSTE, calculatedVat)
                    ScriptManager.RegisterStartupScript(Me.Page, Me.GetType(), "mdrVatDiffWarning" & rowIndex, script, True)
                End If
            End If
        End If

        BindMDRGrid()
    End Sub

    Private Sub SyncMDRGridToModel(Optional isManualTax As Boolean = False)
        Dim lines = CurrentMDRLines
        For i As Integer = 0 To gvMDRDetail.Rows.Count - 1
            Dim row As GridViewRow = gvMDRDetail.Rows(i)
            If i < lines.Count Then
                Dim line = lines(i)

                Dim txtSTCEG As TextBox = CType(row.FindControl("txtSTCEG"), TextBox)
                Dim txtXBLNR As TextBox = CType(row.FindControl("txtXBLNR"), TextBox)
                Dim ddlZFORM As DropDownList = CType(row.FindControl("ddlZFORM_CODE"), DropDownList)
                Dim txtBLDAT As TextBox = CType(row.FindControl("txtBLDAT"), TextBox)
                Dim txtVATDATE As TextBox = CType(row.FindControl("txtVATDATE"), TextBox)
                Dim txtHWBAS As TextBox = CType(row.FindControl("txtHWBAS"), TextBox)
                Dim txtHWSTE As TextBox = CType(row.FindControl("txtHWSTE"), TextBox)
                Dim ddlTAX As DropDownList = CType(row.FindControl("ddlTAX_TYPE"), DropDownList)

                ' 供應商從單頭取得，不再從 GridView 讀取
                If txtSTCEG IsNot Nothing Then line.U_STCEG = txtSTCEG.Text
                If txtXBLNR IsNot Nothing Then line.U_XBLNR = txtXBLNR.Text
                If ddlZFORM IsNot Nothing Then line.U_ZFORM_CODE = ddlZFORM.SelectedValue

                Dim dt As DateTime
                If txtBLDAT IsNot Nothing AndAlso DateTime.TryParse(txtBLDAT.Text, dt) Then line.U_BLDAT = dt
                If txtVATDATE IsNot Nothing AndAlso DateTime.TryParse(txtVATDATE.Text, dt) Then line.U_VATDATE = dt

                If txtHWBAS IsNot Nothing Then Decimal.TryParse(txtHWBAS.Text, line.U_HWBAS)
                If ddlTAX IsNot Nothing Then line.U_TAX_TYPE = ddlTAX.SelectedValue

                ' 讀取 UI 上當前的稅額（用戶可能已手動修改）
                Dim currentHWSTE As Decimal = 0
                If txtHWSTE IsNot Nothing Then
                    Decimal.TryParse(txtHWSTE.Text, currentHWSTE)
                End If

                ' 稅額處理邏輯：
                ' - 非應稅項目（2-零稅、3-免稅）：稅額強制為 0
                ' - 應稅項目（1-應稅）：
                '   * 如果稅額為 0（新增明細），自動計算
                '   * 如果稅額不為 0（已有值），保留用戶輸入的稅額
                '     用戶可能根據實際憑證調整稅額（進位/捨去差異）
                If line.U_TAX_TYPE <> "1" Then
                    ' 非應稅項目，稅額必須為 0
                    line.U_HWSTE = 0
                ElseIf currentHWSTE = 0 AndAlso line.U_HWBAS > 0 Then
                    ' 新增明細或清空稅額，自動計算
                    ' 格式28 (海關代徵營業稅) 使用無條件捨去，其他格式使用四捨五入
                    If line.U_ZFORM_CODE = "28" Then
                        line.U_HWSTE = Math.Floor(line.U_HWBAS * 0.05D)
                    Else
                        line.U_HWSTE = Math.Round(line.U_HWBAS * 0.05D, 0, MidpointRounding.AwayFromZero)
                    End If
                Else
                    ' 已有稅額，保留用戶輸入的值
                    line.U_HWSTE = currentHWSTE
                End If

                ' 供應商從單頭取得 (不再從 GridView 讀取)
                line.U_LIFNR = txtCardCode.Text
            End If
        Next
        CurrentMDRLines = lines
    End Sub

    Protected Sub txtXBLNR_TextChanged(sender As Object, e As EventArgs)
        ' 當發票號碼改變時，自動判斷類型
        hfActiveTab.Value = "mdr" ' 保持在 MDR 頁籤
        
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
        Dim ddlZForm As DropDownList = CType(row.FindControl("ddlZFORM_CODE"), DropDownList)

        Dim invNum As String = txt.Text.Trim().ToUpper()
        Dim formCode As String = "99" ' Default: 其他
        Dim needWarning As Boolean = False
        Dim prefix As String = "" ' 發票字軌 (前2碼)

        ' 判斷是否符合統一發票格式: 2碼英文 + 8碼數字
        If System.Text.RegularExpressions.Regex.IsMatch(invNum, "^[A-Z]{2}\d{8}$") Then
            prefix = invNum.Substring(0, 2)
            formCode = GetFormCodeByPrefix(prefix, needWarning)
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^[A-Z]{3}") Then
            formCode = "28" ' 海關代徵營業稅
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^(BB|BBN)") Then
            formCode = "22" ' 高鐵/二聯收銀機（長條）
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^\d+$") Then
            formCode = "22" ' 高鐵/二聯收銀機（長條）- 純數字
        Else
            formCode = "99" ' 其他
        End If

        If ddlZForm IsNot Nothing Then
            If ddlZForm.Items.FindByValue(formCode) IsNot Nothing Then
                ddlZForm.SelectedValue = formCode
            End If
        End If

        ' 格式28 (海關代徵營業稅) 特殊處理
        If formCode = "28" Then
            HandleFormat28Row(row)
        End If

        ' 如果需要警告（字軌不在當年度表中）
        If needWarning AndAlso Not String.IsNullOrEmpty(prefix) Then
            ' 使用 JavaScript confirm 彈窗
            Dim script As String = String.Format(
                "if(!confirm('發票字軌 {0} 不在114年度發票字軌表中，是否仍要使用此字軌？')) {{ document.getElementById('{1}').value = ''; }}",
                prefix, txt.ClientID)
            ScriptManager.RegisterStartupScript(Me.Page, Me.GetType(), "invWarning_" & txt.ClientID, script, True)
        End If

        SyncMDRGridToModel()
    End Sub

    ''' <summary>
    ''' 格式28 (海關代徵營業稅/進出口報關費) 特殊處理
    ''' - 設定 placeholder 提示使用者
    ''' - 不清空已輸入的金額
    ''' - 同一次進入單據期間，憑證頁籤最多警告一次
    ''' </summary>
    Private Sub HandleFormat28Row(row As GridViewRow)
        Dim txtHWBAS As TextBox = CType(row.FindControl("txtHWBAS"), TextBox)
        Dim txtHWSTE As TextBox = CType(row.FindControl("txtHWSTE"), TextBox)

        ' 只設定 placeholder 提示，不清空已輸入的金額
        If txtHWBAS IsNot Nothing AndAlso String.IsNullOrEmpty(txtHWBAS.Text) Then
            txtHWBAS.Attributes("placeholder") = "請輸入營業稅基"
        End If

        If txtHWSTE IsNot Nothing AndAlso String.IsNullOrEmpty(txtHWSTE.Text) Then
            txtHWSTE.Attributes("placeholder") = "請輸入稅額"
        End If

        ' 同一次進入單據期間，憑證頁籤最多警告一次
        If Not Format28MDRWarningShown Then
            Dim script As String = "alert('此為格式28 (海關代徵營業稅)：\n\n" &
                                  "1. 「未稅金額」欄位請填入「營業稅基」\n" &
                                  "2. 「稅額」欄位請填入實際稅額\n" &
                                  "3. 匯入SAP時將只過帳稅額金額\n\n" &
                                  "請核對海關稅單上的營業稅基與稅額。');"
            ScriptManager.RegisterStartupScript(Me.Page, Me.GetType(), "format28Warning", script, True)
            Format28MDRWarningShown = True
        End If
    End Sub


    ''' <summary>
    ''' 根據發票字軌取得對應的憑證類型代碼
    ''' 114年度發票字軌對照表
    ''' </summary>
    Private Function GetFormCodeByPrefix(prefix As String, ByRef needWarning As Boolean) As String
        needWarning = False
        
        ' 114年度 甲種統一發票字軌 (三聯式手開) -> 21
        Dim typeA As String() = {
            "HT", "HU", "KT", "KU", "MT", "MU", 
            "PT", "PU", "RT", "RU", "TT", "TU"
        }
        
        ' 114年度 乙種統一發票字軌 (二聯式手開) -> 22
        Dim typeB As String() = {
            "HV", "HW", "HX", "KV", "KW", "KX", 
            "MV", "MW", "MX", "PV", "PW", "PX", 
            "RV", "RW", "RX", "TV", "TW", "TX"
        }
        
        ' 114年度 丙種統一發票字軌 (收銀機) -> 22
        Dim typeC As String() = {
            "HY", "KY", "MY", "PY", "RY", "TY"
        }
        
        ' 114年度 丁種統一發票字軌 (電子發票) -> 25
        ' 期別1-2月: HZ, JA-JV, JW-KS
        ' 期別3-4月: KZ, LA-LV, LW-MS
        ' 期別5-6月: MZ, NA-NV, NW-PS
        ' 期別7-8月: PZ, QA-QV, QW-RS
        ' 期別9-10月: RZ, SA-SV, SW-TS
        ' 期別11-12月: TZ, UA-UV, UW-VS
        Dim typeD As String() = {
            "HZ", "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JJ", "JK", "JL", "JM", "JN", "JP", "JQ", "JR", "JS", "JT", "JU", "JV",
            "JW", "JX", "JY", "JZ", "KA", "KB", "KC", "KD", "KE", "KF", "KG", "KH", "KJ", "KK", "KL", "KM", "KN", "KP", "KQ", "KR", "KS",
            "KZ", "LA", "LB", "LC", "LD", "LE", "LF", "LG", "LH", "LJ", "LK", "LL", "LM", "LN", "LP", "LQ", "LR", "LS", "LT", "LU", "LV",
            "LW", "LX", "LY", "LZ", "MA", "MB", "MC", "MD", "ME", "MF", "MG", "MH", "MJ", "MK", "ML", "MM", "MN", "MP", "MQ", "MR", "MS",
            "MZ", "NA", "NB", "NC", "ND", "NE", "NF", "NG", "NH", "NJ", "NK", "NL", "NM", "NN", "NP", "NQ", "NR", "NS", "NT", "NU", "NV",
            "NW", "NX", "NY", "NZ", "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PJ", "PK", "PL", "PM", "PN", "PP", "PQ", "PR", "PS",
            "PZ", "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QJ", "QK", "QL", "QM", "QN", "QP", "QQ", "QR", "QS", "QT", "QU", "QV",
            "QW", "QX", "QY", "QZ", "RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RJ", "RK", "RL", "RM", "RN", "RP", "RQ", "RR", "RS",
            "RZ", "SA", "SB", "SC", "SD", "SE", "SF", "SG", "SH", "SJ", "SK", "SL", "SM", "SN", "SP", "SQ", "SR", "SS", "ST", "SU", "SV",
            "SW", "SX", "SY", "SZ", "TA", "TB", "TC", "TD", "TE", "TF", "TG", "TH", "TJ", "TK", "TL", "TM", "TN", "TP", "TQ", "TR", "TS",
            "TZ", "UA", "UB", "UC", "UD", "UE", "UF", "UG", "UH", "UJ", "UK", "UL", "UM", "UN", "UP", "UQ", "UR", "US", "UT", "UU", "UV",
            "UW", "UX", "UY", "UZ", "VA", "VB", "VC", "VD", "VE", "VF", "VG", "VH", "VJ", "VK", "VL", "VM", "VN", "VP", "VQ", "VR", "VS"
        }

        
        ' 檢查各類型
        If typeA.Contains(prefix) Then
            Return "21" ' 甲種 -> 三聯式統一發票
        ElseIf typeB.Contains(prefix) OrElse typeC.Contains(prefix) Then
            Return "22" ' 乙種/丙種 -> 二聯式/收銀機發票
        ElseIf typeD.Contains(prefix) Then
            Return "25" ' 丁種 -> 電子發票
        Else
            ' 不在114年度字軌表中
            needWarning = True
            Return "99" ' 其他
        End If
    End Function

#End Region

#Region "總計與表尾"
    Private Sub CalculateFooterTotal()
        Dim docTotal As Decimal = CurrentLines.Sum(Function(x) x.LineTotal)
        Dim vatSum As Decimal = CurrentLines.Sum(Function(x) x.VatSum)
        Dim gTotal As Decimal = CurrentLines.Sum(Function(x) x.PriceAfterVat)

        lblDocTotal.Text = docTotal.ToString("N2")
        lblVatSum.Text = vatSum.ToString("N2")
        lblDocTotalWithTax.Text = gTotal.ToString("N2")
    End Sub
#End Region

#Region "存檔與讀取"
    ''' <summary>
    ''' [B] 暫存按鈕 - 已移除草稿狀態，此按鈕保留供編輯已退回的單據使用
    ''' 退回的單據可以修改後再次送審
    ''' </summary>
    Protected Sub btnSave_Click(sender As Object, e As EventArgs)
        ViewState("PendingAction") = "save"
        If Not ValidateAll() Then Return
        ' 若是被退回的單據 (R)，則儲存時維持退回狀態，等待使用者修改後重新送審
        ' 若是待審核的單據 (W)，則維持待審核狀態
        Dim currentStatus As String = txtApprovalStatus.Text
        If String.IsNullOrEmpty(currentStatus) OrElse currentStatus = "P" Then
            currentStatus = "W" ' 新單據直接改為待審核
        End If
        SaveDocument(currentStatus)
    End Sub

    ''' <summary>
    ''' [B] 儲存並送審 - 直接進入待審核狀態
    ''' </summary>
    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs)
        ViewState("PendingAction") = "submit"
        If Not ValidateAll() Then Return
        SaveDocument("W") ' W = 待審核
    End Sub



    Private Property WarningConfirmed As Boolean
        Get
            If ViewState("WarningConfirmed") Is Nothing Then Return False
            Return Convert.ToBoolean(ViewState("WarningConfirmed"))
        End Get
        Set(value As Boolean)
            ViewState("WarningConfirmed") = value
        End Set
    End Property

    ''' <summary>
    ''' [A] 放行時的金額不一致警告確認狀態
    ''' </summary>
    Private Property ApprovalWarningConfirmed As Boolean
        Get
            If ViewState("ApprovalWarningConfirmed") Is Nothing Then Return False
            Return Convert.ToBoolean(ViewState("ApprovalWarningConfirmed"))
        End Get
        Set(value As Boolean)
            ViewState("ApprovalWarningConfirmed") = value
        End Set
    End Property

    ''' <summary>
    ''' E030 (進出口報關費用-其他) 提示警告是否已顯示過
    ''' </summary>
    Private Property E030OtherWarningShown As Boolean
        Get
            If ViewState("E030OtherWarningShown") Is Nothing Then Return False
            Return Convert.ToBoolean(ViewState("E030OtherWarningShown"))
        End Get
        Set(value As Boolean)
            ViewState("E030OtherWarningShown") = value
        End Set
    End Property

    ''' <summary>
    ''' 格式28 (E030A/海關代徵營業稅) 費用頁籤警告是否已顯示過
    ''' </summary>
    Private Property Format28ExpenseWarningShown As Boolean
        Get
            If ViewState("Format28ExpenseWarningShown") Is Nothing Then Return False
            Return Convert.ToBoolean(ViewState("Format28ExpenseWarningShown"))
        End Get
        Set(value As Boolean)
            ViewState("Format28ExpenseWarningShown") = value
        End Set
    End Property

    ''' <summary>
    ''' 格式28 (E030A/海關代徵營業稅) 憑證頁籤警告是否已顯示過
    ''' </summary>
    Private Property Format28MDRWarningShown As Boolean
        Get
            If ViewState("Format28MDRWarningShown") Is Nothing Then Return False
            Return Convert.ToBoolean(ViewState("Format28MDRWarningShown"))
        End Get
        Set(value As Boolean)
            ViewState("Format28MDRWarningShown") = value
        End Set
    End Property

    ''' <summary>
    ''' 驗證單據，收集錯誤與警告，並顯示彈窗
    ''' </summary>
    ''' <returns>True 表示可以儲存，False 表示需要修正或確認</returns>
    Private Function ValidateAll() As Boolean
        ' 先同步 GridView 資料到 Model，確保驗證的是最新資料
        SyncGridDataToModel()
        SyncMDRGridToModel()

        Dim errors As New List(Of String)()
        Dim warnings As New List(Of String)()

        ' 清除欄位旁的錯誤訊息
        lblMessage.Text = ""
        lblErrCardCode.Visible = False
        lblErrDocDate.Visible = False
        lblErrDocDueDate.Visible = False
        lblErrTaxDate.Visible = False

        ' === 錯誤檢查 (必須修正) ===
        If String.IsNullOrEmpty(txtCardCode.Text) Then
            errors.Add("供應商代碼為必填欄位")
        End If

        If String.IsNullOrEmpty(txtDocDate.Text) Then
            errors.Add("過帳日期為必填欄位")
        End If

        If String.IsNullOrEmpty(txtDocDueDate.Text) Then
            errors.Add("到期日為必填欄位")
        End If

        If String.IsNullOrEmpty(txtTaxDate.Text) Then
            errors.Add("文件日期為必填欄位")
        End If

        ' 明細數量檢查
        If CurrentLines.Count = 0 Then
            If CurrentMDRLines.Count = 0 Then
                errors.Add("請至少新增一筆費用明細")
            Else
                errors.Add("禁止僅有憑證明細而無費用明細，請新增費用明細")
            End If
        Else
            ' 有費用明細
            If CurrentMDRLines.Count = 0 Then
                warnings.Add("目前只有費用明細但沒有憑證明細，請確認是否要新增")
            End If
        End If

        Dim docTotal As Decimal = 0
        Decimal.TryParse(lblDocTotalWithTax.Text.Replace(",", ""), docTotal)
        If docTotal <= 0 Then
            errors.Add("單據總額不可為 0，請確認費用明細金額")
        End If

        ' 檢查憑證明細是否有空白的統一編號或憑證號碼 (除非是「99-其他」類型)
        Dim emptyVoucherLines = CurrentMDRLines.Where(Function(x) _
            (String.IsNullOrEmpty(x.U_STCEG) OrElse String.IsNullOrEmpty(x.U_XBLNR)) _
            AndAlso x.U_ZFORM_CODE <> "99")

        If emptyVoucherLines.Any() Then
            errors.Add("憑證明細有空白的統一編號或憑證號碼，請填寫完整。若為非發票類憑證，請選擇憑證類型為「其他」")
        End If

        ' === 警告檢查 (可確認後繼續) ===

        ' 費用明細與憑證明細金額一致性警告
        If CurrentLines.Count > 0 AndAlso CurrentMDRLines.Count > 0 Then
            Dim expenseTotal As Decimal = CurrentLines.Sum(Function(x) x.LineTotal)
            Dim expenseVatSum As Decimal = CurrentLines.Sum(Function(x) x.VatSum)
            Dim mdrTotal As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWBAS)
            Dim mdrVatSum As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWSTE)

            If Math.Abs(expenseTotal - mdrTotal) > 0.01D Then
                warnings.Add(String.Format("費用明細未稅總額 ({0:N0}) 與憑證明細未稅總額 ({1:N0}) 不一致", expenseTotal, mdrTotal))
            End If
            If Math.Abs(expenseVatSum - mdrVatSum) > 0.01D Then
                warnings.Add(String.Format("費用明細稅額 ({0:N0}) 與憑證明細稅額 ({1:N0}) 不一致", expenseVatSum, mdrVatSum))
            End If
        End If

        ' 匯率日期檢查
        Dim curr As String = ddlDocCurrency.SelectedValue
        If curr <> "TWD" AndAlso curr <> "NTD" Then
            If Not String.IsNullOrEmpty(hfRateDate.Value) Then
                Dim rateDate As DateTime
                If DateTime.TryParse(hfRateDate.Value, rateDate) Then
                    Dim daysDiff As Integer = CInt((DateTime.Today - rateDate).TotalDays)
                    If daysDiff > 0 Then
                        warnings.Add(String.Format("匯率資料為 {0:yyyy-MM-dd}，距今已 {1} 天，請確認匯率是否正確", rateDate, daysDiff))
                    End If
                End If
            ElseIf String.IsNullOrEmpty(hfRateDate.Value) AndAlso txtDocRate.Text <> "1.0" Then
                warnings.Add("無法確認匯率日期，請確認匯率是否正確")
            End If
        End If

        ' 檢查是否有 "99-其他" 類型的憑證
        Dim hasOtherType As Boolean = CurrentMDRLines.Any(Function(x) x.U_ZFORM_CODE = "99")
        If hasOtherType Then
            warnings.Add("有「其他」類型的憑證，請確認憑證格式是否正確")
        End If

        ' E030A (海關代徵營業稅) 必須有對應的 28 類型憑證
        Dim hasE030AExpense As Boolean = CurrentLines.Any(Function(x) x.CategoryCode = "E030A")
        Dim hasFormat28MDR As Boolean = CurrentMDRLines.Any(Function(x) x.U_ZFORM_CODE = "28")

        ' 錯誤檢查：E030A 必須有格式28憑證
        If hasE030AExpense AndAlso Not hasFormat28MDR Then
            errors.Add("費用明細「進出口報關費用-28-海關代徵營業稅」項目，必須要有對應的 28 類型憑證明細")
        End If

        ' 格式28 提醒
        If hasE030AExpense OrElse hasFormat28MDR Then
            warnings.Add("【格式28 提醒】此單據含海關代徵營業稅，匯入SAP時將只過帳「稅額」金額，請確認金額正確")
        End If

        ' === 決定是否顯示彈窗 ===

        ' 如果已經確認過警告，且沒有新的錯誤，則允許通過
        If WarningConfirmed AndAlso errors.Count = 0 Then
            Return True
        End If

        ' 如果有錯誤或警告，顯示彈窗
        If errors.Count > 0 OrElse warnings.Count > 0 Then
            ShowValidationPopup(errors, warnings)
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' 顯示驗證結果彈窗
    ''' </summary>
    Private Sub ShowValidationPopup(errors As List(Of String), warnings As List(Of String))
        ' DEBUG: 顯示在 lblMessage
        lblMessage.Text = String.Format("[DEBUG] Errors:{0}, Warnings:{1}, WarningConfirmed:{2}", errors.Count, warnings.Count, WarningConfirmed)
        lblMessage.ForeColor = System.Drawing.Color.Blue

        ' 清空之前的內容
        blErrors.Items.Clear()
        blWarnings.Items.Clear()

        ' 顯示錯誤
        If errors.Count > 0 Then
            pnlErrors.Visible = True
            For Each err As String In errors
                blErrors.Items.Add(err)
            Next
        Else
            pnlErrors.Visible = False
        End If

        ' 顯示警告
        If warnings.Count > 0 Then
            pnlWarnings.Visible = True
            For Each warn As String In warnings
                blWarnings.Items.Add(warn)
            Next
        Else
            pnlWarnings.Visible = False
        End If

        ' 設定彈窗標題顏色和按鈕
        If errors.Count > 0 Then
            ' 有錯誤：紅色標題，只有「返回修改」按鈕
            divValidationHeader.Style("background-color") = "#dc3545"
            btnValidationConfirm.Visible = False
        Else
            ' 只有警告：橘色標題，顯示確認按鈕
            divValidationHeader.Style("background-color") = "#ffc107"
            btnValidationConfirm.Visible = True

            ' 根據操作類型設定按鈕文字
            Dim pendingAction As String = If(ViewState("PendingAction"), "save").ToString()
            Select Case pendingAction
                Case "submit"
                    btnValidationConfirm.Text = "確定送審"
                Case "approve"
                    btnValidationConfirm.Text = "確定放行"
                Case Else
                    btnValidationConfirm.Text = "確定儲存"
            End Select
        End If

        ' 顯示彈窗
        mpeValidation.Show()
    End Sub

    ''' <summary>
    ''' 驗證彈窗 - 返回修改按鈕 (包含 X 按鈕)
    ''' </summary>
    Protected Sub btnValidationBack_Click(sender As Object, e As EventArgs)
        ' 重置警告確認狀態
        WarningConfirmed = False
        ViewState("PendingAction") = Nothing
        ' 使用 JavaScript 隱藏彈窗，避免 ModalPopupExtender 狀態問題
        ScriptManager.RegisterStartupScript(Me.Page, Me.GetType(), "hideValidation",
            "$find('mpeValidationBehavior').hide();", True)
    End Sub

    ''' <summary>
    ''' 驗證彈窗 - 確定按鈕（統一處理送審、儲存、放行）
    ''' </summary>
    Protected Sub btnValidationConfirm_Click(sender As Object, e As EventArgs)
        ' 設定警告已確認
        WarningConfirmed = True
        ApprovalWarningConfirmed = True

        ' 根據之前的操作類型執行對應動作
        Dim pendingAction As String = If(ViewState("PendingAction"), "save").ToString()
        Select Case pendingAction
            Case "submit"
                SaveDocument("W")
            Case "approve"
                ' 放行操作
                UpdateStatus("A")
            Case Else
                ' 一般儲存
                Dim currentStatus As String = txtApprovalStatus.Text
                If String.IsNullOrEmpty(currentStatus) OrElse currentStatus = "P" Then
                    currentStatus = "W"
                End If
                SaveDocument(currentStatus)
        End Select
    End Sub

    Private Function SaveDocument(status As String, Optional isAutoSave As Boolean = False) As Boolean
        Try
            SyncGridDataToModel()
            SyncMDRGridToModel()

            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using trans = conn.BeginTransaction()
                    Try
                        Dim jID As Integer = 0
                        Dim isNewDocument As Boolean = (currentJID = 0)  ' 記錄是否為新增，供稽核日誌使用

                        ' 1. jOPCH (Header)
                        If currentJID = 0 Then
                            ' Insert
                            Dim sqlH As String = "INSERT INTO jOPCH (CardCode, CardName, NumAtCard, InvNum, DeliveryAddrID, AddressName, Address, " &
                                               "DocDate, DocDueDate, TaxDate, DocCurrency, DocRate, DocTotal, VatSum, " &
                                               "GroupNum, PymntGroup, Comments, ApprovalStatus, CreateBy, CreateDate, U_PID, SlpCode) " &
                                               "VALUES (@CardCode, @CardName, @NumAtCard, @InvNum, @DeliveryAddrID, @AddressName, @Address, " &
                                               "@DocDate, @DocDueDate, @TaxDate, @DocCurrency, @DocRate, @DocTotal, @VatSum, " &
                                               "@GroupNum, @PymntGroup, @Comments, @Status, @User, GETDATE(), @UPID, @SlpCode); " &
                                               "SELECT SCOPE_IDENTITY();"

                            Using cmd As New SqlCommand(sqlH, conn, trans)
                                SetHeaderParameters(cmd, status)
                                jID = Convert.ToInt32(cmd.ExecuteScalar())
                            End Using
                            ' 注意：DocEntry/DocNum 欄位不在此設定，留給 SAP 回寫
                            currentJID = jID
                        Else
                            ' Update (使用 jID 作為主鍵)
                            jID = currentJID
                            Dim sqlH As String = "UPDATE jOPCH SET CardCode=@CardCode, CardName=@CardName, NumAtCard=@NumAtCard, InvNum=@InvNum, " &
                                               "DeliveryAddrID=@DeliveryAddrID, AddressName=@AddressName, Address=@Address, " &
                                               "DocDate=@DocDate, DocDueDate=@DocDueDate, TaxDate=@TaxDate, DocCurrency=@DocCurrency, " &
                                               "DocRate=@DocRate, DocTotal=@DocTotal, VatSum=@VatSum, " &
                                               "GroupNum=@GroupNum, PymntGroup=@PymntGroup, Comments=@Comments, ApprovalStatus=@Status, " &
                                               "UpdateBy=@User, UpdateDate=GETDATE(), U_PID=@UPID, SlpCode=@SlpCode WHERE jID=@jID"

                            Using cmd As New SqlCommand(sqlH, conn, trans)
                                cmd.Parameters.AddWithValue("@jID", currentJID)
                                SetHeaderParameters(cmd, status)
                                cmd.ExecuteNonQuery()
                            End Using

                            ' Delete old lines (使用 jID 作為關聯鍵)
                            Using cmd As New SqlCommand("DELETE FROM jPCH1 WHERE jID=@jID", conn, trans)
                                cmd.Parameters.AddWithValue("@jID", currentJID)
                                cmd.ExecuteNonQuery()
                            End Using

                            ' Delete old MDR (使用 jID 作為關聯鍵)
                            Using cmd As New SqlCommand("DELETE FROM jMGUIAP WHERE jID=@jID", conn, trans)
                                cmd.Parameters.AddWithValue("@jID", currentJID)
                                cmd.ExecuteNonQuery()
                            End Using
                            Using cmd As New SqlCommand("DELETE FROM jMGUIAPDetail WHERE jID=@jID", conn, trans)
                                cmd.Parameters.AddWithValue("@jID", currentJID)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If

                        ' 2. 複製附件 (Copy Mode)
                        If isNewDocument AndAlso CopyFromJID > 0 AndAlso CopyAttachments Then
                            CopyAttachmentsFromSource(CopyFromJID, jID, conn, trans)
                        End If

                        ' 3. jPCH1 (Expense Lines)
                        ' 注意：DocEntry/DocNum 欄位不在此設定，留給 SAP 回寫
                        Dim sqlL As String = "INSERT INTO jPCH1 (jID, LineNum, ItemCode, Dscription, AcctCode, " &
                                           "LineTotal, VatGroup, VatPrcnt, LineVat, GTotal, CostingCode, CostingCode2, Currency, Rate) " &
                                           "VALUES (@jID, @LineNum, @ItemCode, @Dscription, @AcctCode, " &
                                           "@LineTotal, @VatGroup, @VatPrcnt, @LineVat, @GTotal, @CostingCode, @CostingCode2, @Currency, @Rate)"

                        For Each line As ExpenseLine In CurrentLines
                            Using cmd As New SqlCommand(sqlL, conn, trans)
                                cmd.Parameters.AddWithValue("@jID", jID) ' FK to Header jID
                                cmd.Parameters.AddWithValue("@LineNum", line.LineNum)
                                cmd.Parameters.AddWithValue("@ItemCode", line.CategoryCode)
                                cmd.Parameters.AddWithValue("@Dscription", line.Description)
                                cmd.Parameters.AddWithValue("@AcctCode", line.AcctCode)
                                cmd.Parameters.AddWithValue("@LineTotal", line.LineTotal)
                                cmd.Parameters.AddWithValue("@VatGroup", line.VatGroup)
                                cmd.Parameters.AddWithValue("@VatPrcnt", line.VatRate)
                                cmd.Parameters.AddWithValue("@LineVat", line.VatSum)
                                cmd.Parameters.AddWithValue("@GTotal", line.PriceAfterVat)
                                cmd.Parameters.AddWithValue("@CostingCode", line.CostingCode)
                                cmd.Parameters.AddWithValue("@CostingCode2", line.CostingCode2)
                                cmd.Parameters.AddWithValue("@Currency", line.Currency)
                                cmd.Parameters.AddWithValue("@Rate", line.Rate)
                                cmd.ExecuteNonQuery()
                            End Using
                        Next

                        ' 4. jMGUIAP & jMGUIAPDetail (MDR)
                        If CurrentMDRLines.Count > 0 Then
                            ' MDR Header (彙總)
                            ' 注意：DocEntry/DocNum 欄位不在此設定，留給 SAP 回寫
                            Dim mdrTotal As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWBAS)
                            Dim mdrVat As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWSTE)

                            Dim sqlMdrH As String = "INSERT INTO jMGUIAP (jID, DocTotal, VatSum, CreateBy, CreateDate) " &
                                                  "VALUES (@jID, @DocTotal, @VatSum, @User, GETDATE())"

                            Using cmd As New SqlCommand(sqlMdrH, conn, trans)
                                cmd.Parameters.AddWithValue("@jID", jID)
                                cmd.Parameters.AddWithValue("@DocTotal", mdrTotal)
                                cmd.Parameters.AddWithValue("@VatSum", mdrVat)
                                cmd.Parameters.AddWithValue("@User", currentUserId)
                                cmd.ExecuteNonQuery()
                            End Using

                            ' MDR Lines
                            ' 注意：jMGUIAPDetail.jID 應該對應 jOPCH.jID，而非 jMGUIAP.ID
                            Dim sqlMdrL As String = "INSERT INTO jMGUIAPDetail (jID, LineNum, U_LIFNR, U_STCEG, U_XBLNR, U_ZFORM_CODE, " &
                                                  "U_BLDAT, U_VATDATE, U_HWBAS, U_HWSTE, U_TAX_TYPE) " &
                                                  "VALUES (@jID, @LineNum, @LIFNR, @STCEG, @XBLNR, @ZFORM, @BLDAT, @VATDATE, @HWBAS, @HWSTE, @TAXTYPE)"

                            For Each line As MDRLine In CurrentMDRLines
                                Using cmd As New SqlCommand(sqlMdrL, conn, trans)
                                    cmd.Parameters.AddWithValue("@jID", jID)  ' 使用 jOPCH.jID 作為關聯鍵
                                    cmd.Parameters.AddWithValue("@LineNum", line.LineNum)
                                    cmd.Parameters.AddWithValue("@LIFNR", line.U_LIFNR)
                                    cmd.Parameters.AddWithValue("@STCEG", line.U_STCEG)
                                    cmd.Parameters.AddWithValue("@XBLNR", line.U_XBLNR)
                                    cmd.Parameters.AddWithValue("@ZFORM", line.U_ZFORM_CODE)
                                    cmd.Parameters.AddWithValue("@BLDAT", If(line.U_BLDAT.HasValue, line.U_BLDAT.Value, DBNull.Value))
                                    cmd.Parameters.AddWithValue("@VATDATE", If(line.U_VATDATE.HasValue, line.U_VATDATE.Value, DBNull.Value))
                                    cmd.Parameters.AddWithValue("@HWBAS", line.U_HWBAS)
                                    cmd.Parameters.AddWithValue("@HWSTE", line.U_HWSTE)
                                    cmd.Parameters.AddWithValue("@TAXTYPE", line.U_TAX_TYPE)
                                    cmd.ExecuteNonQuery()
                                End Using
                            Next
                        End If

                        ' Update U_LastCur in OCRD if needed (選用功能)
                        ' (若表頭幣別為多幣別的業務夥伴，每次新增單據時更新 U_LastCur)
                        ' 注意：需要先在 SAP 中建立 OCRD.U_LastCur UDF，否則會跳過此功能
                        If ddlDocCurrency.Enabled Then ' 表示該業務夥伴為多幣別 (##)
                            Try
                                Dim sqlUpdCur As String = "UPDATE OCRD SET U_LastCur = @Curr WHERE CardCode = @CardCode"
                                Using sapConn As New SqlConnection(sapConnStr)
                                    sapConn.Open()
                                    Using cmdSap As New SqlCommand(sqlUpdCur, sapConn)
                                        cmdSap.Parameters.AddWithValue("@Curr", ddlDocCurrency.SelectedValue)
                                        cmdSap.Parameters.AddWithValue("@CardCode", txtCardCode.Text)
                                        cmdSap.ExecuteNonQuery()
                                    End Using
                                End Using
                            Catch ex As Exception
                                ' U_LastCur UDF 可能不存在，忽略此錯誤 (非核心功能)
                            End Try
                        End If

                        trans.Commit()
                        ShowSuccess("儲存成功")
                        lblDocNum.Text = currentJID.ToString()

                        ' [E] 記錄稽核日誌 (非阻塞)
                        Dim action As String = If(isNewDocument, AuditLogger.Actions.Create, AuditLogger.Actions.Update)
                        AuditLogger.Log("jOPCH", currentJID, action, currentUserId,
                                        changes:=String.Format("Status={0}, Total={1}", status, lblDocTotalWithTax.Text))

                        If Not isAutoSave Then
                            Response.Redirect("ExpenseClaimForm.aspx?jID=" & currentJID)
                        End If
                        Return True
                    Catch ex As Exception
                        trans.Rollback()
                        Throw ex
                    End Try
                End Using
            End Using
        Catch ex As Exception
            ShowError("儲存失敗: " & ex.Message)
            Return False
        End Try
    End Function
    Protected Sub btnDelete_Click(sender As Object, e As EventArgs)
        If currentJID > 0 Then
            Try
                ' 權限檢查：只有草稿(P)的擁有者可以刪除
                Dim docStatus As String = txtApprovalStatus.Text
                Dim docOwner As String = txtOwner.Text

                ' 檢查是否為草稿狀態
                If docStatus <> "P" Then
                    ShowError("只能刪除草稿狀態的單據")
                    Return
                End If

                ' 檢查是否為單據擁有者
                If docOwner <> currentUserId Then
                    ShowError("您只能刪除自己建立的單據")
                    Return
                End If

                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using trans = conn.BeginTransaction()
                        Try
                            ' Delete Lines (使用 jID 作為關聯鍵)
                            Dim cmd As New SqlCommand("DELETE FROM jPCH1 WHERE jID=@jID", conn, trans)
                            cmd.Parameters.AddWithValue("@jID", currentJID)
                            cmd.ExecuteNonQuery()

                            ' Delete MDR (使用 jID 作為關聯鍵)
                            cmd.CommandText = "DELETE FROM jMGUIAP WHERE jID=@jID"
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "DELETE FROM jMGUIAPDetail WHERE jID=@jID"
                            cmd.ExecuteNonQuery()

                            ' Delete Header (使用 jID 作為主鍵)
                            cmd.CommandText = "DELETE FROM jOPCH WHERE jID=@jID"
                            cmd.ExecuteNonQuery()

                            ' Delete Attachments
                            cmd.CommandText = "DELETE FROM jAttach WHERE jID=@jID"
                            cmd.ExecuteNonQuery()

                            trans.Commit()

                            ' [E] 記錄刪除稽核日誌 (非阻塞)
                            AuditLogger.Log("jOPCH", currentJID, AuditLogger.Actions.Delete, currentUserId,
                                            changes:=String.Format("CardCode={0}, Total={1}", txtCardCode.Text, lblDocTotalWithTax.Text))

                            Response.Redirect("Index.aspx")
                        Catch ex As Exception
                            trans.Rollback()
                            Throw ex
                        End Try
                    End Using
                End Using
            Catch ex As Exception
                ShowError("刪除失敗: " & ex.Message)
            End Try
        End If
    End Sub

    Protected Sub btnCancel_Click(sender As Object, e As EventArgs)
        Response.Redirect("ExpenseClaimList.aspx")
    End Sub

    ''' <summary>
    ''' 新增新單據 - 重置表單，開始新的費用申請
    ''' </summary>
    Protected Sub btnNewDocument_Click(sender As Object, e As EventArgs)
        Response.Redirect("ExpenseClaimForm.aspx")
    End Sub

    Protected Sub btnCopyDocument_Click(sender As Object, e As EventArgs)
        If currentJID = 0 Then
            ShowError("無法複製：尚未儲存的單據")
            Return
        End If

        Dim isOwner As Boolean = (txtOwner.Text = currentUserId)
        If Not isApUser AndAlso Not isOwner Then
            ShowError("您沒有權限複製此單據")
            Return
        End If

        Dim copyAttach As String = If(hfCopyAttachment.Value = "1", "1", "0")
        Dim copyMdr As String = If(hfCopyMDR.Value = "1", "1", "0")
        Response.Redirect("ExpenseClaimForm.aspx?CopyFrom=" & currentJID.ToString() & "&CopyAttach=" & copyAttach & "&CopyMDR=" & copyMdr)
    End Sub

    ''' <summary>
    ''' 更新按鈕 - 在檢視模式下更新單據資料（PID、明細、備註等）
    ''' 權限控制：
    ''' - 草稿(P)/待審核(W)/駁回(R)：單據擁有者或審核者可更新完整資料
    ''' - 已核准(A)：僅可更新備註欄
    ''' </summary>
    Protected Sub btnUpdate_Click(sender As Object, e As EventArgs)
        If currentJID = 0 Then
            ShowError("無法更新：尚未儲存的單據")
            Return
        End If

        Dim status As String = txtApprovalStatus.Text
        Dim isOwner As Boolean = (txtOwner.Text = currentUserId)
        ' [暫時修改] 審核者可全權編輯任何狀態的單據，修正完資料後請改回原本邏輯
        ' 原本: Dim canEditFull As Boolean = (status = "P" OrElse status = "W" OrElse status = "R") AndAlso (isOwner OrElse isApUser)
        Dim canEditFull As Boolean = isApUser OrElse ((status = "P" OrElse status = "W" OrElse status = "R") AndAlso isOwner)
        Dim isApproved As Boolean = (status = "A")

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()

                If canEditFull Then
                    ' 草稿/待審核/駁回：更新完整資料（維持原狀態）
                    If SaveDocument(status) Then
                        ShowSuccess("更新成功")
                        LoadDocument(currentJID)
                    End If
                ElseIf isApproved Then
                    ' 已核准：更新備註欄，審核者 (isApUser) 可額外更新 DocEntry
                    Dim sql As String
                    If isApUser Then
                        ' 審核者可更新 DocEntry (AP單號)
                        sql = "UPDATE jOPCH SET Comments=@Comments, DocEntry=@DocEntry, UpdateBy=@User, UpdateDate=GETDATE() WHERE jID=@jID"
                    Else
                        sql = "UPDATE jOPCH SET Comments=@Comments, UpdateBy=@User, UpdateDate=GETDATE() WHERE jID=@jID"
                    End If
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@Comments", txtRemarks.Text)
                        cmd.Parameters.AddWithValue("@User", currentUserId)
                        cmd.Parameters.AddWithValue("@jID", currentJID)
                        If isApUser Then
                            ' 解析 DocEntry，空白時設為 NULL
                            Dim docEntryText As String = txtB1DocEntry.Text.Trim()
                            If String.IsNullOrEmpty(docEntryText) Then
                                cmd.Parameters.AddWithValue("@DocEntry", DBNull.Value)
                            Else
                                Dim docEntryVal As Integer
                                If Integer.TryParse(docEntryText, docEntryVal) Then
                                    cmd.Parameters.AddWithValue("@DocEntry", docEntryVal)
                                Else
                                    cmd.Parameters.AddWithValue("@DocEntry", DBNull.Value)
                                End If
                            End If
                        End If
                        cmd.ExecuteNonQuery()
                    End Using
                    ShowSuccess(If(isApUser, "備註與AP單號更新成功", "備註更新成功"))
                    LoadDocument(currentJID)
                Else
                    ShowError("您沒有權限更新此單據")
                End If
            End Using
        Catch ex As Exception
            ShowError("更新失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 匯出 PDF - 使用 Crystal Report 將費用申請單匯出為 PDF 格式
    ''' </summary>
    Protected Sub btnExportPDF_Click(sender As Object, e As EventArgs)
        Dim jID As String = txtJID.Text.Trim()
        If Not String.IsNullOrEmpty(jID) Then
            ' 使用 Crystal Report Handler 產生 PDF (在新視窗開啟)
            Dim script As String = String.Format("window.open('ExpenseClaimReport.ashx?jID={0}', '_blank');", HttpUtility.UrlEncode(jID))
            ScriptManager.RegisterStartupScript(Me, Me.GetType(), "OpenPDF", script, True)
        Else
            ShowError("無法匯出：尚未儲存的單據或缺少平台單號")
        End If
    End Sub

    ''' <summary>
    ''' 設定為檢視模式 (已存檔的單據)
    ''' 依據狀態決定可編輯的欄位：
    ''' - 草稿(P)/待審核(W)/駁回(R)：可編輯大部分欄位並更新
    ''' - 已核准(A)：僅備註欄可編輯
    ''' </summary>
    Private Sub SetViewMode()
        Dim status As String = txtApprovalStatus.Text
        Dim isOwner As Boolean = (txtOwner.Text = currentUserId)
        ' [暫時修改] 審核者可全權編輯任何狀態的單據，修正完資料後請改回原本邏輯
        ' 原本: Dim canEdit As Boolean = (status = "P" OrElse status = "W" OrElse status = "R") AndAlso (isOwner OrElse isApUser)
        Dim canEdit As Boolean = isApUser OrElse ((status = "P" OrElse status = "W" OrElse status = "R") AndAlso isOwner)
        Dim isApproved As Boolean = (status = "A")

        ' [暫時修改] 審核者可以看到送審按鈕，修正完資料後請改回原本邏輯
        ' 原本: btnSubmit.Visible = False / btnSave.Visible = False
        If canEdit Then
            btnSubmit.Visible = True  ' 送審按鈕
            btnSave.Visible = True    ' 儲存草稿按鈕
        Else
            btnSubmit.Visible = False
            btnSave.Visible = False
        End If
        btnDelete.Visible = False

        ' 顯示檢視模式按鈕
        btnUpdate.Visible = canEdit OrElse isApproved ' 只要能編輯備註就顯示更新按鈕

        ' 審核者使用不同的按鈕樣式和文字
        If isApUser Then
            btnUpdate.Text = "審核者更新"
            btnUpdate.CssClass = "btn btn-warning"  ' 橘色按鈕
        Else
            btnUpdate.Text = "更新 (Update)"
            btnUpdate.CssClass = "btn btn-primary"  ' 藍色按鈕
        End If

        btnExportPDF.Visible = True
        btnNewDocument.Visible = True
        btnCopyDocument.Visible = canEdit OrElse isApUser OrElse isOwner
        btnCancel.Text = "返回列表"

        ' ===== 備註欄：任何狀態都可編輯 =====
        txtRemarks.ReadOnly = False

        If canEdit Then
            ' ===== 草稿/待審核/駁回：允許編輯大部分欄位 =====

            ' 標頭欄位
            txtCardCode.ReadOnly = False
            txtCardName.ReadOnly = False
            txtNumAtCard.ReadOnly = False
            txtDocDate.ReadOnly = False
            txtDocDueDate.ReadOnly = False
            txtTaxDate.ReadOnly = False
            txtDocRate.ReadOnly = False
            txtAddress.ReadOnly = False
            txtPymntGroup.ReadOnly = False
            txtUPID.ReadOnly = False

            ' 下拉選單
            ddlDeliveryAddr.Enabled = True
            ddlDocCurrency.Enabled = True
            ddlGroupNum.Enabled = True
            ddlPurchaser.Enabled = True

            ' 供應商搜尋按鈕
            btnSearchCardCode.Visible = True
            btnSearchCardName.Visible = True
            btnRefreshRate.Visible = True

            ' 明細編輯按鈕
            btnAddLine.Visible = True
            btnDeleteLine.Visible = True
            btnGenerateMDR.Visible = True
            btnAddMDRRow.Visible = True
            btnDeleteMDRRow.Visible = True

            ' 附件上傳
            btnUpload.Visible = True
            fileUpload.Visible = True
        Else
            ' ===== 已核准或無權限：除備註外其他欄位唯讀 =====

            ' 標頭欄位設為唯讀
            txtCardCode.ReadOnly = True
            txtCardName.ReadOnly = True
            txtNumAtCard.ReadOnly = True
            txtDocDate.ReadOnly = True
            txtDocDueDate.ReadOnly = True
            txtTaxDate.ReadOnly = True
            txtDocRate.ReadOnly = True
            txtAddress.ReadOnly = True
            txtPymntGroup.ReadOnly = True
            txtUPID.ReadOnly = True

            ' 停用下拉選單
            ddlDeliveryAddr.Enabled = False
            ddlDocCurrency.Enabled = False
            ddlGroupNum.Enabled = False
            ddlPurchaser.Enabled = False

            ' 停用供應商搜尋按鈕
            btnSearchCardCode.Visible = False
            btnSearchCardName.Visible = False
            btnRefreshRate.Visible = False

            ' 停用明細編輯按鈕
            btnAddLine.Visible = False
            btnDeleteLine.Visible = False
            btnGenerateMDR.Visible = False
            btnAddMDRRow.Visible = False
            btnDeleteMDRRow.Visible = False

            ' 停用附件上傳
            btnUpload.Visible = False
            fileUpload.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' 設定為編輯模式 (新增或可編輯的單據)
    ''' </summary>
    Private Sub SetEditMode()
        ' 顯示新增/編輯模式按鈕
        btnSubmit.Visible = True
        btnSave.Visible = False ' 暫存按鈕預設隱藏
        btnDelete.Visible = (currentJID > 0) ' 只有已存檔的單據才能刪除

        ' 隱藏檢視模式按鈕
        btnExportPDF.Visible = False
        btnNewDocument.Visible = False
        btnCopyDocument.Visible = (currentJID > 0 AndAlso (isApUser OrElse txtOwner.Text = currentUserId))
        btnCancel.Text = "取消 (Cancel)"

        ' 啟用所有輸入欄位
        txtCardCode.ReadOnly = False
        txtCardName.ReadOnly = False
        txtNumAtCard.ReadOnly = False
        txtDocDate.ReadOnly = False
        txtDocDueDate.ReadOnly = False
        txtTaxDate.ReadOnly = False
        txtDocRate.ReadOnly = False
        txtRemarks.ReadOnly = False
        txtAddress.ReadOnly = False
        txtPymntGroup.ReadOnly = False

        ' 啟用下拉選單
        ddlDeliveryAddr.Enabled = True
        ddlGroupNum.Enabled = True
        ddlPurchaser.Enabled = True

        ' 啟用供應商搜尋按鈕
        btnSearchCardCode.Visible = True
        btnSearchCardName.Visible = True
        btnRefreshRate.Visible = True

        ' 啟用明細編輯按鈕
        btnAddLine.Visible = True
        btnDeleteLine.Visible = True
        btnGenerateMDR.Visible = True
        btnAddMDRRow.Visible = True
        btnDeleteMDRRow.Visible = True

        ' 啟用附件上傳
        btnUpload.Visible = True
        fileUpload.Visible = True
    End Sub

    Private Sub SetHeaderParameters(cmd As SqlCommand, status As String)
        cmd.Parameters.AddWithValue("@CardCode", txtCardCode.Text)
        cmd.Parameters.AddWithValue("@CardName", txtCardName.Text)
        cmd.Parameters.AddWithValue("@NumAtCard", txtNumAtCard.Text)
        cmd.Parameters.AddWithValue("@InvNum", "") ' 初次儲存時為空，審核放行後由 SAP 回填
        cmd.Parameters.AddWithValue("@DeliveryAddrID", ddlDeliveryAddr.SelectedValue)
        cmd.Parameters.AddWithValue("@AddressName", ddlDeliveryAddr.SelectedItem.Text)
        cmd.Parameters.AddWithValue("@Address", txtAddress.Text)
        ' [H] 使用安全日期解析，確保格式一致性
        Dim docDate As DateTime? = ParseDateSafe(txtDocDate.Text, DateTime.Today)
        Dim docDueDate As DateTime? = ParseDateSafe(txtDocDueDate.Text, DateTime.Today.AddDays(30))
        Dim taxDate As DateTime? = ParseDateSafe(txtTaxDate.Text)
        cmd.Parameters.AddWithValue("@DocDate", If(docDate.HasValue, docDate.Value, DBNull.Value))
        cmd.Parameters.AddWithValue("@DocDueDate", If(docDueDate.HasValue, docDueDate.Value, DBNull.Value))
        cmd.Parameters.AddWithValue("@TaxDate", If(taxDate.HasValue, taxDate.Value, DBNull.Value))
        cmd.Parameters.AddWithValue("@DocCurrency", ddlDocCurrency.SelectedValue)
        cmd.Parameters.AddWithValue("@DocRate", txtDocRate.Text)
        cmd.Parameters.AddWithValue("@DocTotal", Decimal.Parse(lblDocTotalWithTax.Text.Replace(",", "")))
        cmd.Parameters.AddWithValue("@VatSum", Decimal.Parse(lblVatSum.Text.Replace(",", "")))
        cmd.Parameters.AddWithValue("@GroupNum", ddlGroupNum.SelectedValue)
        cmd.Parameters.AddWithValue("@PymntGroup", If(String.IsNullOrEmpty(txtPymntGroup.Text), DBNull.Value, txtPymntGroup.Text))
        cmd.Parameters.AddWithValue("@Comments", txtRemarks.Text)
        'cmd.Parameters.AddWithValue("@Attachment", If(String.IsNullOrEmpty(lblAttachment.Text), DBNull.Value, lblAttachment.Text))
        cmd.Parameters.AddWithValue("@Status", status)
        cmd.Parameters.AddWithValue("@User", currentUserId)
        cmd.Parameters.AddWithValue("@UPID", If(String.IsNullOrEmpty(txtUPID.Text), DBNull.Value, txtUPID.Text))
        cmd.Parameters.AddWithValue("@SlpCode", ddlPurchaser.SelectedValue)
    End Sub
#End Region

#Region "讀取文件"
    Private Sub LoadDocument(id As Integer)
        Using conn As New SqlConnection(connStr)
            conn.Open()
            ' Load Header (使用 jID 作為主鍵)
            Dim sql As String = "SELECT * FROM jOPCH WHERE jID=@jID"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@jID", id)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        lblDocNum.Text = dr("jID").ToString()  ' 顯示 jID 作為單號
                        txtJID.Text = dr("jID").ToString()

                        txtCardCode.Text = dr("CardCode").ToString()
                        txtCardName.Text = dr("CardName").ToString()
                        txtNumAtCard.Text = dr("NumAtCard").ToString()
                        'txtInvNum.Text = dr("InvNum").ToString()

                        If Not IsDBNull(dr("DeliveryAddrID")) Then ddlDeliveryAddr.SelectedValue = dr("DeliveryAddrID").ToString()
                        txtAddress.Text = dr("Address").ToString()

                        txtDocDate.Text = Convert.ToDateTime(dr("DocDate")).ToString("yyyy-MM-dd")
                        If Not IsDBNull(dr("DocDueDate")) Then txtDocDueDate.Text = Convert.ToDateTime(dr("DocDueDate")).ToString("yyyy-MM-dd")
                        If Not IsDBNull(dr("TaxDate")) Then txtTaxDate.Text = Convert.ToDateTime(dr("TaxDate")).ToString("yyyy-MM-dd")

                        ddlDocCurrency.SelectedValue = dr("DocCurrency").ToString()
                        txtDocRate.Text = dr("DocRate").ToString()

                        If Not IsDBNull(dr("GroupNum")) Then ddlGroupNum.SelectedValue = dr("GroupNum").ToString()
                        If Not IsDBNull(dr("PymntGroup")) Then txtPymntGroup.Text = dr("PymntGroup").ToString()
                        txtRemarks.Text = dr("Comments").ToString()

                        'If Not IsDBNull(dr("Attachment")) Then
                        '    lblAttachment.Text = dr("Attachment").ToString()
                        'End If

                        Dim status As String = dr("ApprovalStatus").ToString()
                        lblDocStatus.Text = GetStatusText(status)
                        lblDocStatus.CssClass = "badge status-" & status
                        txtStatusDisplay.Text = GetStatusText(status)
                        txtApprovalStatus.Text = status

                        txtApprovedBy.Text = dr("ApprovedBy").ToString()
                        txtOwner.Text = dr("CreateBy").ToString()
                        
                        Try
                            ' 直接嘗試讀取，若欄位不存在會被 Catch 捕獲 (SqlDataReader 不支援 Table.Columns)
                            If Not IsDBNull(dr("SlpCode")) Then
                                Dim slp As String = dr("SlpCode").ToString()
                                If ddlPurchaser.Items.FindByValue(slp) IsNot Nothing Then
                                    ddlPurchaser.SelectedValue = slp
                                End If
                            End If
                        Catch
                        End Try

                        If Not IsDBNull(dr("U_PID")) Then txtUPID.Text = dr("U_PID").ToString()

                        ' [SAP Integration] 顯示 DocEntry (AP單號)
                        Try
                            If HasColumn(dr, "DocEntry") AndAlso Not IsDBNull(dr("DocEntry")) Then
                                txtB1DocEntry.Text = dr("DocEntry").ToString()
                            Else
                                txtB1DocEntry.Text = ""
                            End If
                            ' AP單號欄位：審核者 (isApUser) 可編輯
                            txtB1DocEntry.ReadOnly = Not isApUser
                            txtB1DocEntry.CssClass = If(isApUser, "", "readonly-field")
                        Catch
                            ' 忽略 DocEntry 讀取錯誤
                        End Try

                        ' [SAP Integration] 顯示 SAP 過帳狀態
                        Try
                            Dim b1Status As String = ""
                            If HasColumn(dr, "B1PostStatus") AndAlso Not IsDBNull(dr("B1PostStatus")) Then
                                b1Status = dr("B1PostStatus").ToString()
                                If b1Status = "Y" Then
                                    lblDocStatus.Text &= " (SAP: 已過帳)"
                                    lblDocStatus.CssClass &= " sap-success" ' 需確保 CSS 支援，或只改文字
                                ElseIf b1Status = "E" Then
                                    Dim errMsg As String = ""
                                    If HasColumn(dr, "B1ErrMsg") AndAlso Not IsDBNull(dr("B1ErrMsg")) Then
                                        errMsg = dr("B1ErrMsg").ToString()
                                    End If
                                    lblDocStatus.Text &= " (SAP: 失敗)"
                                    lblDocStatus.ToolTip = "SAP 錯誤: " & errMsg
                                    ShowError("SAP 過帳失敗: " & errMsg) ' 提示使用者
                                End If
                            End If
                        Catch ex As Exception
                            ' 忽略 SAP 狀態讀取錯誤
                        End Try

                        ' 顯示審核區塊邏輯:
                        ' 1. 這個區塊要 jtdb 的 User Table 裡的 AP_App 欄位為 1 才可以編輯
                        ' 2. 一般 User 不能編輯也不能按放行/發送意見/退回
                        ' 3. 審核區塊保持顯示（讓一般使用者可看到退回意見），但按鈕完全隱藏

                        pnlApproval.Visible = True ' 只要是既有單據，預設顯示，內容依權限控制

                        txtApprovalComments.ReadOnly = Not isApUser

                        ' 審核按鈕：對一般使用者完全隱藏，只有 AP_App 權限者才能看到
                        btnApprove.Visible = isApUser
                        btnApprove.Enabled = isApUser

                        btnUpdateComment.Visible = isApUser
                        btnUpdateComment.Enabled = isApUser

                        btnReject.Visible = isApUser
                        btnReject.Enabled = isApUser

                        ' 按鈕狀態 (編輯模式)
                        btnSave.Text = "更新 (Update)"
                        btnDelete.Visible = True
                    End If
                End Using
            End Using

            ' Load Attachments (jAttach) - [I] 只讀取未刪除的附件
            Dim attachList As New List(Of AttachmentItem)
            sql = "SELECT * FROM jAttach WHERE jID=@ID AND (IsDeleted=0 OR IsDeleted IS NULL) ORDER BY UploadDate, UploadTime"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ID", id)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        attachList.Add(New AttachmentItem With {
                            .ID = Convert.ToInt32(dr("ID")),
                            .FileName = dr("FilePath").ToString(), ' 這裡暫存 FileName (存檔名)
                            .FilePath = dr("FilePath").ToString(),
                            .UploadDate = Convert.ToDateTime(dr("UploadDate")),
                            .UploadTime = dr("UploadTime").ToString(),
                            .Uploader = dr("Uploader").ToString()
                        })
                    End While
                End Using
            End Using
            CurrentAttachments = attachList
            BindAttachmentGrid()

            ' Load Expense Lines (使用 jID 作為關聯鍵)
            Dim lines As New List(Of ExpenseLine)
            sql = "SELECT * FROM jPCH1 WHERE jID=@jID ORDER BY LineNum"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@jID", id)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        ' 轉換 SAP 稅碼回顯示用代碼: J1→1, J0→2, JX→3
                        Dim vatGroupRaw As String = dr("VatGroup").ToString()
                        Dim vatGroupDisplay As String = vatGroupRaw
                        Select Case vatGroupRaw
                            Case "J1" : vatGroupDisplay = "1"
                            Case "J0" : vatGroupDisplay = "2"
                            Case "JX" : vatGroupDisplay = "3"
                        End Select

                        lines.Add(New ExpenseLine With {
                            .LineNum = Convert.ToInt32(dr("LineNum")),
                            .CategoryCode = dr("ItemCode").ToString(),
                            .Description = dr("Dscription").ToString(),
                            .AcctCode = dr("AcctCode").ToString(),
                            .LineTotal = Convert.ToDecimal(dr("LineTotal")),
                            .VatGroup = vatGroupDisplay,
                            .VatRate = Convert.ToDecimal(dr("VatPrcnt")),
                            .VatSum = Convert.ToDecimal(dr("LineVat")),
                            .PriceAfterVat = Convert.ToDecimal(dr("GTotal")),
                            .CostingCode = dr("CostingCode").ToString(),
                            .CostingCode2 = dr("CostingCode2").ToString()
                        })
                    End While
                End Using
            End Using
            CurrentLines = lines
            BindGrid()

            ' Load MDR Lines (使用 jID 作為關聯鍵)
            Dim mdrLines As New List(Of MDRLine)
            sql = "SELECT * FROM jMGUIAPDetail WHERE jID=@jID ORDER BY LineNum"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@jID", id)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        Dim mdr As New MDRLine With {
                            .LineNum = Convert.ToInt32(dr("LineNum")),
                            .U_LIFNR = dr("U_LIFNR").ToString(),
                            .U_STCEG = dr("U_STCEG").ToString(),
                            .U_XBLNR = dr("U_XBLNR").ToString(),
                            .U_ZFORM_CODE = dr("U_ZFORM_CODE").ToString(),
                            .U_HWBAS = Convert.ToDecimal(dr("U_HWBAS")),
                            .U_HWSTE = Convert.ToDecimal(dr("U_HWSTE")),
                            .U_TAX_TYPE = dr("U_TAX_TYPE").ToString()
                        }
                        If Not IsDBNull(dr("U_BLDAT")) Then mdr.U_BLDAT = Convert.ToDateTime(dr("U_BLDAT"))
                        If Not IsDBNull(dr("U_VATDATE")) Then mdr.U_VATDATE = Convert.ToDateTime(dr("U_VATDATE"))
                        mdrLines.Add(mdr)
                    End While
                End Using
            End Using
            CurrentMDRLines = mdrLines
            BindMDRGrid()

        End Using

        ' 載入完成後設為檢視模式
        SetViewMode()
    End Sub
#End Region

#Region "複製單據"
    Private Sub ApplyCopyOverrides(sourceJID As Integer, copyAttach As Boolean, copyMdr As Boolean)
        Try
            CopyFromJID = sourceJID
            CopyAttachments = copyAttach
            CopyMDR = copyMdr

            currentJID = 0
            lblDocNum.Text = "[新單據]"
            txtJID.Text = ""
            txtB1DocEntry.Text = ""
            txtB1DocEntry.ReadOnly = Not isApUser
            txtB1DocEntry.CssClass = If(isApUser, "", "readonly-field")
            txtUPID.Text = ""
            txtApprovalComments.Text = ""
            txtApprovedBy.Text = ""
            txtApprovalStatus.Text = ""
            txtStatusDisplay.Text = "新增中"
            lblDocStatus.Text = "新增中"
            lblDocStatus.CssClass = "badge status-W"
            txtOwner.Text = currentUserId

            Dim today As String = DateTime.Now.ToString("yyyy-MM-dd")
            txtDocDate.Text = today
            txtTaxDate.Text = today
            CalculateDueDate()

            CurrentAttachments = New List(Of AttachmentItem)()
            BindAttachmentGrid()

            If Not copyMdr Then
                CurrentMDRLines = New List(Of MDRLine)()
                BindMDRGrid()
            End If

            btnApprove.Visible = True
            btnApprove.Enabled = False
            btnUpdateComment.Visible = True
            btnUpdateComment.Enabled = False
            btnReject.Visible = True
            btnReject.Enabled = False
            txtApprovalComments.ReadOnly = True

            SetEditMode()
            If copyAttach AndAlso copyMdr Then
                ShowInfo("已根據原文件資訊複製新單據，附件與憑證明細將於儲存時複製，新增前請檢查更新憑證資訊。")
            ElseIf copyAttach Then
                ShowInfo("已根據原文件資訊複製新單據，附件將於儲存時複製，新增前請檢查更新憑證資訊。")
            ElseIf copyMdr Then
                ShowInfo("已根據原文件資訊複製新單據，憑證明細將於儲存時複製，新增前請檢查更新憑證資訊。")
            Else
                ShowInfo("已根據原文件資訊複製新單據，新增前請檢查更新憑證資訊。")
            End If
        Catch ex As Exception
            ' 複製失敗時靜默
        End Try
    End Sub
#End Region

#Region "附件處理"
    Private Const FORM_NAME As String = "ExpenseClaimForm"

    ''' <summary>
    ''' 取得附件儲存資料夾路徑
    ''' 結構: AttachFile/User/{UserID}/{FormName}/{jID}/
    ''' </summary>
    Private Function GetAttachmentFolder(jID As Integer) As String
        Dim relativePath As String = String.Format("~/AttachFile/User/{0}/{1}/{2}/", currentUserId, FORM_NAME, jID)
        Return Server.MapPath(relativePath)
    End Function

    ''' <summary>
    ''' 取得附件相對路徑 (用於資料庫儲存)
    ''' </summary>
    Private Function GetAttachmentRelativePath(jID As Integer, fileName As String) As String
        Return String.Format("AttachFile/User/{0}/{1}/{2}/{3}", currentUserId, FORM_NAME, jID, fileName)
    End Function

    Private Function GetAttachmentAbsolutePath(relativePath As String) As String
        If String.IsNullOrEmpty(relativePath) Then Return ""
        Return Server.MapPath("~/" & relativePath.TrimStart("/"c))
    End Function

    Private Sub CopyAttachmentsFromSource(sourceJID As Integer, newJID As Integer, conn As SqlConnection, trans As SqlTransaction)
        If sourceJID <= 0 OrElse newJID <= 0 Then Return

        Try
            Dim attachList As New List(Of AttachmentItem)()
            Dim sql As String = "SELECT FilePath, FileName FROM jAttach WHERE jID=@ID AND (IsDeleted=0 OR IsDeleted IS NULL)"
            Using cmd As New SqlCommand(sql, conn, trans)
                cmd.Parameters.AddWithValue("@ID", sourceJID)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        attachList.Add(New AttachmentItem With {
                            .FilePath = dr("FilePath").ToString(),
                            .FileName = dr("FileName").ToString()
                        })
                    End While
                End Using
            End Using

            If attachList.Count = 0 Then Return

            Dim targetFolder As String = GetAttachmentFolder(newJID)
            If Not Directory.Exists(targetFolder) Then
                Directory.CreateDirectory(targetFolder)
            End If

            For Each att In attachList
                Try
                    Dim originName As String = If(String.IsNullOrEmpty(att.FileName), Path.GetFileName(att.FilePath), att.FileName)
                    If String.IsNullOrEmpty(originName) Then Continue For

                    Dim sourcePath As String = GetAttachmentAbsolutePath(att.FilePath)
                    If String.IsNullOrEmpty(sourcePath) OrElse Not File.Exists(sourcePath) Then Continue For

                    Dim savedName As String = DateTime.Now.ToString("yyyyMMddHHmmss") & "_" & originName
                    Dim targetPath As String = Path.Combine(targetFolder, savedName)
                    File.Copy(sourcePath, targetPath, True)

                    Dim relativePath As String = GetAttachmentRelativePath(newJID, savedName)
                    Dim sqlIns As String = "INSERT INTO jAttach (jID, DocEntry, LineNum, FilePath, FileName, Uploader, UploadTime) " &
                                           "VALUES (@jID, @DocEntry, -1, @FilePath, @FileName, @Uploader, @UploadTime)"
                    Using cmdIns As New SqlCommand(sqlIns, conn, trans)
                        cmdIns.Parameters.AddWithValue("@jID", newJID)
                        cmdIns.Parameters.AddWithValue("@DocEntry", newJID)
                        cmdIns.Parameters.AddWithValue("@FilePath", relativePath)
                        cmdIns.Parameters.AddWithValue("@FileName", originName)
                        cmdIns.Parameters.AddWithValue("@Uploader", currentUserId)
                        cmdIns.Parameters.AddWithValue("@UploadTime", DateTime.Now.ToString("HH:mm:ss"))
                        cmdIns.ExecuteNonQuery()
                    End Using
                Catch ex As Exception
                    ' 複製失敗時靜默
                End Try
            Next
        Catch ex As Exception
            ' 複製失敗時靜默
        End Try
    End Sub

    Protected Sub btnUpload_Click(sender As Object, e As EventArgs)
        ' .NET 4.0 相容性修改: 檢查 Request.Files
        If fileUpload.HasFile OrElse Request.Files.Count > 0 Then
            Try
                ' 若尚未存檔 (currentJID=0)，先自動儲存為草稿以取得 jID
                If currentJID = 0 Then
                    If Not SaveDocument("P", True) Then
                        ShowError("上傳附件前自動儲存草稿失敗，請檢查必填欄位。")
                        Return
                    End If
                End If

                ' 建立附件資料夾: AttachFile/User/{UserID}/ExpenseClaimForm/{jID}/
                Dim folder As String = GetAttachmentFolder(currentJID)
                If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)

                Dim list = CurrentAttachments
                Dim successCount As Integer = 0

                ' 遍歷 Request.Files 以支援多檔上傳 (.NET 4.0 Workaround)
                For i As Integer = 0 To Request.Files.Count - 1
                    Dim uploadedFile As HttpPostedFile = Request.Files(i)

                    ' 確保檔案有效且有名稱 (過濾掉空欄位)
                    If uploadedFile.ContentLength > 0 AndAlso Not String.IsNullOrEmpty(uploadedFile.FileName) Then
                        Dim originName As String = Path.GetFileName(uploadedFile.FileName)
                        Dim savedName As String = DateTime.Now.ToString("yyyyMMddHHmmss") & "_" & originName
                        Dim fullPath As String = Path.Combine(folder, savedName)
                        uploadedFile.SaveAs(fullPath)

                        ' 儲存相對路徑到資料庫
                        Dim relativePath As String = GetAttachmentRelativePath(currentJID, savedName)

                        ' 寫入資料庫 jAttach
                        Using conn As New SqlConnection(connStr)
                            conn.Open()
                            Dim sql As String = "INSERT INTO jAttach (jID, DocEntry, LineNum, FilePath, FileName, Uploader, UploadTime) " &
                                              "VALUES (@jID, @DocEntry, -1, @FilePath, @FileName, @Uploader, @UploadTime); SELECT SCOPE_IDENTITY();"
                            Using cmd As New SqlCommand(sql, conn)
                                cmd.Parameters.AddWithValue("@jID", currentJID) ' jOPCH.jID (DocEntry)
                                cmd.Parameters.AddWithValue("@DocEntry", currentJID)
                                cmd.Parameters.AddWithValue("@FilePath", relativePath) ' 儲存相對路徑
                                cmd.Parameters.AddWithValue("@FileName", originName)
                                cmd.Parameters.AddWithValue("@Uploader", currentUserId)
                                cmd.Parameters.AddWithValue("@UploadTime", DateTime.Now.ToString("HH:mm:ss"))

                                Dim newId As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                                list.Add(New AttachmentItem With {
                                    .ID = newId,
                                    .FileName = originName, ' 顯示原檔名
                                    .FilePath = relativePath, ' 儲存相對路徑
                                    .UploadDate = DateTime.Today,
                                    .UploadTime = DateTime.Now.ToString("HH:mm:ss"),
                                    .Uploader = currentUserId
                                })
                                successCount += 1
                            End Using
                        End Using
                    End If
                Next
                
                CurrentAttachments = list
                BindAttachmentGrid()
                ShowSuccess($"成功上傳 {successCount} 個檔案")

            Catch ex As Exception
                ShowError("上傳失敗: " & ex.Message)
            End Try
        Else
             ShowError("請選擇檔案")
        End If
    End Sub

    ''' <summary>
    ''' [I] 附件刪除改用 Soft Delete 機制
    ''' </summary>
    Protected Sub gvAttachments_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "DeleteFile" Then
            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            Dim list = CurrentAttachments

            If index < list.Count Then
                Dim item = list(index)

                Try
                    ' [I] 改用 Soft Delete，不實際刪除資料
                    Using conn As New SqlConnection(connStr)
                        conn.Open()
                        Dim sql As String = "UPDATE jAttach SET IsDeleted=1, DeletedDate=GETDATE(), DeletedBy=@UserId WHERE ID=@ID"
                        Using cmd As New SqlCommand(sql, conn)
                            cmd.Parameters.AddWithValue("@ID", item.ID)
                            cmd.Parameters.AddWithValue("@UserId", currentUserId)
                            cmd.ExecuteNonQuery()
                        End Using
                    End Using

                    ' 不刪除實體檔案，保留備份
                    ' 可設定定期清理 Job 來處理 IsDeleted=1 且超過保留期限的檔案

                    ' [E] 記錄刪除附件稽核日誌
                    AuditLogger.Log("jAttach", currentJID, AuditLogger.Actions.Delete, currentUserId,
                                    changes:=String.Format("FileName={0}, FilePath={1}", item.FileName, item.FilePath))

                    list.RemoveAt(index)
                    CurrentAttachments = list
                    BindAttachmentGrid()
                    ShowSuccess("附件已刪除")
                Catch ex As Exception
                    ShowError("刪除附件失敗: " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub BindAttachmentGrid()
        gvAttachments.DataSource = CurrentAttachments
        gvAttachments.DataBind()
    End Sub
#End Region

#Region "審核"
    Protected Sub btnApprove_Click(sender As Object, e As EventArgs)
        If Not isApUser Then
            ShowError("無權限執行此操作")
            Return
        End If

        ' [A] 放行時檢查金額一致性，使用統一彈窗顯示警告
        Dim warnings = CheckAmountConsistency()

        ' 加入格式28提醒 (E030A 海關代徵營業稅)
        Dim hasE030AExpense As Boolean = CurrentLines.Any(Function(x) x.CategoryCode = "E030A")
        Dim hasFormat28MDR As Boolean = CurrentMDRLines.Any(Function(x) x.U_ZFORM_CODE = "28")
        If hasE030AExpense OrElse hasFormat28MDR Then
            warnings.Add("【格式28 提醒】此單據含海關代徵營業稅，匯入SAP時將只過帳「稅額」金額")
        End If

        If warnings.Count > 0 AndAlso Not ApprovalWarningConfirmed Then
            ' 使用統一彈窗
            ViewState("PendingAction") = "approve"
            ShowValidationPopup(New List(Of String)(), warnings)
            Return
        End If

        ApprovalWarningConfirmed = False ' 重置
        UpdateStatus("A")
    End Sub

    Protected Sub btnReject_Click(sender As Object, e As EventArgs)
        If Not isApUser Then
            ShowError("無權限執行此操作")
            Return
        End If
        UpdateStatus("R")
    End Sub

    Protected Sub btnUpdateComment_Click(sender As Object, e As EventArgs)
        If Not isApUser Then
            ShowError("無權限執行此操作")
            Return
        End If
        ' 僅更新意見，不改狀態
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE jOPCH SET ApprovalComments=@Comm, UpdateBy=@User, UpdateDate=GETDATE() WHERE jID=@jID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@User", currentUserId)
                    cmd.Parameters.AddWithValue("@Comm", txtApprovalComments.Text)
                    cmd.Parameters.AddWithValue("@jID", currentJID)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
            ShowSuccess("意見已更新")
        Catch ex As Exception
            ShowError("更新失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' [B] 狀態轉換驗證與更新
    ''' 狀態轉換規則：
    ''' - W (待審) → A (核准) 或 R (駁回)
    ''' - R (駁回) → W (待審)
    ''' - A (核准) + B1PostStatus='Y' → 終態，不可變更
    ''' - A (核准) + B1PostStatus='E' → 允許重試寫入 SAP
    ''' </summary>
    Private Sub UpdateStatus(newStatus As String)
        Try
            ' 先取得當前狀態和 B1PostStatus (使用 jID 作為主鍵)
            Dim currentStatus As String = ""
            Dim b1PostStatus As String = ""
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("SELECT ApprovalStatus, ISNULL(B1PostStatus,'N') AS B1PostStatus FROM jOPCH WHERE jID=@jID", conn)
                    cmd.Parameters.AddWithValue("@jID", currentJID)
                    Using dr = cmd.ExecuteReader()
                        If dr.Read() Then
                            currentStatus = dr("ApprovalStatus").ToString()
                            b1PostStatus = dr("B1PostStatus").ToString()
                        End If
                    End Using
                End Using
            End Using

            ' 檢查是否已成功寫入 SAP（防止重複寫入）
            If b1PostStatus = "Y" Then
                ShowError("此單據已成功寫入 SAP，無法再次操作")
                Return
            End If

            ' 特殊情況：已核准但寫入失敗，允許重試
            Dim isRetry As Boolean = (currentStatus = "A" AndAlso b1PostStatus = "E" AndAlso newStatus = "A")

            ' [B] 狀態轉換驗證（重試時跳過）
            If Not isRetry AndAlso Not IsValidStatusTransition(currentStatus, newStatus) Then
                ShowError(String.Format("無效的狀態轉換：{0} → {1}", GetStatusText(currentStatus), GetStatusText(newStatus)))
                Return
            End If

            ' 使用樂觀鎖定更新狀態，防止競爭條件
            Dim rowsAffected As Integer = 0
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String
                If isRetry Then
                    ' 重試：只更新審核相關欄位，確保 B1PostStatus 仍為 'E'
                    sql = "UPDATE jOPCH SET ApprovedBy=@User, ApprovalDate=GETDATE(), ApprovalComments=@Comm, B1PostStatus='N', B1ErrMsg=NULL " &
                          "WHERE jID=@jID AND ApprovalStatus='A' AND B1PostStatus='E'"
                Else
                    ' 正常流程：使用樂觀鎖定確保狀態沒有被其他人改過
                    sql = "UPDATE jOPCH SET ApprovalStatus=@Status, ApprovedBy=@User, ApprovalDate=GETDATE(), ApprovalComments=@Comm " &
                          "WHERE jID=@jID AND ApprovalStatus=@OldStatus AND ISNULL(B1PostStatus,'N') <> 'Y'"
                End If

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@jID", currentJID)
                    cmd.Parameters.AddWithValue("@Status", newStatus)
                    cmd.Parameters.AddWithValue("@OldStatus", currentStatus)
                    cmd.Parameters.AddWithValue("@User", currentUserId)
                    cmd.Parameters.AddWithValue("@Comm", txtApprovalComments.Text)
                    rowsAffected = cmd.ExecuteNonQuery()
                End Using
            End Using

            ' 檢查是否成功更新（樂觀鎖定失敗時 rowsAffected = 0）
            If rowsAffected = 0 Then
                ShowError("更新失敗：單據狀態已被其他使用者變更，請重新整理頁面")
                Return
            End If

            ' [E] 記錄狀態變更稽核日誌 (非阻塞)
            Dim auditAction As String = AuditLogger.Actions.StatusChange
            If newStatus = "A" Then auditAction = If(isRetry, "RetryApprove", AuditLogger.Actions.Approve)
            If newStatus = "R" Then auditAction = AuditLogger.Actions.Reject
            AuditLogger.Log("jOPCH", currentJID, auditAction, currentUserId,
                            oldValue:=currentStatus, newValue:=newStatus,
                            changes:=String.Format("{0} → {1}, Comment={2}", GetStatusText(currentStatus), GetStatusText(newStatus), txtApprovalComments.Text))

            ' 如果是放行 (A) 或重試，呼叫 SAP B1 API 建立 AP Invoice
            If newStatus = "A" Then
                ' Call SAP B1 API & MDR Integration
                CreateAPInvoiceInSAP(currentJID)
            End If

            Response.Redirect("ExpenseClaimForm.aspx?jID=" & currentJID)
        Catch ex As Exception
            ShowError("更新狀態失敗: " & ex.Message)
        End Try
    End Sub

#Region "SAP Integration"
    ''' <summary>
    ''' 初始化 SAP 連線 (參考廠長的做法)
    ''' </summary>
    Private Function InitSAPConnection() As Integer
        Dim destIP As String = ConfigurationManager.AppSettings("SapServer")
        Dim dbName As String = ConfigurationManager.AppSettings("SapCompanyDB")
        Dim sapUser As String = ConfigurationManager.AppSettings("SapUserName")
        Dim sapPwd As String = ConfigurationManager.AppSettings("SapPassword")
        Dim dbUser As String = ConfigurationManager.AppSettings("SapDbUserName")
        Dim dbPwd As String = ConfigurationManager.AppSettings("SapDbPassword")

        oCompany.Server = destIP
        oCompany.CompanyDB = dbName
        oCompany.UserName = sapUser
        oCompany.Password = sapPwd
        oCompany.UseTrusted = False
        oCompany.DbUserName = If(String.IsNullOrEmpty(dbUser), "sa", dbUser)
        oCompany.DbPassword = If(String.IsNullOrEmpty(dbPwd), "", dbPwd)
        oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019

        Return oCompany.Connect()
    End Function

    ''' <summary>
    ''' 關閉 SAP 連線 (參考廠長的做法)
    ''' </summary>
    Private Sub CloseSAPConnection()
        If oCompany IsNot Nothing AndAlso oCompany.Connected Then
            oCompany.Disconnect()
        End If
    End Sub

    ''' <summary>
    ''' 建立 SAP AP 發票
    ''' </summary>
    ''' <param name="jID">平台單號 (jOPCH.jID)</param>
    Private Sub CreateAPInvoiceInSAP(jID As Integer)
        ' 使用 Page 層級的 oCompany 物件 (參考廠長的做法)
        Dim oInvoice As SAPbobsCOM.Documents = Nothing
        Dim sapDocEntry As Integer = 0
        Dim errMsg As String = ""

        Try
            ' 0. 雙重檢查：確保尚未成功寫入 SAP（防止競爭條件）
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("SELECT B1PostStatus FROM jOPCH WHERE jID=@jID", conn)
                    cmd.Parameters.AddWithValue("@jID", jID)
                    Dim status = cmd.ExecuteScalar()
                    If status IsNot Nothing AndAlso status.ToString() = "Y" Then
                        ShowWarning("此單據已成功寫入 SAP，跳過重複寫入")
                        Return
                    End If
                End Using
            End Using

            ' 1. 初始化 SAP 連線 (從 Web.config 讀取)
            If String.IsNullOrEmpty(ConfigurationManager.AppSettings("SapServer")) OrElse
               String.IsNullOrEmpty(ConfigurationManager.AppSettings("SapCompanyDB")) Then
                Throw New Exception("Web.config 中缺少 SAP 連線設定 (SapServer/SapCompanyDB)")
            End If

            Dim connResult As Integer = InitSAPConnection()
            If connResult <> 0 Then
                Dim connErrCode As Integer
                Dim connErrMsg As String = ""
                oCompany.GetLastError(connErrCode, connErrMsg)
                Throw New Exception($"SAP 連線失敗 [{connErrCode}]: {connErrMsg}")
            End If

            ' 2. 準備與主要資料
            oInvoice = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)

            ' 設定文件類型為服務型發票 (Service Type)
            oInvoice.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            ' 記錄文件幣別和匯率 (供明細使用)
            Dim docCurrency As String = "NTD"
            Dim docRate As Double = 1.0

            Using conn As New SqlConnection(connStr)
                conn.Open()
                ' Header - jOPCH 主鍵是 jID
                Dim sqlH As String = "SELECT * FROM jOPCH WHERE jID=@jID"
                Using cmdH As New SqlCommand(sqlH, conn)
                    cmdH.Parameters.AddWithValue("@jID", jID)
                    Using drH As SqlDataReader = cmdH.ExecuteReader()
                        If Not drH.Read() Then
                            Throw New Exception("找不到費用申請單: jID=" & jID)
                        End If

                        ' 供應商代碼 (必填)
                        oInvoice.CardCode = drH("CardCode").ToString()

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
                        End If

                        ' 匯率 (若非本幣)
                        If Not IsDBNull(drH("DocRate")) Then
                            docRate = Convert.ToDouble(drH("DocRate"))
                            If docCurrency <> "TWD" AndAlso docCurrency <> "NTD" Then
                                oInvoice.DocRate = docRate
                            End If
                        End If

                        ' 注意: DocTotal 不需手動設定，SAP 會根據明細行自動計算

                        ' 收貨地址
                        If Not IsDBNull(drH("Address")) Then
                            oInvoice.Address = drH("Address").ToString()
                        End If

                        ' 付款條件
                        If Not IsDBNull(drH("GroupNum")) Then
                            oInvoice.GroupNumber = Convert.ToInt32(drH("GroupNum"))
                        End If

                        ' 備註
                        If Not IsDBNull(drH("Comments")) Then
                            oInvoice.Comments = drH("Comments").ToString()
                        End If

                        ' 採購人員 (SlpCode)
                        If Not IsDBNull(drH("SlpCode")) Then
                            oInvoice.SalesPersonCode = Convert.ToInt32(drH("SlpCode"))
                        End If
                    End Using
                End Using

                ' Lines - jPCH1 主鍵是 jID + LineNum
                Dim sqlL As String = "SELECT * FROM jPCH1 WHERE jID=@jID ORDER BY LineNum"
                Using cmdL As New SqlCommand(sqlL, conn)
                    cmdL.Parameters.AddWithValue("@jID", jID)
                    Using drL As SqlDataReader = cmdL.ExecuteReader()
                        Dim lineIndex As Integer = 0
                        While drL.Read()
                            ' 第一行不需要 Add，後續行要 Add
                            If lineIndex > 0 Then oInvoice.Lines.Add()

                            ' 會計科目 (費用類) - 必填欄位
                            Dim acctCode As String = ""
                            If Not IsDBNull(drL("AcctCode")) Then
                                acctCode = drL("AcctCode").ToString().Trim()
                            End If
                            If String.IsNullOrEmpty(acctCode) Then
                                Throw New Exception($"明細行 {lineIndex + 1} 缺少會計科目 (AcctCode)")
                            End If
                            oInvoice.Lines.AccountCode = acctCode

                            ' 說明 (ItemDescription)
                            If Not IsDBNull(drL("Dscription")) AndAlso drL("Dscription").ToString().Trim() <> "" Then
                                oInvoice.Lines.ItemDescription = drL("Dscription").ToString()
                            End If

                            ' 明細幣別 (Currency) - 依據 VBA 規格設定
                            Dim lineCurrency As String = docCurrency
                            If Not IsDBNull(drL("Currency")) AndAlso drL("Currency").ToString() <> "" Then
                                lineCurrency = drL("Currency").ToString()
                            End If
                            oInvoice.Lines.Currency = lineCurrency

                            ' 明細匯率 (Rate)
                            Dim lineRate As Double = docRate
                            If Not IsDBNull(drL("Rate")) Then
                                lineRate = Convert.ToDouble(drL("Rate"))
                            End If
                            oInvoice.Lines.Rate = lineRate

                            ' 取得費用項目代碼 (ItemCode)
                            Dim itemCode As String = ""
                            If Not IsDBNull(drL("ItemCode")) Then
                                itemCode = drL("ItemCode").ToString().Trim()
                            End If

                            ' 金額處理 (本幣 vs 外幣)
                            Dim lineTotal As Double = 0
                            If Not IsDBNull(drL("LineTotal")) Then
                                lineTotal = Convert.ToDouble(drL("LineTotal"))
                            End If

                            ' 格式28 (海關代徵營業稅 E030A) 特殊處理：只過帳稅額，且使用零稅碼
                            ' E030A 在送審時已檢查必須有對應的格式28憑證
                            Dim isFormat28 As Boolean = (itemCode = "E030A")
                            If isFormat28 Then
                                ' 取得稅額 (LineVat) 作為 LineTotal
                                If Not IsDBNull(drL("LineVat")) Then
                                    lineTotal = Convert.ToDouble(drL("LineVat"))
                                Else
                                    lineTotal = 0
                                End If
                            End If

                            ' 明細已設定 Currency，直接使用 LineTotal (SAP 會根據幣別自動處理)
                            oInvoice.Lines.LineTotal = lineTotal

                            ' 稅碼 (VatGroup) - 轉換為 SAP 稅碼: 1=J1(應稅), 2=J0(零稅), 3=JX(免稅)
                            ' 格式28 強制使用 J0 (零稅)，因為稅額本身就是最終金額，不應再課稅
                            If isFormat28 Then
                                oInvoice.Lines.VatGroup = "J0"
                            ElseIf Not IsDBNull(drL("VatGroup")) AndAlso drL("VatGroup").ToString() <> "" Then
                                Dim vatCode As String = drL("VatGroup").ToString()
                                Select Case vatCode
                                    Case "1" : oInvoice.Lines.VatGroup = "J1"  ' 應稅
                                    Case "2" : oInvoice.Lines.VatGroup = "J0"  ' 零稅
                                    Case "3" : oInvoice.Lines.VatGroup = "JX"  ' 免稅
                                    Case Else : oInvoice.Lines.VatGroup = vatCode  ' 若已是 SAP 稅碼則直接使用
                                End Select
                            End If

                            ' 稅額 (TaxTotal) - 寫入用戶輸入的稅額，避免被 SAP 重算
                            ' 格式28 不寫入稅額（因為 LineTotal 已經是稅額本身）
                            If Not isFormat28 Then
                                If Not IsDBNull(drL("LineVat")) Then
                                    Dim lineVat As Double = Convert.ToDouble(drL("LineVat"))
                                    If lineVat > 0 Then
                                        oInvoice.Lines.TaxTotal = lineVat
                                    End If
                                End If
                            End If

                            ' 成本中心1 (產品: 1030-AOI, 1040-ICT)
                            If Not IsDBNull(drL("CostingCode")) AndAlso drL("CostingCode").ToString() <> "" Then
                                oInvoice.Lines.CostingCode = drL("CostingCode").ToString()
                            End If

                            ' 成本中心2 (部門: 1050~1300)
                            If Not IsDBNull(drL("CostingCode2")) AndAlso drL("CostingCode2").ToString() <> "" Then
                                oInvoice.Lines.CostingCode2 = drL("CostingCode2").ToString()
                            End If

                            lineIndex += 1
                        End While
                    End Using
                End Using
            End Using

            ' 3. 新增文件
            If oInvoice.Add() <> 0 Then
                Dim errCode As Integer
                oCompany.GetLastError(errCode, errMsg)
                Throw New Exception($"SAP Error [{errCode}]: {errMsg}")
            Else
                ' 取得 SAP DocEntry
                sapDocEntry = Convert.ToInt32(oCompany.GetNewObjectKey())

                ' 用 DocEntry 查詢 SAP OPCH 取得 DocNum
                Dim sapDocNum As Integer = GetSAPDocNum(sapDocEntry)

                ShowSuccess($"已核准並產生 SAP AP 發票 (DocEntry: {sapDocEntry}, DocNum: {sapDocNum})")

                ' 更新 jOPCH 和 jPCH1 的 DocEntry 與 DocNum
                UpdateSAPPostStatus(jID, sapDocEntry, sapDocNum, "Y", "")
            End If

        Catch ex As Exception
            errMsg = ex.Message
            ShowError(errMsg)
            ' 更新失敗狀態
            UpdateSAPPostStatus(jID, 0, 0, "E", errMsg)
        Finally
            ' 釋放 COM 物件 (參考廠長的簡單做法)
            oInvoice = Nothing
            CloseSAPConnection()
        End Try
    End Sub

    ''' <summary>
    ''' 從 SAP OPCH 取得 DocNum (依據 DocEntry)
    ''' </summary>
    Private Function GetSAPDocNum(sapDocEntry As Integer) As Integer
        Try
            Dim sapConnStr As String = System.Configuration.ConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString
            Using connSap As New SqlConnection(sapConnStr)
                connSap.Open()
                Dim sql As String = "SELECT DocNum FROM OPCH WHERE DocEntry = @DocEntry"
                Using cmd As New SqlCommand(sql, connSap)
                    cmd.Parameters.AddWithValue("@DocEntry", sapDocEntry)
                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToInt32(result)
                    End If
                End Using
            End Using
        Catch
            ' 查詢失敗時回傳 0
        End Try
        Return 0
    End Function

    ''' <summary>
    ''' 更新 SAP 過帳狀態並回寫 DocEntry/DocNum 到 jOPCH 和 jPCH1
    ''' </summary>
    ''' <param name="jID">平台單號 (jOPCH.jID)</param>
    ''' <param name="sapDocEntry">SAP 文件 DocEntry</param>
    ''' <param name="sapDocNum">SAP 文件 DocNum</param>
    ''' <param name="status">過帳狀態 (Y=成功, E=失敗)</param>
    ''' <param name="errMsg">錯誤訊息</param>
    Private Sub UpdateSAPPostStatus(jID As Integer, sapDocEntry As Integer, sapDocNum As Integer, status As String, errMsg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 1. 更新 jOPCH 表頭
                Dim sqlHeader As String = "UPDATE jOPCH SET " &
                                          "B1PostStatus = @Status, " &
                                          "B1ErrMsg = @ErrMsg, " &
                                          "B1PostDate = GETDATE()"

                ' 成功時回寫 DocEntry 和 DocNum
                If sapDocEntry > 0 Then
                    sqlHeader &= ", DocEntry = @SapDocEntry"
                End If
                If sapDocNum > 0 Then
                    sqlHeader &= ", DocNum = @SapDocNum"
                End If

                sqlHeader &= " WHERE jID = @jID"

                Using cmdHeader As New SqlCommand(sqlHeader, conn)
                    cmdHeader.Parameters.AddWithValue("@jID", jID)
                    cmdHeader.Parameters.AddWithValue("@Status", status)
                    cmdHeader.Parameters.AddWithValue("@ErrMsg", If(String.IsNullOrEmpty(errMsg), DBNull.Value, errMsg))
                    If sapDocEntry > 0 Then cmdHeader.Parameters.AddWithValue("@SapDocEntry", sapDocEntry)
                    If sapDocNum > 0 Then cmdHeader.Parameters.AddWithValue("@SapDocNum", sapDocNum)
                    cmdHeader.ExecuteNonQuery()
                End Using

                ' 2. 更新 jPCH1 明細 (回寫 DocEntry 和 DocNum)
                If sapDocEntry > 0 OrElse sapDocNum > 0 Then
                    Dim sqlLines As String = "UPDATE jPCH1 SET "
                    Dim setClauses As New List(Of String)

                    If sapDocEntry > 0 Then
                        setClauses.Add("DocEntry = @SapDocEntry")
                    End If
                    If sapDocNum > 0 Then
                        setClauses.Add("DocNum = @SapDocNum")
                    End If

                    sqlLines &= String.Join(", ", setClauses)
                    sqlLines &= " WHERE jID = @jID"

                    Using cmdLines As New SqlCommand(sqlLines, conn)
                        cmdLines.Parameters.AddWithValue("@jID", jID)
                        If sapDocEntry > 0 Then cmdLines.Parameters.AddWithValue("@SapDocEntry", sapDocEntry)
                        If sapDocNum > 0 Then cmdLines.Parameters.AddWithValue("@SapDocNum", sapDocNum)
                        cmdLines.ExecuteNonQuery()
                    End Using
                End If

            End Using
        Catch ex As Exception
            ' 更新狀態失敗不阻斷流程，但記錄到 Event Log (或忽略)
            System.Diagnostics.Debug.WriteLine("UpdateSAPPostStatus Error: " & ex.Message)
        End Try
    End Sub
#End Region


    ''' <summary>
    ''' [B] 驗證狀態轉換是否有效
    ''' </summary>
    Private Function IsValidStatusTransition(currentStatus As String, newStatus As String) As Boolean
        Select Case currentStatus
            Case "W" ' 待審核 → 可以放行(A)或駁回(R)
                Return newStatus = "A" OrElse newStatus = "R"
            Case "R" ' 駁回 → 可以重新送審(W)
                Return newStatus = "W"
            Case "A" ' 已核准 → 終態，不可變更
                Return False
            Case "P", "" ' 草稿或空 → 可以送審(W) - 相容舊資料
                Return newStatus = "W"
            Case Else
                Return False
        End Select
    End Function
#End Region
#Region "輔助函式"
    ''' <summary>
    ''' [G] 顯示錯誤訊息 (紅色)
    ''' </summary>
    Private Sub ShowError(msg As String)
        lblMessage.Text = msg
        lblMessage.ForeColor = System.Drawing.Color.Red
    End Sub

    ''' <summary>
    ''' [G] 顯示成功訊息 (綠色)
    ''' </summary>
    Private Sub ShowSuccess(msg As String)
        lblMessage.Text = msg
        lblMessage.ForeColor = System.Drawing.Color.Green
    End Sub

    ''' <summary>
    ''' [G] 顯示警告訊息 (橘色)
    ''' </summary>
    Private Sub ShowWarning(msg As String)
        lblMessage.Text = msg
        lblMessage.ForeColor = System.Drawing.Color.FromArgb(255, 152, 0) ' Orange
    End Sub

    ''' <summary>
    ''' [G] 顯示資訊訊息 (藍色)
    ''' </summary>
    Private Sub ShowInfo(msg As String)
        lblMessage.Text = msg
        lblMessage.ForeColor = System.Drawing.Color.Blue
    End Sub

    Private Function GetVatRate(vatCode As String) As Decimal
        If vatCode = "1" Then Return 5
        Return 0
    End Function

    Private Function GetStatusText(status As String) As String
        Select Case status
            Case "P" : Return "草稿"
            Case "W" : Return "待審核"
            Case "A" : Return "已核准"
            Case "R" : Return "駁回"
            Case Else : Return status
        End Select
    End Function

    ''' <summary>
    ''' [A] 檢查費用明細與憑證明細金額一致性，回傳警告訊息列表
    ''' </summary>
    Private Function CheckAmountConsistency() As List(Of String)
        Dim warnings As New List(Of String)()

        If CurrentLines.Count > 0 AndAlso CurrentMDRLines.Count > 0 Then
            Dim expenseTotal As Decimal = CurrentLines.Sum(Function(x) x.LineTotal)
            Dim expenseVatSum As Decimal = CurrentLines.Sum(Function(x) x.VatSum)
            Dim mdrTotal As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWBAS)
            Dim mdrVatSum As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWSTE)

            If Math.Abs(expenseTotal - mdrTotal) > 0.01D Then
                warnings.Add(String.Format("費用明細未稅總額 ({0:N2}) 與憑證明細未稅總額 ({1:N2}) 不一致", expenseTotal, mdrTotal))
            End If
            If Math.Abs(expenseVatSum - mdrVatSum) > 0.01D Then
                warnings.Add(String.Format("費用明細稅額 ({0:N2}) 與憑證明細稅額 ({1:N2}) 不一致", expenseVatSum, mdrVatSum))
            End If
        End If

        Return warnings
    End Function

    ''' <summary>
    ''' 檢查 DataReader 是否包含指定欄位
    ''' </summary>
    Private Function HasColumn(dr As SqlDataReader, columnName As String) As Boolean
        For i As Integer = 0 To dr.FieldCount - 1
            If dr.GetName(i).Equals(columnName, StringComparison.InvariantCultureIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function FindControlRecursive(root As Control, controlId As String) As Control
        If root Is Nothing OrElse String.IsNullOrEmpty(controlId) Then
            Return Nothing
        End If
        Dim ctrl As Control = root.FindControl(controlId)
        If ctrl IsNot Nothing Then
            Return ctrl
        End If
        For Each child As Control In root.Controls
            Dim found As Control = FindControlRecursive(child, controlId)
            If found IsNot Nothing Then
                Return found
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' [H] 安全解析日期字串，確保格式一致性
    ''' 支援格式：yyyy-MM-dd, yyyy/MM/dd, MM/dd/yyyy 等常見格式
    ''' </summary>
    Private Function ParseDateSafe(dateStr As String, Optional defaultValue As DateTime? = Nothing) As DateTime?
        If String.IsNullOrEmpty(dateStr) Then Return defaultValue

        Dim result As DateTime
        ' 嘗試使用精確格式解析
        Dim formats As String() = {"yyyy-MM-dd", "yyyy/MM/dd", "MM/dd/yyyy", "dd/MM/yyyy", "yyyyMMdd"}

        If DateTime.TryParseExact(dateStr, formats, System.Globalization.CultureInfo.InvariantCulture,
                                   System.Globalization.DateTimeStyles.None, result) Then
            Return result
        End If

        ' 若精確格式失敗，嘗試一般解析
        If DateTime.TryParse(dateStr, result) Then
            Return result
        End If

        Return defaultValue
    End Function

    ''' <summary>
    ''' [H] 格式化日期為標準字串格式 (yyyy-MM-dd)
    ''' </summary>
    Private Function FormatDate(dt As DateTime?) As String
        If dt.HasValue Then
            Return dt.Value.ToString("yyyy-MM-dd")
        End If
        Return ""
    End Function
#End Region

End Class
