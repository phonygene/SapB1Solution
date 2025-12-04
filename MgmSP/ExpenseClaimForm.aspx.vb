Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Configuration
Imports System.Web.UI.HtmlControls

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

    Private currentUserId As String = ""
    Private currentDocEntry As Integer = 0
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

            CheckApprovalPermission()

            If Request.QueryString("DocEntry") IsNot Nothing Then
                Integer.TryParse(Request.QueryString("DocEntry"), currentDocEntry)
            End If

            If Not IsPostBack Then
                InitializeDropDowns()

                If currentDocEntry > 0 Then
                    LoadDocument(currentDocEntry)
                Else
                    SetDefaultValues()
                    InitializeGridViews()
                End If
            End If

        Catch ex As Exception
            ShowError("頁面載入錯誤: " & ex.Message)
        End Try
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

    Private Sub SetDefaultValues()
        lblDocNum.Text = "[新單據]"
        txtOwner.Text = currentUserId
        ' If ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedValue = currentUserId
        If ddlPurchaser.SelectedIndex = -1 AndAlso ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedIndex = 0
        lblDocStatus.Text = "草稿"
        lblDocStatus.CssClass = "badge status-P"
        txtStatusDisplay.Text = "草稿"

        Dim today As String = DateTime.Now.ToString("yyyy-MM-dd")
        txtDocDate.Text = today
        txtTaxDate.Text = today
        txtDocDueDate.Text = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd")

        If ddlDocCurrency.Items.FindByValue("NTD") IsNot Nothing Then
            ddlDocCurrency.SelectedValue = "NTD"
        ElseIf ddlDocCurrency.Items.FindByValue("TWD") IsNot Nothing Then
            ddlDocCurrency.SelectedValue = "TWD"
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

        ' 按鈕狀態 (新增模式)
        btnSave.Text = "暫存 (Draft)"
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
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT SlpCode, SlpName FROM OSLP"
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

    ' 在 GridView RowDataBound 時呼叫
    Private Sub LoadExpenseCategories(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("-選擇-", ""))
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT CategoryCode, CategoryName, AcctCode FROM expense_category WHERE Active='Y'"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            Dim item As New ListItem(dr("CategoryName").ToString(), dr("CategoryCode").ToString())
                            item.Attributes.Add("data-acct", dr("AcctCode").ToString())
                            ddl.Items.Add(item)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入費用類別失敗")
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

    Private Sub LoadProducts(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT DimCode, DimDesc FROM ODIM"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddl.Items.Add(New ListItem(dr("DimDesc").ToString(), dr("DimCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入產品失敗")
        End Try
    End Sub

    Private Sub LoadDepartments(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT PrcCode, PrcName FROM OPRC"
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
                                    ElseIf ddlDocCurrency.Items.FindByValue("TWD") IsNot Nothing Then
                                        ddlDocCurrency.SelectedValue = "TWD"
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
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入供應商關聯資料失敗: " & ex.Message)
        End Try
    End Sub
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

    Protected Sub ddlDocCurrency_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim curr As String = ddlDocCurrency.SelectedValue
        Dim rate As Decimal = 1D

        If curr = "TWD" Then
            txtDocRate.Text = "1.0"
        Else
            ' 取得 DocDate，若無值則預設為今天
            Dim docDate As DateTime = DateTime.Today
            If Not String.IsNullOrEmpty(txtDocDate.Text) Then
                DateTime.TryParse(txtDocDate.Text, docDate)
            End If

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 1 Rate FROM ORTT WHERE Currency=@Curr AND RateDate <= @DocDate ORDER BY RateDate DESC"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Curr", curr)
                    cmd.Parameters.AddWithValue("@DocDate", docDate)
                    Dim res = cmd.ExecuteScalar()
                    If res IsNot Nothing Then
                        rate = Convert.ToDecimal(res)
                        txtDocRate.Text = rate.ToString("F4")
                    Else
                        txtDocRate.Text = "1.0"
                    End If
                End Using
            End Using
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

            ' Expense Category
            Dim ddlCat As DropDownList = CType(e.Row.FindControl("ddlExpCategory"), DropDownList)
            LoadExpenseCategories(ddlCat)
            If ddlCat.Items.FindByValue(line.CategoryCode) IsNot Nothing Then ddlCat.SelectedValue = line.CategoryCode

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
            LoadDepartments(ddlCost2)
            If ddlCost2.Items.FindByValue(line.CostingCode2) IsNot Nothing Then ddlCost2.SelectedValue = line.CostingCode2

            ' Values
            CType(e.Row.FindControl("txtDescription"), TextBox).Text = line.Description
            CType(e.Row.FindControl("txtAcctCode"), TextBox).Text = line.AcctCode
            CType(e.Row.FindControl("txtLineTotal"), TextBox).Text = line.LineTotal.ToString("0.##")
            CType(e.Row.FindControl("txtVatSum"), TextBox).Text = line.VatSum.ToString("0.##")
            CType(e.Row.FindControl("txtPriceAfterVat"), TextBox).Text = line.PriceAfterVat.ToString("0.##")

            ' Currency & Rate - 已移至單頭，不再於明細列顯示
        End If
    End Sub

    Protected Sub gvExpenseDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        ' 保留未來擴充
    End Sub

    Protected Sub ddlExpCategory_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim ddl As DropDownList = CType(sender, DropDownList)
        Dim row As GridViewRow = CType(ddl.NamingContainer, GridViewRow)
        Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)

        If ddl.SelectedIndex > 0 Then
            txtAcct.Text = ddl.SelectedItem.Attributes("data-acct")
        Else
            txtAcct.Text = ""
        End If

        SyncGridDataToModel()
    End Sub

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
                ' 重新計算稅額與含稅金額
                If line.VatGroup = "1" Then ' 1-應稅 (5%)
                    line.VatRate = 5
                    line.VatSum = Math.Round(line.LineTotal * 0.05D, 0)
                Else
                    line.VatRate = 0
                    line.VatSum = 0
                End If
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
                line.LineTotal = Math.Round(priceAfterVat / 1.05D, 0)
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
                .U_ZFORM_CODE = "21" ' 預設統一發票
            })
            startNum += 1
        Next

        CurrentMDRLines = mdrLines
        BindMDRGrid()
        ShowError("已產生 " & expenseLines.Count.ToString() & " 筆憑證明細，請填寫統一編號與憑證號碼")
        lblMessage.ForeColor = System.Drawing.Color.Blue
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
        SyncMDRGridToModel(True) ' 手動修改稅額
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
                
                ' 若手動修改，則讀取 UI 上的值
                If isManualTax AndAlso txtHWSTE IsNot Nothing Then
                    Decimal.TryParse(txtHWSTE.Text, line.U_HWSTE)
                End If

                ' 自動計算稅額 (僅在非手動模式下)
                If Not isManualTax Then
                    ' 1-應稅 (5%), 2-零稅 (0), 3-免稅 (0)
                    If line.U_TAX_TYPE = "1" Then
                        line.U_HWSTE = Math.Round(line.U_HWBAS * 0.05D, 0)
                    Else
                        line.U_HWSTE = 0
                    End If
                End If

                ' 供應商從單頭取得 (不再從 GridView 讀取)
                line.U_LIFNR = txtCardCode.Text
            End If
        Next
        CurrentMDRLines = lines
    End Sub

    Protected Sub txtXBLNR_TextChanged(sender As Object, e As EventArgs)
        ' 當發票號碼改變時，自動判斷類型
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
        Dim ddlZForm As DropDownList = CType(row.FindControl("ddlZFORM_CODE"), DropDownList)

        Dim invNum As String = txt.Text.Trim()
        Dim formCode As String = "99" ' Default: 其他

        ' 判斷邏輯
        ' 1. 統一發票: 2碼英文 + 8碼數字
        ' 2. 海關報關單: 3碼英文
        ' 3. 高鐵票: 數字
        ' 4. 公營事業: BB或BBN
        ' 5. 其他: 警告

        If System.Text.RegularExpressions.Regex.IsMatch(invNum, "^[A-Z]{2}\d{8}$") Then
            formCode = "21"
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^[A-Z]{3}") Then
            formCode = "28"
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^(BB|BBN)") Then
            formCode = "22"
        ElseIf System.Text.RegularExpressions.Regex.IsMatch(invNum, "^\d+$") Then
            formCode = "24"
        Else
            formCode = "99"
        End If

        If ddlZForm IsNot Nothing Then
            If ddlZForm.Items.FindByValue(formCode) IsNot Nothing Then
                ddlZForm.SelectedValue = formCode
            End If
        End If

        SyncMDRGridToModel()
    End Sub
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
    Protected Sub btnSave_Click(sender As Object, e As EventArgs)
        If Not ValidateAll() Then Return
        SaveDocument("P") ' P = 草稿
    End Sub

    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs)
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

    Private Function ValidateAll() As Boolean
        Dim isValid As Boolean = True
        lblMessage.Text = ""

        If String.IsNullOrEmpty(txtCardCode.Text) Then
            lblErrCardCode.Text = "必填"
            lblErrCardCode.Visible = True
            isValid = False
        End If

        If String.IsNullOrEmpty(txtDocDate.Text) Then
            lblErrDocDate.Text = "必填"
            lblErrDocDate.Visible = True
            isValid = False
        End If

        If String.IsNullOrEmpty(txtDocDueDate.Text) Then
            lblErrDocDueDate.Text = "必填"
            lblErrDocDueDate.Visible = True
            isValid = False
        End If

        ' 明細檢查
        If CurrentLines.Count = 0 AndAlso CurrentMDRLines.Count = 0 Then
            ShowError("請至少新增一筆費用明細或憑證明細")
            isValid = False
        End If

        ' 檢查憑證明細是否有空白的統一編號或憑證號碼 (除非是「99-其他」類型)
        Dim emptyVoucherLines = CurrentMDRLines.Where(Function(x) _
            (String.IsNullOrEmpty(x.U_STCEG) OrElse String.IsNullOrEmpty(x.U_XBLNR)) _
            AndAlso x.U_ZFORM_CODE <> "99")
        
        If emptyVoucherLines.Any() Then
            ShowError("憑證明細有空白的統一編號或憑證號碼，請填寫完整。若為非發票類憑證，請選擇憑證類型為「其他」。")
            isValid = False
        End If

        ' 檢查是否有 "99-其他" 類型的發票且未確認
        Dim hasOtherType As Boolean = CurrentMDRLines.Any(Function(x) x.U_ZFORM_CODE = "99")
        If hasOtherType AndAlso Not WarningConfirmed Then
            ShowError("警告：有「其他」類型的單據，請確認格式是否正確。若確定要新增，請再次點擊儲存/送出。")
            lblMessage.ForeColor = System.Drawing.Color.Orange
            WarningConfirmed = True ' 設定已確認旗標
            isValid = False
        End If

        Return isValid
    End Function

    Private Function SaveDocument(status As String, Optional isAutoSave As Boolean = False) As Boolean
        Try
            SyncGridDataToModel()
            SyncMDRGridToModel()

            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using trans = conn.BeginTransaction()
                    Try
                        Dim jID As Integer = 0

                        ' 1. jOPCH (Header)
                        If currentDocEntry = 0 Then
                            ' Insert
                            Dim sqlH As String = "INSERT INTO jOPCH (CardCode, CardName, NumAtCard, InvNum, DeliveryAddrID, AddressName, Address, " &
                                               "DocDate, DocDueDate, TaxDate, DocCurrency, DocRate, DocTotal, VatSum, " &
                                               "GroupNum, Comments, ApprovalStatus, CreateBy, CreateDate, U_PID, SlpCode) " &
                                               "VALUES (@CardCode, @CardName, @NumAtCard, @InvNum, @DeliveryAddrID, @AddressName, @Address, " &
                                               "@DocDate, @DocDueDate, @TaxDate, @DocCurrency, @DocRate, @DocTotal, @VatSum, " &
                                               "@GroupNum, @Comments, @Status, @User, GETDATE(), @UPID, @SlpCode); " &
                                               "SELECT SCOPE_IDENTITY();"

                            Using cmd As New SqlCommand(sqlH, conn, trans)
                                SetHeaderParameters(cmd, status)
                                jID = Convert.ToInt32(cmd.ExecuteScalar())
                            End Using

                            ' Update DocEntry = jID
                            Dim sqlUpd As String = "UPDATE jOPCH SET DocEntry=@ID, DocNum=@ID WHERE jID=@ID"
                            Using cmd As New SqlCommand(sqlUpd, conn, trans)
                                cmd.Parameters.AddWithValue("@ID", jID)
                                cmd.ExecuteNonQuery()
                            End Using
                            currentDocEntry = jID
                        Else
                            ' Update
                            jID = currentDocEntry
                            Dim sqlH As String = "UPDATE jOPCH SET CardCode=@CardCode, CardName=@CardName, NumAtCard=@NumAtCard, InvNum=@InvNum, " &
                                               "DeliveryAddrID=@DeliveryAddrID, AddressName=@AddressName, Address=@Address, " &
                                               "DocDate=@DocDate, DocDueDate=@DocDueDate, TaxDate=@TaxDate, DocCurrency=@DocCurrency, " &
                                               "DocRate=@DocRate, DocTotal=@DocTotal, VatSum=@VatSum, " &
                                               "GroupNum=@GroupNum, Comments=@Comments, ApprovalStatus=@Status, " &
                                               "UpdateBy=@User, UpdateDate=GETDATE(), U_PID=@UPID, SlpCode=@SlpCode WHERE DocEntry=@ID"

                            Using cmd As New SqlCommand(sqlH, conn, trans)
                                cmd.Parameters.AddWithValue("@ID", currentDocEntry)
                                SetHeaderParameters(cmd, status)
                                cmd.ExecuteNonQuery()
                            End Using

                            ' Delete old lines
                            Using cmd As New SqlCommand("DELETE FROM jPCH1 WHERE DocEntry=@ID", conn, trans)
                                cmd.Parameters.AddWithValue("@ID", currentDocEntry)
                                cmd.ExecuteNonQuery()
                            End Using

                            ' Delete old MDR
                            Using cmd As New SqlCommand("DELETE FROM jMGUIAP WHERE DocEntry=@ID", conn, trans)
                                cmd.Parameters.AddWithValue("@ID", currentDocEntry)
                                cmd.ExecuteNonQuery()
                            End Using
                            Using cmd As New SqlCommand("DELETE FROM jMGUIAPDetail WHERE DocEntry=@ID", conn, trans)
                                cmd.Parameters.AddWithValue("@ID", currentDocEntry)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If

                        ' 2. jPCH1 (Expense Lines)
                        Dim sqlL As String = "INSERT INTO jPCH1 (jID, DocEntry, LineNum, ItemCode, Dscription, AcctCode, " &
                                           "LineTotal, VatGroup, VatPrcnt, LineVat, GTotal, CostingCode, CostingCode2) " &
                                           "VALUES (@jID, @DocEntry, @LineNum, @ItemCode, @Dscription, @AcctCode, " &
                                           "@LineTotal, @VatGroup, @VatPrcnt, @LineVat, @GTotal, @CostingCode, @CostingCode2)"

                        For Each line As ExpenseLine In CurrentLines
                            Using cmd As New SqlCommand(sqlL, conn, trans)
                                cmd.Parameters.AddWithValue("@jID", jID) ' FK to Header jID
                                cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
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
                                cmd.ExecuteNonQuery()
                            End Using
                        Next

                        ' 3. jMGUIAP & jMGUIAPDetail (MDR)
                        If CurrentMDRLines.Count > 0 Then
                            ' MDR Header (彙總)
                            Dim mdrTotal As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWBAS)
                            Dim mdrVat As Decimal = CurrentMDRLines.Sum(Function(x) x.U_HWSTE)

                            Dim sqlMdrH As String = "INSERT INTO jMGUIAP (jID, DocEntry, DocNum, DocTotal, VatSum, CreateBy, CreateDate) " &
                                                  "VALUES (@jID, @DocEntry, @DocNum, @DocTotal, @VatSum, @User, GETDATE()); SELECT SCOPE_IDENTITY();"

                            Dim mdrID As Integer = 0
                            Using cmd As New SqlCommand(sqlMdrH, conn, trans)
                                cmd.Parameters.AddWithValue("@jID", jID)
                                cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
                                cmd.Parameters.AddWithValue("@DocNum", currentDocEntry)
                                cmd.Parameters.AddWithValue("@DocTotal", mdrTotal)
                                cmd.Parameters.AddWithValue("@VatSum", mdrVat)
                                cmd.Parameters.AddWithValue("@User", currentUserId)
                                mdrID = Convert.ToInt32(cmd.ExecuteScalar())
                            End Using

                            ' MDR Lines
                            Dim sqlMdrL As String = "INSERT INTO jMGUIAPDetail (jID, DocEntry, LineNum, U_LIFNR, U_STCEG, U_XBLNR, U_ZFORM_CODE, " &
                                                  "U_BLDAT, U_VATDATE, U_HWBAS, U_HWSTE, U_TAX_TYPE) " &
                                                  "VALUES (@jID, @DocEntry, @LineNum, @LIFNR, @STCEG, @XBLNR, @ZFORM, @BLDAT, @VATDATE, @HWBAS, @HWSTE, @TAXTYPE)"

                            For Each line As MDRLine In CurrentMDRLines
                                Using cmd As New SqlCommand(sqlMdrL, conn, trans)
                                    cmd.Parameters.AddWithValue("@jID", mdrID)
                                    cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
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

                        ' Update U_LastCur in OCRD if needed
                        ' (若表頭幣別為多幣別的業務夥伴，每次新增單據時更新 U_LastCur)
                        ' 這裡簡化邏輯：直接嘗試更新，若 U_LastCur 不存在會報錯，但我們假設已建立
                        If ddlDocCurrency.Enabled Then ' 表示該業務夥伴為多幣別 (##)
                            Dim sqlUpdCur As String = "UPDATE OCRD SET U_LastCur = @Curr WHERE CardCode = @CardCode"
                            Using cmd As New SqlCommand(sqlUpdCur, conn, trans)
                                ' 注意：OCRD 在 SAP DB，這裡 conn 是 jtdb...
                                ' 跨庫交易較複雜，我們可能需要另開連線到 SAP DB，但不參與此 Transaction (風險可接受)
                            End Using

                            ' 分開執行 SAP DB Update
                            Using sapConn As New SqlConnection(sapConnStr)
                                sapConn.Open()
                                Using cmdSap As New SqlCommand(sqlUpdCur, sapConn)
                                    cmdSap.Parameters.AddWithValue("@Curr", ddlDocCurrency.SelectedValue)
                                    cmdSap.Parameters.AddWithValue("@CardCode", txtCardCode.Text)
                                    cmdSap.ExecuteNonQuery()
                                End Using
                            End Using
                        End If

                        trans.Commit()
                        lblMessage.Text = "儲存成功"
                        lblDocNum.Text = currentDocEntry.ToString()

                        If Not isAutoSave Then
                            Response.Redirect("ExpenseClaimForm.aspx?DocEntry=" & currentDocEntry)
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
        If currentDocEntry > 0 Then
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Using trans = conn.BeginTransaction()
                        Try
                            ' Delete Lines
                            Dim cmd As New SqlCommand("DELETE FROM jPCH1 WHERE DocEntry=@ID", conn, trans)
                            cmd.Parameters.AddWithValue("@ID", currentDocEntry)
                            cmd.ExecuteNonQuery()

                            ' Delete MDR
                            cmd.CommandText = "DELETE FROM jMGUIAP WHERE DocEntry=@ID"
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "DELETE FROM jMGUIAPDetail WHERE DocEntry=@ID"
                            cmd.ExecuteNonQuery()

                            ' Delete Header
                            cmd.CommandText = "DELETE FROM jOPCH WHERE DocEntry=@ID"
                            cmd.ExecuteNonQuery()
                            
                            ' Delete Attachments
                            cmd.CommandText = "DELETE FROM jAttach WHERE jID=@ID"
                            cmd.ExecuteNonQuery()

                            trans.Commit()
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

    Private Sub SetHeaderParameters(cmd As SqlCommand, status As String)
        cmd.Parameters.AddWithValue("@CardCode", txtCardCode.Text)
        cmd.Parameters.AddWithValue("@CardName", txtCardName.Text)
        cmd.Parameters.AddWithValue("@NumAtCard", txtNumAtCard.Text)
        cmd.Parameters.AddWithValue("@InvNum", "") ' 初次儲存時為空，審核放行後由 SAP 回填
        cmd.Parameters.AddWithValue("@DeliveryAddrID", ddlDeliveryAddr.SelectedValue)
        cmd.Parameters.AddWithValue("@AddressName", ddlDeliveryAddr.SelectedItem.Text)
        cmd.Parameters.AddWithValue("@Address", txtAddress.Text)
        cmd.Parameters.AddWithValue("@DocDate", txtDocDate.Text)
        cmd.Parameters.AddWithValue("@DocDueDate", txtDocDueDate.Text)
        cmd.Parameters.AddWithValue("@TaxDate", If(String.IsNullOrEmpty(txtTaxDate.Text), DBNull.Value, txtTaxDate.Text))
        cmd.Parameters.AddWithValue("@DocCurrency", ddlDocCurrency.SelectedValue)
        cmd.Parameters.AddWithValue("@DocRate", txtDocRate.Text)
        cmd.Parameters.AddWithValue("@DocTotal", Decimal.Parse(lblDocTotalWithTax.Text.Replace(",", "")))
        cmd.Parameters.AddWithValue("@VatSum", Decimal.Parse(lblVatSum.Text.Replace(",", "")))
        cmd.Parameters.AddWithValue("@GroupNum", ddlGroupNum.SelectedValue)
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
            ' Load Header
            Dim sql As String = "SELECT * FROM jOPCH WHERE DocEntry=@ID"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ID", id)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        lblDocNum.Text = dr("DocEntry").ToString()
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

                        ' 顯示審核區塊邏輯:
                        ' 1. 這個區塊要 jtdb 的 User Table 裡的 AP_App 欄位為 1 才可以編輯
                        ' 2. 一般 User 不能編輯也不能按放行/發送意見/退回
                        ' 3. 但為了查看退回意見等，建議顯示但唯讀

                        pnlApproval.Visible = True ' 只要是既有單據，預設顯示，內容依權限控制

                        txtApprovalComments.ReadOnly = Not isApUser
                        
                        ' 按鈕保持顯示，但無權限者停用
                        btnApprove.Visible = True
                        btnApprove.Enabled = isApUser
                        
                        btnUpdateComment.Visible = True
                        btnUpdateComment.Enabled = isApUser
                        
                        btnReject.Visible = True
                        btnReject.Enabled = isApUser

                        ' 按鈕狀態 (編輯模式)
                        btnSave.Text = "更新 (Update)"
                        btnDelete.Visible = True
                    End If
                End Using
            End Using

            ' Load Attachments (jAttach)
            Dim attachList As New List(Of AttachmentItem)
            sql = "SELECT * FROM jAttach WHERE jID=@ID ORDER BY UploadDate, UploadTime"
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

            ' Load Expense Lines
            Dim lines As New List(Of ExpenseLine)
            sql = "SELECT * FROM jPCH1 WHERE DocEntry=@ID ORDER BY LineNum"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ID", id)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        lines.Add(New ExpenseLine With {
                            .LineNum = Convert.ToInt32(dr("LineNum")),
                            .CategoryCode = dr("ItemCode").ToString(),
                            .Description = dr("Dscription").ToString(),
                            .AcctCode = dr("AcctCode").ToString(),
                            .LineTotal = Convert.ToDecimal(dr("LineTotal")),
                            .VatGroup = dr("VatGroup").ToString(),
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

            ' Load MDR Lines
            Dim mdrLines As New List(Of MDRLine)
            sql = "SELECT * FROM jMGUIAPDetail WHERE DocEntry=@ID ORDER BY LineNum"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ID", id)
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

    Protected Sub btnUpload_Click(sender As Object, e As EventArgs)
        ' .NET 4.0 相容性修改: 檢查 Request.Files
        If fileUpload.HasFile OrElse Request.Files.Count > 0 Then
            Try
                ' 若尚未存檔 (currentDocEntry=0)，先自動儲存為草稿以取得 jID
                If currentDocEntry = 0 Then
                    If Not SaveDocument("P", True) Then
                        ShowError("上傳附件前自動儲存草稿失敗，請檢查必填欄位。")
                        Return
                    End If
                End If

                ' 建立附件資料夾: AttachFile/User/{UserID}/ExpenseClaimForm/{jID}/
                Dim folder As String = GetAttachmentFolder(currentDocEntry)
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
                        Dim relativePath As String = GetAttachmentRelativePath(currentDocEntry, savedName)

                        ' 寫入資料庫 jAttach
                        Using conn As New SqlConnection(connStr)
                            conn.Open()
                            Dim sql As String = "INSERT INTO jAttach (jID, DocEntry, LineNum, FilePath, FileName, Uploader, UploadTime) " &
                                              "VALUES (@jID, @DocEntry, -1, @FilePath, @FileName, @Uploader, @UploadTime); SELECT SCOPE_IDENTITY();"
                            Using cmd As New SqlCommand(sql, conn)
                                cmd.Parameters.AddWithValue("@jID", currentDocEntry) ' jOPCH.jID (DocEntry)
                                cmd.Parameters.AddWithValue("@DocEntry", currentDocEntry)
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
                lblMessage.Text = $"成功上傳 {successCount} 個檔案"
                lblMessage.ForeColor = System.Drawing.Color.Blue

            Catch ex As Exception
                ShowError("上傳失敗: " & ex.Message)
            End Try
        Else
             ShowError("請選擇檔案")
        End If
    End Sub

    Protected Sub gvAttachments_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "DeleteFile" Then
            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            Dim list = CurrentAttachments
            
            If index < list.Count Then
                Dim item = list(index)
                
                Try
                    ' 刪除 DB
                    Using conn As New SqlConnection(connStr)
                        conn.Open()
                        Dim sql As String = "DELETE FROM jAttach WHERE ID=@ID"
                        Using cmd As New SqlCommand(sql, conn)
                            cmd.Parameters.AddWithValue("@ID", item.ID)
                            cmd.ExecuteNonQuery()
                        End Using
                    End Using
                    
                    ' 刪除實體檔案 (可選，通常保留備份或標記刪除，這裡先不刪實體檔以防誤刪)
                    ' Dim path As String = Server.MapPath("~/Uploads/Expense/" & item.FilePath)
                    ' If File.Exists(path) Then File.Delete(path)

                    list.RemoveAt(index)
                    CurrentAttachments = list
                    BindAttachmentGrid()
                    lblMessage.Text = "附件已刪除"
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
                Dim sql As String = "UPDATE jOPCH SET ApprovalComments=@Comm, UpdateBy=@User, UpdateDate=GETDATE() WHERE DocEntry=@ID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@User", currentUserId)
                    cmd.Parameters.AddWithValue("@Comm", txtApprovalComments.Text)
                    cmd.Parameters.AddWithValue("@ID", currentDocEntry)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
            ShowError("意見已更新") ' 用紅色顯示有點怪，但先這樣
            lblMessage.ForeColor = System.Drawing.Color.Blue
        Catch ex As Exception
            ShowError("更新失敗: " & ex.Message)
        End Try
    End Sub

    Private Sub UpdateStatus(status As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE jOPCH SET ApprovalStatus=@Status, ApprovedBy=@User, ApprovalDate=GETDATE(), ApprovalComments=@Comm WHERE DocEntry=@ID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Status", status)
                    cmd.Parameters.AddWithValue("@User", currentUserId)
                    cmd.Parameters.AddWithValue("@Comm", txtApprovalComments.Text)
                    cmd.Parameters.AddWithValue("@ID", currentDocEntry)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            ' 如果是放行 (A)，這裡應呼叫 SAP B1 API 建立 AP Invoice
            If status = "A" Then
                ' TODO: Call SAP B1 API & MDR Integration
                ' CreateAPInvoiceInSAP(currentDocEntry)
            End If

            Response.Redirect("ExpenseClaimForm.aspx?DocEntry=" & currentDocEntry)
        Catch ex As Exception
            ShowError("更新狀態失敗: " & ex.Message)
        End Try
    End Sub
#End Region
#Region "輔助函式"
    Private Sub ShowError(msg As String)
        lblMessage.Text = msg
        lblMessage.ForeColor = System.Drawing.Color.Red
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
#End Region

End Class
