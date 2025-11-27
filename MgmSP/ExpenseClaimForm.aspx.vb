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
#End Region

#Region "變數宣告"
    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private ReadOnly sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString

    Private currentUserId As String = ""
    Private currentDocEntry As Integer = 0
    Private canApprove As Boolean = False
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
#End Region

#Region "頁面載入"
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
            ' 檢查使用者是否有審核權限
            ' 使用 Approver 欄位 (1=可審核)
            Dim sql As String = "SELECT count(*) FROM [User] WHERE id = @UserId AND Approver = 1"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@UserId", currentUserId)
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                canApprove = (count > 0)
            End Using
        End Using

        ' 只有待審核狀態且有權限者才顯示審核區塊 (LoadDocument 時會再判斷狀態)
        pnlApproval.Visible = False
    End Sub

    Private Sub SetDefaultValues()
        lblDocNum.Text = "[新單據]"
        txtOwner.Text = currentUserId
        If ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedValue = currentUserId
        If ddlPurchaser.SelectedIndex = -1 AndAlso ddlPurchaser.Items.Count > 0 Then ddlPurchaser.SelectedIndex = 0
        lblDocStatus.Text = "草稿"
        lblDocStatus.CssClass = "badge status-P"
        txtStatusDisplay.Text = "草稿"

        Dim today As String = DateTime.Now.ToString("yyyy-MM-dd")
        txtDocDate.Text = today
        txtTaxDate.Text = today
        txtDocDueDate.Text = DateTime.Now.AddDays(30).ToString("yyyy-MM-dd")

        If ddlDocCurrency.Items.FindByValue("TWD") IsNot Nothing Then
            ddlDocCurrency.SelectedValue = "TWD"
        End If
        txtDocRate.Text = "1.0"
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
                Dim sql As String = "SELECT SlpName FROM OSLP"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddlPurchaser.Items.Add(New ListItem(dr("SlpName").ToString(), dr("SlpName").ToString()))
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

            ' 判斷搜尋模式
            Dim isExact As Boolean = (rblSearchMode.SelectedValue = "Exact")

            If Not String.IsNullOrEmpty(keyword) Then
                keyword = keyword.Replace("*", "").Replace("%", "")
                If isExact Then
                    sqlWhere &= " AND CardCode = @Kw"
                Else
                    sqlWhere &= " AND (CardCode LIKE @Kw OR CardName LIKE @Kw)"
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
                ' 幣別
                Dim sqlCurr As String = "SELECT T1.[CurrCode] FROM OCRD T0 INNER JOIN OCRN T1 ON T0.[Currency] = T1.[CurrCode] WHERE T0.[CardCode] = @CardCode"
                Using cmd As New SqlCommand(sqlCurr, conn)
                    cmd.Parameters.AddWithValue("@CardCode", cardCode)
                    Dim curr As Object = cmd.ExecuteScalar()
                    If curr IsNot Nothing Then
                        Dim currCode As String = curr.ToString()
                        If ddlDocCurrency.Items.FindByValue(currCode) IsNot Nothing Then
                            ddlDocCurrency.SelectedValue = currCode
                            ddlDocCurrency_SelectedIndexChanged(Nothing, Nothing) ' Trigger rate update
                        End If
                    End If
                End Using

                ' 付款條件
                Dim sqlPymnt As String = "SELECT T1.[GroupNum] FROM OCRD T0 INNER JOIN OCTG T1 ON T0.GroupNum = T1.GroupNum WHERE T0.[CardCode] = @CardCode"
                Using cmd As New SqlCommand(sqlPymnt, conn)
                    cmd.Parameters.AddWithValue("@CardCode", cardCode)
                    Dim grp As Object = cmd.ExecuteScalar()
                    If grp IsNot Nothing Then
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
        If curr = "TWD" Then
            txtDocRate.Text = "1.0"
        Else
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 1 Rate FROM ORTT WHERE Currency=@Curr AND RateDate <= GETDATE() ORDER BY RateDate DESC"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Curr", curr)
                    Dim res = cmd.ExecuteScalar()
                    If res IsNot Nothing Then
                        txtDocRate.Text = Convert.ToDecimal(res).ToString("F4")
                    Else
                        txtDocRate.Text = "1.0"
                    End If
                End Using
            End Using
        End If
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
        lines.Add(New ExpenseLine With {
            .LineNum = newNum,
            .LineTotal = 0,
            .VatRate = 0,
            .PriceAfterVat = 0
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
            CType(e.Row.FindControl("lblVatSum"), Label).Text = line.VatSum.ToString("N2")
            CType(e.Row.FindControl("txtPriceAfterVat"), TextBox).Text = line.PriceAfterVat.ToString("0.##")
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
        SyncGridDataToModel()
        BindGrid()
    End Sub

    Protected Sub CalculatePriceAfterVat(sender As Object, e As EventArgs)
        ' 當含稅金額改變時觸發 -> 反算未稅金額
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)
        Dim rowIndex As Integer = row.DataItemIndex

        Dim priceAfterVat As Decimal = 0
        Decimal.TryParse(txt.Text, priceAfterVat)

        SyncGridDataToModel()

        Dim lines = CurrentLines
        If rowIndex < lines.Count Then
            Dim line = lines(rowIndex)
            line.PriceAfterVat = priceAfterVat

            ' 反算邏輯
            ' 1-應稅 (5%)
            If line.VatGroup = "1" Then
                line.VatRate = 5
                line.LineTotal = Math.Round(priceAfterVat / 1.05D, 0) ' 或 2 位小數? 通常SAP未稅是反算
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

    Private Sub SyncGridDataToModel()
        Dim lines = CurrentLines
        For i As Integer = 0 To gvExpenseDetail.Rows.Count - 1
            Dim row As GridViewRow = gvExpenseDetail.Rows(i)
            If i < lines.Count Then
                Dim ddlExp As DropDownList = CType(row.FindControl("ddlExpCategory"), DropDownList)
                Dim txtDesc As TextBox = CType(row.FindControl("txtDescription"), TextBox)
                Dim txtAcct As TextBox = CType(row.FindControl("txtAcctCode"), TextBox)
                Dim txtTotal As TextBox = CType(row.FindControl("txtLineTotal"), TextBox)
                Dim ddlVat As DropDownList = CType(row.FindControl("ddlVatGroup"), DropDownList)
                Dim ddlCost As DropDownList = CType(row.FindControl("ddlCostingCode"), DropDownList)
                Dim ddlCost2 As DropDownList = CType(row.FindControl("ddlCostingCode2"), DropDownList)

                Dim line = lines(i)
                If ddlExp IsNot Nothing Then line.CategoryCode = ddlExp.SelectedValue
                If txtDesc IsNot Nothing Then line.Description = txtDesc.Text
                If txtAcct IsNot Nothing Then line.AcctCode = txtAcct.Text
                If txtTotal IsNot Nothing Then Decimal.TryParse(txtTotal.Text, line.LineTotal)
                If ddlVat IsNot Nothing Then line.VatGroup = ddlVat.SelectedValue
                If ddlCost IsNot Nothing Then line.CostingCode = ddlCost.SelectedValue
                If ddlCost2 IsNot Nothing Then line.CostingCode2 = ddlCost2.SelectedValue

                ' 計算
                ' 1-應稅 (5%), 2-零稅 (0), 3-免稅 (0)
                If line.VatGroup = "1" Then
                    line.VatRate = 5
                    line.VatSum = Math.Round(line.LineTotal * 0.05D, 0)
                Else
                    line.VatRate = 0
                    line.VatSum = 0
                End If
                line.PriceAfterVat = line.LineTotal + line.VatSum
            End If
        Next
        CurrentLines = lines
    End Sub
#End Region

#Region "MDR GridView"
    Protected Sub btnAddMDRRow_Click(sender As Object, e As EventArgs)
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

    Private Sub BindMDRGrid()
        gvMDRDetail.DataSource = CurrentMDRLines
        gvMDRDetail.DataBind()
    End Sub

    Protected Sub gvMDRDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        ' 簡單綁定，大部分是 TextBox 直接在前端顯示
    End Sub

    Protected Sub CalculateMDRTotal(sender As Object, e As EventArgs)
        SyncMDRGridToModel()
        BindMDRGrid()
    End Sub

    Private Sub SyncMDRGridToModel()
        Dim lines = CurrentMDRLines
        For i As Integer = 0 To gvMDRDetail.Rows.Count - 1
            Dim row As GridViewRow = gvMDRDetail.Rows(i)
            If i < lines.Count Then
                Dim line = lines(i)

                Dim txtLIFNR As TextBox = CType(row.FindControl("txtLIFNR"), TextBox)
                Dim txtSTCEG As TextBox = CType(row.FindControl("txtSTCEG"), TextBox)
                Dim txtXBLNR As TextBox = CType(row.FindControl("txtXBLNR"), TextBox)
                Dim ddlZFORM As DropDownList = CType(row.FindControl("ddlZFORM_CODE"), DropDownList)
                Dim txtBLDAT As TextBox = CType(row.FindControl("txtBLDAT"), TextBox)
                Dim txtVATDATE As TextBox = CType(row.FindControl("txtVATDATE"), TextBox)
                Dim txtHWBAS As TextBox = CType(row.FindControl("txtHWBAS"), TextBox)
                Dim ddlTAX As DropDownList = CType(row.FindControl("ddlTAX_TYPE"), DropDownList)

                If txtLIFNR IsNot Nothing Then line.U_LIFNR = txtLIFNR.Text
                If txtSTCEG IsNot Nothing Then line.U_STCEG = txtSTCEG.Text
                If txtXBLNR IsNot Nothing Then line.U_XBLNR = txtXBLNR.Text
                If ddlZFORM IsNot Nothing Then line.U_ZFORM_CODE = ddlZFORM.SelectedValue

                Dim dt As DateTime
                If txtBLDAT IsNot Nothing AndAlso DateTime.TryParse(txtBLDAT.Text, dt) Then line.U_BLDAT = dt
                If txtVATDATE IsNot Nothing AndAlso DateTime.TryParse(txtVATDATE.Text, dt) Then line.U_VATDATE = dt

                If txtHWBAS IsNot Nothing Then Decimal.TryParse(txtHWBAS.Text, line.U_HWBAS)
                If ddlTAX IsNot Nothing Then line.U_TAX_TYPE = ddlTAX.SelectedValue

                ' 計算稅額
                ' 1-應稅 (5%), 2-零稅 (0), 3-免稅 (0)
                If line.U_TAX_TYPE = "1" Then
                    line.U_HWSTE = Math.Round(line.U_HWBAS * 0.05D, 0)
                Else
                    line.U_HWSTE = 0
                End If
            End If
        Next
        CurrentMDRLines = lines
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
            ShowError("請至少新增一筆費用明細或發票明細")
            isValid = False
        End If

        Return isValid
    End Function

    Private Sub SaveDocument(status As String)
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
                                               "GroupNum, Comments, ApprovalStatus, CreateBy, CreateDate, U_PID) " &
                                               "VALUES (@CardCode, @CardName, @NumAtCard, @InvNum, @DeliveryAddrID, @AddressName, @Address, " &
                                               "@DocDate, @DocDueDate, @TaxDate, @DocCurrency, @DocRate, @DocTotal, @VatSum, " &
                                               "@GroupNum, @Comments, @Status, @User, GETDATE(), @UPID); " &
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
                                               "UpdateBy=@User, UpdateDate=GETDATE(), U_PID=@UPID WHERE DocEntry=@ID"

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
                        Dim sqlL As String = "INSERT INTO jPCH1 (DocEntry, LineNum, ItemCode, Dscription, AcctCode, " &
                                           "LineTotal, VatGroup, VatPrcnt, LineVat, GTotal, CostingCode, CostingCode2) " &
                                           "VALUES (@DocEntry, @LineNum, @ItemCode, @Dscription, @AcctCode, " &
                                           "@LineTotal, @VatGroup, @VatPrcnt, @LineVat, @GTotal, @CostingCode, @CostingCode2)"

                        For Each line As ExpenseLine In CurrentLines
                            Using cmd As New SqlCommand(sqlL, conn, trans)
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

                            Dim sqlMdrH As String = "INSERT INTO jMGUIAP (DocEntry, DocNum, DocTotal, VatSum, CreateBy, CreateDate) " &
                                                  "VALUES (@DocEntry, @DocNum, @DocTotal, @VatSum, @User, GETDATE()); SELECT SCOPE_IDENTITY();"

                            Dim mdrID As Integer = 0
                            Using cmd As New SqlCommand(sqlMdrH, conn, trans)
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

                        trans.Commit()
                        lblMessage.Text = "儲存成功"
                        lblDocNum.Text = currentDocEntry.ToString()

                        Response.Redirect("ExpenseClaimForm.aspx?DocEntry=" & currentDocEntry)
                    Catch ex As Exception
                        trans.Rollback()
                        Throw ex
                    End Try
                End Using
            End Using
        Catch ex As Exception
            ShowError("儲存失敗: " & ex.Message)
        End Try
    End Sub
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
        Response.Redirect("Index.aspx")
    End Sub

    Private Sub SetHeaderParameters(cmd As SqlCommand, status As String)
        cmd.Parameters.AddWithValue("@CardCode", txtCardCode.Text)
        cmd.Parameters.AddWithValue("@CardName", txtCardName.Text)
        cmd.Parameters.AddWithValue("@NumAtCard", txtNumAtCard.Text)
        cmd.Parameters.AddWithValue("@InvNum", txtInvNum.Text)
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
        cmd.Parameters.AddWithValue("@Attachment", If(String.IsNullOrEmpty(lblAttachment.Text), DBNull.Value, lblAttachment.Text))
        cmd.Parameters.AddWithValue("@Status", status)
        cmd.Parameters.AddWithValue("@User", currentUserId)
        cmd.Parameters.AddWithValue("@UPID", If(String.IsNullOrEmpty(txtUPID.Text), DBNull.Value, txtUPID.Text))
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
                        txtInvNum.Text = dr("InvNum").ToString()

                        If Not IsDBNull(dr("DeliveryAddrID")) Then ddlDeliveryAddr.SelectedValue = dr("DeliveryAddrID").ToString()
                        txtAddress.Text = dr("Address").ToString()

                        txtDocDate.Text = Convert.ToDateTime(dr("DocDate")).ToString("yyyy-MM-dd")
                        If Not IsDBNull(dr("DocDueDate")) Then txtDocDueDate.Text = Convert.ToDateTime(dr("DocDueDate")).ToString("yyyy-MM-dd")
                        If Not IsDBNull(dr("TaxDate")) Then txtTaxDate.Text = Convert.ToDateTime(dr("TaxDate")).ToString("yyyy-MM-dd")

                        ddlDocCurrency.SelectedValue = dr("DocCurrency").ToString()
                        txtDocRate.Text = dr("DocRate").ToString()

                        If Not IsDBNull(dr("GroupNum")) Then ddlGroupNum.SelectedValue = dr("GroupNum").ToString()
                        txtRemarks.Text = dr("Comments").ToString()

                        If Not IsDBNull(dr("Attachment")) Then
                            lblAttachment.Text = dr("Attachment").ToString()
                        End If

                        Dim status As String = dr("ApprovalStatus").ToString()
                        lblDocStatus.Text = GetStatusText(status)
                        lblDocStatus.CssClass = "badge status-" & status
                        txtStatusDisplay.Text = GetStatusText(status)
                        txtApprovalStatus.Text = status

                        txtApprovedBy.Text = dr("ApprovedBy").ToString()
                        txtOwner.Text = dr("CreateBy").ToString()
                        If ddlPurchaser.Items.FindByValue(dr("CreateBy").ToString()) IsNot Nothing Then
                            ddlPurchaser.SelectedValue = dr("CreateBy").ToString()
                        End If

                        If Not IsDBNull(dr("U_PID")) Then txtUPID.Text = dr("U_PID").ToString()

                        ' 顯示審核區塊邏輯
                        If status = "W" AndAlso canApprove Then
                            pnlApproval.Visible = True
                        End If
                    End If
                End Using
            End Using

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
    Protected Sub btnUpload_Click(sender As Object, e As EventArgs)
        If fileUpload.HasFile Then
            Try
                Dim folder As String = Server.MapPath("~/Uploads/Expense/")
                If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)

                Dim fileName As String = DateTime.Now.ToString("yyyyMMddHHmmss") & "_" & fileUpload.FileName
                fileUpload.SaveAs(folder & fileName)
                lblAttachment.Text = fileName
            Catch ex As Exception
                ShowError("上傳失敗: " & ex.Message)
            End Try
        End If
    End Sub
#End Region

#Region "審核"
    Protected Sub btnApprove_Click(sender As Object, e As EventArgs)
        UpdateStatus("A")
    End Sub

    Protected Sub btnReject_Click(sender As Object, e As EventArgs)
        UpdateStatus("R")
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
