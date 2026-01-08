Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Configuration

''' <summary>
''' 請購單 (Purchase Request Form)
''' 建立日期: 2026-01-07
''' 基於費用申請單 (ExpenseClaimForm) 修改
''' </summary>
Partial Public Class PurchaseRequestForm
    Inherits System.Web.UI.Page

#Region "類別定義"
    <Serializable()>
    Public Class PRLine
        Public Property LineNum As Integer
        Public Property ItemCode As String
        Public Property Description As String
        Public Property LineText As String ' 明細摘要
        Public Property Quantity As Decimal
        Public Property Price As Decimal
        Public Property LineTotal As Decimal ' 未稅金額
        Public Property VatGroup As String
        Public Property VatRate As Decimal
        Public Property VatSum As Decimal ' 稅額
        Public Property GTotal As Decimal ' 含稅金額
        Public Property PriceAfVAT As Decimal ' 含稅單價 (SAP B1 命名)
        Public Property WhsCode As String
        Public Property ShipDate As DateTime?
        Public Property CostingCode As String
        Public Property CostingCode2 As String
        ' 顯示用
        Public Property Currency As String
        Public Property Rate As Decimal
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
    Private currentJID As Integer = 0
    Private canApprove As Boolean = False
#End Region

#Region "屬性 (ViewState)"
    Private Property CurrentLines As List(Of PRLine)
        Get
            If ViewState("CurrentLines") Is Nothing Then
                ViewState("CurrentLines") = New List(Of PRLine)()
            End If
            Return CType(ViewState("CurrentLines"), List(Of PRLine))
        End Get
        Set(value As List(Of PRLine))
            ViewState("CurrentLines") = value
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

    ''' <summary>
    ''' 待選供應商代碼 (用於價格更新確認流程)
    ''' </summary>
    Private Property PendingVendorCode As String
        Get
            Return If(ViewState("PendingVendorCode"), "").ToString()
        End Get
        Set(value As String)
            ViewState("PendingVendorCode") = value
        End Set
    End Property

    ''' <summary>
    ''' 待選供應商名稱 (用於價格更新確認流程)
    ''' </summary>
    Private Property PendingVendorName As String
        Get
            Return If(ViewState("PendingVendorName"), "").ToString()
        End Get
        Set(value As String)
            ViewState("PendingVendorName") = value
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
            lblCurrentUser.Text = currentUserId

            CheckApprovalPermission()

            If Request.QueryString("jID") IsNot Nothing Then
                Integer.TryParse(Request.QueryString("jID"), currentJID)
            End If

            If Not IsPostBack Then
                InitializeDropDowns()

                If currentJID > 0 Then
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
        ' 請購單審核權限使用 PU_App 欄位
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Dim sql As String = "SELECT PU_App FROM [User] WHERE id = @UserId"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@UserId", currentUserId)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        canApprove = (Convert.ToInt32(If(IsDBNull(dr("PU_App")), 0, dr("PU_App"))) = 1)
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub SetDefaultValues()
        lblDocNum.Text = "[新單據]"
        txtReqName.Text = currentUserId
        txtOwner.Text = currentUserId
        lblDocStatus.Text = "新增中"
        lblDocStatus.CssClass = "badge status-W"
        txtStatusDisplay.Text = "新增中"

        Dim today As String = DateTime.Now.ToString("yyyy-MM-dd")
        txtDocDate.Text = today

        If ddlDocCurrency.Items.FindByValue("NTD") IsNot Nothing Then
            ddlDocCurrency.SelectedValue = "NTD"
        End If
        txtDocRate.Text = "1.0"

        txtApprovalComments.ReadOnly = True
        btnApprove.Visible = True
        btnApprove.Enabled = False
        btnReject.Visible = True
        btnReject.Enabled = False

        btnDelete.Visible = False
    End Sub
#End Region

#Region "初始化資料"
    Private Sub InitializeDropDowns()
        LoadCurrencies()
        LoadDepartments()
        ' LoadWarehouses, LoadProducts, LoadDepartments2 改為在 RowDataBound 中呼叫 (參照費用申請單模式)
        LoadPurchasers()
        LoadVatGroups()
    End Sub

    Private Sub InitializeGridViews()
        If CurrentLines.Count = 0 Then
            AddNewEmptyLine()
        End If
        BindGrid()
        BindAttachmentGrid()
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

    Private Sub LoadDepartments()
        ddlReqDept.Items.Clear()
        ddlReqDept.Items.Add(New ListItem("- 請選擇 -", ""))
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT EDeptID, EDeptName FROM jDEPT ORDER BY EDeptID"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddlReqDept.Items.Add(New ListItem(dr("EDeptName").ToString(), dr("EDeptID").ToString()))
                        End While
                    End Using
                End Using
            End Using

            ' 嘗試設定使用者預設部門
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT expDEPT FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", currentUserId)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Dim userDept As String = result.ToString()
                        If ddlReqDept.Items.FindByValue(userDept) IsNot Nothing Then
                            ddlReqDept.SelectedValue = userDept
                        End If
                    End If
                End Using
            End Using
        Catch ex As Exception
            ShowError("載入部門失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 載入倉庫下拉選單 (參照費用申請單模式 - 直接填充控制項)
    ''' </summary>
    Private Sub LoadWarehouses(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT WhsCode, WhsName FROM OWHS WHERE Inactive = 'N' ORDER BY WhsCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddl.Items.Add(New ListItem(dr("WhsCode").ToString(), dr("WhsCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 靜默處理，避免影響頁面載入
        End Try
    End Sub

    Private Sub LoadPurchasers()
        ddlPurchaser.Items.Clear()
        ddlPurchaser.Items.Add(New ListItem("-- 請選擇 --", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
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
    ''' 載入產品別下拉選單 (參照費用申請單模式 - 直接填充控制項)
    ''' </summary>
    Private Sub LoadProducts(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT PrcCode, PrcName FROM OPRC WHERE DimCode = 1 AND PrcCode NOT LIKE 'Centr%' ORDER BY PrcCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddl.Items.Add(New ListItem(dr("PrcName").ToString() & " (" & dr("PrcCode").ToString() & ")", dr("PrcCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 靜默處理，避免影響頁面載入
        End Try
    End Sub

    ''' <summary>
    ''' 載入部門別下拉選單 (參照費用申請單模式 - 直接填充控制項)
    ''' </summary>
    Private Sub LoadDepartments2(ddl As DropDownList)
        ddl.Items.Clear()
        ddl.Items.Add(New ListItem("", ""))
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT PrcCode, PrcName FROM OPRC WHERE DimCode = 2 AND PrcCode NOT LIKE 'Centr%' ORDER BY PrcCode"
                Using cmd As New SqlCommand(sql, conn)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        While dr.Read()
                            ddl.Items.Add(New ListItem(dr("PrcName").ToString() & " (" & dr("PrcCode").ToString() & ")", dr("PrcCode").ToString()))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 靜默處理，避免影響頁面載入
        End Try
    End Sub

    Private Sub LoadVatGroups()
        ' 硬編碼稅碼 (參照費用申請單)
        ViewState("VatGroups") = New DataTable()
        Dim dt As DataTable = CType(ViewState("VatGroups"), DataTable)
        dt.Columns.Add("Code", GetType(String))
        dt.Columns.Add("Name", GetType(String))
        dt.Columns.Add("Rate", GetType(Decimal))

        dt.Rows.Add("1", "1-應稅 (5%)", 5D)
        dt.Rows.Add("2", "2-零稅 (0%)", 0D)
        dt.Rows.Add("3", "3-免稅 (0%)", 0D)
    End Sub
#End Region

#Region "GridView 綁定"
    Private Sub BindGrid()
        gvPRDetail.DataSource = CurrentLines
        gvPRDetail.DataBind()
        UpdateTotals()
    End Sub

    Private Sub BindAttachmentGrid()
        gvAttachments.DataSource = CurrentAttachments
        gvAttachments.DataBind()
    End Sub

    Protected Sub gvPRDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim line As PRLine = CType(e.Row.DataItem, PRLine)

            ' 綁定品號
            Dim txtItemCode As TextBox = CType(e.Row.FindControl("txtItemCode"), TextBox)
            If txtItemCode IsNot Nothing Then
                txtItemCode.Text = line.ItemCode
            End If

            ' 綁定說明
            Dim txtDescription As TextBox = CType(e.Row.FindControl("txtDescription"), TextBox)
            If txtDescription IsNot Nothing Then
                txtDescription.Text = line.Description
            End If

            ' 綁定摘要
            Dim txtLineText As TextBox = CType(e.Row.FindControl("txtLineText"), TextBox)
            If txtLineText IsNot Nothing Then
                txtLineText.Text = If(line.LineText, "")
            End If

            ' 綁定數量
            Dim txtQuantity As TextBox = CType(e.Row.FindControl("txtQuantity"), TextBox)
            If txtQuantity IsNot Nothing Then
                txtQuantity.Text = If(line.Quantity > 0, line.Quantity.ToString("N2"), "1")
            End If

            ' 綁定單價 (未稅)
            Dim txtPrice As TextBox = CType(e.Row.FindControl("txtPrice"), TextBox)
            If txtPrice IsNot Nothing Then
                txtPrice.Text = line.Price.ToString("N2")
            End If

            ' 綁定含稅單價
            Dim txtPriceAfVAT As TextBox = CType(e.Row.FindControl("txtPriceAfVAT"), TextBox)
            If txtPriceAfVAT IsNot Nothing Then
                txtPriceAfVAT.Text = line.PriceAfVAT.ToString("N2")
            End If

            ' 綁定稅碼 (顯示名稱，參照費用申請單)
            Dim ddlVatGroup As DropDownList = CType(e.Row.FindControl("ddlVatGroup"), DropDownList)
            If ddlVatGroup IsNot Nothing Then
                ddlVatGroup.Items.Clear()
                Dim dt As DataTable = CType(ViewState("VatGroups"), DataTable)
                If dt IsNot Nothing Then
                    For Each dr As DataRow In dt.Rows
                        Dim item As New ListItem(dr("Name").ToString(), dr("Code").ToString())
                        item.Attributes.Add("data-rate", dr("Rate").ToString())
                        ddlVatGroup.Items.Add(item)
                    Next
                End If
                If Not String.IsNullOrEmpty(line.VatGroup) AndAlso ddlVatGroup.Items.FindByValue(line.VatGroup) IsNot Nothing Then
                    ddlVatGroup.SelectedValue = line.VatGroup
                Else
                    ' 預設選擇應稅
                    If ddlVatGroup.Items.FindByValue("1") IsNot Nothing Then
                        ddlVatGroup.SelectedValue = "1"
                    End If
                End If
            End If

            ' 綁定稅額
            Dim txtVatSum As TextBox = CType(e.Row.FindControl("txtVatSum"), TextBox)
            If txtVatSum IsNot Nothing Then
                txtVatSum.Text = line.VatSum.ToString("N0")
            End If

            ' 綁定含稅金額
            Dim txtGTotal As TextBox = CType(e.Row.FindControl("txtGTotal"), TextBox)
            If txtGTotal IsNot Nothing Then
                txtGTotal.Text = line.GTotal.ToString("N0")
            End If

            ' 綁定倉庫 (參照費用申請單模式 - 直接查詢填充)
            Dim ddlWhsCode As DropDownList = CType(e.Row.FindControl("ddlWhsCode"), DropDownList)
            If ddlWhsCode IsNot Nothing Then
                LoadWarehouses(ddlWhsCode)
                If Not String.IsNullOrEmpty(line.WhsCode) AndAlso ddlWhsCode.Items.FindByValue(line.WhsCode) IsNot Nothing Then
                    ddlWhsCode.SelectedValue = line.WhsCode
                End If
            End If

            ' 綁定交期
            Dim txtShipDate As TextBox = CType(e.Row.FindControl("txtShipDate"), TextBox)
            If txtShipDate IsNot Nothing AndAlso line.ShipDate.HasValue Then
                txtShipDate.Text = line.ShipDate.Value.ToString("yyyy-MM-dd")
            End If

            ' 綁定產品別 (參照費用申請單模式 - 直接查詢填充)
            Dim ddlCostingCode As DropDownList = CType(e.Row.FindControl("ddlCostingCode"), DropDownList)
            If ddlCostingCode IsNot Nothing Then
                LoadProducts(ddlCostingCode)
                If Not String.IsNullOrEmpty(line.CostingCode) AndAlso ddlCostingCode.Items.FindByValue(line.CostingCode) IsNot Nothing Then
                    ddlCostingCode.SelectedValue = line.CostingCode
                End If
            End If

            ' 綁定部門別 (參照費用申請單模式 - 直接查詢填充)
            Dim ddlCostingCode2 As DropDownList = CType(e.Row.FindControl("ddlCostingCode2"), DropDownList)
            If ddlCostingCode2 IsNot Nothing Then
                LoadDepartments2(ddlCostingCode2)
                If Not String.IsNullOrEmpty(line.CostingCode2) AndAlso ddlCostingCode2.Items.FindByValue(line.CostingCode2) IsNot Nothing Then
                    ddlCostingCode2.SelectedValue = line.CostingCode2
                End If
            End If
        End If
    End Sub
#End Region

#Region "明細操作"
    Private Sub AddNewEmptyLine()
        Dim newLine As New PRLine() With {
            .LineNum = CurrentLines.Count + 1,
            .ItemCode = "",
            .Description = "",
            .Quantity = 1,
            .Price = 0,
            .PriceAfVAT = 0,
            .LineTotal = 0,
            .VatGroup = "1",  ' 預設應稅
            .VatRate = 5,     ' 5%
            .VatSum = 0,
            .GTotal = 0,
            .WhsCode = "",
            .ShipDate = Nothing,
            .CostingCode = "",
            .CostingCode2 = ""
        }
        CurrentLines.Add(newLine)
    End Sub

    Protected Sub btnAddLine_Click(sender As Object, e As EventArgs)
        SyncGridToList()
        AddNewEmptyLine()
        BindGrid()
    End Sub

    Protected Sub btnDeleteLine_Click(sender As Object, e As EventArgs)
        SyncGridToList()
        Dim linesToRemove As New List(Of Integer)()

        For i As Integer = 0 To gvPRDetail.Rows.Count - 1
            Dim chk As CheckBox = CType(gvPRDetail.Rows(i).FindControl("chkSelect"), CheckBox)
            If chk IsNot Nothing AndAlso chk.Checked Then
                linesToRemove.Add(i)
            End If
        Next

        ' 從後往前刪除，避免索引錯位
        For i As Integer = linesToRemove.Count - 1 To 0 Step -1
            CurrentLines.RemoveAt(linesToRemove(i))
        Next

        ' 重新編號
        For i As Integer = 0 To CurrentLines.Count - 1
            CurrentLines(i).LineNum = i + 1
        Next

        BindGrid()
    End Sub

    ''' <summary>
    ''' 同步 GridView 資料到 Model (參照費用申請單 SyncGridDataToModel)
    ''' </summary>
    ''' <param name="readPriceAfVAT">是否從含稅單價反推</param>
    Private Sub SyncGridToList(Optional readPriceAfVAT As Boolean = False)
        For i As Integer = 0 To gvPRDetail.Rows.Count - 1
            If i < CurrentLines.Count Then
                Dim row As GridViewRow = gvPRDetail.Rows(i)
                Dim line As PRLine = CurrentLines(i)

                Dim txtItemCode As TextBox = CType(row.FindControl("txtItemCode"), TextBox)
                If txtItemCode IsNot Nothing Then line.ItemCode = txtItemCode.Text.Trim()

                Dim txtDescription As TextBox = CType(row.FindControl("txtDescription"), TextBox)
                If txtDescription IsNot Nothing Then line.Description = txtDescription.Text.Trim()

                Dim txtLineText As TextBox = CType(row.FindControl("txtLineText"), TextBox)
                If txtLineText IsNot Nothing Then line.LineText = txtLineText.Text.Trim()

                Dim txtQuantity As TextBox = CType(row.FindControl("txtQuantity"), TextBox)
                If txtQuantity IsNot Nothing Then
                    Decimal.TryParse(txtQuantity.Text, line.Quantity)
                End If

                Dim txtPrice As TextBox = CType(row.FindControl("txtPrice"), TextBox)
                If txtPrice IsNot Nothing Then
                    Decimal.TryParse(txtPrice.Text, line.Price)
                End If

                ' 讀取含稅單價
                Dim txtPriceAfVAT As TextBox = CType(row.FindControl("txtPriceAfVAT"), TextBox)
                Dim inputPriceAfVAT As Decimal = 0
                If txtPriceAfVAT IsNot Nothing Then
                    Decimal.TryParse(txtPriceAfVAT.Text, inputPriceAfVAT)
                End If

                Dim ddlVatGroup As DropDownList = CType(row.FindControl("ddlVatGroup"), DropDownList)
                If ddlVatGroup IsNot Nothing Then
                    line.VatGroup = ddlVatGroup.SelectedValue
                    ' 取得稅率
                    Dim dt As DataTable = CType(ViewState("VatGroups"), DataTable)
                    If dt IsNot Nothing Then
                        Dim rows = dt.Select("Code = '" & line.VatGroup & "'")
                        If rows.Length > 0 Then
                            line.VatRate = Convert.ToDecimal(rows(0)("Rate"))
                        Else
                            line.VatRate = 0
                        End If
                    End If
                End If

                ' 讀取使用者輸入的稅額
                Dim txtVatSum As TextBox = CType(row.FindControl("txtVatSum"), TextBox)
                Dim userVatSum As Decimal = 0
                If txtVatSum IsNot Nothing Then
                    Decimal.TryParse(txtVatSum.Text, userVatSum)
                End If

                Dim ddlWhsCode As DropDownList = CType(row.FindControl("ddlWhsCode"), DropDownList)
                If ddlWhsCode IsNot Nothing Then line.WhsCode = ddlWhsCode.SelectedValue

                Dim txtShipDate As TextBox = CType(row.FindControl("txtShipDate"), TextBox)
                If txtShipDate IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtShipDate.Text) Then
                    Dim shipDate As DateTime
                    If DateTime.TryParse(txtShipDate.Text, shipDate) Then
                        line.ShipDate = shipDate
                    End If
                Else
                    line.ShipDate = Nothing
                End If

                Dim ddlCostingCode As DropDownList = CType(row.FindControl("ddlCostingCode"), DropDownList)
                If ddlCostingCode IsNot Nothing Then line.CostingCode = ddlCostingCode.SelectedValue

                Dim ddlCostingCode2 As DropDownList = CType(row.FindControl("ddlCostingCode2"), DropDownList)
                If ddlCostingCode2 IsNot Nothing Then line.CostingCode2 = ddlCostingCode2.SelectedValue

                ' ==========================================
                ' 金額計算邏輯 (參照費用申請單)
                ' ==========================================
                If readPriceAfVAT AndAlso inputPriceAfVAT > 0 AndAlso line.Quantity > 0 Then
                    ' 從含稅單價反推 (參照費用申請單 CalculateFromPriceAfterVat)
                    Dim gTotal As Decimal = inputPriceAfVAT * line.Quantity
                    If line.VatGroup = "1" Then
                        ' 應稅：反推未稅金額
                        line.LineTotal = Math.Round(gTotal / 1.05D, 0, MidpointRounding.AwayFromZero)
                        line.VatSum = gTotal - line.LineTotal
                    Else
                        ' 零稅/免稅：含稅金額 = 未稅金額
                        line.LineTotal = gTotal
                        line.VatSum = 0
                    End If
                    line.GTotal = gTotal
                    line.PriceAfVAT = inputPriceAfVAT
                    ' 反推未稅單價
                    line.Price = If(line.Quantity > 0, line.LineTotal / line.Quantity, 0)
                Else
                    ' 正常計算：從未稅單價計算
                    line.LineTotal = line.Quantity * line.Price

                    ' 稅額計算邏輯 (參照費用申請單)
                    If line.VatGroup = "2" OrElse line.VatGroup = "3" Then
                        ' 零稅或免稅
                        line.VatSum = 0
                    ElseIf userVatSum = 0 AndAlso line.LineTotal > 0 AndAlso line.VatGroup = "1" Then
                        ' 使用者未輸入稅額，自動計算 (使用無條件捨去)
                        line.VatSum = Math.Floor(line.LineTotal * 0.05D)
                    Else
                        ' 保留使用者輸入的稅額
                        line.VatSum = userVatSum
                    End If

                    ' 計算含稅金額和含稅單價
                    line.GTotal = line.LineTotal + line.VatSum
                    line.PriceAfVAT = If(line.Quantity > 0, line.GTotal / line.Quantity, 0)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' 數量或未稅單價變更 - 重新計算金額
    ''' </summary>
    Protected Sub CalculateLineTotal(sender As Object, e As EventArgs)
        SyncGridToList(False)
        BindGrid()
    End Sub

    ''' <summary>
    ''' 含稅單價變更 - 從含稅單價反推未稅單價和稅額
    ''' </summary>
    Protected Sub CalculateFromPriceAfVAT(sender As Object, e As EventArgs)
        SyncGridToList(True)
        BindGrid()
    End Sub

    ''' <summary>
    ''' 稅額變更 - 保留使用者輸入的稅額
    ''' </summary>
    Protected Sub CalculateVatSum(sender As Object, e As EventArgs)
        SyncGridToList(False)
        BindGrid()
    End Sub

    Private Sub UpdateTotals()
        Dim totalWithoutTax As Decimal = 0
        Dim totalTax As Decimal = 0

        For Each line As PRLine In CurrentLines
            totalWithoutTax += line.LineTotal
            totalTax += line.VatSum
        Next

        lblDocTotal.Text = totalWithoutTax.ToString("N2")
        lblVatSum.Text = totalTax.ToString("N2")
        lblDocTotalWithTax.Text = (totalWithoutTax + totalTax).ToString("N2")
    End Sub
#End Region

#Region "品號搜尋"
    Protected Sub gvPRDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "SearchItem" Then
            SyncGridToList()
            hfItemSearchRowIndex.Value = e.CommandArgument.ToString()
            txtItemSearchKeyword.Text = ""
            BindItemSearchGrid("")
            mpeItem.Show()
        End If
    End Sub

    Protected Sub btnDoSearchItem_Click(sender As Object, e As EventArgs)
        gvItemSearch.PageIndex = 0  ' 重新搜尋時回到第一頁
        BindItemSearchGrid(txtItemSearchKeyword.Text.Trim())
        mpeItem.Show()
    End Sub

    Protected Sub gvItemSearch_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
        gvItemSearch.PageIndex = e.NewPageIndex
        BindItemSearchGrid(txtItemSearchKeyword.Text.Trim())
        mpeItem.Show()
    End Sub

    Private Sub BindItemSearchGrid(keyword As String)
        Try
            Dim sqlWhere As String = "WHERE frozenFor = 'N' "
            Dim isExact As Boolean = (rblItemSearchMode.SelectedValue = "Exact")

            If Not String.IsNullOrEmpty(keyword) Then
                keyword = keyword.Replace("*", "").Replace("%", "")

                If isExact Then
                    sqlWhere &= " AND (ItemCode LIKE @Kw OR ItemName LIKE @Kw)"
                    keyword = keyword & "%"
                Else
                    sqlWhere &= " AND (ItemCode LIKE @Kw OR ItemName LIKE @Kw)"
                    keyword = "%" & keyword & "%"
                End If
            End If

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = $"SELECT TOP 100 ItemCode, ItemName, LastPurPrc FROM OITM {sqlWhere} ORDER BY ItemCode"
                Using cmd As New SqlCommand(sql, conn)
                    If Not String.IsNullOrEmpty(keyword) Then
                        cmd.Parameters.AddWithValue("@Kw", keyword)
                    End If

                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)
                        gvItemSearch.DataSource = dt
                        gvItemSearch.DataBind()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("品號搜尋失敗: " & ex.Message)
        End Try
    End Sub

    Protected Sub gvItemSearch_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "SelectItem" Then
            Dim args() As String = e.CommandArgument.ToString().Split("|"c)
            Dim itemCode As String = args(0)
            Dim itemName As String = args(1)
            Dim lastPurPrc As Decimal = 0
            If args.Length > 2 Then
                Decimal.TryParse(args(2), lastPurPrc)
            End If

            Dim rowIndex As Integer = Convert.ToInt32(hfItemSearchRowIndex.Value)
            If rowIndex >= 0 AndAlso rowIndex < CurrentLines.Count Then
                Dim line = CurrentLines(rowIndex)
                line.ItemCode = itemCode
                line.Description = itemName
                line.Price = lastPurPrc

                ' 確保稅碼有預設值
                If String.IsNullOrEmpty(line.VatGroup) Then
                    line.VatGroup = "1"  ' 預設應稅
                    line.VatRate = 5
                End If

                ' 重新計算金額 (參照費用申請單邏輯)
                line.LineTotal = line.Quantity * line.Price
                If line.VatGroup = "1" Then
                    line.VatSum = Math.Floor(line.LineTotal * 0.05D)
                Else
                    line.VatSum = 0
                End If
                line.GTotal = line.LineTotal + line.VatSum
                line.PriceAfVAT = If(line.Quantity > 0, line.GTotal / line.Quantity, 0)
            End If

            mpeItem.Hide()
            BindGrid()
        End If
    End Sub
#End Region

#Region "供應商搜尋"
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
    End Sub

    Protected Sub btnDoSearchVendor_Click(sender As Object, e As EventArgs)
        gvVendorSearch.PageIndex = 0  ' 重新搜尋時回到第一頁
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

            Dim searchSource As String = hfSearchSource.Value
            Dim isExact As Boolean = (rblSearchMode.SelectedValue = "Exact")

            If Not String.IsNullOrEmpty(keyword) Then
                keyword = keyword.Replace("*", "").Replace("%", "")

                If isExact Then
                    If searchSource = "Code" Then
                        sqlWhere &= " AND CardCode LIKE @Kw"
                    Else
                        sqlWhere &= " AND CardName LIKE @Kw"
                    End If
                    keyword = keyword & "%"
                Else
                    If searchSource = "Code" Then
                        sqlWhere &= " AND CardCode LIKE @Kw"
                    Else
                        sqlWhere &= " AND CardName LIKE @Kw"
                    End If
                    keyword = "%" & keyword & "%"
                End If
            End If

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = $"SELECT TOP 100 CardCode, CardName FROM OCRD {sqlWhere} ORDER BY CardCode"
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
            Dim args() As String = e.CommandArgument.ToString().Split("|"c)
            Dim selectedCardCode As String = args(0)
            Dim selectedCardName As String = args(1)

            mpeVendor.Hide()

            ' 檢查是否有已輸入品號的明細
            Dim hasItems As Boolean = CurrentLines.Any(Function(l) Not String.IsNullOrEmpty(l.ItemCode))

            If hasItems Then
                ' 有項目，詢問是否更新價格
                PendingVendorCode = selectedCardCode
                PendingVendorName = selectedCardName
                mpePriceUpdate.Show()
            Else
                ' 沒有項目，直接設定供應商
                txtCardCode.Text = selectedCardCode
                txtCardName.Text = selectedCardName
            End If
        End If
    End Sub

    ''' <summary>
    ''' 取消價格更新 - 保持現有價格，只設定供應商
    ''' </summary>
    Protected Sub btnPriceUpdateCancel_Click(sender As Object, e As EventArgs)
        txtCardCode.Text = PendingVendorCode
        txtCardName.Text = PendingVendorName
        PendingVendorCode = ""
        PendingVendorName = ""
        mpePriceUpdate.Hide()
    End Sub

    ''' <summary>
    ''' 確認價格更新 - 設定供應商並更新項目價格
    ''' </summary>
    Protected Sub btnPriceUpdateConfirm_Click(sender As Object, e As EventArgs)
        txtCardCode.Text = PendingVendorCode
        txtCardName.Text = PendingVendorName

        ' 更新所有項目的價格
        For Each line As PRLine In CurrentLines
            If Not String.IsNullOrEmpty(line.ItemCode) Then
                ' 從 OPOR/POR1 取得該供應商對該項目的最後採購價
                Dim vendorPrice As Decimal = GetVendorLastPrice(PendingVendorCode, line.ItemCode)
                line.Price = vendorPrice

                ' 重新計算金額
                line.LineTotal = line.Quantity * line.Price
                If line.VatGroup = "1" Then
                    line.VatSum = Math.Floor(line.LineTotal * 0.05D)
                Else
                    line.VatSum = 0
                End If
                line.GTotal = line.LineTotal + line.VatSum
                line.PriceAfVAT = If(line.Quantity > 0, line.GTotal / line.Quantity, 0)
            End If
        Next

        PendingVendorCode = ""
        PendingVendorName = ""
        mpePriceUpdate.Hide()
        BindGrid()
    End Sub

    ''' <summary>
    ''' 取得供應商對該項目的最後採購價
    ''' 邏輯: 從 OPOR/POR1 找該供應商最後一次下單該項目的折扣後單價 (Price)
    ''' 若找不到則使用該項目的最後採購價 (OITM.LastPurPrc)
    ''' </summary>
    Private Function GetVendorLastPrice(cardCode As String, itemCode As String) As Decimal
        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()

                ' 1. 先從 OPOR/POR1 找供應商最後一次採購該項目的價格
                Dim sql As String = "SELECT TOP 1 T1.Price " &
                                    "FROM OPOR T0 " &
                                    "INNER JOIN POR1 T1 ON T0.DocEntry = T1.DocEntry " &
                                    "WHERE T0.CardCode = @CardCode " &
                                    "AND T1.ItemCode = @ItemCode " &
                                    "AND T0.CANCELED = 'N' " &
                                    "ORDER BY T0.DocDate DESC, T0.DocEntry DESC"

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@CardCode", cardCode)
                    cmd.Parameters.AddWithValue("@ItemCode", itemCode)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToDecimal(result)
                    End If
                End Using

                ' 2. 找不到供應商採購記錄，使用項目的最後採購價
                sql = "SELECT LastPurPrc FROM OITM WHERE ItemCode = @ItemCode"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ItemCode", itemCode)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToDecimal(result)
                    End If
                End Using

                Return 0
            End Using
        Catch ex As Exception
            ' 發生錯誤時返回 0
            Return 0
        End Try
    End Function
#End Region

#Region "幣別匯率"
    Protected Sub ddlDocCurrency_SelectedIndexChanged(sender As Object, e As EventArgs)
        UpdateExchangeRate()
    End Sub

    Protected Sub txtDocRate_TextChanged(sender As Object, e As EventArgs)
        ' 使用者手動輸入匯率
    End Sub

    Protected Sub btnRefreshRate_Click(sender As Object, e As EventArgs)
        UpdateExchangeRate()
    End Sub

    Private Sub UpdateExchangeRate()
        Dim currency As String = ddlDocCurrency.SelectedValue
        If String.IsNullOrEmpty(currency) OrElse currency = "NTD" Then
            txtDocRate.Text = "1.0"
            Return
        End If

        Dim docDate As DateTime = DateTime.Now
        If Not String.IsNullOrEmpty(txtDocDate.Text) Then
            DateTime.TryParse(txtDocDate.Text, docDate)
        End If

        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 1 Rate FROM ORTT WHERE Currency = @Currency AND RateDate <= @RateDate ORDER BY RateDate DESC"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Currency", currency)
                    cmd.Parameters.AddWithValue("@RateDate", docDate)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        txtDocRate.Text = Convert.ToDecimal(result).ToString("N6")
                    Else
                        txtDocRate.Text = "1.0"
                    End If
                End Using
            End Using
        Catch ex As Exception
            ShowError("取得匯率失敗: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "附件上傳"
    Protected Sub btnUpload_Click(sender As Object, e As EventArgs)
        If Not fileUpload.HasFile Then
            ShowError("請選擇要上傳的檔案")
            Return
        End If

        Try
            Dim uploadFolder As String = Server.MapPath("~/Uploads/PR/")
            If Not Directory.Exists(uploadFolder) Then
                Directory.CreateDirectory(uploadFolder)
            End If

            For Each uploadedFile As HttpPostedFile In fileUpload.PostedFiles
                Dim fileName As String = Path.GetFileName(uploadedFile.FileName)
                Dim uniqueName As String = DateTime.Now.ToString("yyyyMMddHHmmss") & "_" & fileName
                Dim filePath As String = Path.Combine(uploadFolder, uniqueName)
                uploadedFile.SaveAs(filePath)

                Dim attachment As New AttachmentItem() With {
                    .ID = CurrentAttachments.Count + 1,
                    .FileName = fileName,
                    .FilePath = filePath,
                    .UploadDate = DateTime.Now,
                    .UploadTime = DateTime.Now.ToString("HH:mm:ss"),
                    .Uploader = currentUserId
                }
                CurrentAttachments.Add(attachment)
            Next

            BindAttachmentGrid()
            lblMessage.Text = "附件上傳成功"
            lblMessage.ForeColor = Drawing.Color.Green
        Catch ex As Exception
            ShowError("附件上傳失敗: " & ex.Message)
        End Try
    End Sub

    Protected Sub gvAttachments_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "DeleteFile" Then
            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            If index >= 0 AndAlso index < CurrentAttachments.Count Then
                ' 刪除實體檔案
                If File.Exists(CurrentAttachments(index).FilePath) Then
                    File.Delete(CurrentAttachments(index).FilePath)
                End If
                CurrentAttachments.RemoveAt(index)
                BindAttachmentGrid()
            End If
        End If
    End Sub
#End Region

#Region "文件儲存"
    Protected Sub btnSubmit_Click(sender As Object, e As EventArgs)
        SyncGridToList()

        Dim errors As New List(Of String)()
        Dim warnings As New List(Of String)()

        ' 驗證
        If String.IsNullOrEmpty(txtDocDate.Text) Then
            errors.Add("請購日期為必填")
        End If

        If CurrentLines.Count = 0 OrElse CurrentLines.All(Function(l) String.IsNullOrEmpty(l.ItemCode)) Then
            errors.Add("請至少新增一筆明細")
        End If

        For i As Integer = 0 To CurrentLines.Count - 1
            Dim line = CurrentLines(i)
            If Not String.IsNullOrEmpty(line.ItemCode) Then
                If line.Quantity <= 0 Then
                    errors.Add($"第 {i + 1} 行: 數量必須大於 0")
                End If
            End If
        Next

        If errors.Count > 0 Then
            ShowValidationErrors(errors, warnings)
            Return
        End If

        Try
            SaveDocument()
            lblMessage.Text = "請購單儲存成功！單號: " & txtJID.Text
            lblMessage.ForeColor = Drawing.Color.Green

            ' 重新載入
            Response.Redirect("PurchaseRequestForm.aspx?jID=" & txtJID.Text)
        Catch ex As Exception
            ShowError("儲存失敗: " & ex.Message)
        End Try
    End Sub

    Private Sub SaveDocument()
        Using conn As New SqlConnection(connStr)
            conn.Open()
            Using trans As SqlTransaction = conn.BeginTransaction()
                Try
                    Dim jID As Integer = currentJID

                    If jID = 0 Then
                        ' 新增 - 先從 OJID 取得全域唯一的 jID
                        Dim ojidSql As String = "INSERT INTO OJID (jUser) VALUES (@jUser); SELECT SCOPE_IDENTITY();"
                        Using cmd As New SqlCommand(ojidSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jUser", currentUserId)
                            jID = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using

                        ' 使用 IDENTITY_INSERT 插入指定的 jID
                        Dim insertSql As String = "SET IDENTITY_INSERT jOPRQ ON; " &
                                                   "INSERT INTO jOPRQ (jID, CardCode, CardName, ReqName, ReqDept, DocDate, ReqDate, DocCurrency, DocRate, DocTotal, VatSum, Comments, DocStatus, ApprovalStatus, U_PID, CreateDate, CreateBy) " &
                                                   "VALUES (@jID, @CardCode, @CardName, @ReqName, @ReqDept, @DocDate, @ReqDate, @DocCurrency, @DocRate, @DocTotal, @VatSum, @Comments, 'O', 'Pending', @U_PID, GETDATE(), @CreateBy); " &
                                                   "SET IDENTITY_INSERT jOPRQ OFF;"

                        Using cmd As New SqlCommand(insertSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jID", jID)
                            cmd.Parameters.AddWithValue("@CardCode", If(String.IsNullOrEmpty(txtCardCode.Text), DBNull.Value, txtCardCode.Text))
                            cmd.Parameters.AddWithValue("@CardName", If(String.IsNullOrEmpty(txtCardName.Text), DBNull.Value, txtCardName.Text))
                            cmd.Parameters.AddWithValue("@ReqName", txtReqName.Text)
                            cmd.Parameters.AddWithValue("@ReqDept", If(String.IsNullOrEmpty(ddlReqDept.SelectedValue), DBNull.Value, ddlReqDept.SelectedValue))
                            cmd.Parameters.AddWithValue("@DocDate", DateTime.Parse(txtDocDate.Text))
                            cmd.Parameters.AddWithValue("@ReqDate", If(String.IsNullOrEmpty(txtReqDate.Text), DBNull.Value, DateTime.Parse(txtReqDate.Text)))
                            cmd.Parameters.AddWithValue("@DocCurrency", ddlDocCurrency.SelectedValue)
                            cmd.Parameters.AddWithValue("@DocRate", Decimal.Parse(txtDocRate.Text))
                            cmd.Parameters.AddWithValue("@DocTotal", Decimal.Parse(lblDocTotalWithTax.Text))
                            cmd.Parameters.AddWithValue("@VatSum", Decimal.Parse(lblVatSum.Text))
                            cmd.Parameters.AddWithValue("@Comments", If(String.IsNullOrEmpty(txtRemarks.Text), DBNull.Value, txtRemarks.Text))
                            cmd.Parameters.AddWithValue("@U_PID", If(String.IsNullOrEmpty(txtUPID.Text), DBNull.Value, txtUPID.Text))
                            cmd.Parameters.AddWithValue("@CreateBy", currentUserId)
                            cmd.ExecuteNonQuery()
                        End Using
                    Else
                        ' 更新
                        Dim updateSql As String = "UPDATE jOPRQ SET CardCode=@CardCode, CardName=@CardName, ReqDept=@ReqDept, DocDate=@DocDate, ReqDate=@ReqDate, " &
                                                   "DocCurrency=@DocCurrency, DocRate=@DocRate, DocTotal=@DocTotal, VatSum=@VatSum, Comments=@Comments, U_PID=@U_PID, UpdateDate=GETDATE(), UpdateBy=@UpdateBy WHERE jID=@jID"

                        Using cmd As New SqlCommand(updateSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jID", jID)
                            cmd.Parameters.AddWithValue("@CardCode", If(String.IsNullOrEmpty(txtCardCode.Text), DBNull.Value, txtCardCode.Text))
                            cmd.Parameters.AddWithValue("@CardName", If(String.IsNullOrEmpty(txtCardName.Text), DBNull.Value, txtCardName.Text))
                            cmd.Parameters.AddWithValue("@ReqDept", If(String.IsNullOrEmpty(ddlReqDept.SelectedValue), DBNull.Value, ddlReqDept.SelectedValue))
                            cmd.Parameters.AddWithValue("@DocDate", DateTime.Parse(txtDocDate.Text))
                            cmd.Parameters.AddWithValue("@ReqDate", If(String.IsNullOrEmpty(txtReqDate.Text), DBNull.Value, DateTime.Parse(txtReqDate.Text)))
                            cmd.Parameters.AddWithValue("@DocCurrency", ddlDocCurrency.SelectedValue)
                            cmd.Parameters.AddWithValue("@DocRate", Decimal.Parse(txtDocRate.Text))
                            cmd.Parameters.AddWithValue("@DocTotal", Decimal.Parse(lblDocTotalWithTax.Text))
                            cmd.Parameters.AddWithValue("@VatSum", Decimal.Parse(lblVatSum.Text))
                            cmd.Parameters.AddWithValue("@Comments", If(String.IsNullOrEmpty(txtRemarks.Text), DBNull.Value, txtRemarks.Text))
                            cmd.Parameters.AddWithValue("@U_PID", If(String.IsNullOrEmpty(txtUPID.Text), DBNull.Value, txtUPID.Text))
                            cmd.Parameters.AddWithValue("@UpdateBy", currentUserId)
                            cmd.ExecuteNonQuery()
                        End Using

                        ' 刪除舊明細
                        Using cmd As New SqlCommand("DELETE FROM jPRQ1 WHERE jID = @jID", conn, trans)
                            cmd.Parameters.AddWithValue("@jID", jID)
                            cmd.ExecuteNonQuery()
                        End Using
                    End If

                    ' 新增明細
                    For i As Integer = 0 To CurrentLines.Count - 1
                        Dim line = CurrentLines(i)
                        If String.IsNullOrEmpty(line.ItemCode) Then Continue For

                        Dim insertLineSql As String = "INSERT INTO jPRQ1 (jID, LineNum, ItemCode, Dscription, U_Linetext, Quantity, Price, LineTotal, GTotal, VatGroup, VatPrcnt, LineVat, WhsCode, ShipDate, CostingCode, CostingCode2, Currency, Rate, LineStatus, CreateDate, CreateBy) " &
                                                       "VALUES (@jID, @LineNum, @ItemCode, @Dscription, @U_Linetext, @Quantity, @Price, @LineTotal, @GTotal, @VatGroup, @VatPrcnt, @LineVat, @WhsCode, @ShipDate, @CostingCode, @CostingCode2, @Currency, @Rate, 'O', GETDATE(), @CreateBy)"

                        Using cmd As New SqlCommand(insertLineSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jID", jID)
                            cmd.Parameters.AddWithValue("@LineNum", i)
                            cmd.Parameters.AddWithValue("@ItemCode", line.ItemCode)
                            cmd.Parameters.AddWithValue("@Dscription", If(String.IsNullOrEmpty(line.Description), DBNull.Value, line.Description))
                            cmd.Parameters.AddWithValue("@U_Linetext", If(String.IsNullOrEmpty(line.LineText), DBNull.Value, line.LineText))
                            cmd.Parameters.AddWithValue("@Quantity", line.Quantity)
                            cmd.Parameters.AddWithValue("@Price", line.Price)
                            cmd.Parameters.AddWithValue("@LineTotal", line.LineTotal)
                            cmd.Parameters.AddWithValue("@GTotal", line.GTotal)
                            cmd.Parameters.AddWithValue("@VatGroup", If(String.IsNullOrEmpty(line.VatGroup), DBNull.Value, line.VatGroup))
                            cmd.Parameters.AddWithValue("@VatPrcnt", line.VatRate)
                            cmd.Parameters.AddWithValue("@LineVat", line.VatSum)
                            cmd.Parameters.AddWithValue("@WhsCode", If(String.IsNullOrEmpty(line.WhsCode), DBNull.Value, line.WhsCode))
                            cmd.Parameters.AddWithValue("@ShipDate", If(line.ShipDate.HasValue, line.ShipDate.Value, DBNull.Value))
                            cmd.Parameters.AddWithValue("@CostingCode", If(String.IsNullOrEmpty(line.CostingCode), DBNull.Value, line.CostingCode))
                            cmd.Parameters.AddWithValue("@CostingCode2", If(String.IsNullOrEmpty(line.CostingCode2), DBNull.Value, line.CostingCode2))
                            cmd.Parameters.AddWithValue("@Currency", ddlDocCurrency.SelectedValue)
                            cmd.Parameters.AddWithValue("@Rate", Decimal.Parse(txtDocRate.Text))
                            cmd.Parameters.AddWithValue("@CreateBy", currentUserId)
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    trans.Commit()
                    txtJID.Text = jID.ToString()
                    currentJID = jID

                Catch ex As Exception
                    trans.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub
#End Region

#Region "文件載入"
    Private Sub LoadDocument(jID As Integer)
        Using conn As New SqlConnection(connStr)
            conn.Open()

            ' 載入表頭
            Dim sql As String = "SELECT * FROM jOPRQ WHERE jID = @jID"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@jID", jID)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        txtJID.Text = jID.ToString()
                        lblDocNum.Text = "PR-" & jID.ToString("D6")
                        txtCardCode.Text = If(IsDBNull(dr("CardCode")), "", dr("CardCode").ToString())
                        txtCardName.Text = If(IsDBNull(dr("CardName")), "", dr("CardName").ToString())
                        txtReqName.Text = dr("ReqName").ToString()

                        If Not IsDBNull(dr("ReqDept")) AndAlso ddlReqDept.Items.FindByValue(dr("ReqDept").ToString()) IsNot Nothing Then
                            ddlReqDept.SelectedValue = dr("ReqDept").ToString()
                        End If

                        txtDocDate.Text = Convert.ToDateTime(dr("DocDate")).ToString("yyyy-MM-dd")

                        If Not IsDBNull(dr("ReqDate")) Then
                            txtReqDate.Text = Convert.ToDateTime(dr("ReqDate")).ToString("yyyy-MM-dd")
                        End If

                        If ddlDocCurrency.Items.FindByValue(dr("DocCurrency").ToString()) IsNot Nothing Then
                            ddlDocCurrency.SelectedValue = dr("DocCurrency").ToString()
                        End If
                        txtDocRate.Text = Convert.ToDecimal(dr("DocRate")).ToString("N6")

                        txtRemarks.Text = If(IsDBNull(dr("Comments")), "", dr("Comments").ToString())
                        txtUPID.Text = If(IsDBNull(dr("U_PID")), "", dr("U_PID").ToString())
                        txtOwner.Text = If(IsDBNull(dr("CreateBy")), "", dr("CreateBy").ToString())

                        ' 狀態
                        Dim approvalStatus As String = dr("ApprovalStatus").ToString()
                        txtApprovalStatus.Text = approvalStatus
                        Select Case approvalStatus
                            Case "Pending"
                                lblDocStatus.Text = "待審核"
                                lblDocStatus.CssClass = "badge status-W"
                                txtStatusDisplay.Text = "待審核"
                            Case "Approved"
                                lblDocStatus.Text = "已核准"
                                lblDocStatus.CssClass = "badge status-A"
                                txtStatusDisplay.Text = "已核准"
                            Case "Rejected"
                                lblDocStatus.Text = "已退回"
                                lblDocStatus.CssClass = "badge status-R"
                                txtStatusDisplay.Text = "已退回"
                        End Select

                        ' 審核區塊
                        If canApprove AndAlso approvalStatus = "Pending" Then
                            txtApprovalComments.ReadOnly = False
                            btnApprove.Enabled = True
                            btnReject.Enabled = True
                        Else
                            txtApprovalComments.ReadOnly = True
                            btnApprove.Enabled = False
                            btnReject.Enabled = False
                        End If

                        If Not IsDBNull(dr("ApprovalComments")) Then
                            txtApprovalComments.Text = dr("ApprovalComments").ToString()
                        End If

                        ' 按鈕狀態
                        btnDelete.Visible = (approvalStatus = "Pending" OrElse approvalStatus = "Rejected")
                        btnUpdate.Visible = (approvalStatus = "Pending")
                        btnSubmit.Visible = (approvalStatus = "Pending" OrElse approvalStatus = "Rejected")
                    End If
                End Using
            End Using

            ' 載入明細
            CurrentLines.Clear()
            sql = "SELECT * FROM jPRQ1 WHERE jID = @jID ORDER BY LineNum"
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@jID", jID)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        Dim qty As Decimal = Convert.ToDecimal(dr("Quantity"))
                        Dim gTotal As Decimal = Convert.ToDecimal(dr("GTotal"))
                        Dim line As New PRLine() With {
                            .LineNum = Convert.ToInt32(dr("LineNum")) + 1,
                            .ItemCode = dr("ItemCode").ToString(),
                            .Description = If(IsDBNull(dr("Dscription")), "", dr("Dscription").ToString()),
                            .LineText = If(IsDBNull(dr("U_Linetext")), "", dr("U_Linetext").ToString()),
                            .Quantity = qty,
                            .Price = Convert.ToDecimal(dr("Price")),
                            .LineTotal = Convert.ToDecimal(dr("LineTotal")),
                            .VatGroup = If(IsDBNull(dr("VatGroup")), "", dr("VatGroup").ToString()),
                            .VatRate = Convert.ToDecimal(dr("VatPrcnt")),
                            .VatSum = Convert.ToDecimal(dr("LineVat")),
                            .GTotal = gTotal,
                            .PriceAfVAT = If(qty > 0, gTotal / qty, 0),  ' 計算含稅單價
                            .WhsCode = If(IsDBNull(dr("WhsCode")), "", dr("WhsCode").ToString()),
                            .ShipDate = If(IsDBNull(dr("ShipDate")), Nothing, Convert.ToDateTime(dr("ShipDate"))),
                            .CostingCode = If(IsDBNull(dr("CostingCode")), "", dr("CostingCode").ToString()),
                            .CostingCode2 = If(IsDBNull(dr("CostingCode2")), "", dr("CostingCode2").ToString())
                        }
                        CurrentLines.Add(line)
                    End While
                End Using
            End Using

            BindGrid()
            BindAttachmentGrid()
        End Using
    End Sub
#End Region

#Region "審核"
    Protected Sub btnApprove_Click(sender As Object, e As EventArgs)
        If currentJID = 0 Then Return

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE jOPRQ SET ApprovalStatus = 'Approved', ApprovedBy = @ApprovedBy, ApprovedDate = GETDATE(), ApprovalComments = @Comments, UpdateDate = GETDATE(), UpdateBy = @UpdateBy WHERE jID = @jID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@jID", currentJID)
                    cmd.Parameters.AddWithValue("@ApprovedBy", currentUserId)
                    cmd.Parameters.AddWithValue("@Comments", txtApprovalComments.Text)
                    cmd.Parameters.AddWithValue("@UpdateBy", currentUserId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            lblMessage.Text = "請購單已核准！"
            lblMessage.ForeColor = Drawing.Color.Green
            Response.Redirect("PurchaseRequestForm.aspx?jID=" & currentJID)
        Catch ex As Exception
            ShowError("核准失敗: " & ex.Message)
        End Try
    End Sub

    Protected Sub btnReject_Click(sender As Object, e As EventArgs)
        If currentJID = 0 Then Return

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE jOPRQ SET ApprovalStatus = 'Rejected', ApprovedBy = @ApprovedBy, ApprovedDate = GETDATE(), ApprovalComments = @Comments, UpdateDate = GETDATE(), UpdateBy = @UpdateBy WHERE jID = @jID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@jID", currentJID)
                    cmd.Parameters.AddWithValue("@ApprovedBy", currentUserId)
                    cmd.Parameters.AddWithValue("@Comments", txtApprovalComments.Text)
                    cmd.Parameters.AddWithValue("@UpdateBy", currentUserId)
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            lblMessage.Text = "請購單已退回！"
            lblMessage.ForeColor = Drawing.Color.Orange
            Response.Redirect("PurchaseRequestForm.aspx?jID=" & currentJID)
        Catch ex As Exception
            ShowError("退回失敗: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "其他按鈕"
    Protected Sub btnDelete_Click(sender As Object, e As EventArgs)
        If currentJID = 0 Then Return

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using trans As SqlTransaction = conn.BeginTransaction()
                    Try
                        ' 刪除明細
                        Using cmd As New SqlCommand("DELETE FROM jPRQ1 WHERE jID = @jID", conn, trans)
                            cmd.Parameters.AddWithValue("@jID", currentJID)
                            cmd.ExecuteNonQuery()
                        End Using

                        ' 刪除表頭
                        Using cmd As New SqlCommand("DELETE FROM jOPRQ WHERE jID = @jID", conn, trans)
                            cmd.Parameters.AddWithValue("@jID", currentJID)
                            cmd.ExecuteNonQuery()
                        End Using

                        trans.Commit()
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

            Response.Redirect("DocumentSearch.aspx")
        Catch ex As Exception
            ShowError("刪除失敗: " & ex.Message)
        End Try
    End Sub

    Protected Sub btnCancel_Click(sender As Object, e As EventArgs)
        Response.Redirect("DocumentSearch.aspx")
    End Sub

    Protected Sub btnUpdate_Click(sender As Object, e As EventArgs)
        btnSubmit_Click(sender, e)
    End Sub

    Protected Sub btnExportPDF_Click(sender As Object, e As EventArgs)
        ' TODO: 實作 PDF 匯出
        lblMessage.Text = "PDF 匯出功能開發中..."
        lblMessage.ForeColor = Drawing.Color.Blue
    End Sub

    Protected Sub btnNewDocument_Click(sender As Object, e As EventArgs)
        Response.Redirect("PurchaseRequestForm.aspx")
    End Sub

    ''' <summary>
    ''' 需求日期變更事件 - 手動修改時檢查是否為假日，詢問使用者是否順延
    ''' </summary>
    Protected Sub txtReqDate_TextChanged(sender As Object, e As EventArgs)
        Try
            lblReqDateHint.Visible = False

            If String.IsNullOrEmpty(txtReqDate.Text) Then Return

            Dim reqDate As DateTime = DateTime.Parse(txtReqDate.Text)

            ' 檢查是否為假日
            If HolidayHelper.IsHoliday(reqDate) Then
                Dim holidayName As String = HolidayHelper.GetHolidayName(reqDate)
                Dim nextWorkday As DateTime = HolidayHelper.GetNextWorkingDay(reqDate)

                ' 儲存原始日期和順延日期，供對話框使用
                hfOriginalReqDate.Value = reqDate.ToString("yyyy-MM-dd")
                hfAdjustedReqDate.Value = nextWorkday.ToString("yyyy-MM-dd")

                ' 設定對話框顯示內容
                lblHolidayOriginalDate.Text = reqDate.ToString("yyyy/MM/dd")
                lblHolidayName.Text = holidayName
                lblHolidayNextWorkday.Text = nextWorkday.ToString("yyyy/MM/dd")

                ' 顯示詢問對話框
                mpeHoliday.Show()
            End If
        Catch ex As Exception
            ' 日期解析失敗不阻斷流程
        End Try
    End Sub

    ''' <summary>
    ''' 假日順延 - 維持原日期
    ''' </summary>
    Protected Sub btnHolidayKeep_Click(sender As Object, e As EventArgs)
        mpeHoliday.Hide()
        ' 保持 txtReqDate 的值不變（使用者選擇的假日）
        lblReqDateHint.Text = "(此日期為假日)"
        lblReqDateHint.Visible = True
    End Sub

    ''' <summary>
    ''' 假日順延 - 順延到下一個工作日
    ''' </summary>
    Protected Sub btnHolidayAdjust_Click(sender As Object, e As EventArgs)
        mpeHoliday.Hide()
        If Not String.IsNullOrEmpty(hfAdjustedReqDate.Value) AndAlso Not String.IsNullOrEmpty(hfOriginalReqDate.Value) Then
            Dim originalDate As DateTime = DateTime.Parse(hfOriginalReqDate.Value)
            Dim holidayName As String = HolidayHelper.GetHolidayName(originalDate)
            txtReqDate.Text = hfAdjustedReqDate.Value
            lblReqDateHint.Text = String.Format("(原 {0:MM/dd} 為{1}，已順延)", originalDate, holidayName)
            lblReqDateHint.Visible = True
        End If
    End Sub
#End Region

#Region "驗證"
    Protected Sub btnValidationBack_Click(sender As Object, e As EventArgs)
        mpeValidation.Hide()
    End Sub

    Protected Sub btnValidationConfirm_Click(sender As Object, e As EventArgs)
        ' 確認後儲存
        mpeValidation.Hide()
        Try
            SaveDocument()
            lblMessage.Text = "請購單儲存成功！單號: " & txtJID.Text
            lblMessage.ForeColor = Drawing.Color.Green
            Response.Redirect("PurchaseRequestForm.aspx?jID=" & txtJID.Text)
        Catch ex As Exception
            ShowError("儲存失敗: " & ex.Message)
        End Try
    End Sub

    Private Sub ShowValidationErrors(errors As List(Of String), warnings As List(Of String))
        blErrors.Items.Clear()
        blWarnings.Items.Clear()

        If errors.Count > 0 Then
            pnlErrors.Visible = True
            For Each err As String In errors
                blErrors.Items.Add(err)
            Next
        Else
            pnlErrors.Visible = False
        End If

        If warnings.Count > 0 Then
            pnlWarnings.Visible = True
            For Each warn As String In warnings
                blWarnings.Items.Add(warn)
            Next
            btnValidationConfirm.Visible = (errors.Count = 0)
        Else
            pnlWarnings.Visible = False
            btnValidationConfirm.Visible = False
        End If

        mpeValidation.Show()
    End Sub
#End Region

#Region "輔助方法"
    Private Sub ShowError(message As String)
        lblMessage.Text = message
        lblMessage.ForeColor = Drawing.Color.Red
    End Sub
#End Region

End Class
