Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.Configuration
Imports SAPbobsCOM

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
        ' 說明欄位追蹤 (用於 Dscription/U_Linetext 邏輯)
        Public Property OriginalDescription As String  ' 原始品名 (來自品號搜尋)
        Public Property DescriptionEdited As Boolean   ' 使用者是否編輯過說明
    End Class

    <Serializable()>
    Public Class AttachmentItem
        Public Property ID As Integer
        Public Property FileName As String
        Public Property FilePath As String
        Public Property UploadDate As DateTime
        Public Property UploadTime As String
        Public Property Uploader As String
        Public Property IsNew As Boolean = False  ' 標記是否為新上傳（尚未寫入資料庫）
    End Class
#End Region

#Region "變數宣告"
    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private ReadOnly sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString

    Private currentUserId As String = ""
    Private currentJID As Integer = 0
    Private isPuUser As Boolean = False  ' PU_App 權限（與費用申請單的 isApUser 對應）
    '
    ' SAP DI API Company 物件
    Public oCompany As New SAPbobsCOM.Company
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
                        isPuUser = (Convert.ToInt32(If(IsDBNull(dr("PU_App")), 0, dr("PU_App"))) = 1)
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub SetDefaultValues()
        lblDocNum.Text = "[新單據]"
        ' 嘗試帶入當前使用者作為請購人
        SetDefaultRequester(currentUserId)
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

        ' 審核區塊：一般使用者隱藏審核按鈕，只有 PU_App 權限者可見
        txtApprovalComments.ReadOnly = True
        btnApprove.Visible = False
        btnApprove.Enabled = False
        btnReject.Visible = False
        btnReject.Enabled = False

        btnDelete.Visible = False
    End Sub
#End Region

#Region "初始化資料"
    Private Sub InitializeDropDowns()
        LoadCurrencies()
        LoadDepartments()
        ' LoadReqName 已移除，改用搜尋彈窗
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

            ' 綁定原始說明和編輯標記 (用於 Dscription/U_Linetext 邏輯)
            Dim hfOriginalDescription As HiddenField = CType(e.Row.FindControl("hfOriginalDescription"), HiddenField)
            If hfOriginalDescription IsNot Nothing Then
                hfOriginalDescription.Value = If(String.IsNullOrEmpty(line.OriginalDescription), line.Description, line.OriginalDescription)
            End If

            Dim hfDescriptionEdited As HiddenField = CType(e.Row.FindControl("hfDescriptionEdited"), HiddenField)
            If hfDescriptionEdited IsNot Nothing Then
                hfDescriptionEdited.Value = line.DescriptionEdited.ToString().ToLower()
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

                ' 讀取原始說明和編輯標記
                Dim hfOriginalDescription As HiddenField = CType(row.FindControl("hfOriginalDescription"), HiddenField)
                If hfOriginalDescription IsNot Nothing AndAlso Not String.IsNullOrEmpty(hfOriginalDescription.Value) Then
                    line.OriginalDescription = hfOriginalDescription.Value
                End If

                Dim hfDescriptionEdited As HiddenField = CType(row.FindControl("hfDescriptionEdited"), HiddenField)
                If hfDescriptionEdited IsNot Nothing Then
                    line.DescriptionEdited = (hfDescriptionEdited.Value.ToLower() = "true")
                End If

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
                line.OriginalDescription = itemName  ' 記錄原始品名
                line.DescriptionEdited = False       ' 重置編輯標記
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

#Region "請購人搜尋"
    ''' <summary>
    ''' 設定預設請購人（從 OHEM 查詢員工資料）
    ''' </summary>
    Private Sub SetDefaultRequester(empCode As String)
        If String.IsNullOrEmpty(empCode) Then Return

        Try
            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = "SELECT Code, lastName + firstName AS EmpName FROM OHEM WHERE Code = @Code AND Active = 'Y'"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Code", empCode)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            txtReqCode.Text = dr("Code").ToString()
                            txtReqName.Text = If(IsDBNull(dr("EmpName")), "", dr("EmpName").ToString())
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 靜默處理
        End Try
    End Sub

    Protected Sub btnSearchReqCode_Click(sender As Object, e As EventArgs)
        PerformReqNameSearch("Code", txtReqCode.Text.Trim())
    End Sub

    Protected Sub btnSearchReqName_Click(sender As Object, e As EventArgs)
        PerformReqNameSearch("Name", txtReqName.Text.Trim())
    End Sub

    Private Sub PerformReqNameSearch(source As String, keyword As String)
        hfReqSearchSource.Value = source
        txtReqNameSearchKeyword.Text = keyword
        BindReqNameSearchGrid(keyword)
        mpeReqName.Show()
    End Sub

    Protected Sub btnDoSearchReqName_Click(sender As Object, e As EventArgs)
        gvReqNameSearch.PageIndex = 0  ' 重新搜尋時回到第一頁
        BindReqNameSearchGrid(txtReqNameSearchKeyword.Text.Trim())
        mpeReqName.Show()
    End Sub

    Protected Sub gvReqNameSearch_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
        gvReqNameSearch.PageIndex = e.NewPageIndex
        BindReqNameSearchGrid(txtReqNameSearchKeyword.Text.Trim())
        mpeReqName.Show()
    End Sub

    Private Sub BindReqNameSearchGrid(keyword As String)
        Try
            Dim sqlWhere As String = "WHERE Active = 'Y' "

            Dim searchSource As String = hfReqSearchSource.Value
            Dim isExact As Boolean = (rblReqSearchMode.SelectedValue = "Exact")

            If Not String.IsNullOrEmpty(keyword) Then
                keyword = keyword.Replace("*", "").Replace("%", "")

                If isExact Then
                    ' 開頭比對
                    If searchSource = "Code" Then
                        sqlWhere &= " AND Code LIKE @Kw"
                    Else
                        ' 按姓名搜尋：lastName + firstName 組合比對
                        sqlWhere &= " AND (lastName + firstName) LIKE @Kw"
                    End If
                    keyword = keyword & "%"
                Else
                    ' 模糊比對
                    If searchSource = "Code" Then
                        sqlWhere &= " AND Code LIKE @Kw"
                    Else
                        ' 按姓名搜尋：lastName + firstName 組合比對
                        sqlWhere &= " AND (lastName + firstName) LIKE @Kw"
                    End If
                    keyword = "%" & keyword & "%"
                End If
            End If

            Using conn As New SqlConnection(sapConnStr)
                conn.Open()
                Dim sql As String = $"SELECT TOP 100 Code, lastName + firstName AS EmpName FROM OHEM {sqlWhere} ORDER BY Code"
                Using cmd As New SqlCommand(sql, conn)
                    If Not String.IsNullOrEmpty(keyword) Then
                        cmd.Parameters.AddWithValue("@Kw", keyword)
                    End If

                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)
                        gvReqNameSearch.DataSource = dt
                        gvReqNameSearch.DataBind()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ShowError("搜尋請購人錯誤: " & ex.Message)
        End Try
    End Sub

    Protected Sub gvReqNameSearch_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "SelectReqName" Then
            Dim args() As String = e.CommandArgument.ToString().Split("|"c)
            If args.Length >= 2 Then
                txtReqCode.Text = args(0)  ' Code
                txtReqName.Text = args(1)  ' EmpName
            End If
            mpeReqName.Hide()
        End If
    End Sub
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
    ''' <summary>
    ''' 取得附件儲存資料夾路徑
    ''' </summary>
    Private Function GetAttachmentFolder() As String
        Return Server.MapPath("~/Uploads/PR/")
    End Function

    ''' <summary>
    ''' 取得附件相對路徑（用於資料庫儲存）
    ''' </summary>
    Private Function GetAttachmentRelativePath(fileName As String) As String
        Return "Uploads/PR/" & fileName
    End Function

    ''' <summary>
    ''' 取得附件絕對路徑
    ''' </summary>
    Private Function GetAttachmentAbsolutePath(relativePath As String) As String
        Return Server.MapPath("~/" & relativePath)
    End Function

    Protected Sub btnUpload_Click(sender As Object, e As EventArgs)
        If Not fileUpload.HasFile Then
            ShowError("請選擇要上傳的檔案")
            Return
        End If

        Try
            Dim uploadFolder As String = GetAttachmentFolder()
            If Not Directory.Exists(uploadFolder) Then
                Directory.CreateDirectory(uploadFolder)
            End If

            For Each uploadedFile As HttpPostedFile In fileUpload.PostedFiles
                Dim fileName As String = Path.GetFileName(uploadedFile.FileName)
                Dim uniqueName As String = DateTime.Now.ToString("yyyyMMddHHmmss") & "_" & fileName
                Dim absolutePath As String = Path.Combine(uploadFolder, uniqueName)
                Dim relativePath As String = GetAttachmentRelativePath(uniqueName)
                uploadedFile.SaveAs(absolutePath)

                Dim attachment As New AttachmentItem() With {
                    .ID = 0,  ' 新上傳的附件 ID 為 0，儲存後會更新
                    .FileName = fileName,
                    .FilePath = relativePath,  ' 使用相對路徑
                    .UploadDate = DateTime.Now,
                    .UploadTime = DateTime.Now.ToString("HH:mm:ss"),
                    .Uploader = currentUserId,
                    .IsNew = True  ' 標記為新上傳
                }
                CurrentAttachments.Add(attachment)
            Next

            BindAttachmentGrid()
            ShowSuccess("附件上傳成功")
        Catch ex As Exception
            ShowError("附件上傳失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 附件刪除 - 使用 Soft Delete 機制
    ''' </summary>
    Protected Sub gvAttachments_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        If e.CommandName = "DeleteFile" Then
            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            If index >= 0 AndAlso index < CurrentAttachments.Count Then
                Dim item = CurrentAttachments(index)

                Try
                    If item.IsNew Then
                        ' 新上傳的附件（尚未寫入資料庫），直接刪除檔案
                        Dim absolutePath As String = GetAttachmentAbsolutePath(item.FilePath)
                        If File.Exists(absolutePath) Then
                            File.Delete(absolutePath)
                        End If
                    Else
                        ' 已存在資料庫的附件，使用 Soft Delete
                        Using conn As New SqlConnection(connStr)
                            conn.Open()
                            Dim sql As String = "UPDATE jAttach SET IsDeleted=1, DeletedDate=GETDATE(), DeletedBy=@UserId WHERE ID=@ID"
                            Using cmd As New SqlCommand(sql, conn)
                                cmd.Parameters.AddWithValue("@ID", item.ID)
                                cmd.Parameters.AddWithValue("@UserId", currentUserId)
                                cmd.ExecuteNonQuery()
                            End Using
                        End Using
                    End If

                    CurrentAttachments.RemoveAt(index)
                    BindAttachmentGrid()
                Catch ex As Exception
                    ShowError("刪除附件失敗: " & ex.Message)
                End Try
            End If
        End If
    End Sub

    ''' <summary>
    ''' 保存新附件到 jAttach 表
    ''' </summary>
    Private Sub SaveAttachments(jID As Integer, conn As SqlConnection, trans As SqlTransaction)
        ' 取得目前最大的 LineNum
        Dim maxLineNum As Integer = 0
        Using cmd As New SqlCommand("SELECT ISNULL(MAX(LineNum), -1) FROM jAttach WHERE jID = @jID AND IsDeleted = 0", conn, trans)
            cmd.Parameters.AddWithValue("@jID", jID)
            Dim result = cmd.ExecuteScalar()
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                maxLineNum = Convert.ToInt32(result) + 1
            End If
        End Using

        ' 只保存新上傳的附件 (IsNew = True)
        For Each attachment As AttachmentItem In CurrentAttachments
            If attachment.IsNew Then
                Dim insertSql As String = "INSERT INTO jAttach (jID, LineNum, FilePath, FileName, Uploader, UploadDate, UploadTime, IsDeleted) " &
                                          "VALUES (@jID, @LineNum, @FilePath, @FileName, @Uploader, @UploadDate, @UploadTime, 0); " &
                                          "SELECT SCOPE_IDENTITY();"

                Using cmd As New SqlCommand(insertSql, conn, trans)
                    cmd.Parameters.AddWithValue("@jID", jID)
                    cmd.Parameters.AddWithValue("@LineNum", maxLineNum)
                    cmd.Parameters.AddWithValue("@FilePath", attachment.FilePath)
                    cmd.Parameters.AddWithValue("@FileName", attachment.FileName)
                    cmd.Parameters.AddWithValue("@Uploader", attachment.Uploader)
                    cmd.Parameters.Add("@UploadDate", SqlDbType.Date).Value = attachment.UploadDate.Date
                    cmd.Parameters.AddWithValue("@UploadTime", attachment.UploadTime)

                    ' 取得新插入的 ID
                    Dim newId = cmd.ExecuteScalar()
                    If newId IsNot Nothing AndAlso Not IsDBNull(newId) Then
                        attachment.ID = Convert.ToInt32(newId)
                        attachment.IsNew = False  ' 標記為已儲存
                    End If
                End Using

                maxLineNum += 1
            End If
        Next
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
            ShowSuccess("請購單儲存成功！單號: " & txtJID.Text)

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

                    ' 取得請購人資訊 (ReqCode 和 ReqName)
                    Dim reqCode As String = txtReqCode.Text.Trim()
                    Dim reqName As String = txtReqName.Text.Trim()

                    If jID = 0 Then
                        ' 新增 - 先從 OJID 取得全域唯一的 jID
                        Dim ojidSql As String = "INSERT INTO OJID (jUser, DocType) VALUES (@jUser, 'jOPRQ'); SELECT SCOPE_IDENTITY();"
                        Using cmd As New SqlCommand(ojidSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jUser", currentUserId)
                            jID = Convert.ToInt32(cmd.ExecuteScalar())
                        End Using

                        ' 使用從 OJID 取得的 jID 插入 jOPRQ（需開啟 IDENTITY_INSERT）
                        Dim insertSql As String = "SET IDENTITY_INSERT jOPRQ ON; " &
                                                   "INSERT INTO jOPRQ (jID, CardCode, CardName, ReqCode, ReqName, ReqDept, SlpCode, DocDate, ReqDate, DocCurrency, DocRate, DocTotal, VatSum, Comments, DocStatus, ApprovalStatus, U_PID, CreateDate, CreateBy) " &
                                                   "VALUES (@jID, @CardCode, @CardName, @ReqCode, @ReqName, @ReqDept, @SlpCode, @DocDate, @ReqDate, @DocCurrency, @DocRate, @DocTotal, @VatSum, @Comments, 'O', 'Pending', @U_PID, GETDATE(), @CreateBy); " &
                                                   "SET IDENTITY_INSERT jOPRQ OFF;"

                        Using cmd As New SqlCommand(insertSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jID", jID)
                            cmd.Parameters.AddWithValue("@CardCode", If(String.IsNullOrEmpty(txtCardCode.Text), DBNull.Value, txtCardCode.Text))
                            cmd.Parameters.AddWithValue("@CardName", If(String.IsNullOrEmpty(txtCardName.Text), DBNull.Value, txtCardName.Text))
                            cmd.Parameters.AddWithValue("@ReqCode", If(String.IsNullOrEmpty(reqCode), DBNull.Value, reqCode))
                            cmd.Parameters.AddWithValue("@ReqName", If(String.IsNullOrEmpty(reqName), DBNull.Value, reqName))
                            cmd.Parameters.AddWithValue("@ReqDept", If(String.IsNullOrEmpty(ddlReqDept.SelectedValue), DBNull.Value, ddlReqDept.SelectedValue))
                            cmd.Parameters.AddWithValue("@SlpCode", If(String.IsNullOrEmpty(ddlPurchaser.SelectedValue), DBNull.Value, ddlPurchaser.SelectedValue))
                            ' 使用 SqlDbType.Date 避免 SqlDateTime 溢位
                            cmd.Parameters.Add("@DocDate", SqlDbType.Date).Value = DateTime.Parse(txtDocDate.Text)
                            cmd.Parameters.Add("@ReqDate", SqlDbType.Date).Value = If(String.IsNullOrEmpty(txtReqDate.Text), DBNull.Value, DateTime.Parse(txtReqDate.Text))
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
                        Dim updateSql As String = "UPDATE jOPRQ SET CardCode=@CardCode, CardName=@CardName, ReqCode=@ReqCode, ReqName=@ReqName, ReqDept=@ReqDept, SlpCode=@SlpCode, DocDate=@DocDate, ReqDate=@ReqDate, " &
                                                   "DocCurrency=@DocCurrency, DocRate=@DocRate, DocTotal=@DocTotal, VatSum=@VatSum, Comments=@Comments, U_PID=@U_PID, UpdateDate=GETDATE(), UpdateBy=@UpdateBy WHERE jID=@jID"

                        Using cmd As New SqlCommand(updateSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jID", jID)
                            cmd.Parameters.AddWithValue("@CardCode", If(String.IsNullOrEmpty(txtCardCode.Text), DBNull.Value, txtCardCode.Text))
                            cmd.Parameters.AddWithValue("@CardName", If(String.IsNullOrEmpty(txtCardName.Text), DBNull.Value, txtCardName.Text))
                            cmd.Parameters.AddWithValue("@ReqCode", If(String.IsNullOrEmpty(reqCode), DBNull.Value, reqCode))
                            cmd.Parameters.AddWithValue("@ReqName", If(String.IsNullOrEmpty(reqName), DBNull.Value, reqName))
                            cmd.Parameters.AddWithValue("@ReqDept", If(String.IsNullOrEmpty(ddlReqDept.SelectedValue), DBNull.Value, ddlReqDept.SelectedValue))
                            cmd.Parameters.AddWithValue("@SlpCode", If(String.IsNullOrEmpty(ddlPurchaser.SelectedValue), DBNull.Value, ddlPurchaser.SelectedValue))
                            ' 使用 SqlDbType.Date 避免 SqlDateTime 溢位
                            cmd.Parameters.Add("@DocDate", SqlDbType.Date).Value = DateTime.Parse(txtDocDate.Text)
                            cmd.Parameters.Add("@ReqDate", SqlDbType.Date).Value = If(String.IsNullOrEmpty(txtReqDate.Text), DBNull.Value, DateTime.Parse(txtReqDate.Text))
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

                        ' Dscription/U_Linetext 邏輯:
                        ' - 如果使用者沒有編輯過說明欄位，Dscription = Description
                        ' - 如果使用者編輯過說明欄位，Dscription = OriginalDescription，U_Linetext = Description
                        Dim dscriptionValue As String = ""
                        Dim uLinetextValue As String = ""

                        If line.DescriptionEdited AndAlso Not String.IsNullOrEmpty(line.OriginalDescription) Then
                            ' 使用者有編輯：原始品名存 Dscription，編輯後的值存 U_Linetext
                            dscriptionValue = line.OriginalDescription
                            uLinetextValue = line.Description
                        Else
                            ' 使用者沒有編輯：直接存 Dscription
                            dscriptionValue = line.Description
                            uLinetextValue = ""
                        End If

                        Dim insertLineSql As String = "INSERT INTO jPRQ1 (jID, LineNum, ItemCode, Dscription, U_Linetext, Quantity, Price, PriceAfVAT, LineTotal, GTotal, VatGroup, VatPrcnt, LineVat, WhsCode, ShipDate, CostingCode, CostingCode2, Currency, Rate, LineStatus, CreateDate, CreateBy) " &
                                                       "VALUES (@jID, @LineNum, @ItemCode, @Dscription, @U_Linetext, @Quantity, @Price, @PriceAfVAT, @LineTotal, @GTotal, @VatGroup, @VatPrcnt, @LineVat, @WhsCode, @ShipDate, @CostingCode, @CostingCode2, @Currency, @Rate, 'O', GETDATE(), @CreateBy)"

                        Using cmd As New SqlCommand(insertLineSql, conn, trans)
                            cmd.Parameters.AddWithValue("@jID", jID)
                            cmd.Parameters.AddWithValue("@LineNum", i)
                            cmd.Parameters.AddWithValue("@ItemCode", line.ItemCode)
                            cmd.Parameters.AddWithValue("@Dscription", If(String.IsNullOrEmpty(dscriptionValue), DBNull.Value, dscriptionValue))
                            cmd.Parameters.AddWithValue("@U_Linetext", If(String.IsNullOrEmpty(uLinetextValue), DBNull.Value, uLinetextValue))
                            cmd.Parameters.AddWithValue("@Quantity", line.Quantity)
                            cmd.Parameters.AddWithValue("@Price", line.Price)
                            cmd.Parameters.AddWithValue("@PriceAfVAT", line.PriceAfVAT)
                            cmd.Parameters.AddWithValue("@LineTotal", line.LineTotal)
                            cmd.Parameters.AddWithValue("@GTotal", line.GTotal)
                            cmd.Parameters.AddWithValue("@VatGroup", If(String.IsNullOrEmpty(line.VatGroup), DBNull.Value, line.VatGroup))
                            cmd.Parameters.AddWithValue("@VatPrcnt", line.VatRate)
                            cmd.Parameters.AddWithValue("@LineVat", line.VatSum)
                            cmd.Parameters.AddWithValue("@WhsCode", If(String.IsNullOrEmpty(line.WhsCode), DBNull.Value, line.WhsCode))
                            ' 使用 SqlDbType.Date 避免 SqlDateTime 溢位，並檢查有效日期
                            If line.ShipDate.HasValue AndAlso line.ShipDate.Value > New DateTime(1753, 1, 1) Then
                                cmd.Parameters.Add("@ShipDate", SqlDbType.Date).Value = line.ShipDate.Value
                            Else
                                cmd.Parameters.Add("@ShipDate", SqlDbType.Date).Value = DBNull.Value
                            End If
                            cmd.Parameters.AddWithValue("@CostingCode", If(String.IsNullOrEmpty(line.CostingCode), DBNull.Value, line.CostingCode))
                            cmd.Parameters.AddWithValue("@CostingCode2", If(String.IsNullOrEmpty(line.CostingCode2), DBNull.Value, line.CostingCode2))
                            cmd.Parameters.AddWithValue("@Currency", ddlDocCurrency.SelectedValue)
                            cmd.Parameters.AddWithValue("@Rate", Decimal.Parse(txtDocRate.Text))
                            cmd.Parameters.AddWithValue("@CreateBy", currentUserId)
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    ' 保存新附件到 jAttach 表
                    SaveAttachments(jID, conn, trans)

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

                        ' 讀取請購人
                        txtReqCode.Text = If(IsDBNull(dr("ReqCode")), "", dr("ReqCode").ToString())
                        txtReqName.Text = If(IsDBNull(dr("ReqName")), "", dr("ReqName").ToString())

                        If Not IsDBNull(dr("ReqDept")) AndAlso ddlReqDept.Items.FindByValue(dr("ReqDept").ToString()) IsNot Nothing Then
                            ddlReqDept.SelectedValue = dr("ReqDept").ToString()
                        End If

                        ' 讀取採購人員
                        If Not IsDBNull(dr("SlpCode")) Then
                            Dim slpCode As String = dr("SlpCode").ToString()
                            If ddlPurchaser.Items.FindByValue(slpCode) IsNot Nothing Then
                                ddlPurchaser.SelectedValue = slpCode
                            End If
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

                        ' 審核區塊：對一般使用者隱藏審核按鈕，只有 PU_App 權限者可見
                        ' 1. 審核意見欄位：PU_App 權限者可編輯
                        txtApprovalComments.ReadOnly = Not isPuUser

                        ' 2. 審核按鈕：只有 PU_App 權限者可見，且僅在 Pending 狀態可操作
                        btnApprove.Visible = isPuUser
                        btnReject.Visible = isPuUser

                        If isPuUser AndAlso approvalStatus = "Pending" Then
                            btnApprove.Enabled = True
                            btnReject.Enabled = True
                        Else
                            btnApprove.Enabled = False
                            btnReject.Enabled = False
                        End If

                        If Not IsDBNull(dr("ApprovalComments")) Then
                            txtApprovalComments.Text = dr("ApprovalComments").ToString()
                        End If

                        ' SAP 單號顯示
                        If HasColumn(dr, "DocNum") AndAlso Not IsDBNull(dr("DocNum")) Then
                            txtSapDocNum.Text = dr("DocNum").ToString()
                        Else
                            txtSapDocNum.Text = ""
                        End If

                        ' SAP 過帳狀態
                        If HasColumn(dr, "B1PostStatus") AndAlso Not IsDBNull(dr("B1PostStatus")) Then
                            Dim postStatus As String = dr("B1PostStatus").ToString()
                            Select Case postStatus
                                Case "Y"
                                    lblSapPostStatus.Text = "(已過帳)"
                                    lblSapPostStatus.ForeColor = Drawing.Color.Green
                                Case "E"
                                    lblSapPostStatus.Text = "(過帳失敗)"
                                    lblSapPostStatus.ForeColor = Drawing.Color.Red
                                    If HasColumn(dr, "B1ErrMsg") AndAlso Not IsDBNull(dr("B1ErrMsg")) Then
                                        Dim errMsg As String = dr("B1ErrMsg").ToString()
                                        lblSapPostStatus.ToolTip = errMsg
                                        ShowError("SAP 過帳失敗: " & errMsg)
                                    End If
                                Case Else
                                    lblSapPostStatus.Text = ""
                            End Select
                        Else
                            lblSapPostStatus.Text = ""
                        End If

                        ' 按鈕狀態
                        btnDelete.Visible = (approvalStatus = "Pending" OrElse approvalStatus = "Rejected")
                        btnUpdate.Visible = (approvalStatus = "Pending")
                        btnSubmit.Visible = (approvalStatus = "Pending" OrElse approvalStatus = "Rejected")
                        btnExportPDF.Visible = True      ' 已儲存的單據可匯出 PDF
                        btnNewDocument.Visible = True    ' 已儲存的單據可新增新單據
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

                        ' 處理 Dscription/U_Linetext 邏輯
                        Dim dscription As String = If(IsDBNull(dr("Dscription")), "", dr("Dscription").ToString())
                        Dim uLinetext As String = ""
                        Dim descriptionEdited As Boolean = False

                        ' 嘗試讀取 U_Linetext (可能不存在)
                        Try
                            If dr.GetOrdinal("U_Linetext") >= 0 Then
                                uLinetext = If(IsDBNull(dr("U_Linetext")), "", dr("U_Linetext").ToString())
                            End If
                        Catch
                            ' 欄位不存在，忽略
                        End Try

                        ' 如果 U_Linetext 有值，表示使用者有編輯過
                        Dim displayDescription As String = dscription
                        If Not String.IsNullOrEmpty(uLinetext) Then
                            displayDescription = uLinetext  ' 顯示使用者編輯的值
                            descriptionEdited = True
                        End If

                        Dim line As New PRLine() With {
                            .LineNum = Convert.ToInt32(dr("LineNum")) + 1,
                            .ItemCode = dr("ItemCode").ToString(),
                            .Description = displayDescription,
                            .OriginalDescription = dscription,
                            .DescriptionEdited = descriptionEdited,
                            .Quantity = qty,
                            .Price = Convert.ToDecimal(dr("Price")),
                            .LineTotal = Convert.ToDecimal(dr("LineTotal")),
                            .VatGroup = If(IsDBNull(dr("VatGroup")), "", dr("VatGroup").ToString()),
                            .VatRate = Convert.ToDecimal(dr("VatPrcnt")),
                            .VatSum = Convert.ToDecimal(dr("LineVat")),
                            .GTotal = gTotal,
                            .PriceAfVAT = If(IsDBNull(dr("PriceAfVAT")), If(qty > 0, gTotal / qty, 0D), Convert.ToDecimal(dr("PriceAfVAT"))),
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

            ' 載入附件
            LoadAttachments(jID, conn)

            BindAttachmentGrid()
        End Using
    End Sub

    ''' <summary>
    ''' 從 jAttach 表載入附件
    ''' </summary>
    Private Sub LoadAttachments(jID As Integer, conn As SqlConnection)
        CurrentAttachments.Clear()

        Dim sql As String = "SELECT ID, LineNum, FilePath, FileName, Uploader, UploadDate, UploadTime " &
                            "FROM jAttach WHERE jID = @jID AND IsDeleted = 0 ORDER BY LineNum"

        Using cmd As New SqlCommand(sql, conn)
            cmd.Parameters.AddWithValue("@jID", jID)
            Using dr As SqlDataReader = cmd.ExecuteReader()
                While dr.Read()
                    Dim attachment As New AttachmentItem() With {
                        .ID = Convert.ToInt32(dr("ID")),
                        .FileName = If(IsDBNull(dr("FileName")), "", dr("FileName").ToString()),
                        .FilePath = If(IsDBNull(dr("FilePath")), "", dr("FilePath").ToString()),
                        .UploadDate = If(IsDBNull(dr("UploadDate")), DateTime.MinValue, Convert.ToDateTime(dr("UploadDate"))),
                        .UploadTime = If(IsDBNull(dr("UploadTime")), "", dr("UploadTime").ToString()),
                        .Uploader = If(IsDBNull(dr("Uploader")), "", dr("Uploader").ToString()),
                        .IsNew = False  ' 從資料庫載入的附件標記為非新增
                    }
                    CurrentAttachments.Add(attachment)
                End While
            End Using
        End Using
    End Sub
#End Region

#Region "審核"
    Protected Sub btnApprove_Click(sender As Object, e As EventArgs)
        If currentJID = 0 Then Return

        Try
            ' 1. 先嘗試寫入 SAP（不先更新狀態）
            Dim sapSuccess As Boolean = CreatePurchaseRequestInSAP(currentJID)

            If sapSuccess Then
                ' 2. SAP 成功後才更新狀態為 Approved
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

                ' 3. 成功才 Redirect
                Response.Redirect("PurchaseRequestForm.aspx?jID=" & currentJID)
            End If
            ' SAP 失敗時不 Redirect，讓錯誤訊息顯示在頁面上

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

            ShowWarning("請購單已退回！")
            Response.Redirect("PurchaseRequestForm.aspx?jID=" & currentJID)
        Catch ex As Exception
            ShowError("退回失敗: " & ex.Message)
        End Try
    End Sub
#End Region


#Region "SAP Integration"
    ''' <summary>
    ''' 初始化 SAP 連線
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
    ''' 關閉 SAP 連線
    ''' </summary>
    Private Sub CloseSAPConnection()
        If oCompany IsNot Nothing AndAlso oCompany.Connected Then
            oCompany.Disconnect()
        End If
    End Sub

    ''' <summary>
    ''' 建立 SAP 請購單 (Purchase Request)
    ''' </summary>
    ''' <returns>True = 成功, False = 失敗</returns>
    Private Function CreatePurchaseRequestInSAP(jID As Integer) As Boolean
        Dim oPR As SAPbobsCOM.Documents = Nothing
        Dim sapDocEntry As Integer = 0
        Dim errMsg As String = ""

        Try
            ' 0. 雙重檢查：確保尚未成功寫入 SAP
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Using cmd As New SqlCommand("SELECT B1PostStatus FROM jOPRQ WHERE jID=@jID", conn)
                    cmd.Parameters.AddWithValue("@jID", jID)
                    Dim status = cmd.ExecuteScalar()
                    If status IsNot Nothing AndAlso status.ToString() = "Y" Then
                        ShowWarning("此單據已成功寫入 SAP，跳過重複寫入")
                        Return True ' 已成功，視為成功
                    End If
                End Using
            End Using

            ' 1. 初始化 SAP 連線
            If String.IsNullOrEmpty(ConfigurationManager.AppSettings("SapServer")) OrElse
               String.IsNullOrEmpty(ConfigurationManager.AppSettings("SapCompanyDB")) Then
                Throw New Exception("Web.config 中缺少 SAP 連線設定 (SapServer/SapCompanyDB)")
            End If

            Dim connResult As Integer = InitSAPConnection()
            If connResult <> 0 Then
                Dim connErrCode As Integer
                Dim connErrMsg As String = ""
                oCompany.GetLastError(connErrCode, connErrMsg)
                Throw New Exception("SAP 連線失敗 [" & connErrCode & "]: " & connErrMsg)
            End If

            ' 2. 建立請購單物件
            oPR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest)

            ' 記錄文件幣別和匯率
            Dim docCurrency As String = "TWD"
            Dim docRate As Double = 1.0

            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 讀取表頭
                Dim sqlH As String = "SELECT * FROM jOPRQ WHERE jID=@jID"
                Using cmdH As New SqlCommand(sqlH, conn)
                    cmdH.Parameters.AddWithValue("@jID", jID)
                    Using drH As SqlDataReader = cmdH.ExecuteReader()
                        If Not drH.Read() Then
                            Throw New Exception("找不到請購單: jID=" & jID)
                        End If

                        ' 供應商代碼 (選填)
                        If Not IsDBNull(drH("CardCode")) AndAlso drH("CardCode").ToString() <> "" Then
                            oPR.CardCode = drH("CardCode").ToString()
                        End If

                        ' 文件日期
                        oPR.DocDate = Convert.ToDateTime(drH("DocDate"))

                        ' 需求日期
                        If Not IsDBNull(drH("ReqDate")) Then
                            oPR.DocDueDate = Convert.ToDateTime(drH("ReqDate"))
                        End If

                        ' 幣別
                        If Not IsDBNull(drH("DocCurrency")) Then
                            docCurrency = drH("DocCurrency").ToString()
                            oPR.DocCurrency = docCurrency
                        End If

                        ' 匯率
                        If Not IsDBNull(drH("DocRate")) Then
                            docRate = Convert.ToDouble(drH("DocRate"))
                            If docCurrency <> "TWD" AndAlso docCurrency <> "NTD" Then
                                oPR.DocRate = docRate
                            End If
                        End If

                        ' 備註
                        If Not IsDBNull(drH("Comments")) Then
                            oPR.Comments = drH("Comments").ToString()
                        End If

                        ' 請購人代碼 -> Requester (員工代碼)
                        If Not IsDBNull(drH("ReqCode")) Then
                            oPR.Requester = drH("ReqCode").ToString()
                        End If

                        ' 請購類型: 12 = 員工
                        oPR.ReqType = 12

                    End Using
                End Using

                ' 讀取明細
                Dim sqlL As String = "SELECT * FROM jPRQ1 WHERE jID=@jID ORDER BY LineNum"
                Using cmdL As New SqlCommand(sqlL, conn)
                    cmdL.Parameters.AddWithValue("@jID", jID)
                    Using drL As SqlDataReader = cmdL.ExecuteReader()
                        Dim lineIndex As Integer = 0
                        While drL.Read()
                            If lineIndex > 0 Then oPR.Lines.Add()

                            ' 料號 (必填)
                            oPR.Lines.ItemCode = drL("ItemCode").ToString()

                            ' 說明 - 若有自訂說明，寫入 U_LineText UDF (不設定 ItemDescription)
                            If Not IsDBNull(drL("U_Linetext")) AndAlso drL("U_Linetext").ToString() <> "" Then
                                oPR.Lines.UserFields.Fields.Item("U_LineText").Value = drL("U_Linetext").ToString()
                            End If

                            ' 數量
                            oPR.Lines.Quantity = Convert.ToDouble(drL("Quantity"))

                            ' 單價 (未稅)
                            If Not IsDBNull(drL("Price")) Then
                                oPR.Lines.UnitPrice = Convert.ToDouble(drL("Price"))
                            End If

                            ' 倉庫
                            If Not IsDBNull(drL("WhsCode")) AndAlso drL("WhsCode").ToString() <> "" Then
                                oPR.Lines.WarehouseCode = drL("WhsCode").ToString()
                            End If

                            ' 需求日期
                            If Not IsDBNull(drL("ShipDate")) Then
                                oPR.Lines.RequiredDate = Convert.ToDateTime(drL("ShipDate"))
                            End If

                            ' 稅碼 (VatGroup) - 轉換為 SAP 稅碼: 1=J1(應稅), 2=J0(零稅), 3=JX(免稅)
                            If Not IsDBNull(drL("VatGroup")) AndAlso drL("VatGroup").ToString() <> "" Then
                                Dim vatCode As String = drL("VatGroup").ToString()
                                Select Case vatCode
                                    Case "1" : oPR.Lines.VatGroup = "J1"  ' 應稅
                                    Case "2" : oPR.Lines.VatGroup = "J0"  ' 零稅
                                    Case "3" : oPR.Lines.VatGroup = "JX"  ' 免稅
                                    Case Else : oPR.Lines.VatGroup = vatCode  ' 若已是 SAP 稅碼則直接使用
                                End Select
                            End If

                            ' 成本中心
                            If Not IsDBNull(drL("CostingCode")) AndAlso drL("CostingCode").ToString() <> "" Then
                                oPR.Lines.CostingCode = drL("CostingCode").ToString()
                            End If
                            If Not IsDBNull(drL("CostingCode2")) AndAlso drL("CostingCode2").ToString() <> "" Then
                                oPR.Lines.CostingCode2 = drL("CostingCode2").ToString()
                            End If

                            ' 專案
                            If Not IsDBNull(drL("Project")) AndAlso drL("Project").ToString() <> "" Then
                                oPR.Lines.ProjectCode = drL("Project").ToString()
                            End If

                            lineIndex += 1
                        End While
                    End Using
                End Using
            End Using

            ' 3. 新增文件
            If oPR.Add() <> 0 Then
                Dim errCode As Integer
                oCompany.GetLastError(errCode, errMsg)
                Throw New Exception("SAP Error [" & errCode & "]: " & errMsg)
            Else
                ' 取得 SAP DocEntry
                sapDocEntry = Convert.ToInt32(oCompany.GetNewObjectKey())

                ' 從 SAP OPRQ 取得 DocNum
                Dim sapDocNum As Integer = GetSAPPRDocNum(sapDocEntry)

                ShowSuccess("已核准並產生 SAP 請購單 (DocEntry: " & sapDocEntry & ", DocNum: " & sapDocNum & ")")

                ' 更新 jOPRQ 的 SAP 單號和狀態
                UpdateSAPPostStatus(jID, sapDocEntry, sapDocNum, "Y", "")

                Return True ' 成功
            End If

        Catch ex As Exception
            errMsg = ex.Message
            ShowError("SAP 過帳失敗: " & errMsg)
            UpdateSAPPostStatus(jID, 0, 0, "E", errMsg)
            Return False ' 失敗
        Finally
            oPR = Nothing
            CloseSAPConnection()
        End Try

        Return False ' 預設失敗
    End Function

    ''' <summary>
    ''' 從 SAP OPRQ 取得 DocNum
    ''' </summary>
    Private Function GetSAPPRDocNum(sapDocEntry As Integer) As Integer
        Try
            Using connSap As New SqlConnection(sapConnStr)
                connSap.Open()
                Dim sql As String = "SELECT DocNum FROM OPRQ WHERE DocEntry = @DocEntry"
                Using cmd As New SqlCommand(sql, connSap)
                    cmd.Parameters.AddWithValue("@DocEntry", sapDocEntry)
                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return Convert.ToInt32(result)
                    End If
                End Using
            End Using
        Catch
        End Try
        Return 0
    End Function

    ''' <summary>
    ''' 更新 SAP 過帳狀態並回寫 DocEntry/DocNum
    ''' </summary>
    Private Sub UpdateSAPPostStatus(jID As Integer, sapDocEntry As Integer, sapDocNum As Integer, status As String, errMsg As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 更新 jOPRQ 表頭
                Dim sqlHeader As String = "UPDATE jOPRQ SET " &
                                          "B1PostStatus = @Status, " &
                                          "B1ErrMsg = @ErrMsg, " &
                                          "B1PostDate = GETDATE()"

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

                ' 更新 jPRQ1 明細
                If sapDocEntry > 0 OrElse sapDocNum > 0 Then
                    Dim sqlLines As String = "UPDATE jPRQ1 SET "
                    Dim setClauses As New List(Of String)

                    If sapDocEntry > 0 Then setClauses.Add("DocEntry = @SapDocEntry")
                    If sapDocNum > 0 Then setClauses.Add("DocNum = @SapDocNum")

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
            System.Diagnostics.Debug.WriteLine("UpdateSAPPostStatus Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 顯示警告訊息
    ''' </summary>
    Private Sub ShowWarning(msg As String)
        lblMessage.Text = msg
        lblMessage.ForeColor = Drawing.Color.Orange
    End Sub

    ''' <summary>
    ''' 顯示成功訊息
    ''' </summary>
    Private Sub ShowSuccess(msg As String)
        lblMessage.Text = msg
        lblMessage.ForeColor = Drawing.Color.Green
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

    ''' <summary>
    ''' 匯出 PDF - 使用 Crystal Report 將請購單匯出為 PDF 格式
    ''' </summary>
    Protected Sub btnExportPDF_Click(sender As Object, e As EventArgs)
        Dim jID As String = txtJID.Text.Trim()
        If String.IsNullOrEmpty(jID) Then
            ShowError("無法匯出：尚未儲存的單據或缺少平台單號")
            Return
        End If

        ' 檢查附件中是否有 PDF 檔案（已儲存的附件，非新上傳）
        Dim pdfAttachments = CurrentAttachments.Where(Function(a) a.FileName.ToLower().EndsWith(".pdf") AndAlso Not a.IsNew).ToList()

        If pdfAttachments.Count > 0 Then
            ' 有 PDF 附件，顯示合併選項彈窗
            rptPdfAttachments.DataSource = pdfAttachments
            rptPdfAttachments.DataBind()
            mpePdfMerge.Show()
        Else
            ' 無 PDF 附件，直接匯出
            ExportPdfDirect()
        End If
    End Sub

    ''' <summary>
    ''' 直接匯出 PDF（無合併）
    ''' </summary>
    Private Sub ExportPdfDirect()
        Dim script As String = String.Format("window.open('PurchaseRequestReport.ashx?jID={0}', '_blank');",
            HttpUtility.UrlEncode(txtJID.Text.Trim()))
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "OpenPDF", script, True)
    End Sub

    ''' <summary>
    ''' PDF 合併彈窗 - 取消按鈕
    ''' </summary>
    Protected Sub btnPdfMergeCancel_Click(sender As Object, e As EventArgs)
        mpePdfMerge.Hide()
    End Sub

    ''' <summary>
    ''' PDF 合併彈窗 - 確認匯出
    ''' </summary>
    Protected Sub btnPdfMergeConfirm_Click(sender As Object, e As EventArgs)
        ' 收集用戶輸入的順序
        Dim mergeList As New List(Of Tuple(Of Integer, Integer))  ' (順序, AttachID)

        For Each item As RepeaterItem In rptPdfAttachments.Items
            If item.ItemType = ListItemType.Item OrElse item.ItemType = ListItemType.AlternatingItem Then
                Dim txtOrder As TextBox = CType(item.FindControl("txtOrder"), TextBox)
                Dim hfAttachId As HiddenField = CType(item.FindControl("hfAttachId"), HiddenField)

                If txtOrder IsNot Nothing AndAlso hfAttachId IsNot Nothing Then
                    Dim order As Integer
                    If Integer.TryParse(txtOrder.Text.Trim(), order) Then
                        mergeList.Add(Tuple.Create(order, Integer.Parse(hfAttachId.Value)))
                    End If
                End If
            End If
        Next

        mpePdfMerge.Hide()

        If mergeList.Count = 0 Then
            ' 沒有輸入任何順序，直接匯出
            ExportPdfDirect()
        Else
            ' 有選擇合併，按順序排列附件並調用合併 Handler
            mergeList = mergeList.OrderBy(Function(t) t.Item1).ToList()
            Dim attachIds As String = String.Join(",", mergeList.Select(Function(t) t.Item2))

            Dim script As String = String.Format(
                "window.open('PurchaseRequestReport.ashx?jID={0}&mergeAttach={1}', '_blank');",
                HttpUtility.UrlEncode(txtJID.Text.Trim()),
                HttpUtility.UrlEncode(attachIds))
            ScriptManager.RegisterStartupScript(Me, Me.GetType(), "OpenPDF", script, True)
        End If
    End Sub

    Protected Sub btnNewDocument_Click(sender As Object, e As EventArgs)
        Response.Redirect("PurchaseRequestForm.aspx")
    End Sub

    ''' <summary>
    ''' 需求日期變更事件 - 檢查是否為假日並自動順延
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

                ' 自動調整到下一個工作日
                txtReqDate.Text = nextWorkday.ToString("yyyy-MM-dd")
                lblReqDateHint.Text = String.Format("(原 {0:MM/dd} 為{1}，已順延)", reqDate, holidayName)
                lblReqDateHint.Visible = True
            End If
        Catch ex As Exception
            ' 日期解析失敗不阻斷流程
        End Try
    End Sub

    ''' <summary>
    ''' 假日彈窗 - 維持原日期
    ''' </summary>
    Protected Sub btnHolidayKeep_Click(sender As Object, e As EventArgs)
        mpeHoliday.Hide()
    End Sub

    ''' <summary>
    ''' 假日彈窗 - 順延到下一個工作日
    ''' </summary>
    Protected Sub btnHolidayAdjust_Click(sender As Object, e As EventArgs)
        Try
            Dim nextWorkday As String = lblHolidayNextWorkday.Text
            If Not String.IsNullOrEmpty(nextWorkday) Then
                txtReqDate.Text = nextWorkday
            End If
        Catch ex As Exception
            ' 忽略錯誤
        End Try
        mpeHoliday.Hide()
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
            ShowSuccess("請購單儲存成功！單號: " & txtJID.Text)
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
#End Region

End Class
