Imports System.Data
Imports System.Data.SqlClient
Imports Newtonsoft.Json.Linq
Imports SAPbobsCOM

''' <summary>
''' 費用申請單表單
''' 繼承 B1TransactionBase，實作費用申請單的 Create 模式
''' </summary>
Public Class ExpenseClaimForm
    Inherits B1TransactionBase

#Region "控制項宣告"

    ' 表頭控制項
    Protected WithEvents lblMode As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents txtID As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtDocEntry As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtDocNum As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtDocTotal As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtVatSum As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCreateDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCreateBy As System.Web.UI.WebControls.TextBox

    ' 工具列按鈕
    Protected WithEvents btnCreate As System.Web.UI.WebControls.Button
    Protected WithEvents btnSearch As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate As System.Web.UI.WebControls.Button
    Protected WithEvents btnSave As System.Web.UI.WebControls.Button
    Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    Protected WithEvents btnAddRow As System.Web.UI.WebControls.Button

    ' 訊息與狀態
    Protected WithEvents lblMessage As System.Web.UI.WebControls.Label
    Protected WithEvents lblUser As System.Web.UI.WebControls.Label
    Protected WithEvents lblTime As System.Web.UI.WebControls.Label
    Protected WithEvents lblStatus As System.Web.UI.HtmlControls.HtmlGenericControl

    ' GridView
    Protected WithEvents gvInvoiceDetail As System.Web.UI.WebControls.GridView

#End Region

#Region "頁面生命週期"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Try
            ' 初始化連線
            InitConnections()

            ' 載入 JSON 配置檔
            Dim configPath As String = Server.MapPath("~/App_Data/ExpenseClaimTransactionConfig.json")
            LoadConfig(configPath)

            If Not IsPostBack Then
                ' 初始化頁面
                InitializePage()

                ' 預設為 Create 模式
                SwitchMode("create")
            End If

        Catch ex As Exception
            ShowMessage("頁面載入失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 初始化頁面
    ''' </summary>
    Private Sub InitializePage()
        ' 顯示當前使用者
        lblUser.Text = GetCurrentUserId()

        ' 顯示當前時間
        lblTime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        ' 初始化明細 GridView
        InitializeDetailGrid()
    End Sub

    ''' <summary>
    ''' 初始化明細 GridView
    ''' </summary>
    Private Sub InitializeDetailGrid()
        Dim dt As New DataTable()

        ' 建立欄位
        dt.Columns.Add("LineId", GetType(Integer))
        dt.Columns.Add("U_LIFNR", GetType(String))
        dt.Columns.Add("U_STCEG", GetType(String))
        dt.Columns.Add("U_XBLNR", GetType(String))
        dt.Columns.Add("U_ZFORM_CODE", GetType(String))
        dt.Columns.Add("U_BLDAT", GetType(DateTime))
        dt.Columns.Add("U_VATDATE", GetType(DateTime))
        dt.Columns.Add("U_HWBAS", GetType(Decimal))
        dt.Columns.Add("U_HWSTE", GetType(Decimal))
        dt.Columns.Add("U_TAX_TYPE", GetType(String))
        dt.Columns.Add("U_CUS_TYPE", GetType(String))
        dt.Columns.Add("U_AM_TYPE", GetType(String))

        ' 新增一列空白資料
        Dim row As DataRow = dt.NewRow()
        row("LineId") = 1
        row("U_ZFORM_CODE") = "21"
        row("U_BLDAT") = DateTime.Today
        row("U_VATDATE") = DateTime.Today
        row("U_HWBAS") = 0
        row("U_HWSTE") = 0
        row("U_TAX_TYPE") = "1"
        row("U_CUS_TYPE") = "0"
        row("U_AM_TYPE") = "1"
        dt.Rows.Add(row)

        ' 綁定資料
        gvInvoiceDetail.DataSource = dt
        gvInvoiceDetail.DataBind()

        ' 儲存到 ViewState
        ViewState("InvoiceDetail") = dt
    End Sub

#End Region

#Region "模式切換"

    Protected Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        SwitchMode("create")
        lblMode.InnerText = "模式：新增（Create）"
        btnUpdate.Enabled = False
        ClearForm()
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        SwitchMode("search")
        lblMode.InnerText = "模式：搜尋（Search）"
        btnUpdate.Enabled = False
    End Sub

    Protected Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        SwitchMode("update")
        lblMode.InnerText = "模式：更新（Update）"
    End Sub

    ''' <summary>
    ''' 清除表單
    ''' </summary>
    Private Sub ClearForm()
        txtID.Text = String.Empty
        txtDocEntry.Text = String.Empty
        txtDocNum.Text = String.Empty
        txtDocTotal.Text = "0.00"
        txtVatSum.Text = "0.00"
        txtCreateDate.Text = String.Empty
        txtCreateBy.Text = String.Empty

        InitializeDetailGrid()
    End Sub

#End Region

#Region "GridView 事件處理"

    ''' <summary>
    ''' 新增列
    ''' </summary>
    Protected Sub btnAddRow_Click(sender As Object, e As EventArgs) Handles btnAddRow.Click
        Try
            Dim dt As DataTable = CType(ViewState("InvoiceDetail"), DataTable)
            If dt Is Nothing Then
                InitializeDetailGrid()
                dt = CType(ViewState("InvoiceDetail"), DataTable)
            End If

            ' 取得新的列號
            Dim newLineId As Integer = dt.Rows.Count + 1

            ' 新增列
            Dim row As DataRow = dt.NewRow()
            row("LineId") = newLineId
            row("U_ZFORM_CODE") = "21"
            row("U_BLDAT") = DateTime.Today
            row("U_VATDATE") = DateTime.Today
            row("U_HWBAS") = 0
            row("U_HWSTE") = 0
            row("U_TAX_TYPE") = "1"
            row("U_CUS_TYPE") = "0"
            row("U_AM_TYPE") = "1"
            dt.Rows.Add(row)

            ' 重新綁定
            gvInvoiceDetail.DataSource = dt
            gvInvoiceDetail.DataBind()

            ViewState("InvoiceDetail") = dt

        Catch ex As Exception
            ShowMessage("新增列失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' GridView 列命令事件
    ''' </summary>
    Protected Sub gvInvoiceDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles gvInvoiceDetail.RowCommand
        If e.CommandName = "Delete" Then
            Try
                Dim rowIndex As Integer = Convert.ToInt32(e.CommandArgument)
                Dim dt As DataTable = CType(ViewState("InvoiceDetail"), DataTable)

                If dt IsNot Nothing AndAlso rowIndex < dt.Rows.Count Then
                    dt.Rows.RemoveAt(rowIndex)

                    ' 重新編號
                    For i As Integer = 0 To dt.Rows.Count - 1
                        dt.Rows(i)("LineId") = i + 1
                    Next

                    gvInvoiceDetail.DataSource = dt
                    gvInvoiceDetail.DataBind()

                    ViewState("InvoiceDetail") = dt

                    ' 重新計算總金額
                    CalculateTotals()
                End If

            Catch ex As Exception
                ShowMessage("刪除列失敗: " & ex.Message)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' GridView 資料綁定事件
    ''' </summary>
    Protected Sub gvInvoiceDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gvInvoiceDetail.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            ' 可在此加入特殊處理，例如欄位顏色設定
        End If
    End Sub

    ''' <summary>
    ''' 未稅金額變更事件（自動計算稅額）
    ''' </summary>
    Protected Sub txtHWBAS_TextChanged(sender As Object, e As EventArgs)
        Try
            Dim txtHWBAS As TextBox = CType(sender, TextBox)
            Dim row As GridViewRow = CType(txtHWBAS.NamingContainer, GridViewRow)

            Dim ddlTAX_TYPE As DropDownList = CType(row.FindControl("ddlTAX_TYPE"), DropDownList)
            Dim txtHWSTE As TextBox = CType(row.FindControl("txtHWSTE"), TextBox)

            Dim hwbas As Decimal = 0
            Decimal.TryParse(txtHWBAS.Text, hwbas)

            ' 計算稅額
            Dim taxType As String = ddlTAX_TYPE.SelectedValue
            Dim hwste As Decimal = CalculateTax(hwbas, taxType)
            txtHWSTE.Text = hwste.ToString("N2")

            ' 更新 ViewState 並重新計算總金額
            UpdateViewStateFromGrid()
            CalculateTotals()

        Catch ex As Exception
            ShowMessage("計算稅額失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 稅別變更事件（重新計算稅額）
    ''' </summary>
    Protected Sub ddlTAX_TYPE_SelectedIndexChanged(sender As Object, e As EventArgs)
        Try
            Dim ddlTAX_TYPE As DropDownList = CType(sender, DropDownList)
            Dim row As GridViewRow = CType(ddlTAX_TYPE.NamingContainer, GridViewRow)

            Dim txtHWBAS As TextBox = CType(row.FindControl("txtHWBAS"), TextBox)
            Dim txtHWSTE As TextBox = CType(row.FindControl("txtHWSTE"), TextBox)

            Dim hwbas As Decimal = 0
            Decimal.TryParse(txtHWBAS.Text, hwbas)

            ' 計算稅額
            Dim taxType As String = ddlTAX_TYPE.SelectedValue
            Dim hwste As Decimal = CalculateTax(hwbas, taxType)
            txtHWSTE.Text = hwste.ToString("N2")

            ' 更新 ViewState 並重新計算總金額
            UpdateViewStateFromGrid()
            CalculateTotals()

        Catch ex As Exception
            ShowMessage("計算稅額失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 從 GridView 更新 ViewState
    ''' </summary>
    Private Sub UpdateViewStateFromGrid()
        Dim dt As DataTable = CType(ViewState("InvoiceDetail"), DataTable)
        If dt Is Nothing Then Return

        For Each row As GridViewRow In gvInvoiceDetail.Rows
            If row.RowType = DataControlRowType.DataRow Then
                Dim lineId As Integer = CInt(gvInvoiceDetail.DataKeys(row.RowIndex).Value)
                Dim dataRow As DataRow = dt.AsEnumerable().FirstOrDefault(Function(r) CInt(r("LineId")) = lineId)

                If dataRow IsNot Nothing Then
                    ' 更新資料
                    Dim txtHWBAS As TextBox = CType(row.FindControl("txtHWBAS"), TextBox)
                    Dim txtHWSTE As TextBox = CType(row.FindControl("txtHWSTE"), TextBox)
                    Dim ddlTAX_TYPE As DropDownList = CType(row.FindControl("ddlTAX_TYPE"), DropDownList)

                    Dim hwbas As Decimal = 0
                    Decimal.TryParse(txtHWBAS.Text, hwbas)
                    dataRow("U_HWBAS") = hwbas

                    Dim hwste As Decimal = 0
                    Decimal.TryParse(txtHWSTE.Text, hwste)
                    dataRow("U_HWSTE") = hwste

                    dataRow("U_TAX_TYPE") = ddlTAX_TYPE.SelectedValue
                End If
            End If
        Next

        ViewState("InvoiceDetail") = dt
    End Sub

    ''' <summary>
    ''' 計算表頭總金額
    ''' </summary>
    Private Sub CalculateTotals()
        Dim dt As DataTable = CType(ViewState("InvoiceDetail"), DataTable)
        If dt Is Nothing Then Return

        Dim docTotal As Decimal = CalculateDocTotal(dt)
        Dim vatSum As Decimal = CalculateVatSum(dt)

        txtDocTotal.Text = docTotal.ToString("N2")
        txtVatSum.Text = vatSum.ToString("N2")
    End Sub

#End Region

#Region "儲存與取消"

    ''' <summary>
    ''' 儲存按鈕事件
    ''' </summary>
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            ' 驗證資料
            Dim errors As Dictionary(Of String, String) = ValidateMasterFields()
            If errors.Count > 0 Then
                Dim errorMsg As String = String.Join(vbCrLf, errors.Values)
                ShowMessage("驗證失敗:" & vbCrLf & errorMsg)
                Return
            End If

            ' 根據當前模式執行對應操作
            Select Case CurrentMode
                Case "create"
                    CreateExpenseClaim()
                Case "update"
                    UpdateExpenseClaim()
                Case Else
                    ShowMessage("當前模式不支援儲存操作")
            End Select

        Catch ex As Exception
            ShowMessage("儲存失敗: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 取消按鈕事件
    ''' </summary>
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Response.Redirect(Request.RawUrl)
    End Sub

#End Region

#Region "業務邏輯 - 建立費用申請單"

    ''' <summary>
    ''' 寫入 jtdb.jOPCH 採購單表頭，產生 jID (IDENTITY)
    ''' </summary>
    ''' <param name="details">發票明細資料表</param>
    ''' <returns>新產生的 jID</returns>
    Private Function SaveToJtdb_jOPCH(details As DataTable) As Integer
        Dim jID As Integer = 0

        Try
            ' 計算總金額
            Dim docTotal As Decimal = 0
            Dim vatSum As Decimal = 0
            For Each row As DataRow In details.Rows
                docTotal += CDec(row("U_HWBAS"))
                vatSum += CDec(row("U_HWSTE"))
            Next

            ' 使用第一列的供應商資訊
            Dim firstRow As DataRow = details.Rows(0)
            Dim cardCode As String = firstRow("U_LIFNR").ToString()
            Dim taxDate As DateTime = CDate(firstRow("U_VATDATE"))

            Dim sql As String = "
                INSERT INTO jOPCH (
                    CardCode, CardName, DocDate, TaxDate,
                    DocTotal, VatSum,
                    DocCurrency, DocRate,
                    DocStatus, Canceled, B1PostStatus,
                    ApprovalStatus,
                    CreateDate, CreateBy
                ) VALUES (
                    @CardCode, @CardName, @DocDate, @TaxDate,
                    @DocTotal, @VatSum,
                    'TWD', 1,
                    'O', 'N', 'N',
                    'Pending',
                    GETDATE(), @CreateBy
                );
                SELECT SCOPE_IDENTITY()
            "

            Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
            conn.Open()

            Try
                Dim cmd As New SqlCommand(sql, conn)
                With cmd.Parameters
                    .AddWithValue("@CardCode", cardCode)
                    .AddWithValue("@CardName", cardCode)
                    .AddWithValue("@DocDate", DateTime.Today)
                    .AddWithValue("@TaxDate", taxDate)
                    .AddWithValue("@DocTotal", docTotal)
                    .AddWithValue("@VatSum", vatSum)
                    .AddWithValue("@CreateBy", GetCurrentUserId())
                End With

                jID = Convert.ToInt32(cmd.ExecuteScalar())

            Finally
                conn.Close()
            End Try

            Return jID

        Catch ex As Exception
            Throw New Exception($"寫入 jOPCH 失敗：{ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' 寫入 jtdb.jPCH1 採購單明細
    ''' </summary>
    ''' <param name="jID">採購單 jID</param>
    ''' <param name="details">發票明細資料表</param>
    Private Sub SaveToJtdb_jPCH1(jID As Integer, details As DataTable)
        Try
            Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
            conn.Open()

            Try
                Dim lineNum As Integer = 0

                For Each row As DataRow In details.Rows
                    Dim sql As String = "
                        INSERT INTO jPCH1 (
                            jID, LineNum,
                            ItemCode, Dscription,
                            Quantity, Price, LineTotal,
                            TaxCode, VatSum,
                            Currency, Rate,
                            CreateDate, CreateBy
                        ) VALUES (
                            @jID, @LineNum,
                            @ItemCode, @Dscription,
                            @Quantity, @Price, @LineTotal,
                            @TaxCode, @VatSum,
                            'TWD', 1,
                            GETDATE(), @CreateBy
                        )
                    "

                    Dim cmd As New SqlCommand(sql, conn)
                    With cmd.Parameters
                        .AddWithValue("@jID", jID)
                        .AddWithValue("@LineNum", lineNum)
                        .AddWithValue("@ItemCode", "_SYS00000000001")
                        .AddWithValue("@Dscription", row("U_XBLNR").ToString())
                        .AddWithValue("@Quantity", 1)
                        .AddWithValue("@Price", CDec(row("U_HWBAS")))
                        .AddWithValue("@LineTotal", CDec(row("U_HWBAS")))
                        .AddWithValue("@TaxCode", "100")
                        .AddWithValue("@VatSum", CDec(row("U_HWSTE")))
                        .AddWithValue("@CreateBy", GetCurrentUserId())
                    End With

                    cmd.ExecuteNonQuery()
                    lineNum += 1
                Next

            Finally
                conn.Close()
            End Try

        Catch ex As Exception
            Throw New Exception($"寫入 jPCH1 失敗：{ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 寫入 jtdb.jMGUIAP 費用申請單表頭
    ''' </summary>
    ''' <param name="jID">採購單 jID（來自 jOPCH）</param>
    ''' <param name="details">發票明細資料表</param>
    Private Sub SaveToJtdb_jMGUIAP(jID As Integer, details As DataTable)
        Try
            ' 計算總金額
            Dim docTotal As Decimal = 0
            Dim vatSum As Decimal = 0
            For Each row As DataRow In details.Rows
                docTotal += CDec(row("U_HWBAS"))
                vatSum += CDec(row("U_HWSTE"))
            Next

            Dim sql As String = "
                INSERT INTO jMGUIAP (
                    jID, DocTotal, VatSum, U_OBJTYPE,
                    MDRPostStatus,
                    CreateDate, CreateBy
                ) VALUES (
                    @jID, @DocTotal, @VatSum, '18',
                    'N',
                    GETDATE(), @CreateBy
                )
            "

            Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
            conn.Open()

            Try
                Dim cmd As New SqlCommand(sql, conn)
                With cmd.Parameters
                    .AddWithValue("@jID", jID)
                    .AddWithValue("@DocTotal", docTotal)
                    .AddWithValue("@VatSum", vatSum)
                    .AddWithValue("@CreateBy", GetCurrentUserId())
                End With

                cmd.ExecuteNonQuery()

            Finally
                conn.Close()
            End Try

        Catch ex As Exception
            Throw New Exception($"寫入 jMGUIAP 失敗：{ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 寫入 jtdb.jMGUIAPDetail 費用申請單明細
    ''' </summary>
    ''' <param name="jID">採購單 jID（來自 jOPCH）</param>
    ''' <param name="details">發票明細資料表</param>
    Private Sub SaveToJtdb_jMGUIAPDetail(jID As Integer, details As DataTable)
        Try
            Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
            conn.Open()

            Try
                Dim lineNum As Integer = 0

                For Each row As DataRow In details.Rows
                    Dim sql As String = "
                        INSERT INTO jMGUIAPDetail (
                            jID, LineNum,
                            U_LIFNR, U_STCEG, U_XBLNR, U_ZFORM_CODE,
                            U_BLDAT, U_VATDATE,
                            U_HWBAS, U_HWSTE,
                            U_TAX_TYPE, U_CUS_TYPE, U_AM_TYPE,
                            U_VATCODE, U_BUKRS, U_MWSKZ,
                            U_FA_DESC, U_FA_QTY, U_FA_USE,
                            U_GatherMark, U_ConsolidQty,
                            CreateDate, CreateBy
                        ) VALUES (
                            @jID, @LineNum,
                            @U_LIFNR, @U_STCEG, @U_XBLNR, @U_ZFORM_CODE,
                            @U_BLDAT, @U_VATDATE,
                            @U_HWBAS, @U_HWSTE,
                            @U_TAX_TYPE, @U_CUS_TYPE, @U_AM_TYPE,
                            @U_VATCODE, @U_BUKRS, @U_MWSKZ,
                            @U_FA_DESC, @U_FA_QTY, @U_FA_USE,
                            @U_GatherMark, @U_ConsolidQty,
                            GETDATE(), @CreateBy
                        )
                    "

                    Dim cmd As New SqlCommand(sql, conn)
                    With cmd.Parameters
                        .AddWithValue("@jID", jID)
                        .AddWithValue("@LineNum", lineNum)
                        .AddWithValue("@U_LIFNR", row("U_LIFNR"))
                        .AddWithValue("@U_STCEG", row("U_STCEG"))
                        .AddWithValue("@U_XBLNR", row("U_XBLNR"))
                        .AddWithValue("@U_ZFORM_CODE", row("U_ZFORM_CODE"))
                        .AddWithValue("@U_BLDAT", CDate(row("U_BLDAT")))
                        .AddWithValue("@U_VATDATE", CDate(row("U_VATDATE")))
                        .AddWithValue("@U_HWBAS", CDec(row("U_HWBAS")))
                        .AddWithValue("@U_HWSTE", CDec(row("U_HWSTE")))
                        .AddWithValue("@U_TAX_TYPE", row("U_TAX_TYPE"))
                        .AddWithValue("@U_CUS_TYPE", row("U_CUS_TYPE"))
                        .AddWithValue("@U_AM_TYPE", row("U_AM_TYPE"))
                        .AddWithValue("@U_VATCODE", "100")
                        .AddWithValue("@U_BUKRS", "100")
                        .AddWithValue("@U_MWSKZ", row("U_ZFORM_CODE"))
                        .AddWithValue("@U_FA_DESC", DBNull.Value)
                        .AddWithValue("@U_FA_QTY", 0)
                        .AddWithValue("@U_FA_USE", DBNull.Value)
                        .AddWithValue("@U_GatherMark", "N")
                        .AddWithValue("@U_ConsolidQty", 0)
                        .AddWithValue("@CreateBy", GetCurrentUserId())
                    End With

                    cmd.ExecuteNonQuery()
                    lineNum += 1
                Next

            Finally
                conn.Close()
            End Try

        Catch ex As Exception
            Throw New Exception($"寫入 jMGUIAPDetail 失敗：{ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 更新 jOPCH 和 jMGUIAP 的 DocEntry（SAP AP Invoice 建立成功後）
    ''' </summary>
    ''' <param name="jID">採購單 jID</param>
    ''' <param name="docEntry">SAP AP Invoice DocEntry</param>
    Private Sub UpdateJtdb_DocEntry(jID As Integer, docEntry As Integer)
        Try
            Dim conn As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
            conn.Open()

            Try
                ' 更新 jOPCH
                Dim sql1 As String = "
                    UPDATE jOPCH
                    SET DocEntry = @DocEntry,
                        DocNum = @DocNum,
                        UpdateDate = GETDATE(),
                        UpdateBy = @UpdateBy
                    WHERE jID = @jID
                "

                Dim cmd1 As New SqlCommand(sql1, conn)
                With cmd1.Parameters
                    .AddWithValue("@DocEntry", docEntry)
                    .AddWithValue("@DocNum", docEntry)
                    .AddWithValue("@jID", jID)
                    .AddWithValue("@UpdateBy", GetCurrentUserId())
                End With
                cmd1.ExecuteNonQuery()

                ' 更新 jMGUIAP
                Dim sql2 As String = "
                    UPDATE jMGUIAP
                    SET DocEntry = @DocEntry,
                        DocNum = @DocNum,
                        UpdateDate = GETDATE(),
                        UpdateBy = @UpdateBy
                    WHERE jID = @jID
                "

                Dim cmd2 As New SqlCommand(sql2, conn)
                With cmd2.Parameters
                    .AddWithValue("@DocEntry", docEntry)
                    .AddWithValue("@DocNum", docEntry)
                    .AddWithValue("@jID", jID)
                    .AddWithValue("@UpdateBy", GetCurrentUserId())
                End With
                cmd2.ExecuteNonQuery()

            Finally
                conn.Close()
            End Try

        Catch ex As Exception
            Throw New Exception($"更新 DocEntry 失敗：{ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 建立費用申請單（完整流程：jtdb → SAP → MDR）
    ''' </summary>
    Private Sub CreateExpenseClaim()
        Dim oCompany As Company = Nothing
        Dim apDocEntry As Integer = 0
        Dim jID As Integer = 0

        Try
            ' 取得明細資料
            Dim dt As DataTable = CType(ViewState("InvoiceDetail"), DataTable)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                ShowMessage("請至少輸入一筆發票明細")
                Return
            End If

            ' ===== 第一步：寫入 jtdb 資料庫 =====
            jID = SaveToJtdb_jOPCH(dt)            ' 產生 jID
            SaveToJtdb_jPCH1(jID, dt)             ' 寫入採購單明細
            SaveToJtdb_jMGUIAP(jID, dt)           ' 寫入費用申請單表頭
            SaveToJtdb_jMGUIAPDetail(jID, dt)     ' 寫入費用申請單明細

            ' ===== 第二步：建立 SAP AP Invoice =====
            oCompany = New Company()
            Dim ret As Integer = InitSAPConnection("127.0.0.1", "JTSTD", "manager", "2408", oCompany)
            If ret <> 0 Then
                Dim errMsg As String = String.Empty
                Dim errCode As Integer = 0
                oCompany.GetLastError(errCode, errMsg)
                Throw New ApplicationException($"連接 SAP 失敗: {errCode} - {errMsg}")
            End If

            Dim firstRow As DataRow = dt.Rows(0)
            Dim cardCode As String = firstRow("U_LIFNR").ToString()

            If Not oCompany.InTransaction Then
                oCompany.StartTransaction()
            End If

            Try
                Dim oInvoice As Documents = CType(oCompany.GetBusinessObject(BoObjectTypes.oPurchaseInvoices), Documents)
                oInvoice.CardCode = cardCode
                oInvoice.DocDate = DateTime.Today

                For Each row As DataRow In dt.Rows
                    Dim hwbas As Decimal = CDec(row("U_HWBAS"))
                    oInvoice.Lines.AccountCode = "_SYS00000000001"
                    oInvoice.Lines.LineTotal = hwbas
                    oInvoice.Lines.TaxCode = "100"
                    oInvoice.Lines.Add()
                Next

                ret = oInvoice.Add()
                If ret <> 0 Then
                    Dim errMsg As String = String.Empty
                    Dim errCode As Integer = 0
                    oCompany.GetLastError(errCode, errMsg)
                    Throw New ApplicationException($"建立 AP Invoice 失敗: {errCode} - {errMsg}")
                End If

                apDocEntry = CInt(oCompany.GetNewObjectKey())
                oCompany.EndTransaction(BoWfTransOpt.wf_Commit)

            Catch ex As Exception
                If oCompany.InTransaction Then
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                End If
                Throw
            End Try

            ' ===== 第三步：更新 DocEntry =====
            UpdateJtdb_DocEntry(jID, apDocEntry)

            ' ===== 第四步：同步至 MDR 資料庫 =====
            Try
                Dim mdrHeaderID As Integer = WriteMDRData(jID)

                If mdrHeaderID > 0 Then
                    ShowMessage($"費用申請單建立成功！jID={jID}, AP DocEntry={apDocEntry}, MDR ID={mdrHeaderID}")
                Else
                    ShowMessage($"費用申請單建立成功，但 MDR 同步失敗。jID={jID}, AP DocEntry={apDocEntry}")
                End If
            Catch ex As Exception
                ' MDR 同步失敗不影響主流程
                ShowMessage($"費用申請單建立成功，但 MDR 同步發生錯誤：{ex.Message}。jID={jID}, AP DocEntry={apDocEntry}")
            End Try

            ' ===== 第五步：更新頁面顯示 =====
            txtID.Text = jID.ToString()
            txtDocEntry.Text = apDocEntry.ToString()
            txtCreateDate.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            txtCreateBy.Text = GetCurrentUserId()

        Catch ex As Exception
            Throw New ApplicationException("建立費用申請單失敗: " & ex.Message, ex)
        Finally
            If oCompany IsNot Nothing Then
                oCompany.Disconnect()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 將營業稅發票資料同步到 MDR 資料庫
    ''' </summary>
    ''' <param name="jID">費用申請單 jID</param>
    ''' <returns>MDR.MGUIAP_Import.ID，失敗回傳 0</returns>
    ''' <remarks>
    ''' 此方法會從 jtdb 讀取 jMGUIAP 和 jMGUIAPDetail 資料，
    ''' 寫入 MDR 資料庫的 MGUIAP_Import 和 MGUIAPDetail_Import 表，
    ''' 供後續 MDRImport.exe 程式處理營業稅過帳。
    ''' </remarks>
    Private Function WriteMDRData(jID As Integer) As Integer
        Dim mdrHeaderID As Integer = 0
        Dim jtdbConn As SqlConnection = Nothing
        Dim mdrConn As SqlConnection = Nothing
        Dim trans As SqlTransaction = Nothing

        Try
            ' ===== 第一步：從 jtdb 讀取資料 =====
            Dim headerData As DataRow = Nothing
            Dim detailData As DataTable = Nothing

            jtdbConn = New SqlConnection(ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
            jtdbConn.Open()

            ' 讀取表頭資料
            Dim sqlHeader As String = "
                SELECT
                    h.jID,
                    h.DocEntry,
                    h.DocNum,
                    d.U_LIFNR,
                    d.U_STCEG,
                    d.U_XBLNR,
                    d.U_BLDAT,
                    d.U_VATDATE,
                    d.U_ZFORM_CODE,
                    d.U_TAX_TYPE,
                    d.U_CUS_TYPE,
                    d.U_AM_TYPE,
                    d.U_BUKRS,
                    d.U_MWSKZ,
                    d.U_VATCODE,
                    SUM(d.U_HWBAS) AS U_HWBAS,
                    SUM(d.U_HWSTE) AS U_HWSTE,
                    SUM(d.U_HWBAS + d.U_HWSTE) AS TotalAmount
                FROM jMGUIAP h
                INNER JOIN jMGUIAPDetail d ON h.jID = d.jID
                WHERE h.jID = @jID
                GROUP BY
                    h.jID, h.DocEntry, h.DocNum,
                    d.U_LIFNR, d.U_STCEG, d.U_XBLNR,
                    d.U_BLDAT, d.U_VATDATE, d.U_ZFORM_CODE,
                    d.U_TAX_TYPE, d.U_CUS_TYPE, d.U_AM_TYPE,
                    d.U_BUKRS, d.U_MWSKZ, d.U_VATCODE
            "

            Dim cmdHeader As New SqlCommand(sqlHeader, jtdbConn)
            cmdHeader.Parameters.AddWithValue("@jID", jID)
            Dim adapterHeader As New SqlDataAdapter(cmdHeader)
            Dim dtHeader As New DataTable()
            adapterHeader.Fill(dtHeader)

            If dtHeader.Rows.Count = 0 Then
                Throw New Exception("找不到 jID=" & jID & " 的發票表頭資料")
            End If
            headerData = dtHeader.Rows(0)

            ' 讀取明細資料
            Dim sqlDetail As String = "
                SELECT
                    jID,
                    LineNum,
                    DocEntry,
                    U_LIFNR,
                    U_STCEG,
                    U_XBLNR,
                    U_ZFORM_CODE,
                    U_BLDAT,
                    U_VATDATE,
                    U_HWBAS,
                    U_HWSTE,
                    U_TAX_TYPE,
                    U_CUS_TYPE,
                    U_AM_TYPE,
                    U_VATCODE,
                    U_BUKRS,
                    U_MWSKZ,
                    U_BELNR,
                    U_FA_DESC,
                    U_FA_QTY,
                    U_FA_USE,
                    U_GatherMark,
                    U_ConsolidQty
                FROM jMGUIAPDetail
                WHERE jID = @jID
                ORDER BY LineNum
            "

            Dim cmdDetail As New SqlCommand(sqlDetail, jtdbConn)
            cmdDetail.Parameters.AddWithValue("@jID", jID)
            Dim adapterDetail As New SqlDataAdapter(cmdDetail)
            detailData = New DataTable()
            adapterDetail.Fill(detailData)

            If detailData.Rows.Count = 0 Then
                Throw New Exception("找不到 jID=" & jID & " 的發票明細資料")
            End If

            ' ===== 第二步：寫入 MDR 資料庫 =====
            mdrConn = New SqlConnection(ConfigurationManager.ConnectionStrings("MDRConnectionString").ConnectionString)
            mdrConn.Open()
            trans = mdrConn.BeginTransaction()

            ' 寫入表頭
            Dim sqlInsertHeader As String = "
                INSERT INTO MGUIAP_Import (
                    jID, DocEntry, DocNum,
                    U_LIFNR, U_STCEG,
                    U_XBLNR, U_BLDAT, U_VATDATE,
                    U_HWBAS, U_HWSTE, TotalAmount,
                    U_ZFORM_CODE, U_TAX_TYPE, U_CUS_TYPE, U_AM_TYPE,
                    U_BUKRS, U_MWSKZ, U_VATCODE,
                    PostStatus, CreateDate, CreateBy
                ) VALUES (
                    @jID, @DocEntry, @DocNum,
                    @U_LIFNR, @U_STCEG,
                    @U_XBLNR, @U_BLDAT, @U_VATDATE,
                    @U_HWBAS, @U_HWSTE, @TotalAmount,
                    @U_ZFORM_CODE, @U_TAX_TYPE, @U_CUS_TYPE, @U_AM_TYPE,
                    @U_BUKRS, @U_MWSKZ, @U_VATCODE,
                    'N', GETDATE(), @CreateBy
                );
                SELECT SCOPE_IDENTITY();
            "

            Dim cmdInsertHeader As New SqlCommand(sqlInsertHeader, mdrConn, trans)
            With cmdInsertHeader.Parameters
                .AddWithValue("@jID", jID)
                .AddWithValue("@DocEntry", If(IsDBNull(headerData("DocEntry")), DBNull.Value, headerData("DocEntry")))
                .AddWithValue("@DocNum", If(IsDBNull(headerData("DocNum")), DBNull.Value, headerData("DocNum")))
                .AddWithValue("@U_LIFNR", If(IsDBNull(headerData("U_LIFNR")), DBNull.Value, headerData("U_LIFNR")))
                .AddWithValue("@U_STCEG", If(IsDBNull(headerData("U_STCEG")), DBNull.Value, headerData("U_STCEG")))
                .AddWithValue("@U_XBLNR", If(IsDBNull(headerData("U_XBLNR")), DBNull.Value, headerData("U_XBLNR")))
                .AddWithValue("@U_BLDAT", If(IsDBNull(headerData("U_BLDAT")), DBNull.Value, headerData("U_BLDAT")))
                .AddWithValue("@U_VATDATE", If(IsDBNull(headerData("U_VATDATE")), DBNull.Value, headerData("U_VATDATE")))
                .AddWithValue("@U_HWBAS", headerData("U_HWBAS"))
                .AddWithValue("@U_HWSTE", headerData("U_HWSTE"))
                .AddWithValue("@TotalAmount", headerData("TotalAmount"))
                .AddWithValue("@U_ZFORM_CODE", If(IsDBNull(headerData("U_ZFORM_CODE")), DBNull.Value, headerData("U_ZFORM_CODE")))
                .AddWithValue("@U_TAX_TYPE", If(IsDBNull(headerData("U_TAX_TYPE")), DBNull.Value, headerData("U_TAX_TYPE")))
                .AddWithValue("@U_CUS_TYPE", If(IsDBNull(headerData("U_CUS_TYPE")), DBNull.Value, headerData("U_CUS_TYPE")))
                .AddWithValue("@U_AM_TYPE", If(IsDBNull(headerData("U_AM_TYPE")), DBNull.Value, headerData("U_AM_TYPE")))
                .AddWithValue("@U_BUKRS", If(IsDBNull(headerData("U_BUKRS")), DBNull.Value, headerData("U_BUKRS")))
                .AddWithValue("@U_MWSKZ", If(IsDBNull(headerData("U_MWSKZ")), DBNull.Value, headerData("U_MWSKZ")))
                .AddWithValue("@U_VATCODE", If(IsDBNull(headerData("U_VATCODE")), DBNull.Value, headerData("U_VATCODE")))
                .AddWithValue("@CreateBy", CommUtil.GetCurrentUserId())
            End With

            ' 執行並取得 IDENTITY
            mdrHeaderID = Convert.ToInt32(cmdInsertHeader.ExecuteScalar())

            ' 寫入明細
            For Each detailRow As DataRow In detailData.Rows
                Dim sqlInsertDetail As String = "
                    INSERT INTO MGUIAPDetail_Import (
                        HeaderID, LineNum, jID, DocEntry,
                        U_LIFNR, U_STCEG, U_XBLNR, U_ZFORM_CODE,
                        U_BLDAT, U_VATDATE,
                        U_HWBAS, U_HWSTE,
                        U_TAX_TYPE, U_CUS_TYPE, U_AM_TYPE, U_VATCODE,
                        U_BUKRS, U_MWSKZ, U_BELNR,
                        U_FA_DESC, U_FA_QTY, U_FA_USE,
                        U_GatherMark, U_ConsolidQty,
                        CreateDate, CreateBy
                    ) VALUES (
                        @HeaderID, @LineNum, @jID, @DocEntry,
                        @U_LIFNR, @U_STCEG, @U_XBLNR, @U_ZFORM_CODE,
                        @U_BLDAT, @U_VATDATE,
                        @U_HWBAS, @U_HWSTE,
                        @U_TAX_TYPE, @U_CUS_TYPE, @U_AM_TYPE, @U_VATCODE,
                        @U_BUKRS, @U_MWSKZ, @U_BELNR,
                        @U_FA_DESC, @U_FA_QTY, @U_FA_USE,
                        @U_GatherMark, @U_ConsolidQty,
                        GETDATE(), @CreateBy
                    )
                "

                Dim cmdInsertDetail As New SqlCommand(sqlInsertDetail, mdrConn, trans)
                With cmdInsertDetail.Parameters
                    .AddWithValue("@HeaderID", mdrHeaderID)
                    .AddWithValue("@LineNum", detailRow("LineNum"))
                    .AddWithValue("@jID", detailRow("jID"))
                    .AddWithValue("@DocEntry", If(IsDBNull(detailRow("DocEntry")), DBNull.Value, detailRow("DocEntry")))
                    .AddWithValue("@U_LIFNR", If(IsDBNull(detailRow("U_LIFNR")), DBNull.Value, detailRow("U_LIFNR")))
                    .AddWithValue("@U_STCEG", If(IsDBNull(detailRow("U_STCEG")), DBNull.Value, detailRow("U_STCEG")))
                    .AddWithValue("@U_XBLNR", If(IsDBNull(detailRow("U_XBLNR")), DBNull.Value, detailRow("U_XBLNR")))
                    .AddWithValue("@U_ZFORM_CODE", If(IsDBNull(detailRow("U_ZFORM_CODE")), DBNull.Value, detailRow("U_ZFORM_CODE")))
                    .AddWithValue("@U_BLDAT", If(IsDBNull(detailRow("U_BLDAT")), DBNull.Value, detailRow("U_BLDAT")))
                    .AddWithValue("@U_VATDATE", If(IsDBNull(detailRow("U_VATDATE")), DBNull.Value, detailRow("U_VATDATE")))
                    .AddWithValue("@U_HWBAS", detailRow("U_HWBAS"))
                    .AddWithValue("@U_HWSTE", detailRow("U_HWSTE"))
                    .AddWithValue("@U_TAX_TYPE", If(IsDBNull(detailRow("U_TAX_TYPE")), DBNull.Value, detailRow("U_TAX_TYPE")))
                    .AddWithValue("@U_CUS_TYPE", If(IsDBNull(detailRow("U_CUS_TYPE")), DBNull.Value, detailRow("U_CUS_TYPE")))
                    .AddWithValue("@U_AM_TYPE", If(IsDBNull(detailRow("U_AM_TYPE")), DBNull.Value, detailRow("U_AM_TYPE")))
                    .AddWithValue("@U_VATCODE", If(IsDBNull(detailRow("U_VATCODE")), DBNull.Value, detailRow("U_VATCODE")))
                    .AddWithValue("@U_BUKRS", If(IsDBNull(detailRow("U_BUKRS")), DBNull.Value, detailRow("U_BUKRS")))
                    .AddWithValue("@U_MWSKZ", If(IsDBNull(detailRow("U_MWSKZ")), DBNull.Value, detailRow("U_MWSKZ")))
                    .AddWithValue("@U_BELNR", If(IsDBNull(detailRow("U_BELNR")), DBNull.Value, detailRow("U_BELNR")))
                    .AddWithValue("@U_FA_DESC", If(IsDBNull(detailRow("U_FA_DESC")), DBNull.Value, detailRow("U_FA_DESC")))
                    .AddWithValue("@U_FA_QTY", detailRow("U_FA_QTY"))
                    .AddWithValue("@U_FA_USE", If(IsDBNull(detailRow("U_FA_USE")), DBNull.Value, detailRow("U_FA_USE")))
                    .AddWithValue("@U_GatherMark", If(IsDBNull(detailRow("U_GatherMark")), DBNull.Value, detailRow("U_GatherMark")))
                    .AddWithValue("@U_ConsolidQty", detailRow("U_ConsolidQty"))
                    .AddWithValue("@CreateBy", CommUtil.GetCurrentUserId())
                End With

                cmdInsertDetail.ExecuteNonQuery()
            Next

            ' ===== 第三步：更新 jtdb.jMGUIAP 的 MDRPostStatus =====
            Dim sqlUpdateStatus As String = "
                UPDATE jMGUIAP
                SET MDRPostStatus = 'P',
                    MDRPostDate = GETDATE()
                WHERE jID = @jID
            "

            Dim cmdUpdateStatus As New SqlCommand(sqlUpdateStatus, jtdbConn)
            cmdUpdateStatus.Parameters.AddWithValue("@jID", jID)
            cmdUpdateStatus.ExecuteNonQuery()

            ' 提交交易
            trans.Commit()

            Return mdrHeaderID

        Catch ex As Exception
            ' 回滾交易
            If trans IsNot Nothing Then
                Try
                    trans.Rollback()
                Catch
                End Try
            End If

            ' 更新錯誤狀態
            If jtdbConn IsNot Nothing AndAlso jtdbConn.State = ConnectionState.Open Then
                Try
                    Dim sqlUpdateError As String = "
                        UPDATE jMGUIAP
                        SET MDRPostStatus = 'E',
                            MDRErrMsg = @ErrorMsg
                        WHERE jID = @jID
                    "
                    Dim cmdUpdateError As New SqlCommand(sqlUpdateError, jtdbConn)
                    cmdUpdateError.Parameters.AddWithValue("@jID", jID)
                    cmdUpdateError.Parameters.AddWithValue("@ErrorMsg", ex.Message.Substring(0, Math.Min(500, ex.Message.Length)))
                    cmdUpdateError.ExecuteNonQuery()
                Catch
                End Try
            End If

            ' 記錄錯誤日誌
            System.Diagnostics.Debug.WriteLine("WriteMDRData Error: " & ex.Message)

            Throw New Exception("MDR 資料同步失敗：" & ex.Message, ex)

        Finally
            If jtdbConn IsNot Nothing AndAlso jtdbConn.State = ConnectionState.Open Then
                jtdbConn.Close()
            End If
            If mdrConn IsNot Nothing AndAlso mdrConn.State = ConnectionState.Open Then
                mdrConn.Close()
            End If
        End Try
    End Function

    ''' <summary>
    ''' 呼叫 MDRImport.exe 程式執行營業稅過帳
    ''' </summary>
    ''' <param name="jID">費用申請單 jID</param>
    ''' <remarks>
    ''' [保留功能] 此功能暫時保留，待確認 MDRImport.exe 路徑與參數後啟用
    ''' </remarks>
    Private Sub CallMDRProgram(jID As Integer)
        ' ===== [保留功能] MDRImport.exe 程式呼叫 =====
        ' 說明: 呼叫 MDRImport.exe 將 MDR 資料寫入 SAP B1 營業稅外掛表
        ' 狀態: 待確認執行檔路徑與參數
        ' TODO: 確認以下資訊
        '   1. MDRImport.exe 完整路徑
        '   2. 呼叫參數格式 (例如: MDRImport.exe /jID:123 或 MDRImport.exe 123)
        '   3. 是否需要等待執行完成
        '   4. 如何取得執行結果 (回傳值、日誌檔案等)
        '
        ' 範例程式碼:
        ' Try
        '     Dim mdrExePath As String = "C:\MDRImport\MDRImport.exe"
        '     Dim args As String = jID.ToString()
        '
        '     Dim psi As New ProcessStartInfo(mdrExePath, args)
        '     psi.WindowStyle = ProcessWindowStyle.Hidden
        '     psi.CreateNoWindow = True
        '
        '     Dim process As Process = Process.Start(psi)
        '     process.WaitForExit()
        '
        '     If process.ExitCode = 0 Then
        '         ' 執行成功
        '     Else
        '         ' 執行失敗
        '     End If
        ' Catch ex As Exception
        '     Throw New Exception("呼叫 MDRImport.exe 失敗：" & ex.Message, ex)
        ' End Try
    End Sub

    ''' <summary>
    ''' 更新費用申請單（暫不實作）
    ''' </summary>
    Private Sub UpdateExpenseClaim()
        ShowMessage("更新功能尚未實作")
    End Sub

    ''' <summary>
    ''' 初始化 SAP 連線
    ''' </summary>
    Private Function InitSAPConnection(server As String, companyDb As String, userName As String, password As String, ByRef oComp As Company) As Integer
        oComp.Server = server
        oComp.CompanyDB = companyDb
        oComp.UserName = userName
        oComp.Password = password
        oComp.UseTrusted = False
        oComp.DbUserName = "sa"
        oComp.DbPassword = "sap19690123"
        oComp.language = BoSuppLangs.ln_English
        oComp.DbServerType = BoDataServerTypes.dst_MSSQL2005

        Return oComp.Connect()
    End Function

#End Region

#Region "實作抽象方法"

    ''' <summary>
    ''' 套用欄位預設值（子類別實作）
    ''' </summary>
    Protected Overrides Sub ApplyFieldDefaults(defaults As JObject)
        ' 根據預設值設定欄位狀態
        ' 實際實作可依需求調整 UI 控制項屬性
    End Sub

    ''' <summary>
    ''' 套用狀態規則（子類別實作）
    ''' </summary>
    Protected Overrides Sub ApplyStateRules(rules As JArray)
        ' 根據規則設定欄位狀態
        ' 實際實作可依需求調整 UI 控制項屬性
    End Sub

    ''' <summary>
    ''' 取得表頭欄位值
    ''' </summary>
    Protected Overrides Function GetMasterFieldValue(fieldName As String) As String
        Select Case fieldName
            Case "ID"
                Return txtID.Text
            Case "U_DocEntry"
                Return txtDocEntry.Text
            Case "DocNum"
                Return txtDocNum.Text
            Case "DocTotal"
                Return txtDocTotal.Text
            Case "VatSum"
                Return txtVatSum.Text
            Case Else
                Return String.Empty
        End Select
    End Function

#End Region

End Class
