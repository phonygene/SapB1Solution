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

    Protected Sub btnCreate_Click(sender As Object, e As EventArgs)
        SwitchMode("create")
        lblMode.InnerText = "模式：新增（Create）"
        btnUpdate.Enabled = False
        ClearForm()
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As EventArgs)
        SwitchMode("search")
        lblMode.InnerText = "模式：搜尋（Search）"
        btnUpdate.Enabled = False
    End Sub

    Protected Sub btnUpdate_Click(sender As Object, e As EventArgs)
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
    Protected Sub btnAddRow_Click(sender As Object, e As EventArgs)
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
    Protected Sub gvInvoiceDetail_RowCommand(sender As Object, e As GridViewCommandEventArgs)
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
    Protected Sub gvInvoiceDetail_RowDataBound(sender As Object, e As GridViewRowEventArgs)
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
    Protected Sub btnSave_Click(sender As Object, e As EventArgs)
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
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs)
        Response.Redirect(Request.RawUrl)
    End Sub

#End Region

#Region "業務邏輯 - 建立費用申請單"

    ''' <summary>
    ''' 建立費用申請單（AP Invoice + MDR）
    ''' </summary>
    Private Sub CreateExpenseClaim()
        Dim oCompany As Company = Nothing
        Dim apDocEntry As Integer = 0

        Try
            ' 取得明細資料
            Dim dt As DataTable = CType(ViewState("InvoiceDetail"), DataTable)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                ShowMessage("請至少輸入一筆發票明細")
                Return
            End If

            ' 1. 連接 SAP DI API
            oCompany = New Company()
            Dim ret As Integer = InitSAPConnection("192.168.1.31", "JTTST1", "manager", "sap123")
            If ret <> 0 Then
                Dim errMsg As String
                Dim errCode As Integer
                oCompany.GetLastError(errCode, errMsg)
                Throw New ApplicationException($"連接 SAP 失敗: {errCode} - {errMsg}")
            End If

            ' 2. 建立 AP Invoice（暫時使用第一列的供應商代碼）
            Dim firstRow As DataRow = dt.Rows(0)
            Dim cardCode As String = firstRow("U_LIFNR").ToString()

            If Not oCompany.InTransaction Then
                oCompany.StartTransaction()
            End If

            Try
                Dim oInvoice As Documents = CType(oCompany.GetBusinessObject(BoObjectTypes.oPurchaseInvoices), Documents)
                oInvoice.CardCode = cardCode
                oInvoice.DocDate = DateTime.Today

                ' 加入明細列（簡化版，實際應依業務需求調整）
                For Each row As DataRow In dt.Rows
                    Dim hwbas As Decimal = CDec(row("U_HWBAS"))
                    Dim hwste As Decimal = CDec(row("U_HWSTE"))

                    oInvoice.Lines.AccountCode = "_SYS00000000001" ' 暫時使用預設科目
                    oInvoice.Lines.LineTotal = hwbas
                    oInvoice.Lines.TaxCode = "100"
                    oInvoice.Lines.Add()
                Next

                ret = oInvoice.Add()
                If ret <> 0 Then
                    Dim errMsg As String
                    Dim errCode As Integer
                    oCompany.GetLastError(errCode, errMsg)
                    Throw New ApplicationException($"建立 AP Invoice 失敗: {errCode} - {errMsg}")
                End If

                ' 取得新建的 DocEntry
                apDocEntry = CInt(oCompany.GetNewObjectKey())

                oCompany.EndTransaction(BoWfTransOpt.wf_Commit)

            Catch ex As Exception
                If oCompany.InTransaction Then
                    oCompany.EndTransaction(BoWfTransOpt.wf_RollBack)
                End If
                Throw
            End Try

            ' 3. 寫入 MDR 中介表
            WriteMDRData(apDocEntry, dt)

            ' 4. 更新頁面顯示
            txtDocEntry.Text = apDocEntry.ToString()
            txtCreateDate.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            txtCreateBy.Text = GetCurrentUserId()

            ShowMessage($"費用申請單建立成功！AP Invoice DocEntry: {apDocEntry}")

        Catch ex As Exception
            Throw New ApplicationException("建立費用申請單失敗: " & ex.Message, ex)
        Finally
            If oCompany IsNot Nothing Then
                oCompany.Disconnect()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 寫入 MDR 中介表
    ''' </summary>
    ''' <param name="apDocEntry">AP Invoice DocEntry</param>
    ''' <param name="details">明細資料</param>
    Private Sub WriteMDRData(apDocEntry As Integer, details As DataTable)
        Try
            ' 產生新的 ID
            Dim newId As Integer = GetNextMDRId()

            ' 寫入表頭 MGUIAP_Import
            Dim sqlHead As String = "
                INSERT INTO MGUIAP_Import (ID, LineId, CreateDate, CreateBy, UpdateDate, UpdateBy,
                                           U_DocEntry, U_OBJTYPE, DocNum, DocTotal, VatSum)
                VALUES (@ID, @LineId, @CreateDate, @CreateBy, @UpdateDate, @UpdateBy,
                        @U_DocEntry, @U_OBJTYPE, @DocNum, @DocTotal, @VatSum)
            "

            Dim docTotal As Decimal = CalculateDocTotal(details)
            Dim vatSum As Decimal = CalculateVatSum(details)
            Dim userId As String = GetCurrentUserId()

            ExecuteLocalNonQuery(sqlHead,
                New SqlParameter("@ID", newId),
                New SqlParameter("@LineId", 1),
                New SqlParameter("@CreateDate", DateTime.Now),
                New SqlParameter("@CreateBy", userId),
                New SqlParameter("@UpdateDate", DateTime.Now),
                New SqlParameter("@UpdateBy", userId),
                New SqlParameter("@U_DocEntry", apDocEntry),
                New SqlParameter("@U_OBJTYPE", "18"),
                New SqlParameter("@DocNum", apDocEntry),
                New SqlParameter("@DocTotal", docTotal),
                New SqlParameter("@VatSum", vatSum))

            ' 寫入明細 MGUIAPDetail_Import
            Dim sqlDetail As String = "
                INSERT INTO MGUIAPDetail_Import
                    (ID, LineId, U_DocEntry, U_OBJTYPE, U_BELNR, U_BLDAT, U_VATDATE,
                     U_STCEG, U_XBLNR, U_ZFORM_CODE, U_HWBAS, U_HWSTE, U_TAX_TYPE,
                     U_CUS_TYPE, U_AM_TYPE, U_VATCODE, U_BUKRS, U_FA_DESC, U_FA_QTY,
                     U_FA_USE, U_GatherMark, U_ConsolidQty, U_MWSKZ, U_LIFNR)
                VALUES
                    (@ID, @LineId, @U_DocEntry, @U_OBJTYPE, @U_BELNR, @U_BLDAT, @U_VATDATE,
                     @U_STCEG, @U_XBLNR, @U_ZFORM_CODE, @U_HWBAS, @U_HWSTE, @U_TAX_TYPE,
                     @U_CUS_TYPE, @U_AM_TYPE, @U_VATCODE, @U_BUKRS, @U_FA_DESC, @U_FA_QTY,
                     @U_FA_USE, @U_GatherMark, @U_ConsolidQty, @U_MWSKZ, @U_LIFNR)
            "

            For Each row As DataRow In details.Rows
                ExecuteLocalNonQuery(sqlDetail,
                    New SqlParameter("@ID", newId),
                    New SqlParameter("@LineId", row("LineId")),
                    New SqlParameter("@U_DocEntry", apDocEntry),
                    New SqlParameter("@U_OBJTYPE", 18),
                    New SqlParameter("@U_BELNR", apDocEntry.ToString()),
                    New SqlParameter("@U_BLDAT", CDate(row("U_BLDAT"))),
                    New SqlParameter("@U_VATDATE", CDate(row("U_VATDATE"))),
                    New SqlParameter("@U_STCEG", row("U_STCEG").ToString()),
                    New SqlParameter("@U_XBLNR", row("U_XBLNR").ToString()),
                    New SqlParameter("@U_ZFORM_CODE", row("U_ZFORM_CODE").ToString()),
                    New SqlParameter("@U_HWBAS", CDec(row("U_HWBAS"))),
                    New SqlParameter("@U_HWSTE", CDec(row("U_HWSTE"))),
                    New SqlParameter("@U_TAX_TYPE", row("U_TAX_TYPE").ToString()),
                    New SqlParameter("@U_CUS_TYPE", row("U_CUS_TYPE").ToString()),
                    New SqlParameter("@U_AM_TYPE", row("U_AM_TYPE").ToString()),
                    New SqlParameter("@U_VATCODE", "100"),
                    New SqlParameter("@U_BUKRS", "100"),
                    New SqlParameter("@U_FA_DESC", DBNull.Value),
                    New SqlParameter("@U_FA_QTY", 0),
                    New SqlParameter("@U_FA_USE", DBNull.Value),
                    New SqlParameter("@U_GatherMark", "N"),
                    New SqlParameter("@U_ConsolidQty", 0),
                    New SqlParameter("@U_MWSKZ", row("U_ZFORM_CODE").ToString()),
                    New SqlParameter("@U_LIFNR", row("U_LIFNR").ToString()))
            Next

        Catch ex As Exception
            Throw New ApplicationException("寫入 MDR 資料失敗: " & ex.Message, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取得下一個 MDR ID
    ''' </summary>
    ''' <returns>新的 ID</returns>
    Private Function GetNextMDRId() As Integer
        Dim sql As String = "SELECT ISNULL(MAX(ID), 0) + 1 FROM MGUIAP_Import"
        Dim dt As DataTable = ExecuteLocalQuery(sql)

        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Return CInt(dt.Rows(0)(0))
        End If

        Return 1
    End Function

    ''' <summary>
    ''' 更新費用申請單（暫不實作）
    ''' </summary>
    Private Sub UpdateExpenseClaim()
        ShowMessage("更新功能尚未實作")
    End Sub

    ''' <summary>
    ''' 初始化 SAP 連線（簡化版）
    ''' </summary>
    Private Function InitSAPConnection(server As String, companyDb As String, userName As String, password As String) As Integer
        Dim oComp As New Company()
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
