Imports System.Data
Imports System.Data.SqlClient

Partial Public Class DocumentSearch
    Inherits System.Web.UI.Page

    Private connStr As String = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private currentUserId As String = ""
    Private isApUser As Boolean = False
    Private Const MAX_RECORDS As Integer = 300

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Session("s_id") Is Nothing OrElse Session("s_id").ToString() = "" Then
            Response.Redirect("~/usermgm/login.aspx")
            Return
        End If

        currentUserId = Session("s_id").ToString()
        
        ' 檢查是否為審核者 (AP_App = 1)
        isApUser = CheckIsApUser()

        If Not IsPostBack Then
            InitializeControls()
        End If
    End Sub

    Private Function CheckIsApUser() As Boolean
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT ISNULL(AP_App, 0) FROM [User] WHERE ID = @ID"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ID", currentUserId)
                    Dim result = cmd.ExecuteScalar()
                    Return result IsNot Nothing AndAlso Convert.ToInt32(result) = 1
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    Private Sub InitializeControls()
        ' 設定使用者欄位權限
        If isApUser Then
            ' 審核者可以查詢所有人的單據
            txtUserCode.ReadOnly = False
            txtUserName.ReadOnly = False
            txtUserCode.CssClass = ""
            txtUserName.CssClass = ""
        Else
            ' 一般使用者只能查詢自己的單據
            txtUserCode.Text = currentUserId
            txtUserCode.ReadOnly = True
            txtUserCode.CssClass = "readonly-field"
            txtUserName.ReadOnly = True
            txtUserName.CssClass = "readonly-field"
            
            ' 嘗試取得使用者名稱
            Try
                Using conn As New SqlConnection(connStr)
                    conn.Open()
                    Dim sql As String = "SELECT Name FROM [User] WHERE ID = @ID"
                    Using cmd As New SqlCommand(sql, conn)
                        cmd.Parameters.AddWithValue("@ID", currentUserId)
                        Dim result = cmd.ExecuteScalar()
                        If result IsNot Nothing Then
                            txtUserName.Text = result.ToString()
                        End If
                    End Using
                End Using
            Catch
            End Try
        End If

        ' 設定分頁大小
        gvResults.PageSize = Convert.ToInt32(ddlPageSize.SelectedValue)
    End Sub
#End Region

#Region "Button Events"
    Protected Sub btnSearch_Click(sender As Object, e As EventArgs)
        gvResults.PageSize = Convert.ToInt32(ddlPageSize.SelectedValue)
        gvResults.PageIndex = 0
        ExecuteSearch()
    End Sub

    Protected Sub btnClear_Click(sender As Object, e As EventArgs)
        ' 清除所有篩選條件
        txtJID.Text = ""
        txtAPNumFrom.Text = ""
        txtAPNumTo.Text = ""
        txtPIDFrom.Text = ""
        txtPIDTo.Text = ""
        ddlDocStatus.SelectedIndex = 0
        ddlApprovalStatus.SelectedIndex = 0
        txtDocDateFrom.Text = ""
        txtDocDateTo.Text = ""
        txtDueDateFrom.Text = ""
        txtDueDateTo.Text = ""
        txtTaxDateFrom.Text = ""
        txtTaxDateTo.Text = ""
        txtCardCodeFrom.Text = ""
        txtCardCodeTo.Text = ""
        txtCardName.Text = ""
        txtComments.Text = ""
        rblCardNameMode.SelectedIndex = 0
        rblCommentsMode.SelectedIndex = 0
        ddlSortBy.SelectedIndex = 0
        rblSortOrder.SelectedIndex = 0
        
        If isApUser Then
            txtUserCode.Text = ""
            txtUserName.Text = ""
        End If
        
        lblMessage.Text = ""
        lblResultCount.Text = ""
        gvResults.DataSource = Nothing
        gvResults.DataBind()
    End Sub
#End Region

#Region "Search Logic"
    Private Sub ExecuteSearch()
        Try
            lblMessage.Text = ""

            Dim docType As String = ddlDocType.SelectedValue
            Dim tableName As String = GetTableName(docType)
            Dim whereClause As String = BuildWhereClause(docType)
            Dim orderClause As String = BuildOrderClause()

            Using conn As New SqlConnection(connStr)
                conn.Open()

                ' 1. 先計算符合條件的總筆數
                Dim countSql As String = String.Format("SELECT COUNT(*) FROM {0} WHERE 1=1 {1}", tableName, whereClause)
                Dim totalCount As Integer = 0

                Using cmd As New SqlCommand(countSql, conn)
                    AddParameters(cmd, docType)
                    totalCount = Convert.ToInt32(cmd.ExecuteScalar())
                End Using

                ' 2. 檢查是否超過 300 筆
                Dim showWarning As Boolean = False
                If totalCount > MAX_RECORDS Then
                    showWarning = True
                    lblMessage.Text = String.Format("⚠️ 查詢結果共 {0} 筆，超過 {1} 筆限制，僅顯示前 {1} 筆。請調整篩選條件。", totalCount, MAX_RECORDS)
                    lblMessage.ForeColor = System.Drawing.Color.DarkOrange
                End If

                ' 3. 執行查詢 (TOP 300) - 根據單據類型選擇欄位
                Dim selectSql As String
                If docType = "PurchaseRequest" Then
                    selectSql = String.Format(
                        "SELECT TOP {0} jID, CardCode, CardName, NULL AS InvNum, U_PID, ApprovalStatus, DocDate, ReqDate AS DocDueDate, " &
                        "Comments, CreateBy, CreateDate, " &
                        "(CASE WHEN ApprovalStatus = 'Approved' THEN 'Y' ELSE 'N' END) AS IsApproved " &
                        "FROM {1} WHERE 1=1 {2} {3}",
                        MAX_RECORDS, tableName, whereClause, orderClause)
                Else
                    selectSql = String.Format(
                        "SELECT TOP {0} jID, CardCode, CardName, InvNum, U_PID, ApprovalStatus, DocDate, DocDueDate, " &
                        "Comments, CreateBy, CreateDate, " &
                        "(CASE WHEN ApprovalStatus = 'A' THEN 'Y' ELSE 'N' END) AS IsApproved " &
                        "FROM {1} WHERE 1=1 {2} {3}",
                        MAX_RECORDS, tableName, whereClause, orderClause)
                End If

                Using cmd As New SqlCommand(selectSql, conn)
                    AddParameters(cmd, docType)

                    Dim da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)

                    gvResults.DataSource = dt
                    gvResults.DataBind()

                    If Not showWarning Then
                        lblResultCount.Text = String.Format("共 {0} 筆資料", totalCount)
                    Else
                        lblResultCount.Text = String.Format("顯示 {0} / {1} 筆", Math.Min(totalCount, MAX_RECORDS), totalCount)
                    End If
                End Using
            End Using

        Catch ex As Exception
            lblMessage.Text = "查詢失敗: " & ex.Message
            lblMessage.ForeColor = System.Drawing.Color.Red
        End Try
    End Sub

    Private Function GetTableName(docType As String) As String
        Select Case docType
            Case "PurchaseRequest"
                Return "jOPRQ"
            Case Else
                Return "jOPCH"
        End Select
    End Function

    Protected Sub ddlDocType_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' 清除查詢結果
        gvResults.DataSource = Nothing
        gvResults.DataBind()
        lblResultCount.Text = ""
        lblMessage.Text = ""
    End Sub

    Private Function BuildWhereClause() As String
        Dim sb As New System.Text.StringBuilder()
        
        ' 使用者權限控制 (非審核者只能看自己的)
        If Not isApUser Then
            sb.Append(" AND CreateBy = @CurrentUser")
        Else
            ' 審核者可指定使用者
            If Not String.IsNullOrEmpty(txtUserCode.Text.Trim()) Then
                sb.Append(" AND CreateBy = @UserCode")
            End If
        End If
        
        ' jID
        If Not String.IsNullOrEmpty(txtJID.Text.Trim()) Then
            Dim jidValue As Integer
            If Integer.TryParse(txtJID.Text.Trim(), jidValue) Then
                sb.Append(" AND jID = @JID")
            End If
        End If
        
        ' AP單號範圍
        If Not String.IsNullOrEmpty(txtAPNumFrom.Text.Trim()) Then
            sb.Append(" AND InvNum >= @APNumFrom")
        End If
        If Not String.IsNullOrEmpty(txtAPNumTo.Text.Trim()) Then
            sb.Append(" AND InvNum <= @APNumTo")
        End If
        
        ' 簽核系統PID範圍
        If Not String.IsNullOrEmpty(txtPIDFrom.Text.Trim()) Then
            sb.Append(" AND U_PID >= @PIDFrom")
        End If
        If Not String.IsNullOrEmpty(txtPIDTo.Text.Trim()) Then
            sb.Append(" AND U_PID <= @PIDTo")
        End If
        
        ' 文件狀態
        If Not String.IsNullOrEmpty(ddlDocStatus.SelectedValue) Then
            sb.Append(" AND ApprovalStatus = @DocStatus")
        End If
        
        ' 放行狀態
        If Not String.IsNullOrEmpty(ddlApprovalStatus.SelectedValue) Then
            If ddlApprovalStatus.SelectedValue = "Y" Then
                sb.Append(" AND ApprovalStatus = 'A'")
            Else
                sb.Append(" AND ApprovalStatus <> 'A'")
            End If
        End If
        
        ' 文件日期範圍
        If Not String.IsNullOrEmpty(txtDocDateFrom.Text) Then
            sb.Append(" AND DocDate >= @DocDateFrom")
        End If
        If Not String.IsNullOrEmpty(txtDocDateTo.Text) Then
            sb.Append(" AND DocDate <= @DocDateTo")
        End If
        
        ' 到期日範圍
        If Not String.IsNullOrEmpty(txtDueDateFrom.Text) Then
            sb.Append(" AND DocDueDate >= @DueDateFrom")
        End If
        If Not String.IsNullOrEmpty(txtDueDateTo.Text) Then
            sb.Append(" AND DocDueDate <= @DueDateTo")
        End If
        
        ' 過帳日期範圍
        If Not String.IsNullOrEmpty(txtTaxDateFrom.Text) Then
            sb.Append(" AND TaxDate >= @TaxDateFrom")
        End If
        If Not String.IsNullOrEmpty(txtTaxDateTo.Text) Then
            sb.Append(" AND TaxDate <= @TaxDateTo")
        End If
        
        ' 供應商代碼範圍
        If Not String.IsNullOrEmpty(txtCardCodeFrom.Text.Trim()) Then
            sb.Append(" AND CardCode >= @CardCodeFrom")
        End If
        If Not String.IsNullOrEmpty(txtCardCodeTo.Text.Trim()) Then
            sb.Append(" AND CardCode <= @CardCodeTo")
        End If
        
        ' 供應商名稱
        If Not String.IsNullOrEmpty(txtCardName.Text.Trim()) Then
            If rblCardNameMode.SelectedValue = "StartsWith" Then
                sb.Append(" AND CardName LIKE @CardName + '%'")
            Else
                sb.Append(" AND CardName LIKE '%' + @CardName + '%'")
            End If
        End If
        
        ' 備註
        If Not String.IsNullOrEmpty(txtComments.Text.Trim()) Then
            If rblCommentsMode.SelectedValue = "StartsWith" Then
                sb.Append(" AND Comments LIKE @Comments + '%'")
            Else
                sb.Append(" AND Comments LIKE '%' + @Comments + '%'")
            End If
        End If
        
        Return sb.ToString()
    End Function

    Private Function BuildOrderClause() As String
        Dim sortColumn As String = ddlSortBy.SelectedValue
        Dim sortOrder As String = rblSortOrder.SelectedValue
        
        ' 驗證排序欄位 (防止 SQL Injection)
        Dim validColumns As String() = {"jID", "DocDate", "DocDueDate", "CardCode", "CardName", "CreateDate"}
        If Not validColumns.Contains(sortColumn) Then
            sortColumn = "jID"
        End If
        
        If sortOrder <> "ASC" AndAlso sortOrder <> "DESC" Then
            sortOrder = "DESC"
        End If
        
        Return String.Format(" ORDER BY {0} {1}", sortColumn, sortOrder)
    End Function

    Private Sub AddParameters(cmd As SqlCommand, Optional docType As String = "")
        cmd.Parameters.AddWithValue("@CurrentUser", currentUserId)
        
        If Not String.IsNullOrEmpty(txtUserCode.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@UserCode", txtUserCode.Text.Trim())
        End If
        
        If Not String.IsNullOrEmpty(txtJID.Text.Trim()) Then
            Dim jidValue As Integer
            If Integer.TryParse(txtJID.Text.Trim(), jidValue) Then
                cmd.Parameters.AddWithValue("@JID", jidValue)
            End If
        End If
        
        If Not String.IsNullOrEmpty(txtAPNumFrom.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@APNumFrom", txtAPNumFrom.Text.Trim())
        End If
        If Not String.IsNullOrEmpty(txtAPNumTo.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@APNumTo", txtAPNumTo.Text.Trim())
        End If
        
        If Not String.IsNullOrEmpty(txtPIDFrom.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@PIDFrom", txtPIDFrom.Text.Trim())
        End If
        If Not String.IsNullOrEmpty(txtPIDTo.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@PIDTo", txtPIDTo.Text.Trim())
        End If
        
        If Not String.IsNullOrEmpty(ddlDocStatus.SelectedValue) Then
            cmd.Parameters.AddWithValue("@DocStatus", ddlDocStatus.SelectedValue)
        End If
        
        If Not String.IsNullOrEmpty(txtDocDateFrom.Text) Then
            cmd.Parameters.AddWithValue("@DocDateFrom", DateTime.Parse(txtDocDateFrom.Text))
        End If
        If Not String.IsNullOrEmpty(txtDocDateTo.Text) Then
            cmd.Parameters.AddWithValue("@DocDateTo", DateTime.Parse(txtDocDateTo.Text))
        End If
        
        If Not String.IsNullOrEmpty(txtDueDateFrom.Text) Then
            cmd.Parameters.AddWithValue("@DueDateFrom", DateTime.Parse(txtDueDateFrom.Text))
        End If
        If Not String.IsNullOrEmpty(txtDueDateTo.Text) Then
            cmd.Parameters.AddWithValue("@DueDateTo", DateTime.Parse(txtDueDateTo.Text))
        End If
        
        If Not String.IsNullOrEmpty(txtTaxDateFrom.Text) Then
            cmd.Parameters.AddWithValue("@TaxDateFrom", DateTime.Parse(txtTaxDateFrom.Text))
        End If
        If Not String.IsNullOrEmpty(txtTaxDateTo.Text) Then
            cmd.Parameters.AddWithValue("@TaxDateTo", DateTime.Parse(txtTaxDateTo.Text))
        End If
        
        If Not String.IsNullOrEmpty(txtCardCodeFrom.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@CardCodeFrom", txtCardCodeFrom.Text.Trim())
        End If
        If Not String.IsNullOrEmpty(txtCardCodeTo.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@CardCodeTo", txtCardCodeTo.Text.Trim())
        End If
        
        If Not String.IsNullOrEmpty(txtCardName.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@CardName", txtCardName.Text.Trim())
        End If
        
        If Not String.IsNullOrEmpty(txtComments.Text.Trim()) Then
            cmd.Parameters.AddWithValue("@Comments", txtComments.Text.Trim())
        End If
    End Sub
#End Region

#Region "GridView Events"
    Protected Sub gvResults_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
        gvResults.PageIndex = e.NewPageIndex
        ExecuteSearch()
    End Sub

    Protected Sub gvResults_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim createBy As String = DataBinder.Eval(e.Row.DataItem, "CreateBy").ToString()
            Dim lbtnCopy As LinkButton = CType(e.Row.FindControl("lbtnCopy"), LinkButton)
            If lbtnCopy IsNot Nothing Then
                Dim ownerId As String = If(createBy, "").Trim()
                Dim canCopy As Boolean = isApUser OrElse String.Equals(ownerId, currentUserId, StringComparison.OrdinalIgnoreCase)
                lbtnCopy.Visible = True
                lbtnCopy.Enabled = canCopy
                lbtnCopy.ToolTip = If(canCopy, "", "僅可複製自己的單據")
            End If
        End If
    End Sub

    Protected Sub gvResults_RowCommand(sender As Object, e As GridViewCommandEventArgs)
        ' 目前複製改由 btnCopyConfirm_Click 處理
    End Sub

    Protected Sub btnCopyConfirm_Click(sender As Object, e As EventArgs)
        Dim rowIndex As Integer
        If Not Integer.TryParse(hfCopyRowIndex.Value, rowIndex) Then Return

        If rowIndex >= gvResults.Rows.Count Then
            rowIndex -= gvResults.PageIndex * gvResults.PageSize
        End If

        If rowIndex < 0 OrElse rowIndex >= gvResults.Rows.Count Then Return

        Dim jID As Integer = Convert.ToInt32(gvResults.DataKeys(rowIndex)("jID"))
        Dim createBy As String = gvResults.DataKeys(rowIndex)("CreateBy").ToString()

        If Not isApUser AndAlso Not String.Equals(createBy, currentUserId, StringComparison.OrdinalIgnoreCase) Then
            Response.Write("<script>alert('您沒有權限複製此單據');</script>")
            Return
        End If

        Dim copyAttach As String = If(hfCopyAttachment.Value = "1", "1", "0")
        Dim copyMdr As String = If(hfCopyMDR.Value = "1", "1", "0")
        Response.Redirect("ExpenseClaimForm.aspx?CopyFrom=" & jID.ToString() & "&CopyAttach=" & copyAttach & "&CopyMDR=" & copyMdr)
    End Sub
#End Region

#Region "Helper Functions"
    Public Function GetStatusText(status As String) As String
        Select Case status
            Case "P" : Return "草稿"
            Case "W" : Return "待審核"
            Case "A" : Return "已核准"
            Case "R" : Return "已退回"
            Case "Pending" : Return "待審核"
            Case "Approved" : Return "已核准"
            Case "Rejected" : Return "已退回"
            Case Else : Return status
        End Select
    End Function

    Public Function TruncateRemarks(text As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & "..."
    End Function
#End Region

End Class
