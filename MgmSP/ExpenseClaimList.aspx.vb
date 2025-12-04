Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Web.UI.WebControls

''' <summary>
''' 費用申請單查詢 (Expense Claim Search)
''' 建立日期: 2025-12-04
''' </summary>
Partial Public Class ExpenseClaimList
    Inherits System.Web.UI.Page

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("s_id") Is Nothing Then
            Response.Redirect("~/usermgm/login.aspx")
            Return
        End If

        If Not IsPostBack Then
            ' 預設查詢本月單據
            txtDateFrom.Text = DateTime.Today.ToString("yyyy-MM-01")
            txtDateTo.Text = DateTime.Today.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
            BindGrid()
        End If
    End Sub

    Protected Sub btnSearch_Click(sender As Object, e As EventArgs)
        gvList.PageIndex = 0
        BindGrid()
    End Sub

    Protected Sub btnClear_Click(sender As Object, e As EventArgs)
        txtJID.Text = ""
        txtDocEntry.Text = ""
        txtUPID.Text = ""
        txtVendor.Text = ""
        ddlStatus.SelectedIndex = 0
        txtRemarks.Text = ""
        txtDateFrom.Text = ""
        txtDateTo.Text = ""
        BindGrid()
    End Sub

    Protected Sub gvList_PageIndexChanging(sender As Object, e As GridViewPageEventArgs)
        gvList.PageIndex = e.NewPageIndex
        BindGrid()
    End Sub

    Private Sub BindGrid()
        Try
            Dim sqlWhere As String = "WHERE 1=1 "
            Dim params As New List(Of SqlParameter)

            ' jID
            If Not String.IsNullOrEmpty(txtJID.Text) Then
                sqlWhere &= " AND jID = @jID"
                params.Add(New SqlParameter("@jID", txtJID.Text.Trim()))
            End If

            ' AP DocEntry
            If Not String.IsNullOrEmpty(txtDocEntry.Text) Then
                sqlWhere &= " AND DocEntry = @DocEntry"
                params.Add(New SqlParameter("@DocEntry", txtDocEntry.Text.Trim()))
            End If

            ' UPID
            If Not String.IsNullOrEmpty(txtUPID.Text) Then
                sqlWhere &= " AND U_PID = @UPID"
                params.Add(New SqlParameter("@UPID", txtUPID.Text.Trim()))
            End If

            ' Vendor
            If Not String.IsNullOrEmpty(txtVendor.Text) Then
                sqlWhere &= " AND (CardCode LIKE @Vendor OR CardName LIKE @Vendor)"
                params.Add(New SqlParameter("@Vendor", "%" & txtVendor.Text.Trim() & "%"))
            End If

            ' Status
            If Not String.IsNullOrEmpty(ddlStatus.SelectedValue) Then
                sqlWhere &= " AND ApprovalStatus = @Status"
                params.Add(New SqlParameter("@Status", ddlStatus.SelectedValue))
            End If

            ' Remarks
            If Not String.IsNullOrEmpty(txtRemarks.Text) Then
                sqlWhere &= " AND Comments LIKE @Remarks"
                params.Add(New SqlParameter("@Remarks", "%" & txtRemarks.Text.Trim() & "%"))
            End If

            ' Date Range
            If Not String.IsNullOrEmpty(txtDateFrom.Text) Then
                Dim dateField As String = ddlDateType.SelectedValue ' DocDate, DocDueDate, TaxDate
                sqlWhere &= $" AND {dateField} >= @DateFrom"
                params.Add(New SqlParameter("@DateFrom", txtDateFrom.Text))
            End If

            If Not String.IsNullOrEmpty(txtDateTo.Text) Then
                Dim dateField As String = ddlDateType.SelectedValue
                sqlWhere &= $" AND {dateField} <= @DateTo"
                params.Add(New SqlParameter("@DateTo", txtDateTo.Text))
            End If

            Dim sql As String = $"SELECT * FROM jOPCH {sqlWhere} ORDER BY jID DESC"

            Using conn As New SqlConnection(connStr)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddRange(params.ToArray())
                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)
                        gvList.DataSource = dt
                        gvList.DataBind()
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' 簡易錯誤處理
            Response.Write("<script>alert('查詢失敗: " & ex.Message.Replace("'", "\'") & "');</script>")
        End Try
    End Sub

    Protected Sub gvList_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim status As String = DataBinder.Eval(e.Row.DataItem, "ApprovalStatus").ToString()
            Dim lblStatus As Label = CType(e.Row.FindControl("lblStatus"), Label)
            
            lblStatus.Text = GetStatusText(status)
            lblStatus.CssClass = "badge status-" & status
        End If
    End Sub

    Private Function GetStatusText(status As String) As String
        Select Case status
            Case "P" : Return "草稿"
            Case "W" : Return "待審核"
            Case "A" : Return "已核准"
            Case "R" : Return "駁回"
            Case Else : Return status
        End Select
    End Function
End Class