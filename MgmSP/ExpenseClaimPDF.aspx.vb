Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class ExpenseClaimPDF
    Inherits System.Web.UI.Page

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
    Private ReadOnly sapConnStr As String = WebConfigurationManager.ConnectionStrings("SapSQLConnection").ConnectionString

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim docEntry As Integer = 0
            If Integer.TryParse(Request.QueryString("DocEntry"), docEntry) AndAlso docEntry > 0 Then
                LoadDocument(docEntry)
            Else
                Response.Write("<h3>錯誤：無效的單據編號</h3>")
                Response.End()
            End If
        End If
    End Sub

    Private Sub LoadDocument(docEntry As Integer)
        Dim slpName As String = ""
        Dim pymntGroupName As String = ""
        Dim slpCode As Integer = 0
        Dim groupNum As Integer = 0

        Using conn As New SqlConnection(connStr)
            conn.Open()

            Dim sql As String = "SELECT * FROM jOPCH WHERE DocEntry = @DocEntry"

            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    If dr.Read() Then
                        litDocNum.Text = dr("DocEntry").ToString()
                        litStatus.Text = GetStatusBadge(dr("ApprovalStatus").ToString())
                        litCardCode.Text = dr("CardCode").ToString()
                        litCardName.Text = dr("CardName").ToString()
                        litNumAtCard.Text = dr("NumAtCard").ToString()

                        If Not IsDBNull(dr("SlpCode")) Then
                            Integer.TryParse(dr("SlpCode").ToString(), slpCode)
                        End If
                        If Not IsDBNull(dr("GroupNum")) Then
                            Integer.TryParse(dr("GroupNum").ToString(), groupNum)
                        End If

                        If Not IsDBNull(dr("DocDate")) Then
                            litDocDate.Text = Convert.ToDateTime(dr("DocDate")).ToString("yyyy-MM-dd")
                        End If
                        If Not IsDBNull(dr("DocDueDate")) Then
                            litDocDueDate.Text = Convert.ToDateTime(dr("DocDueDate")).ToString("yyyy-MM-dd")
                        End If
                        If Not IsDBNull(dr("TaxDate")) Then
                            litTaxDate.Text = Convert.ToDateTime(dr("TaxDate")).ToString("yyyy-MM-dd")
                        End If

                        Dim currency As String = dr("DocCurrency").ToString()
                        Dim rate As String = dr("DocRate").ToString()
                        litCurrency.Text = currency & " / " & rate

                        If Not IsDBNull(dr("PymntGroup")) AndAlso dr("PymntGroup").ToString() <> "" Then
                            litPaymentTerms.Text = dr("PymntGroup").ToString()
                        End If

                        litRemarks.Text = dr("Comments").ToString()

                        Dim docTotal As Decimal = Convert.ToDecimal(dr("DocTotal"))
                        Dim vatSum As Decimal = Convert.ToDecimal(dr("VatSum"))
                        litDocTotal.Text = docTotal.ToString("N0")
                        litVatSum.Text = vatSum.ToString("N0")
                        litGrandTotal.Text = (docTotal + vatSum).ToString("N0")

                        litCreateBy.Text = dr("CreateBy").ToString()
                        If Not IsDBNull(dr("CreateDate")) Then
                            litCreateDate.Text = Convert.ToDateTime(dr("CreateDate")).ToString("yyyy-MM-dd HH:mm")
                        End If
                        litApprovedBy.Text = dr("ApprovedBy").ToString()
                        If Not IsDBNull(dr("ApprovalDate")) Then
                            litApprovalDate.Text = Convert.ToDateTime(dr("ApprovalDate")).ToString("yyyy-MM-dd HH:mm")
                        End If
                        litApprovalComments.Text = dr("ApprovalComments").ToString()
                    End If
                End Using
            End Using

            ' Query SAP for SlpName and PymntGroup name
            Try
                Using sapConn As New SqlConnection(sapConnStr)
                    sapConn.Open()
                    If slpCode > 0 Then
                        Using cmdSlp As New SqlCommand("SELECT SlpName FROM OSLP WHERE SlpCode = @SlpCode", sapConn)
                            cmdSlp.Parameters.AddWithValue("@SlpCode", slpCode)
                            Dim result = cmdSlp.ExecuteScalar()
                            If result IsNot Nothing Then slpName = result.ToString()
                        End Using
                    End If
                    If groupNum > 0 Then
                        Using cmdGrp As New SqlCommand("SELECT PymntGroup FROM OCTG WHERE GroupNum = @GroupNum", sapConn)
                            cmdGrp.Parameters.AddWithValue("@GroupNum", groupNum)
                            Dim result = cmdGrp.ExecuteScalar()
                            If result IsNot Nothing Then pymntGroupName = result.ToString()
                        End Using
                    End If
                End Using
            Catch ex As Exception
                ' Ignore SAP connection errors
            End Try

            litPurchaser.Text = slpName
            If Not String.IsNullOrEmpty(pymntGroupName) Then
                litPaymentTerms.Text = pymntGroupName & " (" & litPaymentTerms.Text & ")"
            End If

            Dim expenseLines As New List(Of Object)
            sql = "SELECT l.*, ISNULL(c.CategoryName, l.ItemCode) AS CategoryName " & _
                  "FROM jPCH1 l " & _
                  "LEFT JOIN expense_category c ON l.ItemCode = c.CategoryCode " & _
                  "WHERE l.DocEntry = @DocEntry ORDER BY l.LineNum"

            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        expenseLines.Add(New With {
                            .LineNum = dr("LineNum"),
                            .CategoryName = dr("CategoryName").ToString(),
                            .Description = dr("Dscription").ToString(),
                            .AcctCode = dr("AcctCode").ToString(),
                            .VatGroupName = GetVatGroupName(dr("VatGroup").ToString()),
                            .LineTotal = Convert.ToDecimal(dr("LineTotal")),
                            .LineVat = Convert.ToDecimal(dr("LineVat")),
                            .GTotal = Convert.ToDecimal(dr("GTotal"))
                        })
                    End While
                End Using
            End Using
            rptExpenseLines.DataSource = expenseLines
            rptExpenseLines.DataBind()

            Dim mdrLines As New List(Of Object)
            sql = "SELECT * FROM jMGUIAPDetail WHERE DocEntry = @DocEntry ORDER BY LineNum"

            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DocEntry", docEntry)
                Using dr As SqlDataReader = cmd.ExecuteReader()
                    While dr.Read()
                        mdrLines.Add(New With {
                            .LineNum = dr("LineNum"),
                            .U_STCEG = dr("U_STCEG").ToString(),
                            .U_XBLNR = dr("U_XBLNR").ToString(),
                            .ZFormName = GetZFormName(dr("U_ZFORM_CODE").ToString()),
                            .U_BLDAT = If(IsDBNull(dr("U_BLDAT")), Nothing, Convert.ToDateTime(dr("U_BLDAT"))),
                            .U_HWBAS = Convert.ToDecimal(dr("U_HWBAS")),
                            .U_HWSTE = Convert.ToDecimal(dr("U_HWSTE"))
                        })
                    End While
                End Using
            End Using

            If mdrLines.Count > 0 Then
                pnlMDR.Visible = True
                rptMDRLines.DataSource = mdrLines
                rptMDRLines.DataBind()
            End If

        End Using
    End Sub

    Private Function GetStatusBadge(status As String) As String
        Dim text As String = ""
        Dim cssClass As String = "status-" & status

        Select Case status
            Case "W"
                text = ChrW(&H5F85) & ChrW(&H5BE9) & ChrW(&H6838)
            Case "A"
                text = ChrW(&H5DF2) & ChrW(&H6838) & ChrW(&H51C6)
            Case "R"
                text = ChrW(&H5DF2) & ChrW(&H9000) & ChrW(&H56DE)
            Case "P"
                text = ChrW(&H8349) & ChrW(&H7A3F)
            Case Else
                text = status
        End Select

        Return String.Format("<span class='status-badge {0}'>{1}</span>", cssClass, text)
    End Function

    Private Function GetVatGroupName(vatGroup As String) As String
        Select Case vatGroup
            Case "1"
                Return ChrW(&H61C9) & ChrW(&H7A05) & " 5%"
            Case "2"
                Return ChrW(&H96F6) & ChrW(&H7A05) & ChrW(&H7387)
            Case "3"
                Return ChrW(&H514D) & ChrW(&H7A05)
            Case Else
                Return vatGroup
        End Select
    End Function

    Private Function GetZFormName(zformCode As String) As String
        Select Case zformCode
            Case "21"
                Return "21-" & ChrW(&H4E09) & ChrW(&H806F) & ChrW(&H624B) & ChrW(&H958B) & ChrW(&H767C) & ChrW(&H7968)
            Case "22"
                Return "22-" & ChrW(&H9AD8) & ChrW(&H9435) & "/" & ChrW(&H4E8C) & ChrW(&H806F) & ChrW(&H6536) & ChrW(&H9280) & ChrW(&H6A5F)
            Case "25"
                Return "25-" & ChrW(&H96FB) & ChrW(&H5B50) & ChrW(&H767C) & ChrW(&H7968) & "/" & ChrW(&H516C) & ChrW(&H71DF)
            Case "28"
                Return "28-" & ChrW(&H6D77) & ChrW(&H95DC) & ChrW(&H4EE3) & ChrW(&H5FB5)
            Case "99"
                Return "99-" & ChrW(&H5176) & ChrW(&H4ED6)
            Case Else
                Return zformCode
        End Select
    End Function

End Class
