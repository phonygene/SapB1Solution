Imports System.Data.SqlClient
Imports System.Web.Configuration

''' <summary>
''' 系統維護頁面
''' 當 OADM.Maintenance = 1 時，使用者會被導向此頁面
''' </summary>
Partial Public Class Maintenance
    Inherits System.Web.UI.Page

    Private ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            LoadMaintenanceMessage()
        End If
    End Sub

    ''' <summary>
    ''' 載入維護訊息
    ''' </summary>
    Private Sub LoadMaintenanceMessage()
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 1 MNote FROM OADM"
                Using cmd As New SqlCommand(sql, conn)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        litMaintenanceNote.Text = Server.HtmlEncode(result.ToString()).Replace(vbCrLf, "<br/>").Replace(vbLf, "<br/>")
                    Else
                        litMaintenanceNote.Text = "系統維護中，請稍後再試。"
                    End If
                End Using
            End Using
        Catch ex As Exception
            litMaintenanceNote.Text = "系統維護中，請稍後再試。"
        End Try
    End Sub

End Class
