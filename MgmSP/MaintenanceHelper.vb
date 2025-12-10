Imports System.Data.SqlClient
Imports System.Web
Imports System.Web.Configuration

''' <summary>
''' 系統維護檢查輔助模組
''' </summary>
Public Class MaintenanceHelper

    Private Shared ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    ''' <summary>
    ''' 檢查系統是否處於維護模式
    ''' </summary>
    ''' <returns>True = 維護中, False = 正常運作</returns>
    Public Shared Function IsUnderMaintenance() As Boolean
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 1 Maintenance FROM OADM"
                Using cmd As New SqlCommand(sql, conn)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return (Convert.ToInt32(result) = 1)
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' 發生錯誤時預設為非維護模式，避免影響正常使用
        End Try
        Return False
    End Function

    ''' <summary>
    ''' 檢查維護狀態並導向維護頁面 (若維護中)
    ''' 在頁面 Page_Load 最開始呼叫此方法
    ''' </summary>
    ''' <param name="response">HttpResponse 物件</param>
    ''' <param name="currentPage">目前頁面路徑 (用於排除維護頁面本身)</param>
    Public Shared Sub CheckAndRedirect(response As HttpResponse, currentPage As String)
        ' 排除維護頁面本身，避免無限迴圈
        If currentPage.ToLower().Contains("maintenance.aspx") Then
            Return
        End If

        If IsUnderMaintenance() Then
            response.Redirect("~/Maintenance.aspx", True)
        End If
    End Sub

End Class
