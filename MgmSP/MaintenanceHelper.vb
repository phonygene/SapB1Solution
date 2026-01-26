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
    ''' 檢查使用者是否為管理員 (admin='Y')
    ''' </summary>
    ''' <param name="userId">使用者 ID</param>
    ''' <returns>True = 是管理員, False = 非管理員</returns>
    Public Shared Function IsAdminUser(userId As String) As Boolean
        If String.IsNullOrEmpty(userId) Then Return False
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT admin FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Dim result = cmd.ExecuteScalar()
                    Return (result IsNot Nothing AndAlso Not IsDBNull(result) AndAlso Convert.ToInt32(result) = 1)
                End Using
            End Using
        Catch ex As Exception
            ' 發生錯誤時預設為非管理員
        End Try
        Return False
    End Function

    ''' <summary>
    ''' 檢查維護狀態並導向維護頁面 (若維護中且非 admin)
    ''' 在頁面 Page_Load 最開始呼叫此方法
    ''' </summary>
    ''' <param name="response">HttpResponse 物件</param>
    ''' <param name="currentPage">目前頁面路徑 (用於排除維護頁面本身)</param>
    Public Shared Sub CheckAndRedirect(response As HttpResponse, currentPage As String)
        ' 呼叫多載方法，不傳入 Session (向下相容)
        CheckAndRedirect(response, currentPage, Nothing)
    End Sub

    ''' <summary>
    ''' 檢查維護狀態並導向維護頁面 (若維護中且非 admin)
    ''' 在頁面 Page_Load 最開始呼叫此方法
    ''' </summary>
    ''' <param name="response">HttpResponse 物件</param>
    ''' <param name="currentPage">目前頁面路徑 (用於排除維護頁面本身)</param>
    ''' <param name="session">HttpSessionState 物件 (用於檢查使用者是否為 admin)</param>
    Public Shared Sub CheckAndRedirect(response As HttpResponse, currentPage As String, session As System.Web.SessionState.HttpSessionState)
        ' 排除維護頁面本身，避免無限迴圈
        If currentPage.ToLower().Contains("maintenance.aspx") Then
            Return
        End If

        ' 排除登入頁面，讓使用者可以登入 (登入後再檢查 admin)
        If currentPage.ToLower().Contains("login.aspx") Then
            Return
        End If

        ' 如果使用者是 admin，跳過維護檢查
        If session IsNot Nothing AndAlso session("s_id") IsNot Nothing Then
            If IsAdminUser(session("s_id").ToString()) Then
                Return
            End If
        End If

        If IsUnderMaintenance() Then
            response.Redirect("~/Maintenance.aspx", True)
        End If
    End Sub

#Region "功能級維護檢查"
    ''' <summary>
    ''' 檢查指定功能是否處於維護模式
    ''' </summary>
    ''' <param name="featureName">功能名稱: ExpenseClaim, PurchaseRequest</param>
    ''' <returns>True = 維護中, False = 正常運作</returns>
    Public Shared Function IsFeatureUnderMaintenance(featureName As String) As Boolean
        If String.IsNullOrEmpty(featureName) Then Return False
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim columnName As String = "Maint_" & featureName
                ' 使用參數化查詢防止 SQL Injection (雖然 columnName 來自程式碼)
                Dim sql As String = String.Format("SELECT TOP 1 [{0}] FROM OADM", columnName)
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
    ''' 取得功能維護訊息
    ''' </summary>
    ''' <param name="featureName">功能名稱</param>
    ''' <returns>維護訊息，若無則回傳預設訊息</returns>
    Public Shared Function GetMaintenanceMessage() As String
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT TOP 1 MNote FROM OADM"
                Using cmd As New SqlCommand(sql, conn)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
        End Try
        Return "功能維護中，請稍後再試。"
    End Function

    ''' <summary>
    ''' 檢查功能維護狀態並導向維護頁面 (若維護中且非 admin)
    ''' 在功能頁面 Page_Load 最開始呼叫此方法
    ''' </summary>
    ''' <param name="response">HttpResponse 物件</param>
    ''' <param name="featureName">功能名稱: ExpenseClaim, PurchaseRequest</param>
    ''' <param name="session">HttpSessionState 物件</param>
    Public Shared Sub CheckFeatureAndRedirect(response As HttpResponse, featureName As String, session As System.Web.SessionState.HttpSessionState)
        ' 如果使用者是 admin，跳過維護檢查
        If session IsNot Nothing AndAlso session("s_id") IsNot Nothing Then
            If IsAdminUser(session("s_id").ToString()) Then
                Return
            End If
        End If

        If IsFeatureUnderMaintenance(featureName) Then
            response.Redirect("~/FeatureMaintenance.aspx?feature=" & HttpUtility.UrlEncode(featureName), True)
        End If
    End Sub
#End Region

End Class

