Imports System.Data.SqlClient
Imports System.Threading
Imports System.Web
Imports System.Web.Configuration

''' <summary>
''' [E] 稽核日誌記錄器
''' 使用非同步方式記錄，不影響主流程效能
''' </summary>
Public Class AuditLogger
    Private Shared ReadOnly connStr As String = WebConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    ''' <summary>
    ''' 記錄操作日誌 (非阻塞)
    ''' </summary>
    Public Shared Sub Log(tableName As String, docEntry As Integer, action As String, userId As String,
                          Optional oldValue As String = Nothing,
                          Optional newValue As String = Nothing,
                          Optional changes As String = Nothing)
        ' 取得 IP 和 UserAgent
        Dim ipAddress As String = ""
        Dim userAgent As String = ""
        If HttpContext.Current IsNot Nothing Then
            ipAddress = GetClientIP()
            userAgent = If(HttpContext.Current.Request.UserAgent, "")
            If userAgent.Length > 500 Then userAgent = userAgent.Substring(0, 500)
        End If

        ' 使用 ThreadPool 非同步寫入，避免阻塞主流程
        ThreadPool.QueueUserWorkItem(Sub(state)
                                         WriteLogInternal(tableName, docEntry, action, userId, oldValue, newValue, changes, ipAddress, userAgent)
                                     End Sub)
    End Sub

    ''' <summary>
    ''' 記錄操作日誌 (同步版本，用於需要確保記錄的場合)
    ''' </summary>
    Public Shared Sub LogSync(tableName As String, docEntry As Integer, action As String, userId As String,
                              Optional oldValue As String = Nothing,
                              Optional newValue As String = Nothing,
                              Optional changes As String = Nothing)
        Dim ipAddress As String = ""
        Dim userAgent As String = ""
        If HttpContext.Current IsNot Nothing Then
            ipAddress = GetClientIP()
            userAgent = If(HttpContext.Current.Request.UserAgent, "")
            If userAgent.Length > 500 Then userAgent = userAgent.Substring(0, 500)
        End If

        WriteLogInternal(tableName, docEntry, action, userId, oldValue, newValue, changes, ipAddress, userAgent)
    End Sub

    ''' <summary>
    ''' 內部寫入方法
    ''' </summary>
    Private Shared Sub WriteLogInternal(tableName As String, docEntry As Integer, action As String, userId As String,
                                        oldValue As String, newValue As String, changes As String,
                                        ipAddress As String, userAgent As String)
        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "INSERT INTO jAuditLog (TableName, DocEntry, Action, UserId, ActionDate, OldValue, NewValue, Changes, IPAddress, UserAgent) " &
                                    "VALUES (@TableName, @DocEntry, @Action, @UserId, GETDATE(), @OldValue, @NewValue, @Changes, @IPAddress, @UserAgent)"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@TableName", tableName)
                    cmd.Parameters.AddWithValue("@DocEntry", If(docEntry > 0, docEntry, DBNull.Value))
                    cmd.Parameters.AddWithValue("@Action", action)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    cmd.Parameters.AddWithValue("@OldValue", If(String.IsNullOrEmpty(oldValue), DBNull.Value, oldValue))
                    cmd.Parameters.AddWithValue("@NewValue", If(String.IsNullOrEmpty(newValue), DBNull.Value, newValue))
                    cmd.Parameters.AddWithValue("@Changes", If(String.IsNullOrEmpty(changes), DBNull.Value, changes))
                    cmd.Parameters.AddWithValue("@IPAddress", If(String.IsNullOrEmpty(ipAddress), DBNull.Value, ipAddress))
                    cmd.Parameters.AddWithValue("@UserAgent", If(String.IsNullOrEmpty(userAgent), DBNull.Value, userAgent))
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            ' 記錄失敗時不拋出例外，避免影響主流程
            ' 可選：寫入 Windows Event Log 或檔案日誌
            System.Diagnostics.Debug.WriteLine("AuditLog Error: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 取得客戶端 IP
    ''' </summary>
    Private Shared Function GetClientIP() As String
        Try
            Dim ip As String = HttpContext.Current.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            If String.IsNullOrEmpty(ip) Then
                ip = HttpContext.Current.Request.ServerVariables("REMOTE_ADDR")
            End If
            If String.IsNullOrEmpty(ip) Then
                ip = HttpContext.Current.Request.UserHostAddress
            End If
            Return If(ip, "")
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 常用操作常數
    ''' </summary>
    Public Class Actions
        Public Const Create As String = "CREATE"
        Public Const Update As String = "UPDATE"
        Public Const Delete As String = "DELETE"
        Public Const StatusChange As String = "STATUS_CHANGE"
        Public Const Approve As String = "APPROVE"
        Public Const Reject As String = "REJECT"
        Public Const Submit As String = "SUBMIT"
        Public Const View As String = "VIEW"
        Public Const Download As String = "DOWNLOAD"
        Public Const Upload As String = "UPLOAD"
    End Class
End Class
