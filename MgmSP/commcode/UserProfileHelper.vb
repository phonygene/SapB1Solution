Imports System.Data.SqlClient

''' <summary>
''' 使用者資料模型
''' </summary>
Public Class UserProfileModel
    Public Property UserId As String
    Public Property UserName As String
    Public Property Password As String
    Public Property EmpSeries As String    ' 工號
    Public Property ExpDept As String      ' 費用部門代碼
    Public Property ExpDeptName As String  ' 費用部門名稱（唯讀）
    Public Property Email As String
End Class

''' <summary>
''' 必填欄位檢查結果
''' </summary>
Public Class RequiredFieldsResult
    Public Property IsComplete As Boolean
    Public Property MissingExpDept As Boolean
    Public Property MissingEmpSeries As Boolean
End Class

''' <summary>
''' 使用者資料存取工具類
''' 用於讀取、更新使用者資料及檢查必填欄位
''' </summary>
Public Class UserProfileHelper

    Private Shared ReadOnly connStr As String = System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString

    ''' <summary>
    ''' 取得使用者資料
    ''' </summary>
    ''' <param name="userId">使用者ID</param>
    ''' <returns>UserProfileModel，若找不到則回傳 Nothing</returns>
    Public Shared Function GetUserProfile(userId As String) As UserProfileModel
        If String.IsNullOrEmpty(userId) Then Return Nothing

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT u.id, u.name, u.pwd, u.EmpSeries, u.expDEPT, u.email, d.EDeptName " &
                                    "FROM [User] u " &
                                    "LEFT JOIN jDEPT d ON u.expDEPT = d.EDeptID " &
                                    "WHERE u.id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            Dim model As New UserProfileModel()
                            model.UserId = If(IsDBNull(dr("id")), "", dr("id").ToString())
                            model.UserName = If(IsDBNull(dr("name")), "", dr("name").ToString())
                            model.Password = If(IsDBNull(dr("pwd")), "", dr("pwd").ToString())
                            model.EmpSeries = If(IsDBNull(dr("EmpSeries")), "", dr("EmpSeries").ToString())
                            model.ExpDept = If(IsDBNull(dr("expDEPT")), "", dr("expDEPT").ToString())
                            model.ExpDeptName = If(IsDBNull(dr("EDeptName")), "", dr("EDeptName").ToString())
                            model.Email = If(IsDBNull(dr("email")), "", dr("email").ToString())
                            Return model
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 回傳 Nothing
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' 更新使用者資料（密碼、工號、費用部門、Email）
    ''' </summary>
    ''' <param name="userId">使用者ID</param>
    ''' <param name="model">要更新的資料</param>
    ''' <returns>True=成功/False=失敗</returns>
    Public Shared Function UpdateUserProfile(userId As String, model As UserProfileModel) As Boolean
        If String.IsNullOrEmpty(userId) OrElse model Is Nothing Then Return False

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE [User] SET pwd = @Pwd, EmpSeries = @EmpSeries, expDEPT = @ExpDept, email = @Email WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Pwd", If(String.IsNullOrEmpty(model.Password), DBNull.Value, model.Password))
                    cmd.Parameters.AddWithValue("@EmpSeries", If(String.IsNullOrEmpty(model.EmpSeries), DBNull.Value, model.EmpSeries))
                    cmd.Parameters.AddWithValue("@ExpDept", If(String.IsNullOrEmpty(model.ExpDept), DBNull.Value, model.ExpDept))
                    cmd.Parameters.AddWithValue("@Email", If(String.IsNullOrEmpty(model.Email), DBNull.Value, model.Email))
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 檢查必填欄位是否完整（ExpDept + EmpSeries）
    ''' </summary>
    ''' <param name="userId">使用者ID</param>
    ''' <returns>RequiredFieldsResult</returns>
    Public Shared Function CheckRequiredFields(userId As String) As RequiredFieldsResult
        Dim result As New RequiredFieldsResult()
        result.IsComplete = False
        result.MissingExpDept = True
        result.MissingEmpSeries = True

        If String.IsNullOrEmpty(userId) Then Return result

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "SELECT expDEPT, EmpSeries FROM [User] WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Using dr As SqlDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            Dim expDept As String = If(IsDBNull(dr("expDEPT")), "", dr("expDEPT").ToString().Trim())
                            Dim empSeries As String = If(IsDBNull(dr("EmpSeries")), "", dr("EmpSeries").ToString().Trim())

                            ' 檢查 ExpDept 是否存在且有效
                            result.MissingExpDept = String.IsNullOrEmpty(expDept)

                            ' 檢查 EmpSeries 是否存在
                            result.MissingEmpSeries = String.IsNullOrEmpty(empSeries)

                            ' 兩者都有值才算完整
                            result.IsComplete = Not result.MissingExpDept AndAlso Not result.MissingEmpSeries
                        End If
                    End Using
                End Using
            End Using

            ' 如果 ExpDept 有值，還要檢查是否存在於 jDEPT 中
            If Not result.MissingExpDept Then
                Dim profile = GetUserProfile(userId)
                If profile IsNot Nothing AndAlso Not String.IsNullOrEmpty(profile.ExpDept) Then
                    Using conn As New SqlConnection(connStr)
                        conn.Open()
                        Dim sql As String = "SELECT COUNT(*) FROM jDEPT WHERE EDeptID = @EDeptID"
                        Using cmd As New SqlCommand(sql, conn)
                            cmd.Parameters.AddWithValue("@EDeptID", profile.ExpDept)
                            If Convert.ToInt32(cmd.ExecuteScalar()) = 0 Then
                                result.MissingExpDept = True
                                result.IsComplete = False
                            End If
                        End Using
                    End Using
                End If
            End If

        Catch ex As Exception
            ' 發生錯誤視為未完整，讓使用者填寫
        End Try

        Return result
    End Function

    ''' <summary>
    ''' 更新費用部門和工號
    ''' </summary>
    ''' <param name="userId">使用者ID</param>
    ''' <param name="expDept">費用部門代碼</param>
    ''' <param name="empSeries">工號</param>
    ''' <returns>True=成功/False=失敗</returns>
    Public Shared Function UpdateExpDeptAndEmpSeries(userId As String, expDept As String, empSeries As String) As Boolean
        If String.IsNullOrEmpty(userId) Then Return False

        Try
            Using conn As New SqlConnection(connStr)
                conn.Open()
                Dim sql As String = "UPDATE [User] SET expDEPT = @ExpDept, EmpSeries = @EmpSeries WHERE id = @UserId"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ExpDept", If(String.IsNullOrEmpty(expDept), DBNull.Value, expDept))
                    cmd.Parameters.AddWithValue("@EmpSeries", If(String.IsNullOrEmpty(empSeries), DBNull.Value, empSeries))
                    cmd.Parameters.AddWithValue("@UserId", userId)
                    Return cmd.ExecuteNonQuery() > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
