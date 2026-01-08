Imports System.Data.SqlClient

''' <summary>
''' 假日處理工具類
''' 用於判斷假日、計算工作日等
''' </summary>
Public Class HolidayHelper

    ''' <summary>
    ''' 判斷指定日期是否為假日
    ''' </summary>
    ''' <param name="checkDate">要檢查的日期</param>
    ''' <returns>True=假日/False=工作日</returns>
    Public Shared Function IsHoliday(checkDate As DateTime) As Boolean
        ' 先檢查是否為週末
        If checkDate.DayOfWeek = DayOfWeek.Saturday OrElse checkDate.DayOfWeek = DayOfWeek.Sunday Then
            ' 週末預設是假日，但要檢查是否為補班日
            If IsWorkdayOverride(checkDate) Then
                Return False ' 補班日，不是假日
            End If
            Return True ' 週末且非補班日，是假日
        End If

        ' 平日檢查是否為國定假日
        Return IsNationalHoliday(checkDate)
    End Function

    ''' <summary>
    ''' 判斷是否為補班日（週末但需上班）
    ''' </summary>
    Private Shared Function IsWorkdayOverride(checkDate As DateTime) As Boolean
        Try
            Using conn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
                conn.Open()
                Dim sql As String = "SELECT COUNT(*) FROM jHolidays WHERE HolidayDate = @Date AND IsWorkday = 1"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Date", checkDate.Date)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 判斷是否為國定假日
    ''' </summary>
    Private Shared Function IsNationalHoliday(checkDate As DateTime) As Boolean
        Try
            Using conn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
                conn.Open()
                Dim sql As String = "SELECT COUNT(*) FROM jHolidays WHERE HolidayDate = @Date AND IsWorkday = 0"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Date", checkDate.Date)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' 取得下一個工作日
    ''' 如果傳入日期是工作日則回傳本身，否則順延到下一個工作日
    ''' </summary>
    ''' <param name="fromDate">起始日期</param>
    ''' <returns>下一個工作日</returns>
    Public Shared Function GetNextWorkingDay(fromDate As DateTime) As DateTime
        Dim result As DateTime = fromDate
        Dim maxAttempts As Integer = 30 ' 防止無限迴圈

        While IsHoliday(result) AndAlso maxAttempts > 0
            result = result.AddDays(1)
            maxAttempts -= 1
        End While

        Return result
    End Function

    ''' <summary>
    ''' 計算到期日並自動跳過假日
    ''' </summary>
    ''' <param name="baseDate">基準日期</param>
    ''' <param name="addMonths">加月數</param>
    ''' <param name="addDays">加天數</param>
    ''' <param name="wasAdjusted">輸出：是否有順延</param>
    ''' <returns>調整後的到期日</returns>
    Public Shared Function CalculateDueDateSkipHoliday(baseDate As DateTime, addMonths As Integer, addDays As Integer, ByRef wasAdjusted As Boolean) As DateTime
        ' 計算原始到期日
        Dim originalDueDate As DateTime = baseDate.AddMonths(addMonths).AddDays(addDays)

        ' 檢查是否需要順延
        Dim adjustedDueDate As DateTime = GetNextWorkingDay(originalDueDate)

        wasAdjusted = (adjustedDueDate <> originalDueDate)

        Return adjustedDueDate
    End Function

    ''' <summary>
    ''' 取得假日名稱（如果是假日）
    ''' </summary>
    ''' <param name="checkDate">要檢查的日期</param>
    ''' <returns>假日名稱，如果不是假日則回傳空字串</returns>
    Public Shared Function GetHolidayName(checkDate As DateTime) As String
        ' 先檢查週末
        If checkDate.DayOfWeek = DayOfWeek.Saturday Then
            If Not IsWorkdayOverride(checkDate) Then
                Return "星期六"
            End If
        ElseIf checkDate.DayOfWeek = DayOfWeek.Sunday Then
            If Not IsWorkdayOverride(checkDate) Then
                Return "星期日"
            End If
        End If

        ' 檢查國定假日
        Try
            Using conn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
                conn.Open()
                Dim sql As String = "SELECT HolidayName FROM jHolidays WHERE HolidayDate = @Date AND IsWorkday = 0"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Date", checkDate.Date)
                    Dim result As Object = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString()
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' 忽略錯誤
        End Try

        Return String.Empty
    End Function

    ''' <summary>
    ''' 取得指定年份的假日列表
    ''' </summary>
    ''' <param name="year">年份</param>
    ''' <returns>假日列表 (日期, 名稱, 是否補班)</returns>
    Public Shared Function GetHolidaysByYear(year As Integer) As DataTable
        Dim dt As New DataTable()
        Try
            Using conn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
                conn.Open()
                Dim sql As String = "SELECT HolidayDate, HolidayName, IsWorkday, Source FROM jHolidays WHERE Year = @Year ORDER BY HolidayDate"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Year", year)
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(dt)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' 回傳空 DataTable
        End Try
        Return dt
    End Function

    ''' <summary>
    ''' 檢查指定年份是否有假日資料
    ''' </summary>
    ''' <param name="year">年份</param>
    ''' <returns>True=有資料/False=無資料</returns>
    Public Shared Function HasHolidayData(year As Integer) As Boolean
        Try
            Using conn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString)
                conn.Open()
                Dim sql As String = "SELECT COUNT(*) FROM jHolidays WHERE Year = @Year"
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Year", year)
                    Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
                End Using
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

End Class
