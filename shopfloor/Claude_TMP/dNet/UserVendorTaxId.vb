'*******************************************************************************
'* 檔案名稱：UserVendorTaxId.vb
'* 用途：使用者常用統編清單的資料存取類別
'*
'* 使用方式：
'* 1. 將本檔案加入到 MgmSP 專案的 commcode 資料夾
'* 2. 在需要使用的頁面加入：Imports UserVendorTaxId
'*
'* 提供的方法：
'* - AddTaxId(userId, taxId, vendorName)：新增常用統編
'* - GetUserTaxIdList(userId)：取得使用者的常用統編清單
'* - UpdateLastUsed(userId, taxId)：更新最後使用時間
'* - DeleteTaxId(userId, taxId)：刪除常用統編
'* - TaxIdExists(userId, taxId)：檢查統編是否已存在
'*
'* 注意事項：
'* - 使用 jtdbConnectionString 連接字串（確認 Web.config 已設定）
'* - 所有方法都使用參數化查詢（防止 SQL Injection）
'* - 使用 Using 語句確保連線正確釋放
'*******************************************************************************

Option Strict On
Option Explicit On

Imports System.Data.SqlClient
Imports System.Configuration

Public Class UserVendorTaxId

    ''' <summary>
    ''' 新增常用統編
    ''' </summary>
    ''' <param name="userId">使用者 ID（對應 User.id）</param>
    ''' <param name="taxId">統一編號（8位數）</param>
    ''' <param name="vendorName">供應商名稱（選填）</param>
    ''' <exception cref="SqlException">資料庫操作失敗時拋出</exception>
    Public Shared Sub AddTaxId(userId As String, taxId As String, Optional vendorName As String = "")
        If String.IsNullOrEmpty(userId) Then
            Throw New ArgumentException("使用者 ID 不可為空", "userId")
        End If

        If String.IsNullOrEmpty(taxId) Then
            Throw New ArgumentException("統一編號不可為空", "taxId")
        End If

        Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        Dim sql As String = "
            INSERT INTO dbo.user_vendor_taxid (id, taxid, vendorname)
            VALUES (@id, @taxid, @vendorname)"

        Using conn As New SqlConnection(connStr)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@id", userId)
                cmd.Parameters.AddWithValue("@taxid", taxId)
                cmd.Parameters.AddWithValue("@vendorname", If(String.IsNullOrEmpty(vendorName), DBNull.Value, CObj(vendorName)))

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 取得使用者的常用統編清單（依最後使用時間排序）
    ''' </summary>
    ''' <param name="userId">使用者 ID</param>
    ''' <returns>包含 num, taxid, vendorname, createdate, lastused 的 DataTable</returns>
    Public Shared Function GetUserTaxIdList(userId As String) As DataTable
        If String.IsNullOrEmpty(userId) Then
            Return New DataTable() ' 回傳空表
        End If

        Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        Dim sql As String = "
            SELECT num, taxid, vendorname, createdate, lastused
            FROM dbo.user_vendor_taxid
            WHERE id = @id
            ORDER BY ISNULL(lastused, createdate) DESC"

        Dim dt As New DataTable()
        Using conn As New SqlConnection(connStr)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@id", userId)

                conn.Open()
                Using adapter As New SqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using

        Return dt
    End Function

    ''' <summary>
    ''' 更新最後使用時間（當使用者選擇某個統編時呼叫）
    ''' </summary>
    ''' <param name="userId">使用者 ID</param>
    ''' <param name="taxId">統一編號</param>
    Public Shared Sub UpdateLastUsed(userId As String, taxId As String)
        If String.IsNullOrEmpty(userId) OrElse String.IsNullOrEmpty(taxId) Then
            Return ' 靜默失敗，不影響主流程
        End If

        Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        Dim sql As String = "
            UPDATE dbo.user_vendor_taxid
            SET lastused = GETDATE()
            WHERE id = @id AND taxid = @taxid"

        Try
            Using conn As New SqlConnection(connStr)
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", userId)
                    cmd.Parameters.AddWithValue("@taxid", taxId)

                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            ' 更新 lastused 失敗不影響主要功能，記錄錯誤即可
            ' 可根據專案需求決定是否記錄到 Log
        End Try
    End Sub

    ''' <summary>
    ''' 刪除常用統編
    ''' </summary>
    ''' <param name="userId">使用者 ID</param>
    ''' <param name="taxId">統一編號</param>
    Public Shared Sub DeleteTaxId(userId As String, taxId As String)
        If String.IsNullOrEmpty(userId) OrElse String.IsNullOrEmpty(taxId) Then
            Throw New ArgumentException("使用者 ID 與統一編號不可為空")
        End If

        Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        Dim sql As String = "
            DELETE FROM dbo.user_vendor_taxid
            WHERE id = @id AND taxid = @taxid"

        Using conn As New SqlConnection(connStr)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@id", userId)
                cmd.Parameters.AddWithValue("@taxid", taxId)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' 檢查統編是否已存在（避免重複新增）
    ''' </summary>
    ''' <param name="userId">使用者 ID</param>
    ''' <param name="taxId">統一編號</param>
    ''' <returns>True 表示已存在，False 表示不存在</returns>
    Public Shared Function TaxIdExists(userId As String, taxId As String) As Boolean
        If String.IsNullOrEmpty(userId) OrElse String.IsNullOrEmpty(taxId) Then
            Return False
        End If

        Dim connStr As String = ConfigurationManager.ConnectionStrings("jtdbConnectionString").ConnectionString
        Dim sql As String = "
            SELECT COUNT(*)
            FROM dbo.user_vendor_taxid
            WHERE id = @id AND taxid = @taxid"

        Using conn As New SqlConnection(connStr)
            Using cmd As New SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@id", userId)
                cmd.Parameters.AddWithValue("@taxid", taxId)

                conn.Open()
                Dim count As Integer = CInt(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

End Class
