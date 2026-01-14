# ASP.NET 資料處理指南

> 適用：所有 Agent（Backend 重點參考）
> 建立：2026-01-13

---

## ADO.NET 參數化查詢

### 型別明確原則

**永遠使用 `Parameters.Add(name, SqlDbType).Value`，禁止使用 `AddWithValue`**

```vb
' 正確做法 - 明確指定型別
cmd.Parameters.Add("@DocDate", SqlDbType.Date).Value = docDate
cmd.Parameters.Add("@Amount", SqlDbType.Decimal).Value = amount
cmd.Parameters.Add("@ItemCode", SqlDbType.NVarChar, 50).Value = itemCode

' 錯誤做法 - 禁止使用（會自動推斷型別，導致問題）
cmd.Parameters.AddWithValue("@DocDate", docDate)    ' [X] DateTime 推斷錯誤
cmd.Parameters.AddWithValue("@Amount", amount)      ' [X] Decimal 精度問題
```

### 常見型別對照表

| VB.NET 型別 | SqlDbType | SQL Server 型別 | 說明 |
|-------------|-----------|-----------------|------|
| Date / DateTime | `SqlDbType.Date` | date | 避免 SqlDateTime overflow (1753-9999) |
| DateTime | `SqlDbType.DateTime2` | datetime2 | 完整時間，範圍 0001-9999 |
| Decimal | `SqlDbType.Decimal` | decimal | 金額必用，需指定精度 |
| String | `SqlDbType.NVarChar` | nvarchar | 中文字串，需指定長度 |
| Integer | `SqlDbType.Int` | int | 整數 |
| Boolean | `SqlDbType.Bit` | bit | 布林值 |

### P003 問題：SqlDateTime Overflow

**症狀**：
```
SqlDateTime 溢位。必須在 1/1/1753 12:00:00 AM 和 12/31/9999 11:59:59 PM 之間
```

**原因**：`AddWithValue` 將 DateTime 推斷為 `SqlDbType.DateTime`（範圍 1753-9999）

**解法**：
```vb
' 明確使用 SqlDbType.Date 或 SqlDbType.DateTime2
cmd.Parameters.Add("@DocDate", SqlDbType.Date).Value = docDate

' 若日期可能為空
If docDate = DateTime.MinValue Then
    cmd.Parameters.Add("@DocDate", SqlDbType.Date).Value = DBNull.Value
Else
    cmd.Parameters.Add("@DocDate", SqlDbType.Date).Value = docDate
End If
```

---

## 資料儲存位置選擇

### ViewState vs Session vs Database

| 資料類型 | 儲存位置 | 生命週期 | 範例 |
|----------|----------|----------|------|
| 頁面暫存狀態 | ViewState | 單一頁面 PostBack | 當前 GridView 排序、展開狀態 |
| 跨頁面狀態 | Session | 使用者 Session | 登入資訊、購物車 |
| 需持久化資料 | Database | 永久 | 文件、附件、業務資料 |

### P004 問題：ViewState 資料遺失

**症狀**：用戶上傳附件後更新單據，附件消失

**原因**：ViewState 在頁面生命週期結束後不會自動持久化

**解法**：
```vb
' 錯誤做法 - 附件只存 ViewState，關閉頁面就遺失
ViewState("Attachments") = attachmentList

' 正確做法 - 附件寫入資料庫
Private Sub SaveAttachments(jID As Integer)
    For Each att In attachmentList
        If att.IsNew Then
            InsertAttachment(jID, att)
        End If
    Next
End Sub
```

### ViewState 使用原則

| 適合 | 不適合 |
|------|--------|
| GridView 當前頁碼 | 附件清單 |
| 下拉選單選中值 | 用戶輸入的業務資料 |
| 暫存的計算結果（僅供顯示） | 需要跨 Session 保留的資料 |
| UI 展開/收合狀態 | 任何需要持久化的資料 |

### ViewState 序列化限制

```vb
' 錯誤 - ListItem 無法序列化
ViewState("Items") = New List(Of ListItem)()

' 正確 - 使用 DataTable
Dim dt As New DataTable()
da.Fill(dt)
ViewState("Items") = dt
```

---

## 連線管理

### Using 模式（強制）

```vb
' 正確 - 確保連線釋放
Using conn As New SqlConnection(connStr)
    conn.Open()
    Using cmd As New SqlCommand(sql, conn)
        ' 操作...
    End Using
End Using

' 錯誤 - 可能導致連線洩漏
Dim conn As New SqlConnection(connStr)
conn.Open()
' 若發生例外，連線不會關閉
```

### Transaction 使用

```vb
Using conn As New SqlConnection(connStr)
    conn.Open()
    Using trans As SqlTransaction = conn.BeginTransaction()
        Try
            ' 多個操作...
            cmd1.Transaction = trans
            cmd1.ExecuteNonQuery()

            cmd2.Transaction = trans
            cmd2.ExecuteNonQuery()

            trans.Commit()
        Catch ex As Exception
            trans.Rollback()
            Throw
        End Try
    End Using
End Using
```

---

## 快速檢查清單

開發資料存取程式碼時：

- [ ] SQL 參數使用 `Parameters.Add(name, SqlDbType).Value`
- [ ] 禁止使用 `AddWithValue`
- [ ] 日期欄位使用 `SqlDbType.Date` 或 `SqlDbType.DateTime2`
- [ ] 金額欄位使用 `SqlDbType.Decimal`
- [ ] 需持久化的資料寫入 Database，非 ViewState
- [ ] 連線使用 `Using` 確保釋放
- [ ] 多表操作使用 Transaction

---

## 相關問題模式

| 編號 | 問題 | 本指南章節 |
|------|------|-----------|
| P003 | ADO.NET DateTime 型別推斷錯誤 | ADO.NET 參數化查詢 |
| P004 | ViewState 暫存資料遺失 | 資料儲存位置選擇 |
