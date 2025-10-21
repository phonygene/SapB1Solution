## 程式碼風格規範 (Coding Style Standards)

### 通用原則

撰寫任何程式代碼（包括但不限於）時，必須遵循當前業界趨勢的優良風格和規範：

- SQL Query
- VB.NET
- C#
- JavaScript/TypeScript
- HTML/CSS
- PowerShell/Batch Script
- CMD Prompt Scripts
- 其他腳本語言

### 程式碼品質要求

#### 可讀性

- **Consistent Naming Conventions**: Establish and follow naming conventions for variables, functions, classes, and files across the codebase
- **Meaningful Names**: Choose descriptive names that reveal intent; avoid abbreviations and single-letter variables except in narrow contexts
- **適當註解**: 複雜邏輯必須加上註解說明（繁體中文）
- **Automated Formatting**: Maintain consistent code style (indenting, line breaks, etc.)
- **Consistent Indentation**: Use consistent indentation (spaces or tabs) and configure your editor/linter to enforce it

#### 現代化實踐

- **Small, Focused Functions**: Keep functions small and focused on a single task for better readability and testability
- **DRY Principle**: Avoid duplication by extracting common logic into reusable functions or modules
- **避免過時寫法**: 不使用已被廢棄或不推薦的語法
- **採用最佳實踐**: 參考官方文件和社群標準
- **效能意識**: 在不犧牲可讀性的前提下優化效能
- **Remove Dead Code**: Delete unused code, commented-out blocks, and imports rather than leaving them as clutter

#### 安全性

- **SQL Injection 防護**: 使用參數化查詢
- **輸入驗證**: 所有使用者輸入必須驗證
- **錯誤處理**: 適當的異常處理機制

#### 向後相容性

- **Backward compatibility only when required:** Unless specifically instructed otherwise, assume you do not need to write additional code logic to handle backward compatibility

### SQL 查詢風格

參考 `@agent-os/standards/backend/database-standards.md` 中的 SQL Server 2005 相容性規範。

**核心原則**：
- 優先使用 CTE 提升可讀性（除非影響效能）
- 明確列出欄位，避免 SELECT *
- 適當使用索引和 JOIN
- 加上繁體中文註解說明業務邏輯

### VB.NET 風格

- **命名慣例**: PascalCase（方法、類別）、camelCase（區域變數）
- **縮排**: 4 空格
- **註解**: 使用 XML 文件註解標註公開方法
- **Option Strict On**: 強制型別檢查

### 範例：優良風格的程式碼

```vb.net
''' <summary>
''' 根據客戶ID取得訂單摘要
''' </summary>
''' <param name="customerId">客戶編號</param>
''' <returns>訂單摘要資料表</returns>
Public Function GetOrderSummary(customerId As String) As DataTable
    Dim query As String = "
        WITH CustomerOrders AS (
            SELECT
                OrderID,
                OrderDate,
                TotalAmount
            FROM Orders
            WHERE CustomerID = @CustomerID
        )
        SELECT * FROM CustomerOrders
        ORDER BY OrderDate DESC
    "

    Using conn As New SqlConnection(connectionString)
        Using cmd As New SqlCommand(query, conn)
            ' 使用參數化查詢防止 SQL Injection
            cmd.Parameters.AddWithValue("@CustomerID", customerId)

            Dim adapter As New SqlDataAdapter(cmd)
            Dim result As New DataTable()
            adapter.Fill(result)

            Return result
        End Using
    End Using
End Function
```

### 持續改進

- **程式碼審查**: 重視程式碼品質而非速度
- **重構意識**: 發現不良實踐時主動提出改進
- **學習新趨勢**: 隨時關注語言和框架的最新最佳實踐
