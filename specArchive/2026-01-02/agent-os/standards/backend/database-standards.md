## 資料庫標準 (Database Standards)

### 資料庫環境

- **DBMS**: Microsoft SQL Server 2005
- **相容性**: 所有 SQL 查詢必須相容於 SQL Server 2005 語法
- **不可使用**: SQL Server 2008 及以後版本的新功能（如 MERGE、DATE/TIME2 等）

### SQL 查詢撰寫規範

#### 優先使用 CTE（Common Table Expression）

當查詢複雜度提升時，優先考慮使用 CTE 架構來優化可讀性和維護性：

**推薦寫法**：
```sql
WITH CustomerOrders AS (
    SELECT
        CustomerID,
        OrderID,
        OrderDate,
        TotalAmount
    FROM Orders
    WHERE OrderDate >= '2024-01-01'
),
OrderSummary AS (
    SELECT
        CustomerID,
        COUNT(OrderID) AS OrderCount,
        SUM(TotalAmount) AS TotalSpent
    FROM CustomerOrders
    GROUP BY CustomerID
)
SELECT
    c.CustomerName,
    os.OrderCount,
    os.TotalSpent
FROM Customers c
INNER JOIN OrderSummary os ON c.CustomerID = os.CustomerID
WHERE os.OrderCount > 5
```

**注意事項**：
- CTE 在 SQL Server 2005 中完全支援
- 僅在不影響效能或結構完整性的情況下使用
- 複雜查詢優先拆分為多個 CTE，提升可讀性

#### SQL Server 2005 相容性檢查清單

**可使用的功能**：
- CTE（Common Table Expressions）
- ROW_NUMBER(), RANK(), DENSE_RANK()
- PIVOT/UNPIVOT
- TRY...CATCH
- XML 資料型別
- VARCHAR(MAX), NVARCHAR(MAX)

**不可使用的功能**（2008+ 版本）**：
- MERGE 語句
- DATE, TIME, DATETIME2, DATETIMEOFFSET 資料型別
- FILESTREAM
- 表值參數（Table-Valued Parameters）
- GROUPING SETS, CUBE, ROLLUP（新語法）

### 程式碼品質要求

- **命名規範**: 使用有意義的表名、欄位名，遵循專案既有命名慣例
- **註解**: 複雜查詢必須加上繁體中文註解說明邏輯
- **格式化**: 使用一致的縮排和換行，提升可讀性
- **效能考量**: 避免 SELECT *，明確指定需要的欄位
- **索引意識**: 在 WHERE、JOIN 條件中使用適當的欄位

### 範例：符合規範的查詢

```sql
-- 查詢 2024 年高價值客戶及其訂單摘要
-- 使用 CTE 優化查詢結構
WITH RecentOrders AS (
    -- 篩選近期訂單
    SELECT
        o.CustomerID,
        o.OrderID,
        o.OrderDate,
        od.ProductID,
        od.Quantity,
        od.UnitPrice,
        (od.Quantity * od.UnitPrice) AS LineTotal
    FROM Orders o
    INNER JOIN OrderDetails od ON o.OrderID = od.OrderID
    WHERE o.OrderDate >= '2024-01-01'
),
CustomerSummary AS (
    -- 計算客戶訂單統計
    SELECT
        CustomerID,
        COUNT(DISTINCT OrderID) AS OrderCount,
        COUNT(DISTINCT ProductID) AS UniqueProducts,
        SUM(LineTotal) AS TotalRevenue
    FROM RecentOrders
    GROUP BY CustomerID
    HAVING SUM(LineTotal) > 100000  -- 高價值客戶門檻
)
-- 最終結果
SELECT
    c.CustomerID,
    c.CustomerName,
    c.ContactName,
    cs.OrderCount AS 訂單數量,
    cs.UniqueProducts AS 產品種類,
    cs.TotalRevenue AS 總營收
FROM Customers c
INNER JOIN CustomerSummary cs ON c.CustomerID = cs.CustomerID
ORDER BY cs.TotalRevenue DESC
```
