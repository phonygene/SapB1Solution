-- ============================================================
-- 請購單資料驗證腳本
-- 用途：定期執行，檢查資料一致性問題
-- 建立：2026-01-14
-- ============================================================

PRINT '=========================================='
PRINT '請購單資料驗證報告'
PRINT '執行時間: ' + CONVERT(VARCHAR, GETDATE(), 120)
PRINT '=========================================='
PRINT ''

-- ============================================================
-- 1. 孤兒明細檢查（明細沒有對應表頭）
-- ============================================================
PRINT '【1】孤兒明細檢查（jPRQ1 沒有對應 jOPRQ）'

SELECT
    D.jID,
    D.LineNum,
    D.ItemCode,
    D.Quantity,
    D.LineTotal
FROM jPRQ1 D
LEFT JOIN jOPRQ H ON D.jID = H.jID
WHERE H.jID IS NULL

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：無孤兒明細'
ELSE
    PRINT '  ✗ 失敗：發現孤兒明細，請檢查上方資料'

PRINT ''

-- ============================================================
-- 2. 表頭沒有明細檢查
-- ============================================================
PRINT '【2】空單據檢查（jOPRQ 沒有 jPRQ1 明細）'

SELECT
    H.jID,
    H.ReqName,
    H.DocDate,
    H.DocTotal,
    H.ApprovalStatus
FROM jOPRQ H
LEFT JOIN jPRQ1 D ON H.jID = D.jID
WHERE D.jID IS NULL
  AND H.Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：無空單據'
ELSE
    PRINT '  ⚠ 警告：發現空單據（無明細），請確認是否正常'

PRINT ''

-- ============================================================
-- 3. 金額一致性檢查（表頭 DocTotal vs 明細 SUM(GTotal)）
-- 注意：DocTotal 是含稅總額，應與 SUM(GTotal) 或 SUM(LineTotal+LineVat) 比對
-- ============================================================
PRINT '【3】金額一致性檢查（DocTotal vs 明細含稅加總）'

SELECT
    H.jID,
    H.ReqName,
    H.DocTotal AS [表頭金額_含稅],
    ISNULL(D.GrossSum, 0) AS [明細加總_含稅],
    H.DocTotal - ISNULL(D.GrossSum, 0) AS [差額]
FROM jOPRQ H
LEFT JOIN (
    SELECT jID, SUM(ISNULL(GTotal, LineTotal + ISNULL(LineVat, 0))) AS GrossSum
    FROM jPRQ1
    GROUP BY jID
) D ON H.jID = D.jID
WHERE ABS(H.DocTotal - ISNULL(D.GrossSum, 0)) > 0.01
  AND H.Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：金額一致'
ELSE
    PRINT '  ✗ 失敗：金額不一致，請檢查上方資料'

PRINT ''

-- ============================================================
-- 4. 稅額一致性檢查（表頭 VatSum vs 明細 SUM）
-- ============================================================
PRINT '【4】稅額一致性檢查（VatSum vs 明細加總）'

SELECT
    H.jID,
    H.ReqName,
    H.VatSum AS [表頭稅額],
    ISNULL(D.VatSum, 0) AS [明細加總],
    H.VatSum - ISNULL(D.VatSum, 0) AS [差額]
FROM jOPRQ H
LEFT JOIN (
    SELECT jID, SUM(LineVat) AS VatSum
    FROM jPRQ1
    GROUP BY jID
) D ON H.jID = D.jID
WHERE ABS(ISNULL(H.VatSum, 0) - ISNULL(D.VatSum, 0)) > 0.01
  AND H.Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：稅額一致'
ELSE
    PRINT '  ✗ 失敗：稅額不一致，請檢查上方資料'

PRINT ''

-- ============================================================
-- 5. 必填欄位檢查
-- ============================================================
PRINT '【5】必填欄位檢查'

-- 5a. 表頭必填欄位
SELECT
    jID,
    CASE WHEN ReqName IS NULL OR ReqName = '' THEN 'ReqName空' ELSE '' END +
    CASE WHEN DocDate IS NULL THEN 'DocDate空' ELSE '' END AS [問題欄位]
FROM jOPRQ
WHERE (ReqName IS NULL OR ReqName = '' OR DocDate IS NULL)
  AND Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：表頭必填欄位完整'
ELSE
    PRINT '  ✗ 失敗：表頭缺少必填欄位'

-- 5b. 明細必填欄位
SELECT
    D.jID,
    D.LineNum,
    CASE WHEN D.ItemCode IS NULL OR D.ItemCode = '' THEN 'ItemCode空' ELSE '' END +
    CASE WHEN D.Quantity IS NULL OR D.Quantity <= 0 THEN 'Quantity異常' ELSE '' END AS [問題欄位]
FROM jPRQ1 D
INNER JOIN jOPRQ H ON D.jID = H.jID
WHERE (D.ItemCode IS NULL OR D.ItemCode = '' OR D.Quantity IS NULL OR D.Quantity <= 0)
  AND H.Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：明細必填欄位完整'
ELSE
    PRINT '  ✗ 失敗：明細缺少必填欄位'

PRINT ''

-- ============================================================
-- 6. 審核狀態一致性檢查
-- ============================================================
PRINT '【6】審核狀態一致性檢查'

-- 已審核但缺少審核人/日期
SELECT
    jID,
    ReqName,
    ApprovalStatus,
    ApprovedBy,
    ApprovedDate
FROM jOPRQ
WHERE ApprovalStatus IN ('A', 'Approved')
  AND (ApprovedBy IS NULL OR ApprovedDate IS NULL)
  AND Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：審核狀態一致'
ELSE
    PRINT '  ✗ 失敗：已審核單據缺少審核人或審核日期'

PRINT ''

-- ============================================================
-- 7. SAP 過帳狀態檢查
-- ============================================================
PRINT '【7】SAP 過帳狀態檢查'

-- 已過帳但缺少 DocEntry
SELECT
    jID,
    ReqName,
    B1PostStatus,
    DocEntry,
    DocNum,
    B1ErrMsg
FROM jOPRQ
WHERE B1PostStatus = 'Y'
  AND (DocEntry IS NULL OR DocEntry = 0)
  AND Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：SAP 過帳狀態一致'
ELSE
    PRINT '  ✗ 失敗：已過帳但缺少 SAP DocEntry'

-- 過帳失敗但沒有錯誤訊息
SELECT
    jID,
    ReqName,
    B1PostStatus,
    B1ErrMsg
FROM jOPRQ
WHERE B1PostStatus = 'E'
  AND (B1ErrMsg IS NULL OR B1ErrMsg = '')
  AND Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：過帳失敗都有錯誤訊息'
ELSE
    PRINT '  ⚠ 警告：過帳失敗但缺少錯誤訊息'

PRINT ''

-- ============================================================
-- 8. 明細計算驗證（LineTotal = Quantity * Price）
-- ============================================================
PRINT '【8】明細計算驗證（LineTotal = Quantity * Price）'

SELECT
    D.jID,
    D.LineNum,
    D.ItemCode,
    D.Quantity,
    D.Price,
    D.LineTotal AS [儲存的LineTotal],
    D.Quantity * D.Price AS [計算的LineTotal],
    D.LineTotal - (D.Quantity * D.Price) AS [差額]
FROM jPRQ1 D
INNER JOIN jOPRQ H ON D.jID = H.jID
WHERE ABS(D.LineTotal - (D.Quantity * D.Price)) > 0.01
  AND H.Canceled = 'N'

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：明細金額計算正確'
ELSE
    PRINT '  ✗ 失敗：明細金額計算不一致'

PRINT ''

-- ============================================================
-- 9. 附件檢查（jAttach 對應）
-- ============================================================
PRINT '【9】附件完整性檢查'

-- 檢查附件記錄是否有對應的請購單
SELECT
    A.jID,
    A.FileName,
    A.FilePath
FROM jAttach A
LEFT JOIN jOPRQ H ON A.jID = H.jID
WHERE A.DocType = 'PR'
  AND H.jID IS NULL
  AND ISNULL(A.IsDeleted, 0) = 0

IF @@ROWCOUNT = 0
    PRINT '  ✓ 通過：附件對應正確'
ELSE
    PRINT '  ⚠ 警告：發現孤兒附件（請購單已刪除但附件仍存在）'

PRINT ''

-- ============================================================
-- 10. 統計摘要
-- ============================================================
PRINT '【10】統計摘要'

SELECT
    '總單據數' AS [項目],
    COUNT(*) AS [數量]
FROM jOPRQ
WHERE Canceled = 'N'

UNION ALL

SELECT
    '待審核',
    COUNT(*)
FROM jOPRQ
WHERE ApprovalStatus IN ('W', 'Pending')
  AND Canceled = 'N'

UNION ALL

SELECT
    '已審核',
    COUNT(*)
FROM jOPRQ
WHERE ApprovalStatus IN ('A', 'Approved')
  AND Canceled = 'N'

UNION ALL

SELECT
    '已駁回',
    COUNT(*)
FROM jOPRQ
WHERE ApprovalStatus IN ('R', 'Rejected')
  AND Canceled = 'N'

UNION ALL

SELECT
    '已過帳SAP',
    COUNT(*)
FROM jOPRQ
WHERE B1PostStatus = 'Y'
  AND Canceled = 'N'

UNION ALL

SELECT
    '過帳失敗',
    COUNT(*)
FROM jOPRQ
WHERE B1PostStatus = 'E'
  AND Canceled = 'N'

PRINT ''
PRINT '=========================================='
PRINT '驗證完成'
PRINT '=========================================='
