-- =============================================================================
-- 資料庫比對與還原腳本
-- 目的：比對 jtdb 與 jtdb_FIX 的差異，並還原資料
-- 日期：2025/12/12
-- =============================================================================

-- =========================================
-- 步驟 1: 比對四個表的欄位差異
-- =========================================

PRINT '========================================';
PRINT '步驟 1: 比對欄位差異';
PRINT '========================================';

-- 1.1 jOPCH 欄位比對
PRINT '--- jOPCH 欄位差異 ---';
SELECT 'jtdb 有, jtdb_FIX 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jOPCH'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jOPCH' AND b.COLUMN_NAME = a.COLUMN_NAME
  )
UNION ALL
SELECT 'jtdb_FIX 有, jtdb 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jOPCH'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jOPCH' AND b.COLUMN_NAME = a.COLUMN_NAME
  );

-- 1.2 jPCH1 欄位比對
PRINT '--- jPCH1 欄位差異 ---';
SELECT 'jtdb 有, jtdb_FIX 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jPCH1'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jPCH1' AND b.COLUMN_NAME = a.COLUMN_NAME
  )
UNION ALL
SELECT 'jtdb_FIX 有, jtdb 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jPCH1'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jPCH1' AND b.COLUMN_NAME = a.COLUMN_NAME
  );

-- 1.3 jMGUIAP 欄位比對
PRINT '--- jMGUIAP 欄位差異 ---';
SELECT 'jtdb 有, jtdb_FIX 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jMGUIAP'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jMGUIAP' AND b.COLUMN_NAME = a.COLUMN_NAME
  )
UNION ALL
SELECT 'jtdb_FIX 有, jtdb 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jMGUIAP'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jMGUIAP' AND b.COLUMN_NAME = a.COLUMN_NAME
  );

-- 1.4 jMGUIAPDETAIL 欄位比對
PRINT '--- jMGUIAPDETAIL 欄位差異 ---';
SELECT 'jtdb 有, jtdb_FIX 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jMGUIAPDETAIL'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jMGUIAPDETAIL' AND b.COLUMN_NAME = a.COLUMN_NAME
  )
UNION ALL
SELECT 'jtdb_FIX 有, jtdb 沒有' AS 差異類型, a.COLUMN_NAME, a.DATA_TYPE
FROM jtdb_FIX.INFORMATION_SCHEMA.COLUMNS a
WHERE a.TABLE_NAME = 'jMGUIAPDETAIL'
  AND NOT EXISTS (
    SELECT 1 FROM jtdb.INFORMATION_SCHEMA.COLUMNS b
    WHERE b.TABLE_NAME = 'jMGUIAPDETAIL' AND b.COLUMN_NAME = a.COLUMN_NAME
  );

-- =========================================
-- 步驟 2: 查看資料筆數差異
-- =========================================

PRINT '========================================';
PRINT '步驟 2: 資料筆數比對';
PRINT '========================================';

SELECT 'jOPCH' AS 表名,
       (SELECT COUNT(*) FROM jtdb.dbo.jOPCH) AS jtdb筆數,
       (SELECT COUNT(*) FROM jtdb_FIX.dbo.jOPCH) AS jtdb_FIX筆數;

SELECT 'jPCH1' AS 表名,
       (SELECT COUNT(*) FROM jtdb.dbo.jPCH1) AS jtdb筆數,
       (SELECT COUNT(*) FROM jtdb_FIX.dbo.jPCH1) AS jtdb_FIX筆數;

SELECT 'jMGUIAP' AS 表名,
       (SELECT COUNT(*) FROM jtdb.dbo.jMGUIAP) AS jtdb筆數,
       (SELECT COUNT(*) FROM jtdb_FIX.dbo.jMGUIAP) AS jtdb_FIX筆數;

SELECT 'jMGUIAPDETAIL' AS 表名,
       (SELECT COUNT(*) FROM jtdb.dbo.jMGUIAPDETAIL) AS jtdb筆數,
       (SELECT COUNT(*) FROM jtdb_FIX.dbo.jMGUIAPDETAIL) AS jtdb_FIX筆數;

-- =========================================
-- 步驟 3: 找出 jtdb 有但 jtdb_FIX 沒有的 jOPCH 資料 (新增的單據)
-- =========================================

PRINT '========================================';
PRINT '步驟 3: jtdb 新增的 jOPCH 資料 (需要補到 jtdb_FIX)';
PRINT '========================================';

SELECT a.jID, a.CardCode, a.CardName, a.DocDate, a.DocTotal, a.ApprovalStatus, a.CreateDate
FROM jtdb.dbo.jOPCH a
WHERE NOT EXISTS (
    SELECT 1 FROM jtdb_FIX.dbo.jOPCH b WHERE b.jID = a.jID
)
ORDER BY a.jID;

-- =========================================
-- 步驟 4: 找出 jtdb 有但 jtdb_FIX 沒有的 jPCH1 資料
-- =========================================

PRINT '========================================';
PRINT '步驟 4: jtdb 新增的 jPCH1 資料 (需要補到 jtdb_FIX)';
PRINT '========================================';

SELECT a.jID, a.LineNum, a.ItemCode, a.Dscription, a.LineTotal
FROM jtdb.dbo.jPCH1 a
WHERE NOT EXISTS (
    SELECT 1 FROM jtdb_FIX.dbo.jPCH1 b WHERE b.jID = a.jID AND b.LineNum = a.LineNum
)
ORDER BY a.jID, a.LineNum;

