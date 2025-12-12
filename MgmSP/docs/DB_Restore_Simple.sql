-- =============================================================================
-- 資料庫修復腳本 (簡化版)
-- 目的：還原 jMGUIAP 和 jMGUIAPDETAIL，並將新增單據狀態改回草稿
-- 日期：2025/12/12
-- =============================================================================

-- =========================================
-- 步驟 0: 確認差異 (先執行這段看結果)
-- =========================================

PRINT '========================================';
PRINT '步驟 0: 確認資料差異';
PRINT '========================================';

-- 查看 jtdb 比 jtdb_FIX 多的 jOPCH (這些是 12/10 之後新增的單據)
SELECT a.jID, a.CardCode, a.CardName, a.DocDate, a.DocTotal, a.ApprovalStatus, a.CreateDate
FROM jtdb.dbo.jOPCH a
WHERE NOT EXISTS (
    SELECT 1 FROM jtdb_FIX.dbo.jOPCH b WHERE b.jID = a.jID
)
ORDER BY a.jID;

-- 查看筆數
SELECT 'jMGUIAP' AS 表名,
       (SELECT COUNT(*) FROM jtdb.dbo.jMGUIAP) AS jtdb目前筆數,
       (SELECT COUNT(*) FROM jtdb_FIX.dbo.jMGUIAP) AS jtdb_FIX備份筆數;

SELECT 'jMGUIAPDETAIL' AS 表名,
       (SELECT COUNT(*) FROM jtdb.dbo.jMGUIAPDETAIL) AS jtdb目前筆數,
       (SELECT COUNT(*) FROM jtdb_FIX.dbo.jMGUIAPDETAIL) AS jtdb_FIX備份筆數;

-- =========================================
-- 步驟 1: 開始交易 (確認步驟0結果正確後再執行)
-- =========================================

BEGIN TRANSACTION;

PRINT '========================================';
PRINT '步驟 1: 清空並還原 jMGUIAP 和 jMGUIAPDETAIL';
PRINT '========================================';

-- 1.1 清空 jtdb 的 jMGUIAPDETAIL (先清子表)
DELETE FROM jtdb.dbo.jMGUIAPDETAIL;
PRINT '已清空 jtdb.jMGUIAPDETAIL: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 1.2 清空 jtdb 的 jMGUIAP (再清主表)
DELETE FROM jtdb.dbo.jMGUIAP;
PRINT '已清空 jtdb.jMGUIAP: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 1.3 從 jtdb_FIX 複製 jMGUIAP 到 jtdb
SET IDENTITY_INSERT jtdb.dbo.jMGUIAP ON;

INSERT INTO jtdb.dbo.jMGUIAP (
    ID, jID, DocEntry, DocNum,
    U_ZESSION, U_ZESSION2, U_LIFNR, U_BUKRS, U_GJAHR,
    U_BESSION, U_BLDAT, U_ZFBDT, U_ZLSCH, U_ZLSPR,
    U_ZFORM, U_ZFORM2, U_WAESSION, U_WAESSION2, U_WESSION, U_WESSION2,
    DocTotal, CreateDate, CreateBy
)
SELECT
    ID, jID, DocEntry, DocNum,
    U_ZESSION, U_ZESSION2, U_LIFNR, U_BUKRS, U_GJAHR,
    U_BESSION, U_BLDAT, U_ZFBDT, U_ZLSCH, U_ZLSPR,
    U_ZFORM, U_ZFORM2, U_WAESSION, U_WAESSION2, U_WESSION, U_WESSION2,
    DocTotal, CreateDate, CreateBy
FROM jtdb_FIX.dbo.jMGUIAP;

SET IDENTITY_INSERT jtdb.dbo.jMGUIAP OFF;
PRINT '已從 jtdb_FIX 複製 jMGUIAP: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 1.4 從 jtdb_FIX 複製 jMGUIAPDETAIL 到 jtdb
INSERT INTO jtdb.dbo.jMGUIAPDETAIL (
    jID, LineNum, DocEntry,
    U_LIFNR, U_HWBAS, U_HWSTE, U_MESSION, U_BESSION,
    U_INVNO, U_INDATE, U_ZFORM, U_ZFORM2,
    CreateDate, CreateBy
)
SELECT
    jID, LineNum, DocEntry,
    U_LIFNR, U_HWBAS, U_HWSTE, U_MESSION, U_BESSION,
    U_INVNO, U_INDATE, U_ZFORM, U_ZFORM2,
    CreateDate, CreateBy
FROM jtdb_FIX.dbo.jMGUIAPDETAIL;

PRINT '已從 jtdb_FIX 複製 jMGUIAPDETAIL: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- =========================================
-- 步驟 2: 將新增單據的狀態改回草稿
-- =========================================

PRINT '========================================';
PRINT '步驟 2: 將 12/10 後新增的單據狀態改回草稿';
PRINT '========================================';

-- 找出 jtdb 比 jtdb_FIX 多的 jID，將狀態改為草稿 (Draft)
UPDATE jtdb.dbo.jOPCH
SET ApprovalStatus = 'Draft',
    ApprovedBy = NULL,
    ApprovedDate = NULL,
    ApprovalDate = NULL,
    ApprovalTime = NULL,
    ApprovalComments = NULL,
    B1PostStatus = NULL,
    B1PostDate = NULL,
    B1ErrMsg = NULL,
    DocEntry = NULL,
    DocNum = NULL
WHERE jID NOT IN (SELECT jID FROM jtdb_FIX.dbo.jOPCH);

PRINT '已將新增單據狀態改回草稿: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 同時清空這些單據在 jPCH1 的 DocEntry/DocNum
UPDATE jtdb.dbo.jPCH1
SET DocEntry = NULL,
    DocNum = NULL
WHERE jID NOT IN (SELECT jID FROM jtdb_FIX.dbo.jOPCH);

PRINT '已清空新增單據的 jPCH1.DocEntry/DocNum: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- =========================================
-- 步驟 3: 驗證結果
-- =========================================

PRINT '========================================';
PRINT '步驟 3: 驗證結果';
PRINT '========================================';

-- 確認還原後的筆數
SELECT 'jMGUIAP 還原後' AS Info, COUNT(*) AS 筆數 FROM jtdb.dbo.jMGUIAP;
SELECT 'jMGUIAPDETAIL 還原後' AS Info, COUNT(*) AS 筆數 FROM jtdb.dbo.jMGUIAPDETAIL;

-- 顯示需要 USER 補憑證明細的單據
SELECT '需要補憑證明細的單據' AS Info,
       a.jID, a.CardCode, a.CardName, a.DocDate, a.DocTotal, a.ApprovalStatus
FROM jtdb.dbo.jOPCH a
WHERE NOT EXISTS (SELECT 1 FROM jtdb_FIX.dbo.jOPCH b WHERE b.jID = a.jID)
ORDER BY a.jID;

PRINT '========================================';
PRINT '修復完成！請檢查上方結果';
PRINT '如果沒問題，請執行 COMMIT';
PRINT '如果有問題，請執行 ROLLBACK';
PRINT '========================================';

-- 確認無誤後執行：
-- COMMIT;

-- 如果有問題執行：
-- ROLLBACK;
