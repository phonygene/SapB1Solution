-- =============================================================================
-- 資料庫修復腳本 (乾淨版)
-- 步驟：先徹底刪除表，再從 jtdb_FIX 複製
-- 日期：2025/12/12
-- =============================================================================

USE jtdb;
GO

-- =========================================
-- 步驟 1: 先刪除外鍵約束和表
-- =========================================

PRINT '========================================';
PRINT '步驟 1: 刪除現有的表和約束';
PRINT '========================================';

-- 1.1 刪除 jMGUIAPDETAIL (子表先刪)
IF OBJECT_ID('dbo.jMGUIAPDETAIL', 'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.jMGUIAPDETAIL;
    PRINT '已刪除 jMGUIAPDETAIL';
END

-- 1.2 刪除 jMGUIAP 的外鍵約束 (如果有的話)
IF EXISTS (SELECT * FROM sys.foreign_keys WHERE name = 'FK_jMGUIAP_jOPCH')
BEGIN
    ALTER TABLE dbo.jMGUIAP DROP CONSTRAINT FK_jMGUIAP_jOPCH;
    PRINT '已刪除外鍵 FK_jMGUIAP_jOPCH';
END

-- 1.3 刪除 jMGUIAP
IF OBJECT_ID('dbo.jMGUIAP', 'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.jMGUIAP;
    PRINT '已刪除 jMGUIAP';
END

PRINT '表刪除完成';
GO

-- =========================================
-- 步驟 2: 從 jtdb_FIX 複製表結構和資料
-- (請在 SSMS 中對 jtdb_FIX 的這兩個表產生 Script，
--  選擇 Schema and Data，然後貼在這裡執行)
-- =========================================

PRINT '========================================';
PRINT '步驟 2: 請從 jtdb_FIX 產生 Script 並執行';
PRINT '========================================';

-- 在 SSMS 中:
-- 1. 展開 jtdb_FIX > Tables
-- 2. 右鍵 jMGUIAP > Script Table as > CREATE To > New Query Window
-- 3. 右鍵 jMGUIAP > Script Table as > INSERT To > (加到同一個 Query)
-- 4. 對 jMGUIAPDETAIL 做同樣的事
-- 5. 把產生的 Script 貼到這裡執行

-- =========================================
-- 步驟 3: 執行完步驟 2 後，再執行這段更新狀態
-- =========================================

/*
PRINT '========================================';
PRINT '步驟 3: 將 12/10 後新增的單據狀態改回草稿';
PRINT '========================================';

-- 找出 jtdb 比 jtdb_FIX 多的 jID，將狀態改為草稿 (Draft)
UPDATE dbo.jOPCH
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
UPDATE dbo.jPCH1
SET DocEntry = NULL,
    DocNum = NULL
WHERE jID NOT IN (SELECT jID FROM jtdb_FIX.dbo.jOPCH);

PRINT '已清空新增單據的 jPCH1.DocEntry/DocNum: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 顯示需要 USER 補憑證明細的單據
SELECT '需要補憑證明細的單據' AS Info,
       a.jID, a.CardCode, a.CardName, a.DocDate, a.DocTotal, a.ApprovalStatus
FROM dbo.jOPCH a
WHERE NOT EXISTS (SELECT 1 FROM jtdb_FIX.dbo.jOPCH b WHERE b.jID = a.jID)
ORDER BY a.jID;
*/
