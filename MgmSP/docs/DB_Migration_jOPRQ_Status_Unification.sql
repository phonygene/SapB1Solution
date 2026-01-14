-- =====================================================
-- jOPRQ 狀態欄位統一遷移腳本 (完整版)
-- 目的：將 jOPRQ 的狀態欄位格式改為與 jOPCH 一致
-- 日期：2026-01-14
-- =====================================================
--
-- 變更說明：
--   1. ApprovalStatus 格式統一：
--      'Pending'  → 'W' (Wait for approval)
--      'Approved' → 'A' (Approved)
--      'Rejected' → 'R' (Rejected)
--
--   2. 新增缺少的 SAP 整合欄位（與 jOPCH 對齊）：
--      B1PostStatus NVARCHAR(1)  - SAP 過帳狀態 (N/Y/E)
--      B1PostDate   DATETIME     - SAP 過帳日期
--      B1ErrMsg     NVARCHAR(500)- SAP 錯誤訊息
--      DocEntry     INT          - SAP 文件 Entry
--      DocNum       INT          - SAP 單據號碼
--
-- =====================================================

USE jtdb;
GO

-- =====================================================
-- 第一部分：新增缺少的欄位
-- =====================================================

PRINT '========================================';
PRINT '第一部分：新增缺少的欄位';
PRINT '========================================';

-- 1. 檢查並新增 B1PostStatus 欄位
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_NAME = 'jOPRQ' AND COLUMN_NAME = 'B1PostStatus')
BEGIN
    ALTER TABLE jOPRQ ADD B1PostStatus NVARCHAR(1) NULL;
    PRINT '✓ 已新增欄位: B1PostStatus NVARCHAR(1)';
END
ELSE
    PRINT '- 欄位已存在: B1PostStatus';
GO

-- 2. 檢查並新增 B1PostDate 欄位
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_NAME = 'jOPRQ' AND COLUMN_NAME = 'B1PostDate')
BEGIN
    ALTER TABLE jOPRQ ADD B1PostDate DATETIME NULL;
    PRINT '✓ 已新增欄位: B1PostDate DATETIME';
END
ELSE
    PRINT '- 欄位已存在: B1PostDate';
GO

-- 3. 檢查並新增 B1ErrMsg 欄位
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_NAME = 'jOPRQ' AND COLUMN_NAME = 'B1ErrMsg')
BEGIN
    ALTER TABLE jOPRQ ADD B1ErrMsg NVARCHAR(500) NULL;
    PRINT '✓ 已新增欄位: B1ErrMsg NVARCHAR(500)';
END
ELSE
    PRINT '- 欄位已存在: B1ErrMsg';
GO

-- 4. 檢查並新增 DocEntry 欄位（SAP 文件 Entry）
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_NAME = 'jOPRQ' AND COLUMN_NAME = 'DocEntry')
BEGIN
    ALTER TABLE jOPRQ ADD DocEntry INT NULL;
    PRINT '✓ 已新增欄位: DocEntry INT';
END
ELSE
    PRINT '- 欄位已存在: DocEntry';
GO

-- 5. 檢查並新增 DocNum 欄位（SAP 單據號碼）
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS
               WHERE TABLE_NAME = 'jOPRQ' AND COLUMN_NAME = 'DocNum')
BEGIN
    ALTER TABLE jOPRQ ADD DocNum INT NULL;
    PRINT '✓ 已新增欄位: DocNum INT';
END
ELSE
    PRINT '- 欄位已存在: DocNum';
GO

PRINT '';
PRINT '欄位新增完成';
PRINT '';

-- =====================================================
-- 第二部分：遷移 ApprovalStatus 值
-- =====================================================

PRINT '========================================';
PRINT '第二部分：遷移 ApprovalStatus 值';
PRINT '========================================';

-- 6. 先備份目前的狀態分布（僅供參考）
PRINT '目前狀態分布：';
SELECT ApprovalStatus, B1PostStatus, COUNT(*) AS RecordCount
FROM jOPRQ
GROUP BY ApprovalStatus, B1PostStatus
ORDER BY ApprovalStatus;
GO

-- 7. 開始交易
BEGIN TRANSACTION;

-- 8. 更新 ApprovalStatus 值（Pending → W, Approved → A, Rejected → R）
UPDATE jOPRQ SET ApprovalStatus = 'W' WHERE ApprovalStatus = 'Pending';
PRINT '✓ Pending → W: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

UPDATE jOPRQ SET ApprovalStatus = 'A' WHERE ApprovalStatus = 'Approved';
PRINT '✓ Approved → A: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

UPDATE jOPRQ SET ApprovalStatus = 'R' WHERE ApprovalStatus = 'Rejected';
PRINT '✓ Rejected → R: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 9. 設定 B1PostStatus 預設值
-- 對於已核准且有 DocNum 的單據，設為 Y（已過帳）
UPDATE jOPRQ SET B1PostStatus = 'Y'
WHERE ApprovalStatus = 'A' AND DocNum IS NOT NULL AND B1PostStatus IS NULL;
PRINT '✓ 已核准且有 DocNum → B1PostStatus=Y: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 對於已核准但無 DocNum 的單據，設為 N（待過帳）
UPDATE jOPRQ SET B1PostStatus = 'N'
WHERE ApprovalStatus = 'A' AND DocNum IS NULL AND B1PostStatus IS NULL;
PRINT '✓ 已核准但無 DocNum → B1PostStatus=N: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 對於其他單據（待審核/已退回），設為 N（未過帳）
UPDATE jOPRQ SET B1PostStatus = 'N' WHERE B1PostStatus IS NULL;
PRINT '✓ 其他 → B1PostStatus=N: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 10. 驗證結果
PRINT '';
PRINT '遷移後狀態分布：';
SELECT ApprovalStatus, B1PostStatus, COUNT(*) AS RecordCount
FROM jOPRQ
GROUP BY ApprovalStatus, B1PostStatus
ORDER BY ApprovalStatus;

-- =====================================================
-- 11. 確認無誤後提交
-- =====================================================
PRINT '';
PRINT '========================================';
PRINT '請確認上方結果是否正確';
PRINT '========================================';
PRINT '- 若正確，請執行: COMMIT';
PRINT '- 若有問題，請執行: ROLLBACK';
PRINT '========================================';

-- COMMIT;
-- ROLLBACK;

-- =====================================================
-- 統一後的狀態欄位對照表（jOPCH 與 jOPRQ）：
-- =====================================================
-- | 欄位           | 型別          | 說明              | 選項值        |
-- |----------------|---------------|-------------------|---------------|
-- | ApprovalStatus | NVARCHAR(1)   | 審核狀態          | W/A/R         |
-- | B1PostStatus   | NVARCHAR(1)   | SAP 過帳狀態      | N/Y/E         |
-- | B1PostDate     | DATETIME      | SAP 過帳日期      | 日期時間      |
-- | B1ErrMsg       | NVARCHAR(500) | SAP 錯誤訊息      | 文字          |
-- | DocEntry       | INT           | SAP 文件 Entry    | 整數          |
-- | DocNum         | INT           | SAP 單據號碼      | 整數          |
-- =====================================================
--
-- ApprovalStatus 選項說明：
--   W = Wait for approval (待審核)
--   A = Approved (已核准)
--   R = Rejected (已退回)
--
-- B1PostStatus 選項說明：
--   N = Not posted (未過帳/待過帳)
--   Y = Posted successfully (已過帳成功)
--   E = Error (過帳失敗)
-- =====================================================
