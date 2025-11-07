-- =============================================
-- 修改 jOPCH 表：新增審核相關欄位
-- 建立日期: 2025-11-05
-- 更新日期: 2025-11-05
-- 說明: 為費用申請單表頭新增審核意見、審核日期時間、審核人欄位
-- =============================================

USE jtdb
GO

-- 檢查並新增審核意見欄位
IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jOPCH]')
    AND name = 'ApprovalComments'
)
BEGIN
    ALTER TABLE [dbo].[jOPCH]
    ADD [ApprovalComments] NVARCHAR(500) NULL  -- 審核意見

    PRINT 'Column ApprovalComments added to jOPCH successfully.'
END
ELSE
BEGIN
    PRINT 'Column ApprovalComments already exists in jOPCH.'
END
GO

-- 檢查並新增審核日期欄位
IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jOPCH]')
    AND name = 'ApprovalDate'
)
BEGIN
    ALTER TABLE [dbo].[jOPCH]
    ADD [ApprovalDate] DATETIME NULL  -- 審核日期

    PRINT 'Column ApprovalDate added to jOPCH successfully.'
END
ELSE
BEGIN
    PRINT 'Column ApprovalDate already exists in jOPCH.'
END
GO

-- 檢查並新增審核時間欄位
IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jOPCH]')
    AND name = 'ApprovalTime'
)
BEGIN
    ALTER TABLE [dbo].[jOPCH]
    ADD [ApprovalTime] CHAR(8) NULL  -- 審核時間 (HH:mm:ss)

    PRINT 'Column ApprovalTime added to jOPCH successfully.'
END
ELSE
BEGIN
    PRINT 'Column ApprovalTime already exists in jOPCH.'
END
GO

-- 檢查並新增審核人欄位
IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jOPCH]')
    AND name = 'ApprovedBy'
)
BEGIN
    ALTER TABLE [dbo].[jOPCH]
    ADD [ApprovedBy] NVARCHAR(20) NULL  -- 審核人 (User ID)

    PRINT 'Column ApprovedBy added to jOPCH successfully.'
END
ELSE
BEGIN
    PRINT 'Column ApprovedBy already exists in jOPCH.'
END
GO

PRINT ''
PRINT '===================================='
PRINT 'jOPCH 表審核相關欄位建立完成！'
PRINT '===================================='
GO
