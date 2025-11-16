-- =============================================
-- 為協作開發的資料表加入稽核欄位
-- 建立日期: 2025-11-07
-- 說明: 統一為協作開發的表加入 CreateBy/UpdateBy 欄位
--       命名規範: CreateDate, CreateBy, UpdateDate, UpdateBy
-- =============================================

USE jtdb
GO

PRINT '=========================================='
PRINT '開始新增稽核欄位到協作開發的資料表'
PRINT '=========================================='
PRINT ''

-- =============================================
-- 1. addr 表
-- =============================================
PRINT '--- 處理 addr 表 ---'

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[addr]')
    AND name = 'CreateBy'
)
BEGIN
    ALTER TABLE [dbo].[addr]
    ADD [CreateBy] NVARCHAR(20) NULL

    PRINT 'Column CreateBy added to addr.'
END
ELSE
BEGIN
    PRINT 'Column CreateBy already exists in addr.'
END
GO

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[addr]')
    AND name = 'UpdateBy'
)
BEGIN
    ALTER TABLE [dbo].[addr]
    ADD [UpdateBy] NVARCHAR(20) NULL

    PRINT 'Column UpdateBy added to addr.'
END
ELSE
BEGIN
    PRINT 'Column UpdateBy already exists in addr.'
END
GO

PRINT ''

-- =============================================
-- 2. expense_category 表
-- =============================================
PRINT '--- 處理 expense_category 表 ---'

-- 檢查表是否存在
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[expense_category]') AND type in (N'U'))
BEGIN
    IF NOT EXISTS (
        SELECT * FROM sys.columns
        WHERE object_id = OBJECT_ID(N'[dbo].[expense_category]')
        AND name = 'CreateBy'
    )
    BEGIN
        ALTER TABLE [dbo].[expense_category]
        ADD [CreateBy] NVARCHAR(20) NULL

        PRINT 'Column CreateBy added to expense_category.'
    END
    ELSE
    BEGIN
        PRINT 'Column CreateBy already exists in expense_category.'
    END

    IF NOT EXISTS (
        SELECT * FROM sys.columns
        WHERE object_id = OBJECT_ID(N'[dbo].[expense_category]')
        AND name = 'UpdateBy'
    )
    BEGIN
        ALTER TABLE [dbo].[expense_category]
        ADD [UpdateBy] NVARCHAR(20) NULL

        PRINT 'Column UpdateBy added to expense_category.'
    END
    ELSE
    BEGIN
        PRINT 'Column UpdateBy already exists in expense_category.'
    END
END
ELSE
BEGIN
    PRINT 'Table expense_category does not exist yet. Will be handled in creation script.'
END
GO

PRINT ''

-- =============================================
-- 3. jPCH1 表 (費用明細)
-- =============================================
PRINT '--- 處理 jPCH1 表 ---'

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jPCH1]')
    AND name = 'CreateDate'
)
BEGIN
    ALTER TABLE [dbo].[jPCH1]
    ADD [CreateDate] DATETIME DEFAULT GETDATE()

    PRINT 'Column CreateDate added to jPCH1.'
END
ELSE
BEGIN
    PRINT 'Column CreateDate already exists in jPCH1.'
END
GO

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jPCH1]')
    AND name = 'CreateBy'
)
BEGIN
    ALTER TABLE [dbo].[jPCH1]
    ADD [CreateBy] NVARCHAR(20) NULL

    PRINT 'Column CreateBy added to jPCH1.'
END
ELSE
BEGIN
    PRINT 'Column CreateBy already exists in jPCH1.'
END
GO

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jPCH1]')
    AND name = 'UpdateDate'
)
BEGIN
    ALTER TABLE [dbo].[jPCH1]
    ADD [UpdateDate] DATETIME NULL

    PRINT 'Column UpdateDate added to jPCH1.'
END
ELSE
BEGIN
    PRINT 'Column UpdateDate already exists in jPCH1.'
END
GO

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jPCH1]')
    AND name = 'UpdateBy'
)
BEGIN
    ALTER TABLE [dbo].[jPCH1]
    ADD [UpdateBy] NVARCHAR(20) NULL

    PRINT 'Column UpdateBy added to jPCH1.'
END
ELSE
BEGIN
    PRINT 'Column UpdateBy already exists in jPCH1.'
END
GO

PRINT ''

-- =============================================
-- 4. jMGUIAPDetail 表 (MDR 發票明細)
-- =============================================
PRINT '--- 處理 jMGUIAPDetail 表 ---'

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAPDetail]')
    AND name = 'CreateDate'
)
BEGIN
    ALTER TABLE [dbo].[jMGUIAPDetail]
    ADD [CreateDate] DATETIME DEFAULT GETDATE()

    PRINT 'Column CreateDate added to jMGUIAPDetail.'
END
ELSE
BEGIN
    PRINT 'Column CreateDate already exists in jMGUIAPDetail.'
END
GO

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAPDetail]')
    AND name = 'CreateBy'
)
BEGIN
    ALTER TABLE [dbo].[jMGUIAPDetail]
    ADD [CreateBy] NVARCHAR(20) NULL

    PRINT 'Column CreateBy added to jMGUIAPDetail.'
END
ELSE
BEGIN
    PRINT 'Column CreateBy already exists in jMGUIAPDetail.'
END
GO

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAPDetail]')
    AND name = 'UpdateDate'
)
BEGIN
    ALTER TABLE [dbo].[jMGUIAPDetail]
    ADD [UpdateDate] DATETIME NULL

    PRINT 'Column UpdateDate added to jMGUIAPDetail.'
END
ELSE
BEGIN
    PRINT 'Column UpdateDate already exists in jMGUIAPDetail.'
END
GO

IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAPDetail]')
    AND name = 'UpdateBy'
)
BEGIN
    ALTER TABLE [dbo].[jMGUIAPDetail]
    ADD [UpdateBy] NVARCHAR(20) NULL

    PRINT 'Column UpdateBy added to jMGUIAPDetail.'
END
ELSE
BEGIN
    PRINT 'Column UpdateBy already exists in jMGUIAPDetail.'
END
GO

PRINT ''
PRINT '=========================================='
PRINT '稽核欄位新增完成！'
PRINT '=========================================='
PRINT ''
PRINT '摘要：'
PRINT '- addr: CreateBy, UpdateBy'
PRINT '- expense_category: CreateBy, UpdateBy (如果表已存在)'
PRINT '- jPCH1: CreateDate, CreateBy, UpdateDate, UpdateBy'
PRINT '- jMGUIAPDetail: CreateDate, CreateBy, UpdateDate, UpdateBy'
PRINT ''
PRINT '注意：jOPCH 和 jMGUIAP 已有完整的稽核欄位，不需修改'
PRINT '=========================================='
GO
