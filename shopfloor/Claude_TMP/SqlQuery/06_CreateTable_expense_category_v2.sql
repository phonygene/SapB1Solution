-- =============================================
-- 建立費用類別資料表（含稽核欄位）
-- 建立日期: 2025-11-05
-- 更新日期: 2025-11-07
-- 說明: 儲存費用類別與總帳科目對應表
--       包含 CreateBy/UpdateBy 稽核欄位
-- =============================================

USE jtdb
GO

-- =============================================
-- 建立 expense_category (費用類別)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[expense_category]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[expense_category] (
        [ID]           INT IDENTITY(1,1) PRIMARY KEY,
        [CategoryCode] NVARCHAR(20) NOT NULL UNIQUE,   -- 類別代碼
        [CategoryName] NVARCHAR(100) NOT NULL,         -- 類別名稱
        [AcctCode]     NVARCHAR(15) NOT NULL,          -- 對應的總帳科目
        [Active]       CHAR(1) NOT NULL DEFAULT 'Y',   -- Y=啟用, N=停用
        [CreateDate]   DATETIME DEFAULT GETDATE(),
        [CreateBy]     NVARCHAR(20) NULL,              -- 建立者
        [UpdateDate]   DATETIME NULL,
        [UpdateBy]     NVARCHAR(20) NULL,              -- 更新者

        CONSTRAINT CK_expense_category_active CHECK (Active IN ('Y', 'N'))
    )

    PRINT 'Table expense_category created successfully.'
END
ELSE
BEGIN
    PRINT 'Table expense_category already exists.'
END
GO

-- 建立索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_expense_category_active')
    CREATE INDEX IX_expense_category_active ON expense_category(Active)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'UX_expense_category_code')
    CREATE UNIQUE INDEX UX_expense_category_code ON expense_category(CategoryCode)
GO

-- 插入測試資料
IF NOT EXISTS (SELECT * FROM expense_category WHERE CategoryCode = 'TRAVEL')
BEGIN
    INSERT INTO expense_category (CategoryCode, CategoryName, AcctCode, CreateBy) VALUES
    ('TRAVEL', '差旅費', '6101', 'SYSTEM'),
    ('MEAL', '膳食費', '6102', 'SYSTEM'),
    ('OFFICE', '辦公用品', '6103', 'SYSTEM'),
    ('TELECOM', '電信費', '6201', 'SYSTEM'),
    ('RENTAL', '租金', '6301', 'SYSTEM'),
    ('UTILITY', '水電費', '6302', 'SYSTEM'),
    ('TRANSPORT', '交通費', '6401', 'SYSTEM'),
    ('TRAINING', '教育訓練費', '6501', 'SYSTEM'),
    ('REPAIR', '修繕費', '6601', 'SYSTEM'),
    ('MISC', '其他雜項', '6999', 'SYSTEM')

    PRINT '測試資料已插入 expense_category 表。'
END
ELSE
BEGIN
    PRINT '測試資料已存在，跳過插入。'
END
GO

PRINT ''
PRINT '===================================='
PRINT 'expense_category 資料表建立完成！'
PRINT '===================================='
GO
