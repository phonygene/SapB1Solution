-- =============================================
-- 修改 User 表：新增審核者權限欄位
-- 建立日期: 2025-11-05
-- 更新日期: 2025-11-05
-- 說明: 為使用者表新增審核者權限欄位
-- =============================================

USE jtdb
GO

-- 檢查並新增審核者權限欄位
IF NOT EXISTS (
    SELECT * FROM sys.columns
    WHERE object_id = OBJECT_ID(N'[dbo].[User]')
    AND name = 'Approver'
)
BEGIN
    ALTER TABLE [dbo].[User]
    ADD [Approver] TINYINT NOT NULL DEFAULT 0  -- 0=非審核者, 1=審核者

    PRINT 'Column Approver added to User successfully.'
END
ELSE
BEGIN
    PRINT 'Column Approver already exists in User.'
END
GO

-- 為測試目的，將特定使用者設定為審核者（請根據實際情況調整）
-- UPDATE [User] SET Approver = 1 WHERE id = 'admin'

PRINT ''
PRINT '===================================='
PRINT 'User 表審核者權限欄位建立完成！'
PRINT '===================================='
PRINT '提醒：請手動設定審核者'
PRINT '範例：UPDATE [User] SET Approver = 1 WHERE id = ''your_user_id'''
PRINT '===================================='
GO
