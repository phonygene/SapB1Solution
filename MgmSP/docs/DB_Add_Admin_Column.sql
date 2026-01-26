-- =============================================
-- Admin Bypass for Maintenance Mode
-- 新增 admin 欄位到 User 表
-- 執行時間: 請在 jtdb 資料庫執行
-- =============================================

-- 檢查欄位是否存在，不存在則新增
IF NOT EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[User]') AND name = 'admin')
BEGIN
    ALTER TABLE [User] ADD admin BIT DEFAULT 0;
    PRINT '已新增 admin 欄位到 User 表';
END
ELSE
BEGIN
    PRINT 'admin 欄位已存在，跳過新增';
END
GO

-- 確保所有現有使用者預設為非 admin
UPDATE [User] SET admin = 0 WHERE admin IS NULL;
GO

-- 範例：將指定使用者設為 admin (請取消註解並修改 ID)
-- UPDATE [User] SET admin = 1 WHERE id = 'your_admin_id';
-- GO

PRINT '完成！請使用 UPDATE 語句設定需要的 admin 使用者 (admin = 1)';

