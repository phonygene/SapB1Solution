-- =============================================
-- 建立收貨地址資料表
-- 建立日期: 2025-11-05
-- 說明: 儲存常用的收貨地址清單
-- =============================================

USE jtdb
GO

-- =============================================
-- 建立 addr (收貨地址)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[addr]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[addr] (
        [ID]         INT IDENTITY(1,1) PRIMARY KEY,
        [addrType]   CHAR(1) NOT NULL DEFAULT 'R',      -- D=交貨, R=收貨
        [addrName]   NVARCHAR(50) NOT NULL,             -- 地址名稱
        [address]    NVARCHAR(254) NOT NULL,            -- 地址
        [active]     CHAR(1) NOT NULL DEFAULT 'Y',      -- Y=啟用, N=停用
        [createDate] DATETIME DEFAULT GETDATE(),
        [updateDate] DATETIME DEFAULT GETDATE(),

        CONSTRAINT CK_addr_addrType CHECK (addrType IN ('D', 'R')),
        CONSTRAINT CK_addr_active CHECK (active IN ('Y', 'N'))
    )

    PRINT 'Table addr created successfully.'
END
ELSE
BEGIN
    PRINT 'Table addr already exists.'
END
GO

-- 建立索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_addr_addrType')
    CREATE INDEX IX_addr_addrType ON addr(addrType)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_addr_active')
    CREATE INDEX IX_addr_active ON addr(active)
GO

-- 插入測試資料
IF NOT EXISTS (SELECT * FROM addr WHERE addrName = '總公司')
BEGIN
    INSERT INTO addr (addrType, addrName, address) VALUES
    ('R', '總公司', '台北市信義區信義路五段7號'),
    ('R', '台中倉庫', '台中市西屯區台灣大道三段99號'),
    ('R', '高雄辦公室', '高雄市前鎮區成功二路88號')

    PRINT '測試資料已插入 addr 表。'
END
ELSE
BEGIN
    PRINT '測試資料已存在，跳過插入。'
END
GO

PRINT ''
PRINT '===================================='
PRINT 'addr 資料表建立完成！'
PRINT '===================================='
GO
