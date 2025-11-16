-- =============================================
-- 建立本地 MDR 測試資料庫
-- 建立日期: 2025-11-16
-- 說明: 在本地環境建立 MDR 資料庫,用於測試營業稅發票資料同步功能
--       模擬正式環境的 MDR 資料庫 (192.168.1.219)
-- =============================================

-- =============================================
-- 第一部分: 建立資料庫
-- =============================================
USE master
GO

-- 檢查資料庫是否存在,若存在則不建立
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'MDR')
BEGIN
    CREATE DATABASE [MDR]
    PRINT 'Database MDR created successfully.'
END
ELSE
BEGIN
    PRINT 'Database MDR already exists.'
END
GO

USE [MDR]
GO

PRINT ''
PRINT '=========================================='
PRINT '開始建立 MDR 資料表'
PRINT '=========================================='
PRINT ''

-- =============================================
-- 第二部分: 建立 MGUIAP_Import (表頭)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MGUIAP_Import]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[MGUIAP_Import] (
        -- 主鍵
        [ID]            INT IDENTITY(1,1) PRIMARY KEY,

        -- 來源系統資訊
        [jID]           INT NULL,                       -- Jet Shopfloor 平台 jID
        [DocEntry]      INT NULL,                       -- SAP B1 AP Invoice DocEntry
        [DocNum]        INT NULL,                       -- SAP B1 AP Invoice DocNum

        -- 供應商資訊
        [U_LIFNR]       NVARCHAR(50) NULL,             -- 供應商代碼 (CardCode)
        [U_STCEG]       NVARCHAR(20) NULL,             -- 供應商統編

        -- 發票資訊
        [U_XBLNR]       NVARCHAR(50) NULL,             -- 發票號碼 (廠商發票號)
        [U_BLDAT]       DATE NULL,                     -- 發票日期
        [U_VATDATE]     DATE NULL,                     -- 營業稅日期

        -- 金額資訊
        [U_HWBAS]       DECIMAL(19,6) DEFAULT 0,       -- 未稅金額 (本幣)
        [U_HWSTE]       DECIMAL(19,6) DEFAULT 0,       -- 營業稅額 (本幣)
        [TotalAmount]   DECIMAL(19,6) DEFAULT 0,       -- 含稅總額 (本幣)

        -- 分類代碼
        [U_ZFORM_CODE]  NVARCHAR(10) NULL,             -- 發票格式代碼
        [U_TAX_TYPE]    NVARCHAR(10) NULL,             -- 稅別
        [U_CUS_TYPE]    NVARCHAR(10) NULL,             -- 客戶類別
        [U_AM_TYPE]     NVARCHAR(10) NULL,             -- 金額類別
        [U_BUKRS]       NVARCHAR(10) NULL,             -- 公司代碼
        [U_MWSKZ]       NVARCHAR(10) NULL,             -- 稅碼
        [U_VATCODE]     NVARCHAR(20) NULL,             -- 營業稅碼

        -- SAP 回寫資訊
        [U_BELNR]       NVARCHAR(50) NULL,             -- SAP 憑證編號 (回寫)
        [PostStatus]    CHAR(1) DEFAULT 'N',           -- 過帳狀態 (N=未過帳, Y=已過帳, E=錯誤)
        [PostDate]      DATETIME NULL,                 -- 過帳日期時間
        [ErrorMsg]      NVARCHAR(500) NULL,            -- 錯誤訊息

        -- 稽核欄位
        [CreateDate]    DATETIME DEFAULT GETDATE(),
        [CreateBy]      NVARCHAR(20) NULL,
        [UpdateDate]    DATETIME NULL,
        [UpdateBy]      NVARCHAR(20) NULL,

        CONSTRAINT CK_MGUIAP_Import_PostStatus CHECK (PostStatus IN ('N', 'Y', 'E'))
    )

    -- 建立索引
    CREATE INDEX IX_MGUIAP_Import_jID ON MGUIAP_Import(jID)
    CREATE INDEX IX_MGUIAP_Import_DocEntry ON MGUIAP_Import(DocEntry)
    CREATE INDEX IX_MGUIAP_Import_PostStatus ON MGUIAP_Import(PostStatus)
    CREATE INDEX IX_MGUIAP_Import_LIFNR ON MGUIAP_Import(U_LIFNR)
    CREATE INDEX IX_MGUIAP_Import_XBLNR ON MGUIAP_Import(U_XBLNR)

    PRINT 'Table MGUIAP_Import created successfully.'
END
ELSE
BEGIN
    PRINT 'Table MGUIAP_Import already exists.'
END
GO

-- =============================================
-- 第三部分: 建立 MGUIAPDetail_Import (明細)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[MGUIAPDetail_Import]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[MGUIAPDetail_Import] (
        -- 主鍵
        [ID]            INT IDENTITY(1,1) PRIMARY KEY,
        [HeaderID]      INT NOT NULL,                  -- 關聯到 MGUIAP_Import.ID
        [LineNum]       INT NOT NULL,                  -- 列號

        -- 來源系統資訊
        [jID]           INT NULL,                      -- Jet Shopfloor 平台 jID
        [DocEntry]      INT NULL,                      -- SAP B1 AP Invoice DocEntry

        -- 發票明細資訊
        [U_LIFNR]       NVARCHAR(50) NULL,            -- 供應商代碼
        [U_STCEG]       NVARCHAR(20) NULL,            -- 供應商統編
        [U_XBLNR]       NVARCHAR(50) NULL,            -- 發票號碼
        [U_ZFORM_CODE]  NVARCHAR(10) NULL,            -- 發票格式代碼
        [U_BLDAT]       DATE NULL,                    -- 發票日期
        [U_VATDATE]     DATE NULL,                    -- 營業稅日期

        -- 金額資訊
        [U_HWBAS]       DECIMAL(19,6) DEFAULT 0,      -- 未稅金額
        [U_HWSTE]       DECIMAL(19,6) DEFAULT 0,      -- 營業稅額

        -- 分類代碼
        [U_TAX_TYPE]    NVARCHAR(10) NULL,            -- 稅別
        [U_CUS_TYPE]    NVARCHAR(10) NULL,            -- 客戶類別
        [U_AM_TYPE]     NVARCHAR(10) NULL,            -- 金額類別
        [U_VATCODE]     NVARCHAR(20) NULL,            -- 營業稅碼
        [U_BUKRS]       NVARCHAR(10) NULL,            -- 公司代碼
        [U_MWSKZ]       NVARCHAR(10) NULL,            -- 稅碼

        -- SAP 回寫資訊
        [U_BELNR]       NVARCHAR(50) NULL,            -- SAP 憑證編號

        -- 固定資產相關欄位
        [U_FA_DESC]     NVARCHAR(200) NULL,           -- 固定資產品名
        [U_FA_QTY]      DECIMAL(19,6) DEFAULT 0,      -- 固定資產數量
        [U_FA_USE]      NVARCHAR(200) NULL,           -- 固定資產用途

        -- 彙總標記
        [U_GatherMark]  NVARCHAR(1) DEFAULT 'N',      -- 彙總註記 (Y=彙總, N=非彙總)
        [U_ConsolidQty] DECIMAL(19,6) DEFAULT 0,      -- 彙總數量

        -- 稽核欄位
        [CreateDate]    DATETIME DEFAULT GETDATE(),
        [CreateBy]      NVARCHAR(20) NULL,
        [UpdateDate]    DATETIME NULL,
        [UpdateBy]      NVARCHAR(20) NULL,

        CONSTRAINT FK_MGUIAPDetail_Header FOREIGN KEY (HeaderID)
            REFERENCES MGUIAP_Import(ID) ON DELETE CASCADE,
        CONSTRAINT CK_MGUIAPDetail_GatherMark CHECK (U_GatherMark IN ('Y', 'N'))
    )

    -- 建立索引
    CREATE INDEX IX_MGUIAPDetail_Import_HeaderID ON MGUIAPDetail_Import(HeaderID)
    CREATE INDEX IX_MGUIAPDetail_Import_jID ON MGUIAPDetail_Import(jID)
    CREATE INDEX IX_MGUIAPDetail_Import_DocEntry ON MGUIAPDetail_Import(DocEntry)
    CREATE INDEX IX_MGUIAPDetail_Import_LineNum ON MGUIAPDetail_Import(LineNum)

    PRINT 'Table MGUIAPDetail_Import created successfully.'
END
ELSE
BEGIN
    PRINT 'Table MGUIAPDetail_Import already exists.'
END
GO

-- =============================================
-- 第四部分: 插入測試資料
-- =============================================
PRINT ''
PRINT '--- 插入測試資料 ---'
PRINT ''

-- 插入測試表頭資料
IF NOT EXISTS (SELECT * FROM MGUIAP_Import WHERE jID = 9999)
BEGIN
    INSERT INTO MGUIAP_Import (
        jID, DocEntry, DocNum,
        U_LIFNR, U_STCEG,
        U_XBLNR, U_BLDAT, U_VATDATE,
        U_HWBAS, U_HWSTE, TotalAmount,
        U_ZFORM_CODE, U_TAX_TYPE, U_CUS_TYPE, U_AM_TYPE,
        U_BUKRS, U_MWSKZ, U_VATCODE,
        PostStatus, CreateBy
    ) VALUES (
        9999,           -- jID (測試用)
        NULL,           -- DocEntry (尚未產生 AP Invoice)
        NULL,           -- DocNum
        'V00001',       -- 供應商代碼
        '12345678',     -- 供應商統編
        'TEST-INV-001', -- 發票號碼
        '2025-11-16',   -- 發票日期
        '2025-11-16',   -- 營業稅日期
        10000.00,       -- 未稅金額
        500.00,         -- 營業稅額
        10500.00,       -- 含稅總額
        '31',           -- 發票格式代碼 (二聯式)
        'V0',           -- 稅別 (一般稅率)
        '1',            -- 客戶類別
        '1',            -- 金額類別
        '1000',         -- 公司代碼
        'V0',           -- 稅碼
        'V0',           -- 營業稅碼
        'N',            -- 未過帳
        'SYSTEM'        -- 建立者
    )

    -- 取得剛插入的 HeaderID
    DECLARE @HeaderID INT = SCOPE_IDENTITY()

    -- 插入測試明細資料
    INSERT INTO MGUIAPDetail_Import (
        HeaderID, LineNum, jID, DocEntry,
        U_LIFNR, U_STCEG, U_XBLNR, U_ZFORM_CODE,
        U_BLDAT, U_VATDATE,
        U_HWBAS, U_HWSTE,
        U_TAX_TYPE, U_CUS_TYPE, U_AM_TYPE, U_VATCODE,
        U_BUKRS, U_MWSKZ,
        U_FA_DESC, U_FA_QTY, U_FA_USE,
        U_GatherMark, CreateBy
    ) VALUES (
        @HeaderID,      -- HeaderID
        1,              -- LineNum
        9999,           -- jID
        NULL,           -- DocEntry
        'V00001',       -- 供應商代碼
        '12345678',     -- 供應商統編
        'TEST-INV-001', -- 發票號碼
        '31',           -- 發票格式代碼
        '2025-11-16',   -- 發票日期
        '2025-11-16',   -- 營業稅日期
        5000.00,        -- 未稅金額
        250.00,         -- 營業稅額
        'V0',           -- 稅別
        '1',            -- 客戶類別
        '1',            -- 金額類別
        'V0',           -- 營業稅碼
        '1000',         -- 公司代碼
        'V0',           -- 稅碼
        '測試商品A',    -- 固定資產品名
        1.00,           -- 數量
        '辦公使用',      -- 用途
        'N',            -- 非彙總
        'SYSTEM'        -- 建立者
    ), (
        @HeaderID,      -- HeaderID
        2,              -- LineNum
        9999,           -- jID
        NULL,           -- DocEntry
        'V00001',       -- 供應商代碼
        '12345678',     -- 供應商統編
        'TEST-INV-001', -- 發票號碼
        '31',           -- 發票格式代碼
        '2025-11-16',   -- 發票日期
        '2025-11-16',   -- 營業稅日期
        5000.00,        -- 未稅金額
        250.00,         -- 營業稅額
        'V0',           -- 稅別
        '1',            -- 客戶類別
        '1',            -- 金額類別
        'V0',           -- 營業稅碼
        '1000',         -- 公司代碼
        'V0',           -- 稅碼
        '測試商品B',    -- 固定資產品名
        2.00,           -- 數量
        '生產使用',      -- 用途
        'N',            -- 非彙總
        'SYSTEM'        -- 建立者
    )

    PRINT 'Test data inserted successfully (1 header + 2 details).'
END
ELSE
BEGIN
    PRINT 'Test data already exists.'
END
GO

PRINT ''
PRINT '=========================================='
PRINT 'MDR 本地測試資料庫建立完成！'
PRINT '=========================================='
PRINT ''
PRINT '資料庫資訊：'
PRINT '- 資料庫名稱: MDR'
PRINT '- 連線字串範例: Server=.\SQLEXPRESS2008R2;Database=MDR;Integrated Security=True'
PRINT ''
PRINT '資料表：'
PRINT '1. MGUIAP_Import (表頭) - 22 欄位'
PRINT '2. MGUIAPDetail_Import (明細) - 25 欄位'
PRINT ''
PRINT '測試資料：'
PRINT '- 1 筆表頭 (jID=9999)'
PRINT '- 2 筆明細 (LineNum=1,2)'
PRINT ''
PRINT '⚠️ 注意事項：'
PRINT '- 此為本地測試環境,模擬正式 MDR 資料庫'
PRINT '- 正式環境: Server=192.168.1.219;Database=MDR'
PRINT '- 測試時請使用本地資料庫'
PRINT '- 程式碼中需判斷環境並切換連線字串'
PRINT ''
PRINT '下一步：'
PRINT '1. 在 Web.config 加入 MDR 連線字串'
PRINT '2. 實作 WriteMDRData() 方法'
PRINT '3. 測試資料寫入'
PRINT '4. 確認 MDRImport.exe 路徑與參數'
PRINT '=========================================='
GO
