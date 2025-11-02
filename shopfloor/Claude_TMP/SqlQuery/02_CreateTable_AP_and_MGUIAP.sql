-- =============================================
-- 建立平台 AP 發票與營業稅發票資料表
-- 建立日期: 2025-10-30
-- 說明:
--   jOPCH: AP單頭（對應SAP B1的OPCH）
--   jPCH1: AP單身（對應SAP B1的PCH1）
--   jMGUIAP: 營業稅發票表頭
--   jMGUIAPDetail: 營業稅發票明細（與AP單身1:1對應）
-- =============================================

USE jtdb
GO

-- =============================================
-- 1. 建立 jOPCH (AP單頭)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jOPCH]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[jOPCH] (
        -- 主鍵
        [DocEntry]      INT IDENTITY(1,1) PRIMARY KEY,  -- 單據內部編號
        [DocNum]        INT NULL,                        -- 單據號碼（對應SAP B1）

        -- 供應商資訊
        [CardCode]      NVARCHAR(50) NOT NULL,          -- 供應商代碼
        [CardName]      NVARCHAR(100) NULL,             -- 供應商名稱
        [NumAtCard]     NVARCHAR(100) NULL,             -- 業務夥伴參考號碼

        -- 日期資訊
        [DocDate]       DATE NOT NULL,                  -- 單據日期
        [DocDueDate]    DATE NULL,                      -- 到期日
        [TaxDate]       DATE NULL,                      -- 憑證日期

        -- 金額資訊
        [DocCurrency]   NVARCHAR(3) DEFAULT 'TWD',      -- 幣別
        [DocRate]       DECIMAL(19,6) DEFAULT 1,        -- 匯率
        [DocTotal]      DECIMAL(19,6) DEFAULT 0,        -- 單據總額（含稅）
        [VatSum]        DECIMAL(19,6) DEFAULT 0,        -- 稅額
        [DocTotalFC]    DECIMAL(19,6) DEFAULT 0,        -- 外幣總額

        -- 發票與地址資訊
        [InvNum]        NVARCHAR(50) NULL,              -- 發票號碼
        [AddressName]   NVARCHAR(100) NULL,             -- 送貨地址名稱
        [Address]       NVARCHAR(254) NULL,             -- 送貨地址

        -- 付款條件
        [GroupNum]      INT NULL,                       -- 付款條件

        -- 其他資訊
        [Comments]      NVARCHAR(254) NULL,             -- 備註
        [JrnlMemo]      NVARCHAR(50) NULL,              -- 日記帳備註

        -- 狀態
        [DocStatus]     NVARCHAR(1) DEFAULT 'O',        -- 單據狀態 (O=Open, C=Closed)
        [Canceled]      NVARCHAR(1) DEFAULT 'N',        -- 是否取消 (Y/N)

        -- SAP B1 關聯
        [B1DocEntry]    INT NULL,                       -- 對應SAP B1的DocEntry
        [B1DocNum]      INT NULL,                       -- 對應SAP B1的DocNum
        [B1PostStatus]  NVARCHAR(1) DEFAULT 'N',        -- SAP B1過帳狀態 (Y/N)
        [B1PostDate]    DATETIME NULL,                  -- SAP B1過帳時間
        [B1ErrMsg]      NVARCHAR(500) NULL,             -- SAP B1錯誤訊息

        -- 審核流程
        [ApprovalStatus] NVARCHAR(20) DEFAULT 'Pending', -- 審核狀態 (Pending/Approved/Rejected)
        [ApprovedBy]     NVARCHAR(50) NULL,              -- 核准人
        [ApprovedDate]   DATETIME NULL,                  -- 核准日期

        -- 系統欄位
        [CreateDate]    DATETIME DEFAULT GETDATE(),     -- 建立日期
        [CreateBy]      NVARCHAR(50) NULL,              -- 建立人
        [UpdateDate]    DATETIME NULL,                  -- 更新日期
        [UpdateBy]      NVARCHAR(50) NULL               -- 更新人
    )

    PRINT 'Table jOPCH created successfully.'
END
ELSE
BEGIN
    PRINT 'Table jOPCH already exists.'
END
GO

-- =============================================
-- 2. 建立 jPCH1 (AP單身)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jPCH1]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[jPCH1] (
        -- 主鍵
        [DocEntry]      INT NOT NULL,                   -- 表頭DocEntry（FK to jOPCH）
        [LineNum]       INT NOT NULL,                   -- 列號（從0開始）

        -- 物料或科目
        [ItemCode]      NVARCHAR(50) NULL,              -- 品號（若為物料）
        [Dscription]    NVARCHAR(200) NULL,             -- 說明
        [AcctCode]      NVARCHAR(50) NULL,              -- 會計科目（若為費用）

        -- 數量與金額
        [Quantity]      DECIMAL(19,6) DEFAULT 0,        -- 數量
        [Price]         DECIMAL(19,6) DEFAULT 0,        -- 單價
        [Currency]      NVARCHAR(3) DEFAULT 'TWD',      -- 幣別
        [Rate]          DECIMAL(19,6) DEFAULT 1,        -- 匯率
        [LineTotal]     DECIMAL(19,6) DEFAULT 0,        -- 列總額（未稅）
        [GTotal]        DECIMAL(19,6) DEFAULT 0,        -- 列總額（含稅）

        -- 稅務資訊
        [TaxCode]       NVARCHAR(20) NULL,              -- 稅率代碼
        [VatGroup]      NVARCHAR(20) NULL,              -- 稅別
        [VatPrcnt]      DECIMAL(19,6) DEFAULT 0,        -- 稅率%
        [LineVat]       DECIMAL(19,6) DEFAULT 0,        -- 該列稅額

        -- 成本中心與專案
        [CostingCode]   NVARCHAR(50) NULL,              -- 成本中心
        [CostingCode2]  NVARCHAR(50) NULL,              -- 成本中心2
        [CostingCode3]  NVARCHAR(50) NULL,              -- 成本中心3
        [Project]       NVARCHAR(50) NULL,              -- 專案代碼

        -- 明細備註與附件
        [LineMemo]      NVARCHAR(254) NULL,             -- 明細備註
        [Attachment]    NVARCHAR(500) NULL,             -- 附件路徑（C:\jAttach\）

        -- 其他
        [WhsCode]       NVARCHAR(50) NULL,              -- 倉庫代碼
        [UomCode]       NVARCHAR(20) NULL,              -- 單位

        -- SAP B1 關聯
        [B1DocEntry]    INT NULL,                       -- 對應SAP B1的DocEntry
        [B1LineNum]     INT NULL,                       -- 對應SAP B1的LineNum

        -- 主鍵約束
        CONSTRAINT PK_jPCH1 PRIMARY KEY (DocEntry, LineNum),

        -- 外鍵約束
        CONSTRAINT FK_jPCH1_jOPCH FOREIGN KEY (DocEntry)
            REFERENCES jOPCH(DocEntry) ON DELETE CASCADE
    )

    PRINT 'Table jPCH1 created successfully.'
END
ELSE
BEGIN
    PRINT 'Table jPCH1 already exists.'
END
GO

-- =============================================
-- 3. 建立 jMGUIAP (營業稅發票表頭)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAP]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[jMGUIAP] (
        -- 主鍵
        [ID]            INT IDENTITY(1,1) PRIMARY KEY,  -- 營業稅單據ID

        -- 關聯到AP發票
        [DocEntry]      INT NOT NULL,                   -- 關聯到jOPCH.DocEntry
        [B1DocEntry]    INT NULL,                       -- 對應SAP B1的DocEntry

        -- 基本資訊
        [DocNum]        INT NULL,                       -- 單據號碼
        [DocTotal]      DECIMAL(19,6) DEFAULT 0,        -- 單據總額
        [VatSum]        DECIMAL(19,6) DEFAULT 0,        -- 稅額總額

        -- SAP B1 MDR 資訊
        [U_OBJTYPE]     NVARCHAR(20) DEFAULT '18',      -- 物件類型（18=AP Invoice）
        [MDRPostStatus] NVARCHAR(1) DEFAULT 'N',        -- MDR過帳狀態 (Y/N)
        [MDRPostDate]   DATETIME NULL,                  -- MDR過帳時間
        [MDRErrMsg]     NVARCHAR(500) NULL,             -- MDR錯誤訊息

        -- 系統欄位
        [CreateDate]    DATETIME DEFAULT GETDATE(),     -- 建立日期
        [CreateBy]      NVARCHAR(50) NULL,              -- 建立人
        [UpdateDate]    DATETIME NULL,                  -- 更新日期
        [UpdateBy]      NVARCHAR(50) NULL,              -- 更新人

        -- 外鍵約束
        CONSTRAINT FK_jMGUIAP_jOPCH FOREIGN KEY (DocEntry)
            REFERENCES jOPCH(DocEntry) ON DELETE CASCADE
    )

    PRINT 'Table jMGUIAP created successfully.'
END
ELSE
BEGIN
    PRINT 'Table jMGUIAP already exists.'
END
GO

-- =============================================
-- 4. 建立 jMGUIAPDetail (營業稅發票明細)
-- =============================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAPDetail]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[jMGUIAPDetail] (
        -- 主鍵（與AP單身對應）
        [DocEntry]      INT NOT NULL,                   -- 對應jPCH1.DocEntry
        [LineNum]       INT NOT NULL,                   -- 對應jPCH1.LineNum

        -- 營業稅發票基本資訊
        [U_LIFNR]       NVARCHAR(50) NULL,              -- 供應商代碼
        [U_STCEG]       NVARCHAR(20) NULL,              -- 統一編號
        [U_XBLNR]       NVARCHAR(50) NULL,              -- 發票號碼
        [U_ZFORM_CODE]  NVARCHAR(10) NULL,              -- 發票類型（21,22,25,26,27,28）

        -- 日期
        [U_BLDAT]       DATE NULL,                      -- 憑證日期
        [U_VATDATE]     DATE NULL,                      -- 營業稅日期

        -- 金額
        [U_HWBAS]       DECIMAL(19,6) DEFAULT 0,        -- 未稅金額
        [U_HWSTE]       DECIMAL(19,6) DEFAULT 0,        -- 稅額

        -- 稅別
        [U_TAX_TYPE]    NVARCHAR(10) NULL,              -- 稅別（1=應稅, 2=零稅, 3=免稅）
        [U_CUS_TYPE]    NVARCHAR(10) NULL,              -- 客戶類型
        [U_AM_TYPE]     NVARCHAR(10) NULL,              -- 金額類型

        -- SAP B1 資訊
        [U_VATCODE]     NVARCHAR(20) NULL,              -- 稅碼
        [U_BUKRS]       NVARCHAR(10) NULL,              -- 公司代碼
        [U_MWSKZ]       NVARCHAR(10) NULL,              -- 稅碼
        [U_BELNR]       NVARCHAR(50) NULL,              -- 憑證號碼

        -- 固定資產相關
        [U_FA_DESC]     NVARCHAR(200) NULL,             -- 固定資產說明
        [U_FA_QTY]      DECIMAL(19,6) DEFAULT 0,        -- 固定資產數量
        [U_FA_USE]      NVARCHAR(200) NULL,             -- 固定資產用途

        -- 其他
        [U_GatherMark]  NVARCHAR(1) DEFAULT 'N',        -- 彙總標記
        [U_ConsolidQty] DECIMAL(19,6) DEFAULT 0,        -- 合併數量

        -- SAP B1 關聯
        [B1DocEntry]    INT NULL,                       -- 對應SAP B1的DocEntry
        [B1LineNum]     INT NULL,                       -- 對應SAP B1的LineNum

        -- 主鍵約束
        CONSTRAINT PK_jMGUIAPDetail PRIMARY KEY (DocEntry, LineNum),

        -- 外鍵約束（關聯到AP單身）
        CONSTRAINT FK_jMGUIAPDetail_jPCH1 FOREIGN KEY (DocEntry, LineNum)
            REFERENCES jPCH1(DocEntry, LineNum) ON DELETE CASCADE
    )

    PRINT 'Table jMGUIAPDetail created successfully.'
END
ELSE
BEGIN
    PRINT 'Table jMGUIAPDetail already exists.'
END
GO

-- =============================================
-- 5. 建立索引
-- =============================================

-- jOPCH 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jOPCH_CardCode')
    CREATE INDEX IX_jOPCH_CardCode ON jOPCH(CardCode)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jOPCH_DocDate')
    CREATE INDEX IX_jOPCH_DocDate ON jOPCH(DocDate)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jOPCH_B1DocEntry')
    CREATE INDEX IX_jOPCH_B1DocEntry ON jOPCH(B1DocEntry)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jOPCH_InvNum')
    CREATE INDEX IX_jOPCH_InvNum ON jOPCH(InvNum)
GO

-- jPCH1 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jPCH1_ItemCode')
    CREATE INDEX IX_jPCH1_ItemCode ON jPCH1(ItemCode)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jPCH1_AcctCode')
    CREATE INDEX IX_jPCH1_AcctCode ON jPCH1(AcctCode)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jPCH1_Project')
    CREATE INDEX IX_jPCH1_Project ON jPCH1(Project)
GO

-- jMGUIAPDetail 索引
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jMGUIAPDetail_XBLNR')
    CREATE INDEX IX_jMGUIAPDetail_XBLNR ON jMGUIAPDetail(U_XBLNR)
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_jMGUIAPDetail_STCEG')
    CREATE INDEX IX_jMGUIAPDetail_STCEG ON jMGUIAPDetail(U_STCEG)
GO

PRINT ''
PRINT '===================================='
PRINT '所有資料表建立完成！'
PRINT '===================================='
PRINT 'jOPCH        - AP單頭'
PRINT 'jPCH1        - AP單身'
PRINT 'jMGUIAP      - 營業稅表頭'
PRINT 'jMGUIAPDetail - 營業稅明細'
PRINT '===================================='
GO
