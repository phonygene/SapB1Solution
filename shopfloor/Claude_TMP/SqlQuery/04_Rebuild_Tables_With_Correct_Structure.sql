-- =============================================
-- 重建資料表結構（修正主鍵為 jID）
-- 建立日期: 2025-10-30
-- 說明: 將主鍵從 DocEntry 改為 jID (IDENTITY)
-- =============================================

USE jtdb
GO

PRINT '===== 開始重建資料表 ====='
PRINT ''

-- =============================================
-- 1. 刪除所有相關表（依相依順序）
-- =============================================
PRINT '1. 刪除現有資料表...'

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAPDetail]'))
BEGIN
    DROP TABLE [dbo].[jMGUIAPDetail]
    PRINT '  - jMGUIAPDetail 已刪除'
END

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jMGUIAP]'))
BEGIN
    DROP TABLE [dbo].[jMGUIAP]
    PRINT '  - jMGUIAP 已刪除'
END

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jPCH1]'))
BEGIN
    DROP TABLE [dbo].[jPCH1]
    PRINT '  - jPCH1 已刪除'
END

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[jOPCH]'))
BEGIN
    DROP TABLE [dbo].[jOPCH]
    PRINT '  - jOPCH 已刪除'
END

PRINT ''
GO

-- =============================================
-- 2. 重建 jOPCH（單頭）- 主鍵改為 jID
-- =============================================
PRINT '2. 建立 jOPCH (AP單頭)...'

CREATE TABLE [dbo].[jOPCH] (
    -- 主鍵（平台唯一流水號）
    [jID]           INT IDENTITY(1,1) PRIMARY KEY,  -- 平台主鍵

    -- SAP B1 對應欄位
    [DocEntry]      INT NULL,                       -- SAP B1 的 DocEntry（回寫）
    [DocNum]        INT NULL,                       -- SAP B1 的 DocNum（回寫）

    -- 供應商資訊
    [CardCode]      NVARCHAR(50) NOT NULL,
    [CardName]      NVARCHAR(100) NULL,
    [NumAtCard]     NVARCHAR(100) NULL,

    -- 日期資訊
    [DocDate]       DATE NOT NULL,
    [DocDueDate]    DATE NULL,
    [TaxDate]       DATE NULL,

    -- 金額資訊
    [DocCurrency]   NVARCHAR(3) DEFAULT 'TWD',
    [DocRate]       DECIMAL(19,6) DEFAULT 1,
    [DocTotal]      DECIMAL(19,6) DEFAULT 0,
    [VatSum]        DECIMAL(19,6) DEFAULT 0,
    [DocTotalFC]    DECIMAL(19,6) DEFAULT 0,

    -- 發票與地址資訊
    [InvNum]        NVARCHAR(50) NULL,
    [AddressName]   NVARCHAR(100) NULL,
    [Address]       NVARCHAR(254) NULL,

    -- 付款條件
    [GroupNum]      INT NULL,

    -- 其他資訊
    [Comments]      NVARCHAR(254) NULL,
    [JrnlMemo]      NVARCHAR(50) NULL,

    -- 稅務設定
    [VatInclude]    NVARCHAR(1) DEFAULT 'N',        -- 單價是否含稅

    -- 狀態
    [DocStatus]     NVARCHAR(1) DEFAULT 'O',
    [Canceled]      NVARCHAR(1) DEFAULT 'N',
    [LineStatus]    NVARCHAR(1) DEFAULT 'O',

    -- SAP B1 關聯與狀態
    [B1DocEntry]    INT NULL,                       -- (已廢棄，使用 DocEntry)
    [B1DocNum]      INT NULL,                       -- (已廢棄，使用 DocNum)
    [B1PostStatus]  NVARCHAR(1) DEFAULT 'N',
    [B1PostDate]    DATETIME NULL,
    [B1ErrMsg]      NVARCHAR(500) NULL,

    -- 多公司支援
    [CompanyDB]     NVARCHAR(50) NULL,              -- 對應的 SAP B1 資料庫名稱

    -- 審核流程
    [ApprovalStatus] NVARCHAR(20) DEFAULT 'Pending',
    [ApprovedBy]     NVARCHAR(50) NULL,
    [ApprovedDate]   DATETIME NULL,

    -- 系統欄位
    [CreateDate]    DATETIME DEFAULT GETDATE(),
    [CreateBy]      NVARCHAR(50) NULL,              -- 建立者（原 Creator）
    [Creator]       NVARCHAR(50) NULL,              -- 登入帳號
    [UpdateDate]    DATETIME NULL,
    [UpdateBy]      NVARCHAR(50) NULL
)

PRINT '  - jOPCH 建立完成'
GO

-- =============================================
-- 3. 重建 jPCH1（單身）- 主鍵改為 (jID, LineNum)
-- =============================================
PRINT '3. 建立 jPCH1 (AP單身)...'

CREATE TABLE [dbo].[jPCH1] (
    -- 主鍵
    [jID]           INT NOT NULL,                   -- FK to jOPCH.jID
    [LineNum]       INT NOT NULL,                   -- 列號

    -- SAP B1 對應欄位
    [DocEntry]      INT NULL,                       -- SAP B1 的 DocEntry（回寫）
    [DocNum]        INT NULL,                       -- SAP B1 的 DocNum（回寫）

    -- 物料或科目
    [ItemCode]      NVARCHAR(50) NULL,
    [Dscription]    NVARCHAR(200) NULL,
    [AcctCode]      NVARCHAR(50) NULL,

    -- 數量與金額
    [Quantity]      DECIMAL(19,6) DEFAULT 0,
    [OpenQty]       DECIMAL(19,6) DEFAULT 0,        -- 未交數量
    [Price]         DECIMAL(19,6) DEFAULT 0,
    [Currency]      NVARCHAR(3) DEFAULT 'TWD',
    [Rate]          DECIMAL(19,6) DEFAULT 1,
    [LineTotal]     DECIMAL(19,6) DEFAULT 0,
    [GTotal]        DECIMAL(19,6) DEFAULT 0,

    -- 稅務資訊
    [TaxCode]       NVARCHAR(20) NULL,
    [VatGroup]      NVARCHAR(20) NULL,
    [VatPrcnt]      DECIMAL(19,6) DEFAULT 0,
    [LineVat]       DECIMAL(19,6) DEFAULT 0,
    [VatInclude]    NVARCHAR(1) DEFAULT 'N',        -- 該列是否含稅

    -- 交貨資訊
    [ShipDate]      DATE NULL,                      -- 交貨日期
    [WhsCode]       NVARCHAR(50) NULL,
    [UomCode]       NVARCHAR(20) NULL,

    -- 成本中心與專案
    [CostingCode]   NVARCHAR(50) NULL,
    [CostingCode2]  NVARCHAR(50) NULL,
    [CostingCode3]  NVARCHAR(50) NULL,
    [Project]       NVARCHAR(50) NULL,

    -- 明細備註與附件
    [LineMemo]      NVARCHAR(254) NULL,
    [Attachment]    NVARCHAR(500) NULL,

    -- 狀態
    [LineStatus]    NVARCHAR(1) DEFAULT 'O',

    -- SAP B1 關聯
    [B1DocEntry]    INT NULL,                       -- (已廢棄，使用 DocEntry)
    [B1LineNum]     INT NULL,

    -- 主鍵約束
    CONSTRAINT PK_jPCH1 PRIMARY KEY (jID, LineNum),

    -- 外鍵約束
    CONSTRAINT FK_jPCH1_jOPCH FOREIGN KEY (jID)
        REFERENCES jOPCH(jID) ON DELETE CASCADE
)

PRINT '  - jPCH1 建立完成'
GO

-- =============================================
-- 4. 重建 jMGUIAP（營業稅表頭）
-- =============================================
PRINT '4. 建立 jMGUIAP (營業稅表頭)...'

CREATE TABLE [dbo].[jMGUIAP] (
    -- 主鍵
    [ID]            INT IDENTITY(1,1) PRIMARY KEY,

    -- 關聯到 AP 發票（使用 jID）
    [jID]           INT NOT NULL,                   -- FK to jOPCH.jID

    -- SAP B1 對應欄位
    [DocEntry]      INT NULL,                       -- SAP B1 的 DocEntry（回寫）
    [B1DocEntry]    INT NULL,                       -- (已廢棄，使用 DocEntry)

    -- 基本資訊
    [DocNum]        INT NULL,
    [DocTotal]      DECIMAL(19,6) DEFAULT 0,
    [VatSum]        DECIMAL(19,6) DEFAULT 0,

    -- SAP B1 MDR 資訊
    [U_OBJTYPE]     NVARCHAR(20) DEFAULT '18',
    [MDRPostStatus] NVARCHAR(1) DEFAULT 'N',
    [MDRPostDate]   DATETIME NULL,
    [MDRErrMsg]     NVARCHAR(500) NULL,

    -- 系統欄位
    [CreateDate]    DATETIME DEFAULT GETDATE(),
    [CreateBy]      NVARCHAR(50) NULL,
    [UpdateDate]    DATETIME NULL,
    [UpdateBy]      NVARCHAR(50) NULL,

    -- 外鍵約束
    CONSTRAINT FK_jMGUIAP_jOPCH FOREIGN KEY (jID)
        REFERENCES jOPCH(jID) ON DELETE CASCADE
)

PRINT '  - jMGUIAP 建立完成'
GO

-- =============================================
-- 5. 重建 jMGUIAPDetail（營業稅明細）
-- =============================================
PRINT '5. 建立 jMGUIAPDetail (營業稅明細)...'

CREATE TABLE [dbo].[jMGUIAPDetail] (
    -- 主鍵（與 AP 單身對應）
    [jID]           INT NOT NULL,                   -- FK to jPCH1.jID
    [LineNum]       INT NOT NULL,                   -- FK to jPCH1.LineNum

    -- SAP B1 對應欄位
    [DocEntry]      INT NULL,                       -- SAP B1 的 DocEntry（回寫）

    -- 營業稅發票基本資訊
    [U_LIFNR]       NVARCHAR(50) NULL,
    [U_STCEG]       NVARCHAR(20) NULL,
    [U_XBLNR]       NVARCHAR(50) NULL,
    [U_ZFORM_CODE]  NVARCHAR(10) NULL,

    -- 日期
    [U_BLDAT]       DATE NULL,
    [U_VATDATE]     DATE NULL,

    -- 金額
    [U_HWBAS]       DECIMAL(19,6) DEFAULT 0,
    [U_HWSTE]       DECIMAL(19,6) DEFAULT 0,

    -- 稅別
    [U_TAX_TYPE]    NVARCHAR(10) NULL,
    [U_CUS_TYPE]    NVARCHAR(10) NULL,
    [U_AM_TYPE]     NVARCHAR(10) NULL,

    -- SAP B1 資訊
    [U_VATCODE]     NVARCHAR(20) NULL,
    [U_BUKRS]       NVARCHAR(10) NULL,
    [U_MWSKZ]       NVARCHAR(10) NULL,
    [U_BELNR]       NVARCHAR(50) NULL,

    -- 固定資產相關
    [U_FA_DESC]     NVARCHAR(200) NULL,
    [U_FA_QTY]      DECIMAL(19,6) DEFAULT 0,
    [U_FA_USE]      NVARCHAR(200) NULL,

    -- 其他
    [U_GatherMark]  NVARCHAR(1) DEFAULT 'N',
    [U_ConsolidQty] DECIMAL(19,6) DEFAULT 0,

    -- SAP B1 關聯
    [B1DocEntry]    INT NULL,                       -- (已廢棄，使用 DocEntry)
    [B1LineNum]     INT NULL,

    -- 主鍵約束
    CONSTRAINT PK_jMGUIAPDetail PRIMARY KEY (jID, LineNum),

    -- 外鍵約束（關聯到 AP 單身）
    CONSTRAINT FK_jMGUIAPDetail_jPCH1 FOREIGN KEY (jID, LineNum)
        REFERENCES jPCH1(jID, LineNum) ON DELETE CASCADE
)

PRINT '  - jMGUIAPDetail 建立完成'
GO

-- =============================================
-- 6. 建立索引
-- =============================================
PRINT ''
PRINT '6. 建立索引...'

-- jOPCH 索引
CREATE INDEX IX_jOPCH_DocEntry ON jOPCH(DocEntry)
CREATE INDEX IX_jOPCH_CardCode ON jOPCH(CardCode)
CREATE INDEX IX_jOPCH_DocDate ON jOPCH(DocDate)
CREATE INDEX IX_jOPCH_CompanyDB ON jOPCH(CompanyDB)
CREATE INDEX IX_jOPCH_InvNum ON jOPCH(InvNum)
PRINT '  - jOPCH 索引建立完成'

-- jPCH1 索引
CREATE INDEX IX_jPCH1_DocEntry ON jPCH1(DocEntry)
CREATE INDEX IX_jPCH1_ItemCode ON jPCH1(ItemCode)
CREATE INDEX IX_jPCH1_AcctCode ON jPCH1(AcctCode)
CREATE INDEX IX_jPCH1_Project ON jPCH1(Project)
PRINT '  - jPCH1 索引建立完成'

-- jMGUIAP 索引
CREATE INDEX IX_jMGUIAP_DocEntry ON jMGUIAP(DocEntry)
PRINT '  - jMGUIAP 索引建立完成'

-- jMGUIAPDetail 索引
CREATE INDEX IX_jMGUIAPDetail_DocEntry ON jMGUIAPDetail(DocEntry)
CREATE INDEX IX_jMGUIAPDetail_XBLNR ON jMGUIAPDetail(U_XBLNR)
CREATE INDEX IX_jMGUIAPDetail_STCEG ON jMGUIAPDetail(U_STCEG)
PRINT '  - jMGUIAPDetail 索引建立完成'

PRINT ''
PRINT '===================================='
PRINT '資料表重建完成！'
PRINT '===================================='
PRINT '結構變更:'
PRINT '  1. jOPCH 主鍵: DocEntry → jID (IDENTITY)'
PRINT '  2. jPCH1 主鍵: (DocEntry, LineNum) → (jID, LineNum)'
PRINT '  3. 新增欄位: Creator, CompanyDB'
PRINT '  4. 新增欄位: VatInclude, LineStatus, ShipDate, OpenQty'
PRINT '  5. DocEntry/DocNum 改為可 NULL（由 SAP B1 回寫）'
PRINT '===================================='
GO
