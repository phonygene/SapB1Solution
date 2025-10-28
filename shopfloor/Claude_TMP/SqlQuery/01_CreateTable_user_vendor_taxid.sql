/*******************************************************************************
 * 檔案說明：建立使用者常用統編清單表
 * 資料庫：jtdb
 * 用途：儲存使用者在費用申請單中常用的供應商統一編號
 *
 * 使用方式：
 * 1. 在 SSMS 連接到 SQL Server
 * 2. 確認連接到 jtdb 資料庫
 * 3. 執行本腳本
 *
 * 注意事項：
 * - 表名稱：user_vendor_taxid（如需修改請全文搜尋替換）
 * - 外鍵約束：會檢查 User.id 是否存在
 * - 如果表已存在會報錯（正常情況）
 ******************************************************************************/

USE [jtdb]
GO

/****** Object:  Table [dbo].[user_vendor_taxid] ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[user_vendor_taxid](
    [num] [int] IDENTITY(1,1) NOT NULL,
    [id] [varchar](50) NOT NULL,           -- 對應 User.id
    [taxid] [varchar](10) NOT NULL,        -- 統一編號（8位數）
    [vendorname] [nvarchar](100) NULL,     -- 供應商名稱（選填）
    [createdate] [datetime] NOT NULL,      -- 建立時間
    [lastused] [datetime] NULL,            -- 最後使用時間（預留給 Phase 2）
 CONSTRAINT [PK_user_vendor_taxid] PRIMARY KEY CLUSTERED
(
    [num] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF,
       ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ_user_vendor_taxid_id_taxid] UNIQUE NONCLUSTERED
(
    [id] ASC,
    [taxid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF,
       ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO

-- 預設值約束：建立時間自動填入當前時間
ALTER TABLE [dbo].[user_vendor_taxid] ADD CONSTRAINT [DF_user_vendor_taxid_createdate]
    DEFAULT (GETDATE()) FOR [createdate]
GO

-- 外鍵約束：確保 id 存在於 User 表
-- 注意：如果不想要外鍵約束（例如方便測試），可以註解掉以下 3 行
ALTER TABLE [dbo].[user_vendor_taxid] WITH CHECK ADD CONSTRAINT [FK_user_vendor_taxid_User]
    FOREIGN KEY([id]) REFERENCES [dbo].[User] ([id])
GO

ALTER TABLE [dbo].[user_vendor_taxid] CHECK CONSTRAINT [FK_user_vendor_taxid_User]
GO

-- 驗證建立成功
SELECT 'Table created successfully' AS Result
GO
