/*
=============================================================================
AI 輔助功能資料庫結構
目標資料庫：jtdb
建立日期：2025-12-24
=============================================================================
*/

-- ============================================================================
-- 1. AI_IssueType - Issue 類型代碼表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_IssueType]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_IssueType] (
        [TypeID]        INT IDENTITY(1,1)   NOT NULL,
        [TypeCode]      VARCHAR(20)         NOT NULL,
        [TypeName]      NVARCHAR(50)        NOT NULL,
        [IsActive]      BIT                 NOT NULL DEFAULT 1,
        [SortOrder]     INT                 NOT NULL DEFAULT 0,
        CONSTRAINT [PK_AI_IssueType] PRIMARY KEY CLUSTERED ([TypeID]),
        CONSTRAINT [UQ_AI_IssueType_Code] UNIQUE ([TypeCode])
    );
    PRINT 'Created table: AI_IssueType';
END
GO

-- ============================================================================
-- 2. AI_IssueStatus - Issue 狀態代碼表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_IssueStatus]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_IssueStatus] (
        [StatusID]      INT IDENTITY(1,1)   NOT NULL,
        [StatusCode]    VARCHAR(20)         NOT NULL,
        [StatusName]    NVARCHAR(50)        NOT NULL,
        [IsActive]      BIT                 NOT NULL DEFAULT 1,
        [SortOrder]     INT                 NOT NULL DEFAULT 0,
        CONSTRAINT [PK_AI_IssueStatus] PRIMARY KEY CLUSTERED ([StatusID]),
        CONSTRAINT [UQ_AI_IssueStatus_Code] UNIQUE ([StatusCode])
    );
    PRINT 'Created table: AI_IssueStatus';
END
GO

-- ============================================================================
-- 3. AI_Tag - Tag 標籤主表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_Tag]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_Tag] (
        [TagID]         INT IDENTITY(1,1)   NOT NULL,
        [TagName]       NVARCHAR(50)        NOT NULL,
        [Category]      VARCHAR(20)         NULL,           -- TYPE/MODULE/AREA/SEVERITY/STATUS/SYSTEM
        [Color]         VARCHAR(10)         NULL,           -- Hex 顏色碼，如 #FF0000
        [UseCount]      INT                 NOT NULL DEFAULT 0,
        [IsSystem]      BIT                 NOT NULL DEFAULT 0,  -- 系統預設不可刪除
        [IsActive]      BIT                 NOT NULL DEFAULT 1,
        [SortOrder]     INT                 NOT NULL DEFAULT 0,
        [CreateTime]    DATETIME            NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_AI_Tag] PRIMARY KEY CLUSTERED ([TagID]),
        CONSTRAINT [UQ_AI_Tag_Name] UNIQUE ([TagName])
    );
    PRINT 'Created table: AI_Tag';
END
GO

-- ============================================================================
-- 4. AI_Issue - Issue 主表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_Issue]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_Issue] (
        [IssueID]           INT IDENTITY(1,1)   NOT NULL,
        [IssueTypeID]       INT                 NOT NULL,
        [IssueStatusID]     INT                 NOT NULL,
        [PageURL]           NVARCHAR(500)       NULL,
        [AreaCode]          VARCHAR(20)         NULL,           -- HEADER/DETAIL/FOOTER
        [AreaTabName]       NVARCHAR(50)        NULL,           -- 單身頁籤名稱
        [ElementID]         VARCHAR(100)        NULL,
        [Description]       NVARCHAR(MAX)       NOT NULL,
        [Analysis]          NVARCHAR(MAX)       NULL,
        [ScreenshotPath]    NVARCHAR(500)       NULL,
        [SubmitUserID]      VARCHAR(50)         NOT NULL,
        [SubmitTime]        DATETIME            NOT NULL DEFAULT GETDATE(),
        [SubmitRemark]      NVARCHAR(MAX)       NULL,
        [DevRemark]         NVARCHAR(MAX)       NULL,
        [Priority]          INT                 NOT NULL DEFAULT 99,
        [AssignedTo]        VARCHAR(50)         NULL,
        [RelatedIssueID]    INT                 NULL,
        [IsAutoSubmit]      BIT                 NOT NULL DEFAULT 0,
        [AutoSubmitReason]  VARCHAR(20)         NULL,           -- LIMIT_HOURLY/LIMIT_ISSUE/VIOLATION
        [ConversationID]    INT                 NULL,
        [CreateTime]        DATETIME            NOT NULL DEFAULT GETDATE(),
        [UpdateTime]        DATETIME            NULL,
        CONSTRAINT [PK_AI_Issue] PRIMARY KEY CLUSTERED ([IssueID]),
        CONSTRAINT [FK_AI_Issue_Type] FOREIGN KEY ([IssueTypeID]) REFERENCES [dbo].[AI_IssueType]([TypeID]),
        CONSTRAINT [FK_AI_Issue_Status] FOREIGN KEY ([IssueStatusID]) REFERENCES [dbo].[AI_IssueStatus]([StatusID]),
        CONSTRAINT [FK_AI_Issue_Related] FOREIGN KEY ([RelatedIssueID]) REFERENCES [dbo].[AI_Issue]([IssueID])
    );

    -- 索引
    CREATE NONCLUSTERED INDEX [IX_AI_Issue_SubmitUserID] ON [dbo].[AI_Issue]([SubmitUserID]);
    CREATE NONCLUSTERED INDEX [IX_AI_Issue_Status] ON [dbo].[AI_Issue]([IssueStatusID]);
    CREATE NONCLUSTERED INDEX [IX_AI_Issue_Type] ON [dbo].[AI_Issue]([IssueTypeID]);
    CREATE NONCLUSTERED INDEX [IX_AI_Issue_SubmitTime] ON [dbo].[AI_Issue]([SubmitTime] DESC);
    CREATE NONCLUSTERED INDEX [IX_AI_Issue_Priority] ON [dbo].[AI_Issue]([Priority], [IssueStatusID]);

    PRINT 'Created table: AI_Issue';
END
GO

-- ============================================================================
-- 5. AI_IssueTag - Issue 與 Tag 關聯表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_IssueTag]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_IssueTag] (
        [IssueID]       INT                 NOT NULL,
        [TagID]         INT                 NOT NULL,
        [CreateTime]    DATETIME            NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_AI_IssueTag] PRIMARY KEY CLUSTERED ([IssueID], [TagID]),
        CONSTRAINT [FK_AI_IssueTag_Issue] FOREIGN KEY ([IssueID]) REFERENCES [dbo].[AI_Issue]([IssueID]) ON DELETE CASCADE,
        CONSTRAINT [FK_AI_IssueTag_Tag] FOREIGN KEY ([TagID]) REFERENCES [dbo].[AI_Tag]([TagID])
    );

    -- 反向索引，方便從 Tag 查 Issue
    CREATE NONCLUSTERED INDEX [IX_AI_IssueTag_TagID] ON [dbo].[AI_IssueTag]([TagID]);

    PRINT 'Created table: AI_IssueTag';
END
GO

-- ============================================================================
-- 6. AI_ConversationLog - 對話紀錄表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_ConversationLog]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_ConversationLog] (
        [LogID]             INT IDENTITY(1,1)   NOT NULL,
        [ConversationID]    INT                 NOT NULL,       -- 同一次對話共用
        [UserID]            VARCHAR(50)         NOT NULL,
        [PageURL]           NVARCHAR(500)       NULL,
        [Sequence]          INT                 NOT NULL,       -- 對話順序
        [Role]              VARCHAR(10)         NOT NULL,       -- USER/AI
        [Content]           NVARCHAR(MAX)       NOT NULL,
        [IsViolation]       BIT                 NOT NULL DEFAULT 0,
        [TokenUsed]         INT                 NULL,
        [CreateTime]        DATETIME            NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_AI_ConversationLog] PRIMARY KEY CLUSTERED ([LogID])
    );

    -- 索引
    CREATE NONCLUSTERED INDEX [IX_AI_ConversationLog_ConvID] ON [dbo].[AI_ConversationLog]([ConversationID], [Sequence]);
    CREATE NONCLUSTERED INDEX [IX_AI_ConversationLog_UserID] ON [dbo].[AI_ConversationLog]([UserID], [CreateTime] DESC);

    PRINT 'Created table: AI_ConversationLog';
END
GO

-- ============================================================================
-- 7. AI_UserQuota - 使用者配額與禁用狀態表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_UserQuota]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_UserQuota] (
        [UserID]            VARCHAR(50)         NOT NULL,
        [HourlyCount]       INT                 NOT NULL DEFAULT 0,
        [HourlyResetTime]   DATETIME            NOT NULL DEFAULT GETDATE(),
        [IsBanned]          BIT                 NOT NULL DEFAULT 0,
        [BanTime]           DATETIME            NULL,
        [BanReason]         NVARCHAR(500)       NULL,
        [BanIssueID]        INT                 NULL,
        [UnbanTime]         DATETIME            NULL,
        [UnbanBy]           VARCHAR(50)         NULL,
        [UpdateTime]        DATETIME            NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_AI_UserQuota] PRIMARY KEY CLUSTERED ([UserID])
    );
    PRINT 'Created table: AI_UserQuota';
END
GO

-- ============================================================================
-- 8. AI_ViolationLog - 違規紀錄表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_ViolationLog]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_ViolationLog] (
        [ViolationID]       INT IDENTITY(1,1)   NOT NULL,
        [UserID]            VARCHAR(50)         NOT NULL,
        [ConversationID]    INT                 NOT NULL,
        [ViolationType]     VARCHAR(20)         NOT NULL,       -- KEYWORD/AI_JUDGE
        [TriggerContent]    NVARCHAR(500)       NULL,
        [MatchedKeyword]    NVARCHAR(100)       NULL,
        [CreateTime]        DATETIME            NOT NULL DEFAULT GETDATE(),
        CONSTRAINT [PK_AI_ViolationLog] PRIMARY KEY CLUSTERED ([ViolationID])
    );

    -- 索引：用於查詢「5 分鐘內違規次數」
    CREATE NONCLUSTERED INDEX [IX_AI_ViolationLog_UserTime] ON [dbo].[AI_ViolationLog]([UserID], [CreateTime] DESC);

    PRINT 'Created table: AI_ViolationLog';
END
GO

-- ============================================================================
-- 9. AI_FilterKeyword - 關鍵字過濾表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_FilterKeyword]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_FilterKeyword] (
        [KeywordID]     INT IDENTITY(1,1)   NOT NULL,
        [Keyword]       NVARCHAR(100)       NOT NULL,
        [MatchType]     VARCHAR(10)         NOT NULL DEFAULT 'CONTAINS',  -- EXACT/CONTAINS
        [IsActive]      BIT                 NOT NULL DEFAULT 1,
        [CreateTime]    DATETIME            NOT NULL DEFAULT GETDATE(),
        [CreateBy]      VARCHAR(50)         NOT NULL,
        [Remark]        NVARCHAR(200)       NULL,
        CONSTRAINT [PK_AI_FilterKeyword] PRIMARY KEY CLUSTERED ([KeywordID])
    );
    PRINT 'Created table: AI_FilterKeyword';
END
GO

-- ============================================================================
-- 10. AI_PageArea - 頁面區域定義表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_PageArea]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_PageArea] (
        [AreaID]        INT IDENTITY(1,1)   NOT NULL,
        [PageURL]       NVARCHAR(500)       NOT NULL,
        [AreaCode]      VARCHAR(20)         NOT NULL,           -- HEADER/DETAIL/FOOTER
        [AreaName]      NVARCHAR(50)        NOT NULL,
        [ContainerID]   VARCHAR(100)        NOT NULL,           -- HTML 容器 ID
        [SortOrder]     INT                 NOT NULL DEFAULT 0,
        [IsActive]      BIT                 NOT NULL DEFAULT 1,
        CONSTRAINT [PK_AI_PageArea] PRIMARY KEY CLUSTERED ([AreaID])
    );

    -- 索引
    CREATE NONCLUSTERED INDEX [IX_AI_PageArea_PageURL] ON [dbo].[AI_PageArea]([PageURL]);

    PRINT 'Created table: AI_PageArea';
END
GO

-- ============================================================================
-- 11. AI_FieldHelp - 欄位說明表
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_FieldHelp]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_FieldHelp] (
        [HelpID]        INT IDENTITY(1,1)   NOT NULL,
        [PageURL]       NVARCHAR(500)       NOT NULL,
        [AreaCode]      VARCHAR(20)         NOT NULL,           -- HEADER/DETAIL/FOOTER
        [TabName]       NVARCHAR(50)        NULL,               -- 單身頁籤名稱
        [ElementID]     VARCHAR(100)        NULL,               -- NULL 表示區域層級說明
        [HelpTitle]     NVARCHAR(100)       NOT NULL,
        [HelpContent]   NVARCHAR(MAX)       NOT NULL,
        [IsActive]      BIT                 NOT NULL DEFAULT 1,
        [UpdateTime]    DATETIME            NOT NULL DEFAULT GETDATE(),
        [UpdateBy]      VARCHAR(50)         NOT NULL,
        CONSTRAINT [PK_AI_FieldHelp] PRIMARY KEY CLUSTERED ([HelpID])
    );

    -- 索引
    CREATE NONCLUSTERED INDEX [IX_AI_FieldHelp_Page] ON [dbo].[AI_FieldHelp]([PageURL], [AreaCode]);
    CREATE NONCLUSTERED INDEX [IX_AI_FieldHelp_Element] ON [dbo].[AI_FieldHelp]([PageURL], [ElementID]) WHERE [ElementID] IS NOT NULL;

    PRINT 'Created table: AI_FieldHelp';
END
GO

-- ============================================================================
-- 12. AI_ConversationSession - 對話 Session 表 (用於產生 ConversationID)
-- ============================================================================
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[AI_ConversationSession]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[AI_ConversationSession] (
        [ConversationID]    INT IDENTITY(1,1)   NOT NULL,
        [UserID]            VARCHAR(50)         NOT NULL,
        [PageURL]           NVARCHAR(500)       NULL,
        [MessageCount]      INT                 NOT NULL DEFAULT 0,
        [Status]            VARCHAR(20)         NOT NULL DEFAULT 'ACTIVE',  -- ACTIVE/CLOSED/SUBMITTED
        [IssueID]           INT                 NULL,           -- 若已提交 Issue
        [StartTime]         DATETIME            NOT NULL DEFAULT GETDATE(),
        [LastActiveTime]    DATETIME            NOT NULL DEFAULT GETDATE(),
        [EndTime]           DATETIME            NULL,
        CONSTRAINT [PK_AI_ConversationSession] PRIMARY KEY CLUSTERED ([ConversationID])
    );

    -- 索引
    CREATE NONCLUSTERED INDEX [IX_AI_ConversationSession_User] ON [dbo].[AI_ConversationSession]([UserID], [Status]);

    PRINT 'Created table: AI_ConversationSession';
END
GO

PRINT '';
PRINT '=============================================================================';
PRINT 'AI Helper 資料表建立完成';
PRINT '=============================================================================';
