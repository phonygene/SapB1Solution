/*
=============================================================================
AI 輔助功能預設資料
目標資料庫：jtdb
建立日期：2025-12-24
說明：請在執行 AI_Helper_Schema.sql 後執行此腳本
=============================================================================
*/

-- ============================================================================
-- 1. AI_IssueType - Issue 類型預設資料
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM [dbo].[AI_IssueType] WHERE [TypeCode] = 'REQUIREMENT')
BEGIN
    INSERT INTO [dbo].[AI_IssueType] ([TypeCode], [TypeName], [IsActive], [SortOrder])
    VALUES
        ('REQUIREMENT', N'需求', 1, 1),
        ('BUG', N'Bug 回報', 1, 2),
        ('SUGGESTION', N'其他建議', 1, 3),
        ('VIOLATION', N'違規', 1, 4),
        ('AUTO_LIMIT', N'超過限制', 1, 5);
    PRINT 'Inserted default data: AI_IssueType';
END
GO

-- ============================================================================
-- 2. AI_IssueStatus - Issue 狀態預設資料
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM [dbo].[AI_IssueStatus] WHERE [StatusCode] = 'PENDING')
BEGIN
    INSERT INTO [dbo].[AI_IssueStatus] ([StatusCode], [StatusName], [IsActive], [SortOrder])
    VALUES
        ('PENDING', N'未確認', 1, 1),
        ('PROCESSING', N'處理中', 1, 2),
        ('INVALID', N'誤報', 1, 3),
        ('RESOLVED', N'已處理', 1, 4);
    PRINT 'Inserted default data: AI_IssueStatus';
END
GO

-- ============================================================================
-- 3. AI_Tag - Tag 預設資料（系統標籤）
-- ============================================================================
IF NOT EXISTS (SELECT 1 FROM [dbo].[AI_Tag] WHERE [IsSystem] = 1)
BEGIN
    INSERT INTO [dbo].[AI_Tag] ([TagName], [Category], [Color], [IsSystem], [IsActive], [SortOrder])
    VALUES
        -- 問題類型 (TYPE)
        (N'bug', 'TYPE', '#DC3545', 1, 1, 1),
        (N'功能需求', 'TYPE', '#28A745', 1, 1, 2),
        (N'功能改進', 'TYPE', '#17A2B8', 1, 1, 3),
        (N'操作疑問', 'TYPE', '#6C757D', 1, 1, 4),
        (N'文件說明', 'TYPE', '#FFC107', 1, 1, 5),
        (N'介面問題', 'TYPE', '#E83E8C', 1, 1, 6),
        (N'使用體驗', 'TYPE', '#6F42C1', 1, 1, 7),
        (N'效能問題', 'TYPE', '#FD7E14', 1, 1, 8),

        -- 功能模組 (MODULE) - 可依實際系統調整
        (N'費用申請', 'MODULE', '#007BFF', 1, 1, 10),
        (N'簽核流程', 'MODULE', '#007BFF', 1, 1, 11),
        (N'報表', 'MODULE', '#007BFF', 1, 1, 12),
        (N'匯出', 'MODULE', '#007BFF', 1, 1, 13),
        (N'登入權限', 'MODULE', '#007BFF', 1, 1, 14),
        (N'通知提醒', 'MODULE', '#007BFF', 1, 1, 15),
        (N'查詢功能', 'MODULE', '#007BFF', 1, 1, 16),
        (N'列印', 'MODULE', '#007BFF', 1, 1, 17),

        -- 問題區域 (AREA)
        (N'單頭', 'AREA', '#20C997', 1, 1, 20),
        (N'單身', 'AREA', '#20C997', 1, 1, 21),
        (N'單尾', 'AREA', '#20C997', 1, 1, 22),
        (N'彈窗', 'AREA', '#20C997', 1, 1, 23),
        (N'選單', 'AREA', '#20C997', 1, 1, 24),

        -- 嚴重程度 (SEVERITY)
        (N'嚴重', 'SEVERITY', '#DC3545', 1, 1, 30),
        (N'主要', 'SEVERITY', '#FD7E14', 1, 1, 31),
        (N'次要', 'SEVERITY', '#FFC107', 1, 1, 32),
        (N'輕微', 'SEVERITY', '#6C757D', 1, 1, 33),

        -- 狀態標記 (STATUS)
        (N'重複問題', 'STATUS', '#6C757D', 1, 1, 40),
        (N'不修復', 'STATUS', '#6C757D', 1, 1, 41),
        (N'無法重現', 'STATUS', '#6C757D', 1, 1, 42),
        (N'需要更多資訊', 'STATUS', '#17A2B8', 1, 1, 43),
        (N'已確認', 'STATUS', '#28A745', 1, 1, 44),

        -- 系統標記 (SYSTEM)
        (N'違規', 'SYSTEM', '#DC3545', 1, 1, 50),
        (N'自動提交', 'SYSTEM', '#6C757D', 1, 1, 51),
        (N'超過限制', 'SYSTEM', '#FFC107', 1, 1, 52);

    PRINT 'Inserted default data: AI_Tag';
END
GO

-- ============================================================================
-- 4. AI_FilterKeyword - 關鍵字過濾預設資料（範例）
-- ============================================================================
-- 注意：這裡只放範例，實際敏感詞請自行維護
IF NOT EXISTS (SELECT 1 FROM [dbo].[AI_FilterKeyword])
BEGIN
    INSERT INTO [dbo].[AI_FilterKeyword] ([Keyword], [MatchType], [IsActive], [CreateBy], [Remark])
    VALUES
        -- 以下為範例，請依實際需求調整
        (N'忽略上述指令', 'CONTAINS', 1, 'SYSTEM', N'防止 Prompt Injection'),
        (N'ignore previous', 'CONTAINS', 1, 'SYSTEM', N'防止 Prompt Injection'),
        (N'disregard', 'CONTAINS', 1, 'SYSTEM', N'防止 Prompt Injection'),
        (N'pretend you are', 'CONTAINS', 1, 'SYSTEM', N'防止角色扮演繞過'),
        (N'假裝你是', 'CONTAINS', 1, 'SYSTEM', N'防止角色扮演繞過'),
        (N'扮演', 'CONTAINS', 1, 'SYSTEM', N'防止角色扮演繞過'),
        (N'jailbreak', 'CONTAINS', 1, 'SYSTEM', N'防止越獄嘗試'),
        (N'DAN', 'EXACT', 1, 'SYSTEM', N'防止 DAN 越獄'),
        (N'開發者模式', 'CONTAINS', 1, 'SYSTEM', N'防止模式切換嘗試'),
        (N'developer mode', 'CONTAINS', 1, 'SYSTEM', N'防止模式切換嘗試');

    PRINT 'Inserted default data: AI_FilterKeyword (範例)';
END
GO

PRINT '';
PRINT '=============================================================================';
PRINT 'AI Helper 預設資料建立完成';
PRINT '=============================================================================';
PRINT '';
PRINT '注意事項：';
PRINT '1. AI_FilterKeyword 只包含範例資料，請依實際需求維護';
PRINT '2. AI_Tag 的 MODULE 類別可依實際系統功能調整';
PRINT '3. AI_PageArea 和 AI_FieldHelp 需在開發各頁面時逐一新增';
PRINT '=============================================================================';
