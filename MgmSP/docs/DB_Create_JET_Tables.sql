-- =====================================================
-- JET 自有資料表建立腳本
-- 包含：jOPCH（費用申請單）、jPCH1（費用明細）
--       jOPRQ（請購單）、jPRQ1（請購明細）
-- 日期：2026-01-15
-- =====================================================
-- 執行前請確認：
--   1. 已連接到 jtdb 資料庫
--   2. 有足夠權限建立資料表
-- =====================================================

USE jtdb;
GO

-- =====================================================
-- 第一部分：jOPCH 費用申請單表頭
-- =====================================================

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'jOPCH')
BEGIN
    CREATE TABLE jOPCH (
        -- 主鍵
        jID INT IDENTITY(1,1) PRIMARY KEY,

        -- 供應商資訊
        CardCode NVARCHAR(50) NULL,          -- 供應商代碼
        CardName NVARCHAR(200) NULL,         -- 供應商名稱
        NumAtCard NVARCHAR(50) NULL,         -- 供應商參考編號
        InvNum NVARCHAR(50) NULL,            -- 發票號碼

        -- 地址資訊
        DeliveryAddrID NVARCHAR(50) NULL,    -- 收貨地址 ID
        AddressName NVARCHAR(100) NULL,      -- 地址名稱
        Address NVARCHAR(500) NULL,          -- 完整地址

        -- 日期
        DocDate DATE NULL,                   -- 單據日期
        DocDueDate DATE NULL,                -- 到期日期
        TaxDate DATE NULL,                   -- 稅務日期

        -- 金額
        DocCurrency NVARCHAR(10) NULL,       -- 幣別
        DocRate DECIMAL(18,6) NULL,          -- 匯率
        DocTotal DECIMAL(18,2) NULL,         -- 單據總額（未稅）
        VatSum DECIMAL(18,2) NULL,           -- 稅額

        -- 付款條件
        GroupNum INT NULL,                   -- 付款條件群組
        PymntGroup INT NULL,                 -- 付款群組

        -- 備註
        Comments NVARCHAR(MAX) NULL,         -- 備註
        U_PID NVARCHAR(50) NULL,             -- 專案代碼
        SlpCode INT NULL,                    -- 業務員代碼

        -- 狀態（與 jOPRQ 統一）
        ApprovalStatus NVARCHAR(1) DEFAULT 'W',  -- W=待審核, A=已核准, R=已退回

        -- 審核資訊
        ApprovedBy NVARCHAR(50) NULL,        -- 審核人
        ApprovalDate DATETIME NULL,          -- 審核日期
        ApprovalComments NVARCHAR(500) NULL, -- 審核意見

        -- SAP 整合
        B1PostStatus NVARCHAR(1) DEFAULT 'N', -- N=未過帳, Y=已過帳, E=錯誤
        B1PostDate DATETIME NULL,            -- SAP 過帳日期
        B1ErrMsg NVARCHAR(500) NULL,         -- SAP 錯誤訊息
        DocEntry INT NULL,                   -- SAP DocEntry

        -- 異動紀錄
        CreateDate DATETIME DEFAULT GETDATE(),
        CreateBy NVARCHAR(50) NULL,
        UpdateDate DATETIME NULL,
        UpdateBy NVARCHAR(50) NULL
    );

    PRINT '✓ jOPCH 表已建立';
END
ELSE
    PRINT '- jOPCH 表已存在';
GO

-- =====================================================
-- 第二部分：jPCH1 費用申請單明細
-- =====================================================

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'jPCH1')
BEGIN
    CREATE TABLE jPCH1 (
        -- 主鍵
        ID INT IDENTITY(1,1) PRIMARY KEY,
        jID INT NOT NULL,                    -- 關聯 jOPCH.jID
        LineNum INT NOT NULL,                -- 行號

        -- 品項資訊
        ItemCode NVARCHAR(50) NULL,          -- 品號
        Dscription NVARCHAR(200) NULL,       -- 說明
        AcctCode NVARCHAR(50) NULL,          -- 會計科目

        -- 金額
        LineTotal DECIMAL(18,2) NULL,        -- 行小計（未稅）
        VatGroup NVARCHAR(10) NULL,          -- 稅碼
        VatPrcnt DECIMAL(5,2) NULL,          -- 稅率
        LineVat DECIMAL(18,2) NULL,          -- 行稅額
        GTotal DECIMAL(18,2) NULL,           -- 行小計（含稅）

        -- 成本中心
        CostingCode NVARCHAR(50) NULL,       -- 成本中心 1
        CostingCode2 NVARCHAR(50) NULL,      -- 成本中心 2

        -- 幣別
        Currency NVARCHAR(10) NULL,          -- 幣別
        Rate DECIMAL(18,6) NULL,             -- 匯率

        -- SAP 回寫
        DocEntry INT NULL                    -- SAP DocEntry
    );

    -- 建立索引
    CREATE INDEX IX_jPCH1_jID ON jPCH1(jID);

    PRINT '✓ jPCH1 表已建立';
END
ELSE
    PRINT '- jPCH1 表已存在';
GO

-- =====================================================
-- 第三部分：jOPRQ 請購單表頭
-- =====================================================

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'jOPRQ')
BEGIN
    CREATE TABLE jOPRQ (
        -- 主鍵
        jID INT IDENTITY(1,1) PRIMARY KEY,

        -- 供應商資訊
        CardCode NVARCHAR(50) NULL,          -- 供應商代碼
        CardName NVARCHAR(200) NULL,         -- 供應商名稱

        -- 請購人資訊
        ReqCode NVARCHAR(50) NULL,           -- 請購人代碼
        ReqName NVARCHAR(100) NULL,          -- 請購人姓名
        ReqDept NVARCHAR(100) NULL,          -- 請購部門
        SlpCode INT NULL,                    -- 採購人員代碼

        -- 日期
        DocDate DATE NULL,                   -- 單據日期
        ReqDate DATE NULL,                   -- 需求日期

        -- 金額
        DocCurrency NVARCHAR(10) NULL,       -- 幣別
        DocRate DECIMAL(18,6) NULL,          -- 匯率
        DocTotal DECIMAL(18,2) NULL,         -- 單據總額（未稅）
        VatSum DECIMAL(18,2) NULL,           -- 稅額

        -- 備註
        Comments NVARCHAR(MAX) NULL,         -- 備註
        U_PID NVARCHAR(50) NULL,             -- 專案代碼

        -- 狀態（與 jOPCH 統一）
        DocStatus NVARCHAR(1) DEFAULT 'O',   -- O=Open, C=Closed
        ApprovalStatus NVARCHAR(1) DEFAULT 'W',  -- W=待審核, A=已核准, R=已退回

        -- 審核資訊
        ApprovedBy NVARCHAR(50) NULL,        -- 審核人
        ApprovedDate DATETIME NULL,          -- 審核日期
        ApprovalComments NVARCHAR(500) NULL, -- 審核意見

        -- SAP 整合（與 jOPCH 統一）
        B1PostStatus NVARCHAR(1) DEFAULT 'N', -- N=未過帳, Y=已過帳, E=錯誤
        B1PostDate DATETIME NULL,            -- SAP 過帳日期
        B1ErrMsg NVARCHAR(500) NULL,         -- SAP 錯誤訊息
        DocEntry INT NULL,                   -- SAP DocEntry
        DocNum INT NULL,                     -- SAP DocNum

        -- 異動紀錄
        CreateDate DATETIME DEFAULT GETDATE(),
        CreateBy NVARCHAR(50) NULL,
        UpdateDate DATETIME NULL,
        UpdateBy NVARCHAR(50) NULL
    );

    PRINT '✓ jOPRQ 表已建立';
END
ELSE
    PRINT '- jOPRQ 表已存在';
GO

-- =====================================================
-- 第四部分：jPRQ1 請購單明細
-- =====================================================

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'jPRQ1')
BEGIN
    CREATE TABLE jPRQ1 (
        -- 主鍵
        ID INT IDENTITY(1,1) PRIMARY KEY,
        jID INT NOT NULL,                    -- 關聯 jOPRQ.jID
        LineNum INT NOT NULL,                -- 行號

        -- 品項資訊
        ItemCode NVARCHAR(50) NULL,          -- 品號
        Dscription NVARCHAR(200) NULL,       -- 品名（系統）
        U_Linetext NVARCHAR(500) NULL,       -- 自訂說明

        -- 數量與金額
        Quantity DECIMAL(18,6) NULL,         -- 數量
        Price DECIMAL(18,6) NULL,            -- 未稅單價
        PriceAfVAT DECIMAL(18,6) NULL,       -- 含稅單價
        LineTotal DECIMAL(18,2) NULL,        -- 行小計（未稅）
        GTotal DECIMAL(18,2) NULL,           -- 行小計（含稅）

        -- 稅務
        VatGroup NVARCHAR(10) NULL,          -- 稅碼
        VatPrcnt DECIMAL(5,2) NULL,          -- 稅率
        LineVat DECIMAL(18,2) NULL,          -- 行稅額

        -- 倉庫與交期
        WhsCode NVARCHAR(20) NULL,           -- 倉庫
        ShipDate DATE NULL,                  -- 交貨日期

        -- 成本中心
        CostingCode NVARCHAR(50) NULL,       -- 成本中心 1
        CostingCode2 NVARCHAR(50) NULL,      -- 成本中心 2

        -- 幣別
        Currency NVARCHAR(10) NULL,          -- 幣別
        Rate DECIMAL(18,6) NULL,             -- 匯率

        -- 狀態
        LineStatus NVARCHAR(1) DEFAULT 'O',  -- O=Open, C=Closed

        -- SAP 回寫
        DocEntry INT NULL,                   -- SAP DocEntry
        DocNum INT NULL,                     -- SAP DocNum

        -- 異動紀錄
        CreateDate DATETIME DEFAULT GETDATE(),
        CreateBy NVARCHAR(50) NULL
    );

    -- 建立索引
    CREATE INDEX IX_jPRQ1_jID ON jPRQ1(jID);

    PRINT '✓ jPRQ1 表已建立';
END
ELSE
    PRINT '- jPRQ1 表已存在';
GO

-- =====================================================
-- 第五部分：建立外鍵關聯
-- =====================================================

-- jPCH1 -> jOPCH
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS
               WHERE CONSTRAINT_NAME = 'FK_jPCH1_jOPCH')
BEGIN
    ALTER TABLE jPCH1
    ADD CONSTRAINT FK_jPCH1_jOPCH
    FOREIGN KEY (jID) REFERENCES jOPCH(jID) ON DELETE CASCADE;

    PRINT '✓ 外鍵 FK_jPCH1_jOPCH 已建立';
END
ELSE
    PRINT '- 外鍵 FK_jPCH1_jOPCH 已存在';
GO

-- jPRQ1 -> jOPRQ
IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS
               WHERE CONSTRAINT_NAME = 'FK_jPRQ1_jOPRQ')
BEGIN
    ALTER TABLE jPRQ1
    ADD CONSTRAINT FK_jPRQ1_jOPRQ
    FOREIGN KEY (jID) REFERENCES jOPRQ(jID) ON DELETE CASCADE;

    PRINT '✓ 外鍵 FK_jPRQ1_jOPRQ 已建立';
END
ELSE
    PRINT '- 外鍵 FK_jPRQ1_jOPRQ 已存在';
GO

-- =====================================================
-- 完成
-- =====================================================

PRINT '';
PRINT '========================================';
PRINT '建表完成！';
PRINT '========================================';
PRINT '';
PRINT '資料表摘要：';
PRINT '  jOPCH  - 費用申請單表頭';
PRINT '  jPCH1  - 費用申請單明細';
PRINT '  jOPRQ  - 請購單表頭';
PRINT '  jPRQ1  - 請購單明細';
PRINT '';
PRINT '狀態欄位對照：';
PRINT '  ApprovalStatus: W=待審核, A=已核准, R=已退回';
PRINT '  B1PostStatus:   N=未過帳, Y=已過帳, E=錯誤';
PRINT '========================================';

-- =====================================================
-- 欄位對照表
-- =====================================================
/*
jOPCH vs jOPRQ 共同欄位：
| 欄位             | jOPCH | jOPRQ | 說明                |
|------------------|-------|-------|---------------------|
| jID              | ✓     | ✓     | 主鍵                |
| CardCode         | ✓     | ✓     | 供應商代碼          |
| CardName         | ✓     | ✓     | 供應商名稱          |
| DocDate          | ✓     | ✓     | 單據日期            |
| DocCurrency      | ✓     | ✓     | 幣別                |
| DocRate          | ✓     | ✓     | 匯率                |
| DocTotal         | ✓     | ✓     | 總額（未稅）        |
| VatSum           | ✓     | ✓     | 稅額                |
| Comments         | ✓     | ✓     | 備註                |
| U_PID            | ✓     | ✓     | 專案代碼            |
| SlpCode          | ✓     | ✓     | 業務員/採購員代碼   |
| ApprovalStatus   | ✓     | ✓     | W/A/R               |
| ApprovedBy       | ✓     | ✓     | 審核人              |
| ApprovalDate     | ✓     | ✓*    | 審核日期 (*jOPRQ用ApprovedDate) |
| ApprovalComments | ✓     | ✓     | 審核意見            |
| B1PostStatus     | ✓     | ✓     | N/Y/E               |
| B1PostDate       | ✓     | ✓     | SAP 過帳日期        |
| B1ErrMsg         | ✓     | ✓     | SAP 錯誤訊息        |
| DocEntry         | ✓     | ✓     | SAP DocEntry        |
| CreateDate       | ✓     | ✓     | 建立日期            |
| CreateBy         | ✓     | ✓     | 建立者              |
| UpdateDate       | ✓     | ✓     | 更新日期            |
| UpdateBy         | ✓     | ✓     | 更新者              |

jOPCH 特有欄位：
| 欄位             | 說明                |
|------------------|---------------------|
| NumAtCard        | 供應商參考編號      |
| InvNum           | 發票號碼            |
| DeliveryAddrID   | 收貨地址 ID         |
| AddressName      | 地址名稱            |
| Address          | 完整地址            |
| DocDueDate       | 到期日期            |
| TaxDate          | 稅務日期            |
| GroupNum         | 付款條件群組        |
| PymntGroup       | 付款群組            |

jOPRQ 特有欄位：
| 欄位             | 說明                |
|------------------|---------------------|
| ReqCode          | 請購人代碼          |
| ReqName          | 請購人姓名          |
| ReqDept          | 請購部門            |
| ReqDate          | 需求日期            |
| DocStatus        | 文件狀態 O/C        |
| DocNum           | SAP 單據號碼        |
*/
