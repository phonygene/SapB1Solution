IF OBJECT_ID('dbo.UEDL', 'U') IS NULL
BEGIN
    CREATE TABLE dbo.UEDL (
        UserId NVARCHAR(50) NOT NULL,
        AcctCode NVARCHAR(50) NOT NULL,
        ItemCode NVARCHAR(50) NULL,
        CostingCode2 NVARCHAR(50) NULL,
        jID INT NULL,
        LineNum INT NULL,
        expDate DATETIME NOT NULL CONSTRAINT DF_UEDL_expDate DEFAULT (GETDATE()),
        CONSTRAINT PK_UEDL PRIMARY KEY (UserId, AcctCode)
    );
END;
