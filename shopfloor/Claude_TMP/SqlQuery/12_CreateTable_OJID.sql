IF OBJECT_ID('dbo.OJID','U') IS NULL
BEGIN
    CREATE TABLE dbo.OJID (
        jID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        jDate DATE NOT NULL CONSTRAINT DF_OJID_jDate DEFAULT (CONVERT(date, GETDATE())),
        jTime CHAR(8) NOT NULL CONSTRAINT DF_OJID_jTime DEFAULT (CONVERT(char(8), GETDATE(), 108)),
        jUser NVARCHAR(20) NOT NULL,
        UserIP NVARCHAR(45) NULL
    );
END;
GO

IF NOT EXISTS (
    SELECT 1
    FROM sys.indexes
    WHERE name = 'IX_OJID_jUser_jDate'
      AND object_id = OBJECT_ID('dbo.OJID')
)
BEGIN
    CREATE INDEX IX_OJID_jUser_jDate ON dbo.OJID (jUser, jDate);
END;
GO
