-- =============================================
-- 2026-01-12 OJID 外鍵約束
-- 目的：防止單據 jID 繞過 OJID 序號控制
-- =============================================

-- 1. 檢查資料完整性（執行約束前先確認）
SELECT 'jOPRQ' AS TableName, jID FROM jOPRQ WHERE jID NOT IN (SELECT jID FROM OJID)
UNION ALL
SELECT 'jOPCH' AS TableName, jID FROM jOPCH WHERE jID NOT IN (SELECT jID FROM OJID);
-- 如果有結果，需要先修復這些資料再繼續

-- 2. 建立外鍵約束
ALTER TABLE jOPRQ
ADD CONSTRAINT FK_jOPRQ_OJID
FOREIGN KEY (jID) REFERENCES OJID(jID);

ALTER TABLE jOPCH
ADD CONSTRAINT FK_jOPCH_OJID
FOREIGN KEY (jID) REFERENCES OJID(jID);

-- 3. 驗證約束已建立
SELECT
    fk.name AS ConstraintName,
    OBJECT_NAME(fk.parent_object_id) AS TableName,
    COL_NAME(fkc.parent_object_id, fkc.parent_column_id) AS ColumnName,
    OBJECT_NAME(fk.referenced_object_id) AS ReferencedTable
FROM sys.foreign_keys fk
INNER JOIN sys.foreign_key_columns fkc ON fk.object_id = fkc.constraint_object_id
WHERE fk.name LIKE 'FK_%_OJID';
