-- 移除錯誤的外鍵約束
-- jMGUIAPDetail.jID 應指向 jMGUIAP，而非 jPCH1

-- 先查看現有的外鍵約束
SELECT 
    fk.name AS FK_Name,
    tp.name AS Parent_Table,
    cp.name AS Parent_Column,
    tr.name AS Referenced_Table,
    cr.name AS Referenced_Column
FROM sys.foreign_keys fk
INNER JOIN sys.tables tp ON fk.parent_object_id = tp.object_id
INNER JOIN sys.tables tr ON fk.referenced_object_id = tr.object_id
INNER JOIN sys.foreign_key_columns fkc ON fk.object_id = fkc.constraint_object_id
INNER JOIN sys.columns cp ON fkc.parent_column_id = cp.column_id AND fkc.parent_object_id = cp.object_id
INNER JOIN sys.columns cr ON fkc.referenced_column_id = cr.column_id AND fkc.referenced_object_id = cr.object_id
WHERE tp.name = 'jMGUIAPDetail'
ORDER BY FK_Name;

-- 移除錯誤的外鍵約束 (FK_jMGUIAPDetail_jPCH1)
IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = 'FK_jMGUIAPDetail_jPCH1')
BEGIN
    ALTER TABLE jMGUIAPDetail DROP CONSTRAINT FK_jMGUIAPDetail_jPCH1;
    PRINT 'FK_jMGUIAPDetail_jPCH1 has been dropped successfully.';
END
ELSE
BEGIN
    PRINT 'FK_jMGUIAPDetail_jPCH1 does not exist.';
END

-- 如果需要，新增正確的外鍵約束 (指向 jMGUIAP)
-- ALTER TABLE jMGUIAPDetail 
-- ADD CONSTRAINT FK_jMGUIAPDetail_jMGUIAP 
-- FOREIGN KEY (jID) REFERENCES jMGUIAP(ID);

