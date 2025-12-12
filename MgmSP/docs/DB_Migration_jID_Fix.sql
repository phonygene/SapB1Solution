-- =============================================================================
-- 資料庫遷移腳本：修正 jID 關聯結構
--
-- 目的：將四個資料表統一使用 jID 作為關聯鍵，DocEntry/DocNum 留給 SAP 回寫
-- 日期：2025/12/12
--
-- 執行順序：
-- 1. 先在測試環境執行並驗證
-- 2. 備份生產環境資料庫
-- 3. 在生產環境執行
-- =============================================================================

-- 開始交易
BEGIN TRANSACTION;

PRINT '========================================';
PRINT '步驟 0: 備份現有資料（查詢用）';
PRINT '========================================';

-- 備份查詢（不實際建立備份表，僅供確認）
SELECT 'jOPCH 現有資料' AS Info, COUNT(*) AS Cnt FROM jOPCH;
SELECT 'jPCH1 現有資料' AS Info, COUNT(*) AS Cnt FROM jPCH1;
SELECT 'jMGUIAP 現有資料' AS Info, COUNT(*) AS Cnt FROM jMGUIAP;
SELECT 'jMGUIAPDETAIL 現有資料' AS Info, COUNT(*) AS Cnt FROM jMGUIAPDETAIL;

PRINT '========================================';
PRINT '步驟 1: 刪除孤兒資料';
PRINT '========================================';

-- 刪除 jMGUIAPDETAIL 中找不到對應 jOPCH 的資料
-- 這些資料的 DocEntry 在 jOPCH 中不存在
DELETE FROM jMGUIAPDETAIL
WHERE DocEntry NOT IN (SELECT DocEntry FROM jOPCH WHERE DocEntry IS NOT NULL)
  AND DocEntry IS NOT NULL;

PRINT '已刪除 jMGUIAPDETAIL 孤兒資料: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 刪除 jMGUIAP 中找不到對應 jOPCH 的資料
DELETE FROM jMGUIAP
WHERE jID NOT IN (SELECT jID FROM jOPCH);

PRINT '已刪除 jMGUIAP 孤兒資料: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

PRINT '========================================';
PRINT '步驟 2: 修正 jMGUIAPDETAIL 的 jID';
PRINT '========================================';

-- 根據 DocEntry 找到對應的 jOPCH.jID，更新 jMGUIAPDETAIL.jID
-- 目前 jMGUIAPDETAIL.jID 存的是 jMGUIAP.ID（錯誤），應該改為 jOPCH.jID
UPDATE d
SET d.jID = o.jID
FROM jMGUIAPDETAIL d
INNER JOIN jOPCH o ON d.DocEntry = o.DocEntry
WHERE d.DocEntry IS NOT NULL;

PRINT '已修正 jMGUIAPDETAIL.jID: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

PRINT '========================================';
PRINT '步驟 3: 清空 DocEntry/DocNum 欄位';
PRINT '（這些欄位應該由 SAP 回寫，不應該預設值）';
PRINT '========================================';

-- 清空 jOPCH 的 DocEntry 和 DocNum
UPDATE jOPCH SET DocEntry = NULL, DocNum = NULL;
PRINT '已清空 jOPCH.DocEntry/DocNum: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 清空 jPCH1 的 DocEntry 和 DocNum
UPDATE jPCH1 SET DocEntry = NULL, DocNum = NULL;
PRINT '已清空 jPCH1.DocEntry/DocNum: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 清空 jMGUIAP 的 DocEntry 和 DocNum
UPDATE jMGUIAP SET DocEntry = NULL, DocNum = NULL;
PRINT '已清空 jMGUIAP.DocEntry/DocNum: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

-- 清空 jMGUIAPDETAIL 的 DocEntry
UPDATE jMGUIAPDETAIL SET DocEntry = NULL;
PRINT '已清空 jMGUIAPDETAIL.DocEntry: ' + CAST(@@ROWCOUNT AS VARCHAR(10)) + ' 筆';

PRINT '========================================';
PRINT '步驟 4: 修改 jMGUIAP 資料表結構';
PRINT '（刪除 ID 欄位，讓 jID 成為主鍵）';
PRINT '========================================';

-- 4.1 先刪除 ID 欄位上的主鍵約束
IF EXISTS (SELECT * FROM sys.key_constraints WHERE parent_object_id = OBJECT_ID('jMGUIAP') AND type = 'PK')
BEGIN
    DECLARE @pkName NVARCHAR(200);
    SELECT @pkName = name FROM sys.key_constraints WHERE parent_object_id = OBJECT_ID('jMGUIAP') AND type = 'PK';
    EXEC('ALTER TABLE jMGUIAP DROP CONSTRAINT ' + @pkName);
    PRINT '已刪除 jMGUIAP 主鍵約束: ' + @pkName;
END

-- 4.2 刪除 ID 欄位
IF EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID('jMGUIAP') AND name = 'ID')
BEGIN
    ALTER TABLE jMGUIAP DROP COLUMN ID;
    PRINT '已刪除 jMGUIAP.ID 欄位';
END

-- 4.3 將 jID 設為主鍵
ALTER TABLE jMGUIAP ADD CONSTRAINT PK_jMGUIAP PRIMARY KEY (jID);
PRINT '已將 jMGUIAP.jID 設為主鍵';

PRINT '========================================';
PRINT '步驟 5: 驗證結果';
PRINT '========================================';

-- 驗證 jMGUIAPDETAIL 的 jID 都能對應到 jOPCH
SELECT 'jMGUIAPDETAIL 無法對應的筆數' AS Info, COUNT(*) AS Cnt
FROM jMGUIAPDETAIL d
WHERE NOT EXISTS (SELECT 1 FROM jOPCH o WHERE o.jID = d.jID);

-- 驗證 jMGUIAP 的 jID 都能對應到 jOPCH
SELECT 'jMGUIAP 無法對應的筆數' AS Info, COUNT(*) AS Cnt
FROM jMGUIAP m
WHERE NOT EXISTS (SELECT 1 FROM jOPCH o WHERE o.jID = m.jID);

-- 顯示修正後的資料
SELECT 'jOPCH 修正後' AS Info, jID, DocEntry, DocNum, CardCode, ApprovalStatus FROM jOPCH ORDER BY jID;
SELECT 'jPCH1 修正後' AS Info, jID, LineNum, DocEntry, DocNum, AcctCode FROM jPCH1 ORDER BY jID, LineNum;
SELECT 'jMGUIAP 修正後' AS Info, jID, DocEntry, DocNum, DocTotal FROM jMGUIAP ORDER BY jID;
SELECT 'jMGUIAPDETAIL 修正後' AS Info, jID, LineNum, DocEntry, U_LIFNR, U_HWBAS FROM jMGUIAPDETAIL ORDER BY jID, LineNum;

PRINT '========================================';
PRINT '遷移完成！請檢查上方查詢結果';
PRINT '如果沒問題，請執行 COMMIT';
PRINT '如果有問題，請執行 ROLLBACK';
PRINT '========================================';

-- 確認無誤後執行：
-- COMMIT;

-- 如果有問題執行：
-- ROLLBACK;
