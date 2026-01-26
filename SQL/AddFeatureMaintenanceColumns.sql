-- =====================================================
-- 功能級維護模式 - OADM 表擴展
-- 建立日期: 2026-01-19
-- =====================================================

-- 新增費用申請單維護開關
ALTER TABLE OADM ADD Maint_ExpenseClaim TINYINT NOT NULL DEFAULT 0;
GO

-- 新增請購單維護開關  
ALTER TABLE OADM ADD Maint_PurchaseRequest TINYINT NOT NULL DEFAULT 0;
GO

-- 驗證欄位
SELECT * FROM OADM;
