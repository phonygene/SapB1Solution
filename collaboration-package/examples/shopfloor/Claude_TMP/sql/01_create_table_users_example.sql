-- =====================================================================
--   建立使用者資料表
-- =====================================================================
-- 更新時間：2025-11-14
--
-- 用途說明：
-- 建立 users 資料表，用於儲存使用者帳號資訊
--
-- 使用方式：
-- 1. 在資料庫管理工具（pgAdmin / DBeaver）中開啟此檔案
-- 2. 連線到目標資料庫
-- 3. 執行此 SQL 腳本
--
-- 注意事項：
-- - 此範例適用於 PostgreSQL
-- - 如使用 MySQL，請將 SERIAL 改為 AUTO_INCREMENT
-- - 如使用 SQLite，請將 SERIAL 改為 INTEGER PRIMARY KEY AUTOINCREMENT
-- =====================================================================

-- 如果表格已存在則刪除（開發環境使用，生產環境請小心）
-- DROP TABLE IF EXISTS users CASCADE;

-- 建立 users 資料表
CREATE TABLE users (
    id SERIAL PRIMARY KEY,
    username VARCHAR(50) NOT NULL UNIQUE,
    email VARCHAR(100) NOT NULL UNIQUE,
    hashed_password VARCHAR(255) NOT NULL,
    is_active BOOLEAN NOT NULL DEFAULT TRUE,
    is_admin BOOLEAN NOT NULL DEFAULT FALSE,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

-- 建立索引以提升查詢效能
CREATE INDEX idx_users_username ON users(username);
CREATE INDEX idx_users_email ON users(email);
CREATE INDEX idx_users_active ON users(is_active);

-- 新增註解
COMMENT ON TABLE users IS '使用者帳號資料表';
COMMENT ON COLUMN users.id IS '使用者 ID（主鍵）';
COMMENT ON COLUMN users.username IS '使用者名稱（唯一）';
COMMENT ON COLUMN users.email IS '電子郵件（唯一）';
COMMENT ON COLUMN users.hashed_password IS '雜湊後的密碼';
COMMENT ON COLUMN users.is_active IS '帳號是否啟用';
COMMENT ON COLUMN users.is_admin IS '是否為管理員';
COMMENT ON COLUMN users.created_at IS '建立時間';
COMMENT ON COLUMN users.updated_at IS '更新時間';

-- 插入測試資料（可選）
-- INSERT INTO users (username, email, hashed_password)
-- VALUES ('admin', 'admin@example.com', 'hashed_password_here');

-- 驗證建立成功
SELECT * FROM users;

-- =====================================================================
-- 執行完成後，應該會看到：
-- - 資料表建立成功
-- - 3 個索引建立成功
-- - 查詢結果為空（如果沒有插入測試資料）
-- =====================================================================
