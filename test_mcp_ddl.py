#!/usr/bin/env python3
"""
測試 MCP DDL 功能
直接模擬 MCP Server 的 DDL 執行邏輯
"""

import pyodbc
import os

# === 連線設定（從 .env 讀取）===
DB_DRIVER = "FreeTDS"
DB_SERVER = "172.17.16.1"
DB_PORT = "12948"
DB_NAME = "jtdb"
DB_USER = "sa"
DB_PASSWORD = "sap19690123"

def test_ddl():
    """測試 DDL 執行"""

    # 建立連線字串（與 MCP Server 相同）
    conn_str = f"DRIVER={{{DB_DRIVER}}};SERVER={DB_SERVER};"

    if DB_PORT:
        conn_str += f"PORT={DB_PORT};"

    conn_str += (
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASSWORD};"
        f"Encrypt=no;"
    )

    print("=" * 70)
    print("測試 MCP DDL 功能")
    print("=" * 70)
    print(f"\n連線字串: {conn_str}\n")

    try:
        # 連線到資料庫
        print("⏳ 連接資料庫...")
        conn = pyodbc.connect(conn_str, timeout=10)
        print("✅ 連線成功\n")

        # === 測試 1: 檢查 addr 表是否存在 ===
        print("-" * 70)
        print("測試 1: 檢查 addr 表是否存在")
        print("-" * 70)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.tables WHERE name = 'addr'")
        result = cursor.fetchall()

        if result:
            print(f"❌ addr 表已存在，先刪除...")
            cursor.execute("DROP TABLE addr")
            conn.commit()
            print("✅ addr 表已刪除\n")
        else:
            print("✅ addr 表不存在，可以建立\n")

        # === 測試 2: 建立 addr 表（MCP 方式）===
        print("-" * 70)
        print("測試 2: 建立 addr 表（模擬 MCP DDL）")
        print("-" * 70)

        ddl_query = """CREATE TABLE [dbo].[addr] (
    [ID]         INT IDENTITY(1,1) PRIMARY KEY,
    [addrType]   CHAR(1) NOT NULL DEFAULT 'R',
    [addrName]   NVARCHAR(50) NOT NULL,
    [address]    NVARCHAR(254) NOT NULL,
    [active]     CHAR(1) NOT NULL DEFAULT 'Y',
    [createDate] DATETIME DEFAULT GETDATE(),
    [updateDate] DATETIME DEFAULT GETDATE(),

    CONSTRAINT CK_addr_addrType CHECK (addrType IN ('D', 'R')),
    CONSTRAINT CK_addr_active CHECK (active IN ('Y', 'N'))
)"""

        print(f"執行 DDL: {ddl_query[:100]}...")
        cursor = conn.cursor()
        cursor.execute(ddl_query)

        print("⏳ 執行 commit()...")
        conn.commit()
        print("✅ commit() 完成")

        print("⏳ 關閉連線...")
        conn.close()
        print("✅ 連線已關閉\n")

        # === 測試 3: 重新連線並驗證 ===
        print("-" * 70)
        print("測試 3: 重新連線並驗證表是否存在")
        print("-" * 70)

        print("⏳ 建立新連線...")
        conn = pyodbc.connect(conn_str, timeout=10)
        print("✅ 新連線建立成功")

        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sys.tables WHERE name = 'addr'")
        result = cursor.fetchall()

        if result:
            print(f"✅ 驗證成功！addr 表已建立: {result}")
        else:
            print("❌ 驗證失敗！addr 表不存在")
            return False

        # === 測試 4: 查詢表結構 ===
        print("\n" + "-" * 70)
        print("測試 4: 查詢表結構")
        print("-" * 70)

        cursor.execute("""
            SELECT COLUMN_NAME, DATA_TYPE, IS_NULLABLE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = 'addr'
            ORDER BY ORDINAL_POSITION
        """)
        columns = cursor.fetchall()

        print("addr 表欄位:")
        for col in columns:
            print(f"  - {col[0]}: {col[1]} (Nullable: {col[2]})")

        # === 測試 5: 插入資料 ===
        print("\n" + "-" * 70)
        print("測試 5: 插入測試資料")
        print("-" * 70)

        insert_query = """INSERT INTO addr (addrType, addrName, address) VALUES
('R', '總公司', '台北市信義區信義路五段7號'),
('R', '台中倉庫', '台中市西屯區台灣大道三段99號'),
('R', '高雄辦公室', '高雄市前鎮區成功二路88號')"""

        cursor.execute(insert_query)
        affected = cursor.rowcount
        conn.commit()
        print(f"✅ 插入成功，影響 {affected} 筆記錄")

        # === 測試 6: 查詢資料 ===
        print("\n" + "-" * 70)
        print("測試 6: 查詢資料")
        print("-" * 70)

        cursor.execute("SELECT ID, addrName, address FROM addr")
        rows = cursor.fetchall()

        print(f"addr 表資料 ({len(rows)} 筆):")
        for row in rows:
            print(f"  ID={row[0]}, 名稱={row[1]}, 地址={row[2]}")

        # 關閉連線
        conn.close()

        print("\n" + "=" * 70)
        print("✅ 所有測試通過！")
        print("=" * 70)

        return True

    except Exception as e:
        print(f"\n❌ 錯誤: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_ddl()
    exit(0 if success else 1)
