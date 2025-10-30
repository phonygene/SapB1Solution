"""測試 MCP Server 功能

在正式配置到 Claude Code 前，使用此腳本測試各項功能。
"""

import asyncio
import sys
from pathlib import Path

# 加入專案路徑
sys.path.insert(0, str(Path(__file__).parent))

from src.database import DatabaseManager
from src.backup_manager import BackupManager


def print_section(title: str):
    """列印測試區段標題"""
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60)


async def test_connection():
    """測試資料庫連線"""
    print_section("測試 1: 資料庫連線")

    try:
        db = DatabaseManager()
        conn = db.connect()
        print("✓ 連線成功")

        size = db.get_db_size()
        print(f"✓ 資料庫大小: {size:.2f} MB")

        db.disconnect()
        print("✓ 連線關閉成功")

        return True
    except Exception as e:
        print(f"✗ 連線失敗: {e}")
        return False


async def test_query():
    """測試查詢功能"""
    print_section("測試 2: 查詢功能")

    try:
        db = DatabaseManager()

        # 查詢資料表列表
        print("\n查詢資料表列表（前 5 個）...")
        results = db.execute_query("""
            SELECT TOP 5 TABLE_NAME
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_TYPE = 'BASE TABLE'
            ORDER BY TABLE_NAME
        """)

        print(f"✓ 查詢成功，返回 {len(results)} 筆記錄")
        for row in results:
            print(f"  - {row['TABLE_NAME']}")

        db.disconnect()
        return True

    except Exception as e:
        print(f"✗ 查詢失敗: {e}")
        return False


async def test_table_info():
    """測試資料表資訊"""
    print_section("測試 3: 資料表資訊")

    try:
        db = DatabaseManager()

        # 先取得第一個資料表
        tables = db.list_tables()
        if not tables:
            print("✗ 沒有找到資料表")
            return False

        table_name = tables[0]
        print(f"\n取得資料表資訊: {table_name}")

        info = db.get_table_info(table_name)
        print(f"✓ 資料表: {info['table_name']}")
        print(f"  欄位數: {info['column_count']}")
        print(f"  主鍵: {', '.join(info['primary_keys']) if info['primary_keys'] else '無'}")

        print(f"\n  前 5 個欄位:")
        for col in info['columns'][:5]:
            pk_mark = " [PK]" if col['IS_PRIMARY_KEY'] else ""
            nullable = "NULL" if col['IS_NULLABLE'] == 'YES' else "NOT NULL"
            print(f"    - {col['COLUMN_NAME']}{pk_mark}: {col['DATA_TYPE']} {nullable}")

        db.disconnect()
        return True

    except Exception as e:
        print(f"✗ 測試失敗: {e}")
        return False


async def test_backup():
    """測試備份功能"""
    print_section("測試 4: 備份功能")

    try:
        db = DatabaseManager()
        backup_mgr = BackupManager()

        conn = db.connect()
        db_size = db.get_db_size()

        print(f"\n資料庫大小: {db_size:.2f} MB")

        # 取得當前策略
        strategy = backup_mgr.get_current_strategy(db_size)
        print(f"當前策略: {strategy.get('description', '未知')}")
        print(f"  流水備份上限: {strategy['rolling_limit']} 個")
        print(f"  每日備份保留: {strategy['daily_retain_days']} 天")

        # 測試建立備份
        print("\n建立測試備份...")
        success, msg = backup_mgr.create_backup(conn, db_size, "rolling")

        if success:
            print(f"✓ 備份建立成功")
            print(f"  檔案: {Path(msg).name}")
        else:
            print(f"✗ 備份失敗: {msg}")
            return False

        # 列出備份
        print("\n現有備份:")
        backups = backup_mgr.list_backups()

        if backups['rolling']:
            print(f"  流水備份 ({len(backups['rolling'])} 個):")
            for b in backups['rolling'][:3]:
                print(f"    - {b['name']} ({b['size_mb']:.2f} MB)")

        if backups['daily']:
            print(f"  每日備份 ({len(backups['daily'])} 個):")
            for b in backups['daily'][:3]:
                print(f"    - {b['name']} ({b['size_mb']:.2f} MB)")

        db.disconnect()
        return True

    except Exception as e:
        print(f"✗ 備份測試失敗: {e}")
        import traceback
        traceback.print_exc()
        return False


async def test_list_tables():
    """測試列出資料表"""
    print_section("測試 5: 列出所有資料表")

    try:
        db = DatabaseManager()
        tables = db.list_tables()

        print(f"✓ 找到 {len(tables)} 個資料表")

        if len(tables) > 10:
            print(f"\n前 10 個資料表:")
            for table in tables[:10]:
                print(f"  - {table}")
            print(f"  ... 還有 {len(tables) - 10} 個")
        else:
            print(f"\n所有資料表:")
            for table in tables:
                print(f"  - {table}")

        db.disconnect()
        return True

    except Exception as e:
        print(f"✗ 測試失敗: {e}")
        return False


async def main():
    """執行所有測試"""
    print("\n" + "╔" + "═" * 58 + "╗")
    print("║" + " " * 10 + "SAP B1 SQL MCP Server 測試程式" + " " * 17 + "║")
    print("╚" + "═" * 58 + "╝")

    # 檢查環境變數
    import os
    from dotenv import load_dotenv

    load_dotenv()

    required_vars = ["DB_SERVER", "DB_NAME", "DB_USER", "DB_PASSWORD"]
    missing_vars = [var for var in required_vars if not os.getenv(var)]

    if missing_vars:
        print("\n⚠️  缺少環境變數:")
        for var in missing_vars:
            print(f"  - {var}")
        print("\n請建立 .env 檔案並設定資料庫連線資訊")
        print("參考 .env.example 檔案")
        return

    # 執行測試
    tests = [
        ("連線測試", test_connection),
        ("查詢測試", test_query),
        ("資料表資訊測試", test_table_info),
        ("列出資料表測試", test_list_tables),
        ("備份測試", test_backup),
    ]

    results = []
    for name, test_func in tests:
        try:
            result = await test_func()
            results.append((name, result))
        except Exception as e:
            print(f"\n✗ {name} 發生未預期的錯誤: {e}")
            import traceback
            traceback.print_exc()
            results.append((name, False))

    # 測試結果總結
    print_section("測試結果總結")

    passed = 0
    failed = 0

    for name, result in results:
        status = "✓ 通過" if result else "✗ 失敗"
        print(f"{status} - {name}")
        if result:
            passed += 1
        else:
            failed += 1

    print(f"\n總計: {passed}/{len(results)} 測試通過")

    if failed == 0:
        print("\n🎉 所有測試通過！MCP Server 已準備好使用。")
        print("\n下一步:")
        print("1. 配置 .env 檔案（如果還沒有）")
        print("2. 將 MCP Server 配置到 Claude Code")
        print("3. 重啟 Claude Code")
    else:
        print("\n⚠️  部分測試失敗，請檢查錯誤訊息並修正問題。")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\n測試已中斷")
    except Exception as e:
        print(f"\n\n測試程式執行失敗: {e}")
        import traceback
        traceback.print_exc()
