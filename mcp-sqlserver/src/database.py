"""資料庫操作模組

提供安全的 SQL Server 資料庫操作功能。
支援多資料庫動態切換。
"""

import pyodbc
import os
import json
import logging
from typing import List, Dict, Any, Tuple, Optional
from dotenv import load_dotenv

load_dotenv()
logger = logging.getLogger(__name__)


class DatabaseManager:
    """SQL Server 資料庫管理器

    特點：
    - 支援多資料庫動態切換
    - 安全的參數化查詢
    - 自動事務管理
    - SQL 注入防護
    - 操作記錄
    """

    def __init__(self, config_path: str = "config.json"):
        """初始化資料庫管理器

        Args:
            config_path: 配置檔案路徑
        """
        # 載入配置
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

        self.safe_mode = self.config["safe_mode"]

        # 多資料庫連線池
        self._connections: Dict[str, pyodbc.Connection] = {}
        self._current_db: Optional[str] = None

        # 資料庫配置
        self.databases = self.config.get("databases", {})
        self.default_database = self.config.get("default_database", "jtdb")

        logger.info(f"資料庫管理器初始化完成，可用資料庫: {list(self.databases.keys())}")

    def get_available_databases(self) -> List[Dict[str, str]]:
        """取得所有可用的資料庫列表

        Returns:
            資料庫資訊列表
        """
        return [
            {
                "name": name,
                "description": cfg.get("description", ""),
                "server": cfg.get("server", ""),
                "database": cfg.get("database", "")
            }
            for name, cfg in self.databases.items()
        ]

    def _get_db_config(self, db_name: Optional[str] = None) -> Dict[str, Any]:
        """取得資料庫配置

        Args:
            db_name: 資料庫名稱（必填）

        Returns:
            資料庫配置字典

        Raises:
            ValueError: 未指定資料庫或資料庫不存在時拋出
        """
        if db_name is None:
            available = list(self.databases.keys())
            raise ValueError(f"必須指定資料庫（db 參數）。可用選項: {available}")

        if db_name not in self.databases:
            available = list(self.databases.keys())
            raise ValueError(f"未知的資料庫: {db_name}，可用選項: {available}")

        return self.databases[db_name]

    def connect(self, db_name: Optional[str] = None) -> pyodbc.Connection:
        """建立或返回指定資料庫的連線

        Args:
            db_name: 資料庫名稱（必填）

        Returns:
            pyodbc 連線物件

        Raises:
            ValueError: 未指定資料庫時拋出
            Exception: 連線失敗時拋出異常
        """
        if db_name is None:
            available = list(self.databases.keys())
            raise ValueError(f"必須指定資料庫（db 參數）。可用選項: {available}")

        self._current_db = db_name

        # 檢查是否已有連線
        if db_name in self._connections:
            try:
                # 測試連線是否還有效
                self._connections[db_name].cursor().execute("SELECT 1")
                return self._connections[db_name]
            except:
                # 連線失效，關閉並重新連線
                try:
                    self._connections[db_name].close()
                except:
                    pass
                del self._connections[db_name]

        # 取得資料庫配置
        db_config = self._get_db_config(db_name)

        driver = db_config.get("driver", "ODBC Driver 17 for SQL Server")
        server = db_config.get("server", "localhost")
        port = db_config.get("port")
        database = db_config.get("database")
        username = db_config.get("username")
        password = db_config.get("password")

        if not all([database, username, password]):
            raise ValueError(f"資料庫 {db_name} 缺少必要的連線資訊")

        # 建立連線字串
        conn_str = f"DRIVER={{{driver}}};SERVER={server};"

        # 如果有指定 PORT，加入連線字串
        if port:
            conn_str += f"PORT={port};"

        conn_str += (
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password};"
            f"Encrypt=no;"
        )

        try:
            self._connections[db_name] = pyodbc.connect(conn_str, timeout=10)
            logger.info(f"成功連線至資料庫: {database} @ {server} (別名: {db_name})")
            return self._connections[db_name]
        except Exception as e:
            logger.error(f"資料庫連線失敗 ({db_name}): {str(e)}")
            raise

    def disconnect(self, db_name: Optional[str] = None):
        """關閉資料庫連線

        Args:
            db_name: 資料庫名稱，None 則關閉所有連線
        """
        if db_name is None:
            # 關閉所有連線
            for name, conn in list(self._connections.items()):
                try:
                    conn.close()
                    logger.info(f"資料庫連線已關閉: {name}")
                except Exception as e:
                    logger.warning(f"關閉連線時發生錯誤 ({name}): {str(e)}")
            self._connections.clear()
        elif db_name in self._connections:
            try:
                self._connections[db_name].close()
                logger.info(f"資料庫連線已關閉: {db_name}")
            except Exception as e:
                logger.warning(f"關閉連線時發生錯誤 ({db_name}): {str(e)}")
            finally:
                del self._connections[db_name]

    def get_db_size(self, db_name: Optional[str] = None) -> float:
        """取得資料庫大小（MB）

        Args:
            db_name: 資料庫名稱，None 則使用預設

        Returns:
            資料庫大小（MB）
        """
        try:
            conn = self.connect(db_name)
            cursor = conn.cursor()

            db_config = self._get_db_config(db_name)
            actual_db_name = db_config.get("database")
            query = f"""
                SELECT
                    SUM(CAST(size AS BIGINT)) * 8.0 / 1024 AS SizeMB
                FROM sys.master_files
                WHERE database_id = DB_ID('{actual_db_name}')
            """

            cursor.execute(query)
            row = cursor.fetchone()
            size_mb = float(row[0]) if row and row[0] else 0.0

            logger.debug(f"資料庫大小 ({db_name}): {size_mb:.2f} MB")
            return size_mb

        except Exception as e:
            logger.error(f"取得資料庫大小失敗: {str(e)}")
            return 0.0

    def execute_query(self, query: str, params: Optional[tuple] = None, db_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """執行查詢（SELECT）

        Args:
            query: SQL 查詢語句
            params: 參數（用於參數化查詢，防止 SQL Injection）
            db_name: 資料庫名稱，None 則使用預設

        Returns:
            查詢結果列表（字典格式）

        Raises:
            Exception: 查詢執行失敗時拋出異常
        """
        conn = self.connect(db_name)
        cursor = conn.cursor()

        try:
            logger.debug(f"執行查詢: {query[:100]}...")

            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)

            # 取得欄位名稱
            if cursor.description:
                columns = [column[0] for column in cursor.description]

                # 轉換為字典列表
                results = []
                for row in cursor.fetchall():
                    row_dict = {}
                    for i, value in enumerate(row):
                        # 處理特殊類型
                        if value is None:
                            row_dict[columns[i]] = None
                        elif isinstance(value, (bytes, bytearray)):
                            row_dict[columns[i]] = value.hex()
                        else:
                            row_dict[columns[i]] = value
                    results.append(row_dict)

                logger.info(f"查詢成功，返回 {len(results)} 筆記錄")
                return results
            else:
                logger.info("查詢成功，無返回結果")
                return []

        except Exception as e:
            logger.error(f"查詢執行失敗: {str(e)}")
            raise

    def execute_write(
        self,
        query: str,
        params: Optional[tuple] = None,
        backup_manager=None,
        db_name: Optional[str] = None
    ) -> Tuple[bool, str, int]:
        """執行寫入操作（INSERT, UPDATE, DELETE）

        在寫入前會自動建立備份。

        Args:
            query: SQL 語句
            params: 參數
            backup_manager: 備份管理器實例
            db_name: 資料庫名稱，None 則使用預設

        Returns:
            (是否成功, 訊息, 影響行數)
        """
        # 安全檢查
        if not self._is_safe_query(query):
            msg = "查詢包含不安全的關鍵字，操作已拒絕"
            logger.warning(f"{msg}: {query[:100]}")
            return False, msg, 0

        conn = self.connect(db_name)

        try:
            # 取得資料庫配置，檢查是否啟用備份
            db_config = self._get_db_config(db_name)
            backup_enabled = db_config.get("backup_enabled", True)

            # 寫入前自動備份
            if backup_manager and backup_enabled:
                db_size = self.get_db_size(db_name)

                # 檢查是否需要每日備份
                if backup_manager.should_create_daily_backup():
                    success, msg = backup_manager.create_backup(conn, db_size, "daily")
                    if success:
                        logger.info(f"已建立每日備份: {msg}")
                    else:
                        logger.warning(f"每日備份失敗: {msg}")

                # 建立流水備份
                success, msg = backup_manager.create_backup(conn, db_size, "rolling")
                if not success:
                    logger.warning(f"流水備份失敗: {msg}")

            # 執行寫入
            logger.info(f"執行寫入操作: {query[:100]}...")

            cursor = conn.cursor()

            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)

            affected_rows = cursor.rowcount

            # 檢查影響行數
            max_rows = self.safe_mode.get("max_affected_rows", 1000)
            if affected_rows > max_rows:
                conn.rollback()
                msg = f"影響行數 ({affected_rows}) 超過安全上限 ({max_rows})，操作已回滾"
                logger.warning(msg)
                return False, msg, 0

            # 提交事務
            conn.commit()

            msg = f"操作成功，影響 {affected_rows} 筆記錄"
            logger.info(msg)

            return True, msg, affected_rows

        except Exception as e:
            # 回滾事務
            try:
                conn.rollback()
            except:
                pass

            error_msg = f"寫入失敗: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return False, error_msg, 0

    def _is_safe_query(self, query: str) -> bool:
        """檢查查詢是否安全

        Args:
            query: SQL 查詢語句

        Returns:
            True 如果安全
        """
        query_upper = query.upper()

        # 檢查黑名單關鍵字
        blacklist = self.safe_mode.get("blacklist_keywords", [])
        for keyword in blacklist:
            if keyword.upper() in query_upper:
                logger.warning(f"查詢包含黑名單關鍵字: {keyword}")
                return False

        return True

    def get_table_info(self, table_name: str, db_name: Optional[str] = None) -> Dict[str, Any]:
        """取得資料表結構資訊

        Args:
            table_name: 資料表名稱
            db_name: 資料庫名稱，None 則使用預設

        Returns:
            包含表名和欄位資訊的字典
        """
        try:
            query = """
                SELECT
                    c.COLUMN_NAME,
                    c.DATA_TYPE,
                    c.IS_NULLABLE,
                    c.CHARACTER_MAXIMUM_LENGTH,
                    c.NUMERIC_PRECISION,
                    c.NUMERIC_SCALE,
                    c.COLUMN_DEFAULT
                FROM INFORMATION_SCHEMA.COLUMNS c
                WHERE c.TABLE_NAME = ?
                ORDER BY c.ORDINAL_POSITION
            """

            columns = self.execute_query(query, (table_name,), db_name)

            # 取得主鍵資訊
            pk_query = """
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE
                WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + CONSTRAINT_NAME), 'IsPrimaryKey') = 1
                AND TABLE_NAME = ?
            """

            pk_columns = self.execute_query(pk_query, (table_name,), db_name)
            pk_column_names = [row['COLUMN_NAME'] for row in pk_columns]

            # 標記主鍵
            for col in columns:
                col['IS_PRIMARY_KEY'] = col['COLUMN_NAME'] in pk_column_names

            logger.info(f"取得資料表資訊: {table_name} ({len(columns)} 個欄位)")

            return {
                "table_name": table_name,
                "column_count": len(columns),
                "columns": columns,
                "primary_keys": pk_column_names
            }

        except Exception as e:
            logger.error(f"取得資料表資訊失敗: {str(e)}")
            raise

    def list_tables(self, schema: str = "dbo", db_name: Optional[str] = None) -> List[str]:
        """列出所有資料表

        Args:
            schema: 結構描述名稱（預設 dbo）
            db_name: 資料庫名稱，None 則使用預設

        Returns:
            資料表名稱列表
        """
        try:
            query = """
                SELECT TABLE_NAME
                FROM INFORMATION_SCHEMA.TABLES
                WHERE TABLE_TYPE = 'BASE TABLE'
                AND TABLE_SCHEMA = ?
                ORDER BY TABLE_NAME
            """

            results = self.execute_query(query, (schema,), db_name)
            tables = [row['TABLE_NAME'] for row in results]

            logger.info(f"找到 {len(tables)} 個資料表 ({db_name or self.default_database})")
            return tables

        except Exception as e:
            logger.error(f"列出資料表失敗: {str(e)}")
            raise

    def execute_ddl(self, query: str, db_name: Optional[str] = None) -> Tuple[bool, str]:
        """執行 DDL 操作（CREATE, DROP, ALTER）

        Args:
            query: DDL SQL 語句
            db_name: 資料庫名稱，None 則使用預設

        Returns:
            (是否成功, 訊息)
        """
        conn = self.connect(db_name)

        try:
            logger.info(f"執行 DDL 操作: {query[:100]}...")

            cursor = conn.cursor()
            cursor.execute(query)

            # 提交 DDL 變更（pyodbc + FreeTDS 需要手動 commit）
            conn.commit()

            # 關閉連線，強制下次操作建立新連線以讀取 schema 變更
            try:
                conn.close()
                self._connection = None
            except:
                pass

            msg = f"DDL 操作成功"
            logger.info(msg)

            return True, msg

        except Exception as e:
            # 回滾事務
            try:
                conn.rollback()
            except:
                pass

            error_msg = f"DDL 執行失敗: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return False, error_msg
