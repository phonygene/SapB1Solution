"""MCP Server 主程式

提供 SQL Server 資料庫操作的 MCP 工具集。
"""

import asyncio
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any
from mcp.server import Server
from mcp.types import Tool, TextContent
from mcp.server.stdio import stdio_server

# 加入專案根目錄到 Python 路徑
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.database import DatabaseManager
from src.backup_manager import BackupManager

# 設定日誌
log_dir = Path("logs")
log_dir.mkdir(exist_ok=True)

logging.basicConfig(
    level=getattr(logging, os.getenv("LOG_LEVEL", "INFO")),
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / 'mcp_server.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 建立 MCP Server
app = Server("sapb1-sql-mcp")

# 初始化管理器（延遲初始化）
db_manager: DatabaseManager = None
backup_manager: BackupManager = None


def init_managers():
    """初始化管理器"""
    global db_manager, backup_manager

    if db_manager is None:
        try:
            db_manager = DatabaseManager()
            backup_manager = BackupManager()
            logger.info("管理器初始化成功")
        except Exception as e:
            logger.error(f"管理器初始化失敗: {str(e)}", exc_info=True)
            raise


def get_db_param_schema() -> dict:
    """取得 db 參數的 schema 定義"""
    return {
        "type": "string",
        "description": "目標資料庫（可選）。jtdb=JET自有資料(j開頭表)、sapb1=SAP B1資料(O開頭表)。預設使用 jtdb。",
        "enum": ["jtdb", "sapb1"]
    }


@app.list_tools()
async def list_tools() -> list[Tool]:
    """列出所有可用工具"""
    db_param = get_db_param_schema()

    return [
        Tool(
            name="sql_query",
            description="執行 SQL 查詢（SELECT）。用於讀取資料，不會修改資料庫。支援參數化查詢以防止 SQL Injection。",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "SQL SELECT 查詢語句。例如：SELECT * FROM OITM WHERE ItemCode = ?"
                    },
                    "params": {
                        "type": "array",
                        "description": "查詢參數列表（可選）。用於參數化查詢，提高安全性。",
                        "items": {
                            "type": ["string", "number", "null", "boolean"]
                        }
                    },
                    "db": db_param
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="sql_write",
            description="執行 SQL 寫入操作（INSERT, UPDATE, DELETE）。會在寫入前自動建立備份，確保資料安全。",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "SQL 寫入語句。例如：INSERT INTO TestTable (Name, Value) VALUES (?, ?)"
                    },
                    "params": {
                        "type": "array",
                        "description": "查詢參數列表（可選）",
                        "items": {
                            "type": ["string", "number", "null", "boolean"]
                        }
                    },
                    "db": db_param
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="sql_ddl",
            description="執行 DDL 操作（CREATE, DROP, ALTER, TRUNCATE）。用於建立、刪除、修改資料表結構。DDL 語句會自動提交。",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "DDL SQL 語句。例如：CREATE TABLE test (id INT, name NVARCHAR(50))"
                    },
                    "db": db_param
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="get_table_info",
            description="取得資料表結構資訊，包括欄位名稱、型別、是否可為 NULL、主鍵等。",
            inputSchema={
                "type": "object",
                "properties": {
                    "table_name": {
                        "type": "string",
                        "description": "資料表名稱。例如：OITM（物料主檔）、OCRD（業務夥伴主檔）"
                    },
                    "db": db_param
                },
                "required": ["table_name"]
            }
        ),
        Tool(
            name="list_tables",
            description="列出資料庫中所有資料表名稱",
            inputSchema={
                "type": "object",
                "properties": {
                    "schema": {
                        "type": "string",
                        "description": "結構描述名稱（可選，預設為 dbo）",
                        "default": "dbo"
                    },
                    "db": db_param
                }
            }
        ),
        Tool(
            name="list_databases",
            description="列出所有可用的資料庫配置",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="list_backups",
            description="列出所有可用的備份檔案，包括流水備份和每日備份",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="restore_backup",
            description="從指定的備份檔案還原資料庫。⚠️ 注意：此操作會覆蓋當前資料庫！",
            inputSchema={
                "type": "object",
                "properties": {
                    "backup_name": {
                        "type": "string",
                        "description": "備份檔案名稱。例如：rolling_20250429_143022.bak 或 daily_20250429.bak"
                    }
                },
                "required": ["backup_name"]
            }
        ),
        Tool(
            name="get_db_status",
            description="取得資料庫狀態資訊，包括資料庫大小、當前備份策略、備份數量等",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="create_backup",
            description="手動建立資料庫備份",
            inputSchema={
                "type": "object",
                "properties": {
                    "backup_type": {
                        "type": "string",
                        "enum": ["rolling", "daily"],
                        "description": "備份類型：rolling（流水備份）或 daily（每日備份）",
                        "default": "rolling"
                    }
                }
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: Any) -> list[TextContent]:
    """處理工具調用"""
    try:
        # 確保管理器已初始化
        init_managers()

        # 取得目標資料庫（如有指定）
        db_name = arguments.get("db", None)

        if name == "sql_query":
            query = arguments["query"]
            params = tuple(arguments.get("params", []))

            results = db_manager.execute_query(query, params if params else None, db_name)

            return [TextContent(
                type="text",
                text=json.dumps(results, ensure_ascii=False, indent=2, default=str)
            )]

        elif name == "sql_write":
            query = arguments["query"]
            params = tuple(arguments.get("params", []))

            success, message, affected_rows = db_manager.execute_write(
                query,
                params if params else None,
                backup_manager,
                db_name
            )

            result = {
                "success": success,
                "message": message,
                "affected_rows": affected_rows,
                "database": db_name or db_manager.default_database
            }

            return [TextContent(
                type="text",
                text=json.dumps(result, ensure_ascii=False, indent=2)
            )]

        elif name == "sql_ddl":
            query = arguments["query"]

            success, message = db_manager.execute_ddl(query, db_name)

            result = {
                "success": success,
                "message": message,
                "database": db_name or db_manager.default_database
            }

            return [TextContent(
                type="text",
                text=json.dumps(result, ensure_ascii=False, indent=2)
            )]

        elif name == "get_table_info":
            table_name = arguments["table_name"]
            info = db_manager.get_table_info(table_name, db_name)
            info["database"] = db_name or db_manager.default_database

            return [TextContent(
                type="text",
                text=json.dumps(info, ensure_ascii=False, indent=2)
            )]

        elif name == "list_tables":
            schema = arguments.get("schema", "dbo")
            tables = db_manager.list_tables(schema, db_name)

            result = {
                "database": db_name or db_manager.default_database,
                "schema": schema,
                "table_count": len(tables),
                "tables": tables
            }

            return [TextContent(
                type="text",
                text=json.dumps(result, ensure_ascii=False, indent=2)
            )]

        elif name == "list_databases":
            databases = db_manager.get_available_databases()

            result = {
                "default": db_manager.default_database,
                "available": databases
            }

            return [TextContent(
                type="text",
                text=json.dumps(result, ensure_ascii=False, indent=2)
            )]

        elif name == "list_backups":
            backups = backup_manager.list_backups()

            result = {
                "rolling_backups": backups["rolling"],
                "daily_backups": backups["daily"],
                "total_count": len(backups["rolling"]) + len(backups["daily"])
            }

            return [TextContent(
                type="text",
                text=json.dumps(result, ensure_ascii=False, indent=2)
            )]

        elif name == "restore_backup":
            backup_name = arguments["backup_name"]
            conn = db_manager.connect()
            success, message = backup_manager.restore_backup(backup_name, conn)

            result = {
                "success": success,
                "message": message,
                "backup_name": backup_name
            }

            return [TextContent(
                type="text",
                text=json.dumps(result, ensure_ascii=False, indent=2)
            )]

        elif name == "get_db_status":
            db_size = db_manager.get_db_size()
            strategy = backup_manager.get_current_strategy(db_size)
            backups = backup_manager.list_backups()

            # 計算備份總大小
            total_backup_size = (
                sum(b["size_mb"] for b in backups["rolling"]) +
                sum(b["size_mb"] for b in backups["daily"])
            )

            status = {
                "database": {
                    "name": os.getenv("DB_NAME"),
                    "server": os.getenv("DB_SERVER"),
                    "size_mb": round(db_size, 2)
                },
                "current_strategy": {
                    "name": next(
                        (k for k, v in backup_manager.config["backup_strategies"].items() if v == strategy),
                        "unknown"
                    ),
                    "rolling_limit": strategy["rolling_limit"],
                    "daily_retain_days": strategy["daily_retain_days"],
                    "enabled": strategy["enabled"],
                    "description": strategy.get("description", "")
                },
                "backup_counts": {
                    "rolling": len(backups["rolling"]),
                    "daily": len(backups["daily"]),
                    "total": len(backups["rolling"]) + len(backups["daily"])
                },
                "total_backup_size_mb": round(total_backup_size, 2),
                "backup_enabled": os.getenv("BACKUP_ENABLED", "true").lower() == "true"
            }

            return [TextContent(
                type="text",
                text=json.dumps(status, ensure_ascii=False, indent=2)
            )]

        elif name == "create_backup":
            backup_type = arguments.get("backup_type", "rolling")
            db_size = db_manager.get_db_size()
            conn = db_manager.connect()

            success, message = backup_manager.create_backup(conn, db_size, backup_type)

            result = {
                "success": success,
                "message": message,
                "backup_type": backup_type,
                "database_size_mb": round(db_size, 2)
            }

            return [TextContent(
                type="text",
                text=json.dumps(result, ensure_ascii=False, indent=2)
            )]

        else:
            error_result = {"error": f"未知的工具: {name}"}
            return [TextContent(
                type="text",
                text=json.dumps(error_result, ensure_ascii=False)
            )]

    except Exception as e:
        logger.error(f"工具調用失敗 ({name}): {str(e)}", exc_info=True)

        error_result = {
            "error": str(e),
            "tool": name,
            "arguments": arguments
        }

        return [TextContent(
            type="text",
            text=json.dumps(error_result, ensure_ascii=False, indent=2)
        )]


async def main():
    """啟動 MCP Server"""
    logger.info("=" * 60)
    logger.info("正在啟動 SAP B1 SQL MCP Server...")
    logger.info(f"Python 版本: {sys.version}")
    logger.info(f"工作目錄: {os.getcwd()}")
    logger.info("=" * 60)

    try:
        # 初始化管理器
        init_managers()

        # 啟動 stdio server
        async with stdio_server() as (read_stream, write_stream):
            logger.info("MCP Server 已啟動，等待連線...")
            await app.run(
                read_stream,
                write_stream,
                app.create_initialization_options()
            )

    except KeyboardInterrupt:
        logger.info("收到中斷信號，正在關閉...")
    except Exception as e:
        logger.error(f"Server 執行失敗: {str(e)}", exc_info=True)
        raise
    finally:
        if db_manager:
            db_manager.disconnect()
        logger.info("MCP Server 已停止")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Server 已停止")
    except Exception as e:
        logger.error(f"啟動失敗: {str(e)}", exc_info=True)
        sys.exit(1)
