"""備份管理模組

提供智能備份策略，根據資料庫大小自動調整備份方案。
"""

import os
import json
import logging
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Tuple, Dict, List

logger = logging.getLogger(__name__)


class BackupManager:
    """資料庫備份管理器

    特點：
    - 流水備份（Rolling Backup）：保留最近 N 個備份
    - 每日備份（Daily Backup）：每天首次操作時建立，保留 N 天
    - 智能策略：根據資料庫大小自動調整備份數量
    """

    def __init__(self, config_path: str = "config.json"):
        """初始化備份管理器

        Args:
            config_path: 配置檔案路徑
        """
        # 載入配置
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

        # 設定備份目錄
        self.backup_dir = Path(os.getenv("BACKUP_DIR", "./backups"))
        self.rolling_dir = self.backup_dir / "rolling"
        self.daily_dir = self.backup_dir / "daily"

        # 確保目錄存在
        self.rolling_dir.mkdir(parents=True, exist_ok=True)
        self.daily_dir.mkdir(parents=True, exist_ok=True)

        # 記錄最後一次每日備份日期
        self._last_daily_backup_date = self._get_last_daily_backup_date()

        logger.info(f"備份管理器初始化完成，備份目錄: {self.backup_dir}")

    def get_current_strategy(self, db_size_mb: float) -> Dict:
        """根據資料庫大小選擇適合的備份策略

        Args:
            db_size_mb: 資料庫大小（MB）

        Returns:
            備份策略配置
        """
        strategies = self.config["backup_strategies"]

        if db_size_mb < strategies["small"]["max_size_mb"]:
            strategy_name = "small"
        elif db_size_mb < strategies["medium"]["max_size_mb"]:
            strategy_name = "medium"
        elif strategies["large"]["max_size_mb"] and db_size_mb < strategies["large"]["max_size_mb"]:
            strategy_name = "large"
        else:
            strategy_name = "very_large"

        strategy = strategies[strategy_name]
        logger.info(f"資料庫大小 {db_size_mb:.2f} MB，使用策略: {strategy_name}")

        return strategy

    def should_create_daily_backup(self) -> bool:
        """檢查今天是否需要建立每日備份

        Returns:
            True 如果需要建立每日備份
        """
        today = datetime.now().date()
        need_backup = self._last_daily_backup_date != today

        if need_backup:
            logger.info("今日尚未建立每日備份")

        return need_backup

    def create_backup(
        self,
        db_connection,
        db_size_mb: float,
        backup_type: str = "rolling"
    ) -> Tuple[bool, str]:
        """建立資料庫備份

        Args:
            db_connection: pyodbc 資料庫連線
            db_size_mb: 資料庫大小（MB）
            backup_type: "rolling"（流水備份）或 "daily"（每日備份）

        Returns:
            (是否成功, 訊息或備份檔案路徑)
        """
        try:
            # 取得當前策略
            strategy = self.get_current_strategy(db_size_mb)

            # 檢查策略是否啟用
            if not strategy.get("enabled", True):
                msg = f"當前策略已停用自動備份（{strategy.get('description', '')}）"
                logger.warning(msg)
                return False, msg

            # 生成備份檔名
            if backup_type == "daily":
                timestamp = datetime.now().strftime("%Y%m%d")
                backup_file = self.daily_dir / f"daily_{timestamp}.bak"

                # 檢查今日備份是否已存在
                if backup_file.exists():
                    msg = f"今日備份已存在: {backup_file.name}"
                    logger.info(msg)
                    return True, str(backup_file)
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_file = self.rolling_dir / f"rolling_{timestamp}.bak"

            # 執行 SQL Server 備份
            db_name = os.getenv("DB_NAME")
            # 使用字串格式而非 Path.absolute()，因為 SQL Server 運行在 Windows 上
            # 移除 COMPRESSION（Express Edition 不支援）
            backup_sql = f"""
                BACKUP DATABASE [{db_name}]
                TO DISK = '{str(backup_file)}'
                WITH FORMAT, STATS = 10;
            """

            logger.info(f"開始建立備份: {backup_file.name}")

            # 設置 autocommit 以避免事務中執行備份的錯誤
            old_autocommit = db_connection.autocommit
            db_connection.autocommit = True

            cursor = db_connection.cursor()
            cursor.execute(backup_sql)

            # 等待備份完成
            while cursor.nextset():
                pass

            # 恢復原始 autocommit 設定
            db_connection.autocommit = old_autocommit

            # 嘗試取得檔案大小（如果在 WSL 無法存取 Windows 路徑則跳過）
            try:
                file_size_mb = backup_file.stat().st_size / (1024*1024)
                logger.info(f"備份建立成功: {backup_file.name} ({file_size_mb:.2f} MB)")
            except (FileNotFoundError, OSError):
                logger.info(f"備份建立成功: {backup_file.name}")

            # 更新最後每日備份日期
            if backup_type == "daily":
                self._last_daily_backup_date = datetime.now().date()

            # 清理舊備份
            self._cleanup_old_backups(strategy, backup_type)

            return True, str(backup_file)

        except Exception as e:
            error_msg = f"備份失敗: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return False, error_msg

    def _cleanup_old_backups(self, strategy: Dict, backup_type: str):
        """清理舊備份檔案

        Args:
            strategy: 備份策略配置
            backup_type: 備份類型
        """
        if backup_type == "rolling":
            limit = strategy["rolling_limit"]
            if limit > 0:
                self._cleanup_rolling_backups(limit)
        elif backup_type == "daily":
            retain_days = strategy["daily_retain_days"]
            if retain_days > 0:
                self._cleanup_daily_backups(retain_days)

    def _cleanup_rolling_backups(self, limit: int):
        """清理流水備份，只保留最近 N 個

        Args:
            limit: 保留數量
        """
        backups = sorted(
            self.rolling_dir.glob("rolling_*.bak"),
            key=lambda p: p.stat().st_mtime
        )

        if len(backups) > limit:
            for backup in backups[:-limit]:
                size_mb = backup.stat().st_size / (1024 * 1024)
                backup.unlink()
                logger.info(f"已刪除舊的流水備份: {backup.name} ({size_mb:.2f} MB)")

    def _cleanup_daily_backups(self, retain_days: int):
        """清理每日備份，只保留 N 天內的

        Args:
            retain_days: 保留天數
        """
        cutoff_date = datetime.now() - timedelta(days=retain_days)

        for backup in self.daily_dir.glob("daily_*.bak"):
            # 從檔名提取日期
            date_str = backup.stem.replace("daily_", "")
            try:
                backup_date = datetime.strptime(date_str, "%Y%m%d")
                if backup_date < cutoff_date:
                    size_mb = backup.stat().st_size / (1024 * 1024)
                    backup.unlink()
                    logger.info(f"已刪除過期的每日備份: {backup.name} ({size_mb:.2f} MB)")
            except ValueError:
                logger.warning(f"無法解析備份日期，跳過: {backup.name}")

    def _get_last_daily_backup_date(self) -> Optional[datetime]:
        """取得最後一次每日備份的日期

        Returns:
            最後備份日期，如果沒有備份則返回 None
        """
        daily_backups = list(self.daily_dir.glob("daily_*.bak"))
        if not daily_backups:
            return None

        # 找到最新的備份
        latest = max(daily_backups, key=lambda p: p.stat().st_mtime)
        date_str = latest.stem.replace("daily_", "")

        try:
            return datetime.strptime(date_str, "%Y%m%d").date()
        except ValueError:
            logger.warning(f"無法解析備份日期: {latest.name}")
            return None

    def list_backups(self) -> Dict[str, List[Dict]]:
        """列出所有備份檔案

        Returns:
            包含 rolling 和 daily 兩個清單的字典
        """
        rolling_backups = [
            {
                "name": b.name,
                "size_mb": round(b.stat().st_size / (1024 * 1024), 2),
                "created": datetime.fromtimestamp(b.stat().st_ctime).isoformat(),
                "path": str(b)
            }
            for b in sorted(self.rolling_dir.glob("rolling_*.bak"), key=lambda p: p.stat().st_mtime, reverse=True)
        ]

        daily_backups = [
            {
                "name": b.name,
                "size_mb": round(b.stat().st_size / (1024 * 1024), 2),
                "created": datetime.fromtimestamp(b.stat().st_ctime).isoformat(),
                "path": str(b)
            }
            for b in sorted(self.daily_dir.glob("daily_*.bak"), key=lambda p: p.stat().st_mtime, reverse=True)
        ]

        return {
            "rolling": rolling_backups,
            "daily": daily_backups
        }

    def restore_backup(self, backup_name: str, db_connection) -> Tuple[bool, str]:
        """從備份還原資料庫

        Args:
            backup_name: 備份檔案名稱（如：rolling_20250429_143022.bak）
            db_connection: pyodbc 資料庫連線

        Returns:
            (是否成功, 訊息)
        """
        try:
            # 尋找備份檔案
            backup_file = None
            if backup_name.startswith("rolling_"):
                backup_file = self.rolling_dir / backup_name
            elif backup_name.startswith("daily_"):
                backup_file = self.daily_dir / backup_name
            else:
                return False, f"無效的備份檔案名稱: {backup_name}"

            if not backup_file.exists():
                return False, f"找不到備份檔案: {backup_name}"

            # 執行還原
            db_name = os.getenv("DB_NAME")

            logger.info(f"開始還原資料庫 {db_name} 從備份: {backup_name}")

            # SQL Server 還原步驟
            # 使用字串格式而非 Path.absolute()，因為 SQL Server 運行在 Windows 上
            restore_sql = f"""
                USE master;
                ALTER DATABASE [{db_name}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;

                RESTORE DATABASE [{db_name}]
                FROM DISK = '{str(backup_file)}'
                WITH REPLACE, STATS = 10;

                ALTER DATABASE [{db_name}] SET MULTI_USER;
            """

            # 設置 autocommit 以避免事務中執行還原的錯誤
            old_autocommit = db_connection.autocommit
            db_connection.autocommit = True

            cursor = db_connection.cursor()
            cursor.execute(restore_sql)

            # 等待還原完成
            while cursor.nextset():
                pass

            # 恢復原始 autocommit 設定
            db_connection.autocommit = old_autocommit

            msg = f"資料庫還原成功，已還原至備份: {backup_name}"
            logger.info(msg)

            return True, msg

        except Exception as e:
            error_msg = f"還原失敗: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return False, error_msg
