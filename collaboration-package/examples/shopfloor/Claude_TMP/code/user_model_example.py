"""
=====================================================================
  user_model.py - 使用者資料模型範例
=====================================================================
更新時間：2025-11-14

一、用途說明
---------------------------------------------------------------------
這是 SQLAlchemy 使用者資料模型的範例檔案，展示如何定義資料表結構。

二、目標位置
---------------------------------------------------------------------
應加入到：app/models/user.py
位置：models 目錄下作為獨立模組

三、使用方式
---------------------------------------------------------------------
1. 將此檔案複製到 app/models/ 目錄
2. 根據你的資料庫調整欄位
3. 在 __init__.py 中 import 此模型
4. 執行 alembic 遷移建立資料表

四、需要檢查的點
---------------------------------------------------------------------
- [ ] 檢查資料庫連線字串是否正確
- [ ] 確認 SQLAlchemy 版本（建議 2.0+）
- [ ] 調整 table name 符合你的命名規範
- [ ] 檢查欄位長度是否符合需求

五、完整代碼
---------------------------------------------------------------------
"""

from datetime import datetime
from sqlalchemy import Column, Integer, String, Boolean, DateTime
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()


class User(Base):
    """使用者資料模型"""

    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, nullable=False, index=True)
    email = Column(String(100), unique=True, nullable=False, index=True)
    hashed_password = Column(String(255), nullable=False)
    is_active = Column(Boolean, default=True)
    is_admin = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def __repr__(self):
        return f"<User(id={self.id}, username='{self.username}', email='{self.email}')>"


"""
=====================================================================
使用範例：

from app.models.user import User
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

# 建立資料庫連線
engine = create_engine("postgresql://user:password@localhost/dbname")
SessionLocal = sessionmaker(bind=engine)

# 建立表格（首次使用）
Base.metadata.create_all(bind=engine)

# 新增使用者
session = SessionLocal()
new_user = User(
    username="john_doe",
    email="john@example.com",
    hashed_password="hashed_password_here"
)
session.add(new_user)
session.commit()
=====================================================================
"""
