from sqlalchemy import create_engine, Column, Integer, String, Boolean, DateTime, Float, Text, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship, scoped_session
from datetime import datetime
import os

Base = declarative_base()

class User(Base):
    __tablename__ = 'users'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    username = Column(String(50), unique=True, nullable=False)
    email = Column(String(100), unique=True, nullable=False)
    password_hash = Column(String(255), nullable=False)
    full_name = Column(String(100))
    role = Column(String(20), default='user')
    is_active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    last_login = Column(DateTime)
    
    report_runs = relationship("ReportRun", back_populates="user")

class ReportRun(Base):
    __tablename__ = 'report_runs'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    user_id = Column(Integer, ForeignKey('users.id'), nullable=False)
    customer_name = Column(String(100))
    report_type = Column(String(20))
    template_name = Column(String(255))
    status = Column(String(20))
    created_at = Column(DateTime, default=datetime.utcnow)
    duration_seconds = Column(Float)
    log_data = Column(Text)
    
    user = relationship("User", back_populates="report_runs")
    report_files = relationship("ReportFile", back_populates="report_run", cascade="all, delete-orphan")

class ReportFile(Base):
    __tablename__ = 'report_files'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    run_id = Column(Integer, ForeignKey('report_runs.id'), nullable=False)
    filename = Column(String(255), nullable=False)
    file_path = Column(String(500), nullable=False)
    file_size = Column(Integer)
    step = Column(String(20))
    created_at = Column(DateTime, default=datetime.utcnow)
    
    report_run = relationship("ReportRun", back_populates="report_files")

def get_database_url():
    """Get database URL from environment variables"""
    return os.getenv('DATABASE_URL')

_engine = None
_SessionFactory = None

def get_engine():
    """Get or create the global database engine (singleton pattern)"""
    global _engine
    if _engine is None:
        database_url = get_database_url()
        if not database_url:
            raise ValueError("DATABASE_URL environment variable not set")
        _engine = create_engine(
            database_url,
            pool_size=10,
            max_overflow=20,
            pool_pre_ping=True,
            pool_recycle=3600,
            echo=False
        )
    return _engine

def get_session_factory():
    """Get or create the global session factory (singleton pattern)"""
    global _SessionFactory
    if _SessionFactory is None:
        engine = get_engine()
        _SessionFactory = scoped_session(sessionmaker(bind=engine))
    return _SessionFactory

def create_db_engine():
    """Legacy function for backward compatibility - returns the global engine"""
    return get_engine()

def create_tables():
    """Create all tables in the database"""
    engine = get_engine()
    Base.metadata.create_all(engine)
    print("Database tables created successfully!")

def get_session():
    """Get a new database session from the connection pool"""
    SessionFactory = get_session_factory()
    return SessionFactory()

if __name__ == "__main__":
    create_tables()
