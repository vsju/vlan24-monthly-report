from sqlalchemy import create_engine, Column, Integer, String, Boolean, DateTime, Float, Text, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
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

def create_db_engine():
    """Create database engine"""
    database_url = get_database_url()
    if not database_url:
        raise ValueError("DATABASE_URL environment variable not set")
    return create_engine(database_url)

def create_tables():
    """Create all tables in the database"""
    engine = create_db_engine()
    Base.metadata.create_all(engine)
    print("Database tables created successfully!")

def get_session():
    """Get a new database session"""
    engine = create_db_engine()
    Session = sessionmaker(bind=engine)
    return Session()

if __name__ == "__main__":
    create_tables()
