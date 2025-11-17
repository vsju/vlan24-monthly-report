from db_models import User, ReportRun, ReportFile, get_session
from datetime import datetime
import bcrypt
import os

def hash_password(password):
    """Hash a password using bcrypt"""
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def verify_password(password, hashed):
    """Verify a password against a hash"""
    return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))

def create_user(username, email, password, full_name, role='user'):
    """Create a new user"""
    session = get_session()
    try:
        existing_user = session.query(User).filter(
            (User.username == username) | (User.email == email)
        ).first()
        
        if existing_user:
            raise ValueError("Username or email already exists")
        
        user = User(
            username=username,
            email=email,
            password_hash=hash_password(password),
            full_name=full_name,
            role=role
        )
        session.add(user)
        session.commit()
        return user.id
    finally:
        session.close()

def get_user_by_username(username):
    """Get user by username"""
    session = get_session()
    try:
        return session.query(User).filter(User.username == username).first()
    finally:
        session.close()

def get_all_users():
    """Get all users"""
    session = get_session()
    try:
        return session.query(User).all()
    finally:
        session.close()

def update_user_active_status(username, is_active):
    """Update user active status"""
    session = get_session()
    try:
        user = session.query(User).filter(User.username == username).first()
        if user:
            user.is_active = is_active
            session.commit()
            return True
        return False
    finally:
        session.close()

def update_user_password(username, new_password):
    """Update user password"""
    session = get_session()
    try:
        user = session.query(User).filter(User.username == username).first()
        if user:
            user.password_hash = hash_password(new_password)
            session.commit()
            return True
        return False
    finally:
        session.close()

def update_last_login(username):
    """Update user's last login time"""
    session = get_session()
    try:
        user = session.query(User).filter(User.username == username).first()
        if user:
            user.last_login = datetime.utcnow()
            session.commit()
    finally:
        session.close()

def create_report_run(user_id, customer_name, report_type, template_name, status, duration_seconds=None, log_data=None):
    """Create a new report run record"""
    session = get_session()
    try:
        report_run = ReportRun(
            user_id=user_id,
            customer_name=customer_name,
            report_type=report_type,
            template_name=template_name,
            status=status,
            duration_seconds=duration_seconds,
            log_data=log_data
        )
        session.add(report_run)
        session.commit()
        return report_run.id
    finally:
        session.close()

def add_report_file(run_id, filename, file_path, file_size, step):
    """Add a report file record"""
    session = get_session()
    try:
        report_file = ReportFile(
            run_id=run_id,
            filename=filename,
            file_path=file_path,
            file_size=file_size,
            step=step
        )
        session.add(report_file)
        session.commit()
        return report_file.id
    finally:
        session.close()

def get_user_report_runs(user_id, limit=50):
    """Get report runs for a specific user"""
    session = get_session()
    try:
        return session.query(ReportRun).filter(
            ReportRun.user_id == user_id
        ).order_by(ReportRun.created_at.desc()).limit(limit).all()
    finally:
        session.close()

def get_all_report_runs(limit=100):
    """Get all report runs (for admin)"""
    session = get_session()
    try:
        return session.query(ReportRun).order_by(
            ReportRun.created_at.desc()
        ).limit(limit).all()
    finally:
        session.close()

def get_report_files_by_run_id(run_id):
    """Get all files for a specific report run"""
    session = get_session()
    try:
        return session.query(ReportFile).filter(
            ReportFile.run_id == run_id
        ).all()
    finally:
        session.close()
