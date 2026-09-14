"""
数据库模型 - 使用 SQLAlchemy 统一 SQLite / MySQL
"""

from datetime import datetime
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class User(db.Model):
    """用户表"""
    __tablename__ = 'users'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    username = db.Column(db.String(100), unique=True, nullable=False, index=True)
    password = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default='user')  # admin / user
    real_name = db.Column(db.String(50), default='')
    phone = db.Column(db.String(20), default='')
    company = db.Column(db.String(200), default='')
    created_at = db.Column(db.String(30), default='')
    last_login = db.Column(db.String(30), default='')


class Customer(db.Model):
    """客户资料表（企业信息）"""
    __tablename__ = 'customers'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    company_name = db.Column(db.String(255), index=True, default='')
    legal_person = db.Column(db.String(100), default='')
    legal_phone = db.Column(db.Text, default='')
    registered_capital = db.Column(db.String(100), default='')
    established_date = db.Column(db.String(30), default='')
    business_status = db.Column(db.String(50), default='')
    industry = db.Column(db.String(200), default='')
    address = db.Column(db.String(500), default='')
    business_scope = db.Column(db.Text, default='')
    credit_code = db.Column(db.String(50), default='')
    province = db.Column(db.String(50), default='')
    city = db.Column(db.String(50), default='')
    district = db.Column(db.String(50), default='')
    insured_count = db.Column(db.String(50), default='')
    source = db.Column(db.String(100), default='')
    batch_name = db.Column(db.String(200), default='')
    status = db.Column(db.String(20), default='public')  # public / private
    owner_id = db.Column(db.Integer, default=0)
    claimed_at = db.Column(db.String(30), default='')
    created_at = db.Column(db.String(30), default='')
    remark = db.Column(db.String(500), default='')
    verify_result = db.Column(db.String(20), default='pass')  # pass / fail
    fail_reason = db.Column(db.String(200), default='')
    shareholder_info = db.Column(db.Text, default='')
    website = db.Column(db.String(500), default='')

    __table_args__ = (
        db.Index('idx_company_name', 'company_name'),
        db.Index('idx_verify_result', 'verify_result'),
        db.Index('idx_industry', 'industry'),
    )


class Batch(db.Model):
    """批次表"""
    __tablename__ = 'batches'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    name = db.Column(db.String(200), default='')
    source = db.Column(db.String(100), default='')
    total_companies = db.Column(db.Integer, default=0)
    total_phones = db.Column(db.Integer, default=0)
    verified = db.Column(db.Integer, default=0)
    no_exception = db.Column(db.Integer, default=0)
    status = db.Column(db.String(30), default='step1')  # step1_done / processing / completed / failed
    upload_file = db.Column(db.String(500), default='')
    map_file = db.Column(db.String(500), default='')
    step1_files = db.Column(db.Text, default='')  # JSON array
    step3_file = db.Column(db.String(500), default='')
    failed_file = db.Column(db.String(500), default='')
    created_by = db.Column(db.Integer, default=0)
    created_at = db.Column(db.String(30), default='')
    remark = db.Column(db.String(500), default='')
    shareholder_file = db.Column(db.String(500), default='')


class OperationLog(db.Model):
    """操作日志表"""
    __tablename__ = 'operation_logs'

    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    user_id = db.Column(db.Integer, default=0, index=True)
    action = db.Column(db.String(50), default='')
    target = db.Column(db.String(200), default='')
    detail = db.Column(db.Text, default='')
    created_at = db.Column(db.String(30), default='')


def init_db(app):
    """初始化数据库连接和表"""
    db.init_app(app)
    with app.app_context():
        db.create_all()


def get_raw_connection():
    """获取原始数据库连接（用于复杂查询和性能优化）"""
    return db.engine.raw_connection()
