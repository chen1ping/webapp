"""
配置模块 - 支持环境变量和 .env 文件配置
本地运行和云端部署统一读取配置
"""

import os
from dotenv import load_dotenv

# 加载 .env 文件（如果存在）
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if os.path.exists(_env_path):
    load_dotenv(_env_path)


def _get_bool(key, default=False):
    val = os.environ.get(key, '')
    return str(val).lower() in ('1', 'true', 'yes', 'on')


def _get_int(key, default=0):
    val = os.environ.get(key, '')
    try:
        return int(val)
    except (ValueError, TypeError):
        return default


# ========== 应用基础配置 ==========
APP_DIR = os.path.dirname(os.path.abspath(__file__))
SECRET_KEY = os.environ.get('SECRET_KEY', 'your-secret-key-change-in-production')
DEBUG = _get_bool('DEBUG', False)
PORT = _get_int('PORT', 443)

# ========== 数据库配置 ==========
# DB_TYPE: sqlite / mysql / postgres
DB_TYPE = os.environ.get('DB_TYPE', 'sqlite')

if DB_TYPE == 'postgres':
    # PostgreSQL 配置（Supabase / Render 等）
    # 支持直接传入完整 DATABASE_URL，也支持分开配置
    DATABASE_URL = os.environ.get('DATABASE_URL', '')
    if DATABASE_URL:
        SQLALCHEMY_DATABASE_URI = DATABASE_URL
    else:
        DB_HOST = os.environ.get('DB_HOST', 'localhost')
        DB_PORT = _get_int('DB_PORT', 5432)
        DB_USER = os.environ.get('DB_USER', 'postgres')
        DB_PASSWORD = os.environ.get('DB_PASSWORD', '')
        DB_NAME = os.environ.get('DB_NAME', 'webapp')
        SQLALCHEMY_DATABASE_URI = (
            f"postgresql://{DB_USER}:{DB_PASSWORD}"
            f"@{DB_HOST}:{DB_PORT}/{DB_NAME}"
        )
elif DB_TYPE == 'mysql':
    # MySQL 配置（腾讯云 CDB）
    DB_HOST = os.environ.get('DB_HOST', 'localhost')
    DB_PORT = _get_int('DB_PORT', 3306)
    DB_USER = os.environ.get('DB_USER', 'root')
    DB_PASSWORD = os.environ.get('DB_PASSWORD', '')
    DB_NAME = os.environ.get('DB_NAME', 'webapp')
    DB_CHARSET = os.environ.get('DB_CHARSET', 'utf8mb4')
    SQLALCHEMY_DATABASE_URI = (
        f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}"
        f"@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset={DB_CHARSET}"
    )
else:
    # SQLite（本地默认）
    DB_PATH = os.environ.get('DB_PATH', os.path.join(APP_DIR, 'data', 'system.db'))
    SQLALCHEMY_DATABASE_URI = f"sqlite:///{DB_PATH}"

SQLALCHEMY_TRACK_MODIFICATIONS = False
SQLALCHEMY_POOL_SIZE = _get_int('DB_POOL_SIZE', 10)
SQLALCHEMY_MAX_OVERFLOW = _get_int('DB_MAX_OVERFLOW', 20)
SQLALCHEMY_POOL_RECYCLE = _get_int('DB_POOL_RECYCLE', 3600)

# ========== 文件存储配置 ==========
# STORAGE_TYPE: local / cos (腾讯云COS) / supabase
STORAGE_TYPE = os.environ.get('STORAGE_TYPE', 'local')

if STORAGE_TYPE == 'cos':
    # 腾讯云 COS 配置
    COS_SECRET_ID = os.environ.get('COS_SECRET_ID', '')
    COS_SECRET_KEY = os.environ.get('COS_SECRET_KEY', '')
    COS_REGION = os.environ.get('COS_REGION', 'ap-guangzhou')
    COS_BUCKET = os.environ.get('COS_BUCKET', '')
    COS_DOMAIN = os.environ.get('COS_DOMAIN', '')  # 自定义域名（可选）
elif STORAGE_TYPE == 'supabase':
    # Supabase Storage 配置
    SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
    SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')
    SUPABASE_BUCKET = os.environ.get('SUPABASE_BUCKET', 'webapp-files')
else:
    # 本地存储
    UPLOAD_DIR = os.environ.get('UPLOAD_DIR', os.path.join(APP_DIR, 'data', 'uploads'))
    RESULT_DIR = os.environ.get('RESULT_DIR', os.path.join(APP_DIR, 'data', 'results'))
    EXPORT_DIR = os.environ.get('EXPORT_DIR', os.path.join(APP_DIR, 'data', 'exports'))
    BACKUP_DIR = os.environ.get('BACKUP_DIR', os.path.join(APP_DIR, 'data', 'backups'))

# 本地数据目录（缓存、临时文件等始终本地）
DATA_DIR = os.path.join(APP_DIR, 'data')

# ========== 其他配置 ==========
BATCH_SIZE = _get_int('BATCH_SIZE', 2000)
MAX_CONTENT_LENGTH = _get_int('MAX_CONTENT_LENGTH', 200) * 1024 * 1024  # 200MB
SESSION_COOKIE_SECURE = _get_bool('SESSION_COOKIE_SECURE', True)
