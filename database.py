"""
数据库连接包装器
统一 SQLite 和 MySQL 的连接接口，保持 raw SQL 风格
最小化对 app.py 的改动
"""

import sqlite3
import os
import threading

from config import DB_TYPE

# MySQL 支持（可选，未安装不报错）
try:
    import pymysql
    HAS_MYSQL = True
except ImportError:
    HAS_MYSQL = False

# PostgreSQL 支持（可选，未安装不报错）
try:
    import psycopg2
    import psycopg2.extras
    HAS_POSTGRES = True
except ImportError:
    HAS_POSTGRES = False

_local = threading.local()


def get_db():
    """获取数据库连接（线程内复用）"""
    conn = getattr(_local, 'conn', None)
    if conn is not None:
        try:
            # 测试连接是否有效
            if DB_TYPE == 'mysql':
                conn.ping(reconnect=True)
            elif DB_TYPE == 'postgres':
                conn.rollback()  # postgres 用 rollback 检测连接
                cur = conn.cursor()
                cur.execute("SELECT 1")
                cur.close()
            return conn
        except:
            pass

    if DB_TYPE == 'mysql' and HAS_MYSQL:
        from config import DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME, DB_CHARSET
        conn = pymysql.connect(
            host=DB_HOST,
            port=DB_PORT,
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME,
            charset=DB_CHARSET,
            cursorclass=pymysql.cursors.DictCursor,
            autocommit=False,
        )
    elif DB_TYPE == 'postgres' and HAS_POSTGRES:
        from config import DATABASE_URL, DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME
        if DATABASE_URL:
            conn = psycopg2.connect(DATABASE_URL)
        else:
            conn = psycopg2.connect(
                host=DB_HOST, port=DB_PORT, user=DB_USER,
                password=DB_PASSWORD, dbname=DB_NAME,
            )
        # 使用字典游标
        conn.cursor_factory = psycopg2.extras.RealDictCursor
        conn.autocommit = False
    else:
        from config import DB_PATH
        conn = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
        conn.row_factory = sqlite3.Row

    _local.conn = conn
    return conn


def close_db():
    """关闭当前线程的数据库连接"""
    conn = getattr(_local, 'conn', None)
    if conn is not None:
        try:
            conn.close()
        except:
            pass
        _local.conn = None


def init_db_tables():
    """初始化数据库表结构（兼容 SQLite / MySQL / PostgreSQL）"""
    conn = get_db()
    cursor = conn.cursor()

    if DB_TYPE == 'mysql':
        _init_mysql_tables(cursor)
    elif DB_TYPE == 'postgres':
        _init_postgres_tables(cursor)
    else:
        _init_sqlite_tables(cursor)

    conn.commit()


def _init_sqlite_tables(cursor):
    """SQLite 表结构"""
    cursor.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT DEFAULT 'user',
            real_name TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            company TEXT DEFAULT '',
            created_at TEXT DEFAULT '',
            last_login TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            company_name TEXT DEFAULT '',
            legal_person TEXT DEFAULT '',
            legal_phone TEXT DEFAULT '',
            registered_capital TEXT DEFAULT '',
            established_date TEXT DEFAULT '',
            business_status TEXT DEFAULT '',
            industry TEXT DEFAULT '',
            address TEXT DEFAULT '',
            business_scope TEXT DEFAULT '',
            credit_code TEXT DEFAULT '',
            province TEXT DEFAULT '',
            city TEXT DEFAULT '',
            district TEXT DEFAULT '',
            insured_count TEXT DEFAULT '',
            source TEXT DEFAULT '',
            batch_name TEXT DEFAULT '',
            status TEXT DEFAULT 'public',
            owner_id INTEGER DEFAULT 0,
            claimed_at TEXT DEFAULT '',
            created_at TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            verify_result TEXT DEFAULT 'pass',
            fail_reason TEXT DEFAULT '',
            shareholder_info TEXT DEFAULT '',
            website TEXT DEFAULT ''
        );

        CREATE INDEX IF NOT EXISTS idx_customers_company ON customers(company_name);
        CREATE INDEX IF NOT EXISTS idx_customers_verify ON customers(verify_result);
        CREATE INDEX IF NOT EXISTS idx_customers_industry ON customers(industry);

        CREATE TABLE IF NOT EXISTS batches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT DEFAULT '',
            source TEXT DEFAULT '',
            total_companies INTEGER DEFAULT 0,
            total_phones INTEGER DEFAULT 0,
            verified INTEGER DEFAULT 0,
            no_exception INTEGER DEFAULT 0,
            status TEXT DEFAULT 'step1',
            upload_file TEXT DEFAULT '',
            map_file TEXT DEFAULT '',
            step1_files TEXT DEFAULT '',
            step3_file TEXT DEFAULT '',
            failed_file TEXT DEFAULT '',
            created_by INTEGER DEFAULT 0,
            created_at TEXT DEFAULT '',
            remark TEXT DEFAULT '',
            shareholder_file TEXT DEFAULT ''
        );

        CREATE TABLE IF NOT EXISTS operation_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER DEFAULT 0,
            action TEXT DEFAULT '',
            target TEXT DEFAULT '',
            detail TEXT DEFAULT '',
            created_at TEXT DEFAULT ''
        );

        CREATE INDEX IF NOT EXISTS idx_logs_user ON operation_logs(user_id);
    """)


def _init_mysql_tables(cursor):
    """MySQL 表结构"""
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INT AUTO_INCREMENT PRIMARY KEY,
            username VARCHAR(100) NOT NULL UNIQUE,
            password VARCHAR(255) NOT NULL,
            role VARCHAR(20) DEFAULT 'user',
            real_name VARCHAR(50) DEFAULT '',
            phone VARCHAR(20) DEFAULT '',
            company VARCHAR(200) DEFAULT '',
            created_at VARCHAR(30) DEFAULT '',
            last_login VARCHAR(30) DEFAULT '',
            INDEX idx_username (username)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            id INT AUTO_INCREMENT PRIMARY KEY,
            company_name VARCHAR(255) DEFAULT '',
            legal_person VARCHAR(100) DEFAULT '',
            legal_phone TEXT,
            registered_capital VARCHAR(100) DEFAULT '',
            established_date VARCHAR(30) DEFAULT '',
            business_status VARCHAR(50) DEFAULT '',
            industry VARCHAR(200) DEFAULT '',
            address VARCHAR(500) DEFAULT '',
            business_scope TEXT,
            credit_code VARCHAR(50) DEFAULT '',
            province VARCHAR(50) DEFAULT '',
            city VARCHAR(50) DEFAULT '',
            district VARCHAR(50) DEFAULT '',
            insured_count VARCHAR(50) DEFAULT '',
            source VARCHAR(100) DEFAULT '',
            batch_name VARCHAR(200) DEFAULT '',
            status VARCHAR(20) DEFAULT 'public',
            owner_id INT DEFAULT 0,
            claimed_at VARCHAR(30) DEFAULT '',
            created_at VARCHAR(30) DEFAULT '',
            remark VARCHAR(500) DEFAULT '',
            verify_result VARCHAR(20) DEFAULT 'pass',
            fail_reason VARCHAR(200) DEFAULT '',
            shareholder_info TEXT,
            website VARCHAR(500) DEFAULT '',
            INDEX idx_company_name (company_name),
            INDEX idx_verify_result (verify_result),
            INDEX idx_industry (industry)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS batches (
            id INT AUTO_INCREMENT PRIMARY KEY,
            name VARCHAR(200) DEFAULT '',
            source VARCHAR(100) DEFAULT '',
            total_companies INT DEFAULT 0,
            total_phones INT DEFAULT 0,
            verified INT DEFAULT 0,
            no_exception INT DEFAULT 0,
            status VARCHAR(30) DEFAULT 'step1',
            upload_file VARCHAR(500) DEFAULT '',
            map_file VARCHAR(500) DEFAULT '',
            step1_files TEXT,
            step3_file VARCHAR(500) DEFAULT '',
            failed_file VARCHAR(500) DEFAULT '',
            created_by INT DEFAULT 0,
            created_at VARCHAR(30) DEFAULT '',
            remark VARCHAR(500) DEFAULT '',
            shareholder_file VARCHAR(500) DEFAULT ''
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS operation_logs (
            id INT AUTO_INCREMENT PRIMARY KEY,
            user_id INT DEFAULT 0,
            action VARCHAR(50) DEFAULT '',
            target VARCHAR(200) DEFAULT '',
            detail TEXT,
            created_at VARCHAR(30) DEFAULT '',
            INDEX idx_user_id (user_id)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """)


def _init_postgres_tables(cursor):
    """PostgreSQL 表结构"""
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username VARCHAR(100) NOT NULL UNIQUE,
            password VARCHAR(255) NOT NULL,
            role VARCHAR(20) DEFAULT 'user',
            real_name VARCHAR(50) DEFAULT '',
            phone VARCHAR(20) DEFAULT '',
            company VARCHAR(200) DEFAULT '',
            created_at VARCHAR(30) DEFAULT '',
            last_login VARCHAR(30) DEFAULT ''
        )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS customers (
            id SERIAL PRIMARY KEY,
            company_name VARCHAR(255) DEFAULT '',
            legal_person VARCHAR(100) DEFAULT '',
            legal_phone TEXT DEFAULT '',
            registered_capital VARCHAR(100) DEFAULT '',
            established_date VARCHAR(30) DEFAULT '',
            business_status VARCHAR(50) DEFAULT '',
            industry VARCHAR(200) DEFAULT '',
            address VARCHAR(500) DEFAULT '',
            business_scope TEXT DEFAULT '',
            credit_code VARCHAR(50) DEFAULT '',
            province VARCHAR(50) DEFAULT '',
            city VARCHAR(50) DEFAULT '',
            district VARCHAR(50) DEFAULT '',
            insured_count VARCHAR(50) DEFAULT '',
            source VARCHAR(100) DEFAULT '',
            batch_name VARCHAR(200) DEFAULT '',
            status VARCHAR(20) DEFAULT 'public',
            owner_id INTEGER DEFAULT 0,
            claimed_at VARCHAR(30) DEFAULT '',
            created_at VARCHAR(30) DEFAULT '',
            remark VARCHAR(500) DEFAULT '',
            verify_result VARCHAR(20) DEFAULT 'pass',
            fail_reason VARCHAR(200) DEFAULT '',
            shareholder_info TEXT DEFAULT '',
            website VARCHAR(500) DEFAULT ''
        )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_customers_company ON customers(company_name)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_customers_verify ON customers(verify_result)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_customers_industry ON customers(industry)")

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS batches (
            id SERIAL PRIMARY KEY,
            name VARCHAR(200) DEFAULT '',
            source VARCHAR(100) DEFAULT '',
            total_companies INTEGER DEFAULT 0,
            total_phones INTEGER DEFAULT 0,
            verified INTEGER DEFAULT 0,
            no_exception INTEGER DEFAULT 0,
            status VARCHAR(30) DEFAULT 'step1',
            upload_file VARCHAR(500) DEFAULT '',
            map_file VARCHAR(500) DEFAULT '',
            step1_files TEXT DEFAULT '',
            step3_file VARCHAR(500) DEFAULT '',
            failed_file VARCHAR(500) DEFAULT '',
            created_by INTEGER DEFAULT 0,
            created_at VARCHAR(30) DEFAULT '',
            remark VARCHAR(500) DEFAULT '',
            shareholder_file VARCHAR(500) DEFAULT ''
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS operation_logs (
            id SERIAL PRIMARY KEY,
            user_id INTEGER DEFAULT 0,
            action VARCHAR(50) DEFAULT '',
            target VARCHAR(200) DEFAULT '',
            detail TEXT DEFAULT '',
            created_at VARCHAR(30) DEFAULT ''
        )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_logs_user_id ON operation_logs(user_id)")


def exec_sql(sql, params=None):
    """执行 SQL 并返回 cursor（兼容三种数据库）"""
    conn = get_db()
    cursor = conn.cursor()
    if params:
        # MySQL 和 PostgreSQL 用 %s，SQLite 用 ?
        if DB_TYPE in ('mysql', 'postgres'):
            cursor.execute(sql, params)
        else:
            # 把 %s 替换成 ?
            sqlite_sql = sql.replace('%s', '?')
            cursor.execute(sqlite_sql, params)
    else:
        cursor.execute(sql)
    return cursor
