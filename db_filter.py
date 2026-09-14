"""
数据库过滤工具 - 使用数据库查询替代内存集合，大幅降低内存占用

支持 SQLite / MySQL / PostgreSQL
"""

import sqlite3
import os
import re
import gc
import tempfile
import threading

_local = threading.local()


def _get_main_conn(db_path, db_type='sqlite'):
    """获取主库连接（线程内复用，db_filter专用）"""
    key = f'main_{db_path}'
    conn = getattr(_local, key, None)
    if conn is not None:
        try:
            if db_type == 'mysql':
                conn.ping(reconnect=True)
            elif db_type == 'postgres':
                conn.rollback()
                cur = conn.cursor()
                cur.execute("SELECT 1")
                cur.close()
            return conn
        except:
            pass

    if db_type == 'mysql':
        import pymysql
        from urllib.parse import urlparse
        # db_path 可能是 URL
        parsed = urlparse(db_path)
        conn = pymysql.connect(
            host=parsed.hostname or 'localhost',
            port=parsed.port or 3306,
            user=parsed.username or 'root',
            password=parsed.password or '',
            database=parsed.path.lstrip('/'),
            charset='utf8mb4',
            cursorclass=pymysql.cursors.DictCursor,
            autocommit=False,
        )
    elif db_type == 'postgres':
        import psycopg2
        import psycopg2.extras
        conn = psycopg2.connect(db_path)
        conn.cursor_factory = psycopg2.extras.RealDictCursor
        conn.autocommit = False
    else:
        conn = sqlite3.connect(db_path, timeout=30, isolation_level=None)
        conn.row_factory = sqlite3.Row

    setattr(_local, key, conn)
    return conn


def _create_temp_db():
    """创建临时数据库（SQLite 本地文件，不占内存）"""
    fd, path = tempfile.mkstemp(suffix='.db')
    os.close(fd)
    conn = sqlite3.connect(path)
    conn.execute('PRAGMA journal_mode=OFF')
    conn.execute('PRAGMA synchronous=OFF')
    conn.execute('PRAGMA temp_store=FILE')
    return conn, path


def _close_temp_db(conn, path):
    """关闭并删除临时数据库"""
    try:
        conn.close()
    except:
        pass
    try:
        if os.path.exists(path):
            os.remove(path)
    except:
        pass


def _detect_db_type(db_path):
    """根据 db_path 自动判断数据库类型"""
    if db_path.startswith('postgresql://') or db_path.startswith('postgres://'):
        return 'postgres'
    if db_path.startswith('mysql://') or db_path.startswith('mysql+pymysql://'):
        return 'mysql'
    # 本地文件路径 -> sqlite
    if os.path.exists(db_path) or db_path.endswith('.db') or db_path.endswith('.sqlite'):
        return 'sqlite'
    # 默认 sqlite
    return 'sqlite'


def filter_verified_companies(db_path, companies_list, db_type=None):
    """
    过滤已核验企业，返回已核验的企业名称集合
    
    参数:
        db_path: 主数据库路径或连接URL
        companies_list: 企业名称列表
        db_type: sqlite / mysql / postgres（不传则自动判断）
    
    返回:
        (已核验企业名称集合, 已核验数量)
    """
    if db_type is None:
        db_type = _detect_db_type(db_path)

    unique_names = list(set(n.strip() for n in companies_list if n and n.strip()))

    if db_type == 'sqlite':
        # SQLite 模式：用 ATTACH + JOIN
        temp_conn, temp_path = _create_temp_db()
        try:
            temp_conn.execute('CREATE TABLE temp_companies (name TEXT PRIMARY KEY)')
            temp_conn.executemany(
                'INSERT OR IGNORE INTO temp_companies (name) VALUES (?)',
                [(n,) for n in unique_names]
            )
            temp_conn.commit()

            temp_conn.execute(f"ATTACH DATABASE '{db_path}' AS main_db")
            cursor = temp_conn.execute(
                """
                SELECT t.name 
                FROM temp_companies t
                INNER JOIN main_db.customers c 
                    ON t.name = c.company_name
                WHERE c.verify_result = 'pass' AND c.company_name != ''
                """
            )
            verified_names = set(row[0] for row in cursor.fetchall())
            temp_conn.execute("DETACH DATABASE main_db")
            return verified_names, len(verified_names)
        finally:
            _close_temp_db(temp_conn, temp_path)
    else:
        # MySQL / PostgreSQL 模式：直接 IN 查询
        # 分批查询，避免 SQL 过长
        verified_names = set()
        batch_size = 500
        main_conn = _get_main_conn(db_path, db_type)

        for i in range(0, len(unique_names), batch_size):
            batch = unique_names[i:i+batch_size]
            placeholders = ', '.join(['%s'] * len(batch))
            cursor = main_conn.cursor()
            cursor.execute(
                f"""
                SELECT company_name FROM customers
                WHERE verify_result = 'pass'
                  AND company_name != ''
                  AND company_name IN ({placeholders})
                """,
                batch
            )
            for row in cursor.fetchall():
                name = row['company_name'] if hasattr(row, 'keys') else row[0]
                verified_names.add(name)
            cursor.close()

        return verified_names, len(verified_names)


def filter_verified_phones_fast(db_path, phones_series, db_type=None):
    """
    过滤已核验电话，返回已核验的电话号码集合
    
    参数:
        db_path: 主数据库路径或连接URL
        phones_series: 电话号码列表/Series
        db_type: sqlite / mysql / postgres（不传则自动判断）
    
    返回:
        (已核验电话号码集合, 已核验数量)
    """
    if db_type is None:
        db_type = _detect_db_type(db_path)

    unique_phones = list(set(p for p in phones_series if p and p.strip()))

    if db_type == 'sqlite':
        # SQLite 模式：临时表 + ATTACH + JOIN
        temp_conn, temp_path = _create_temp_db()
        try:
            temp_conn.execute('CREATE TABLE temp_phones (phone TEXT PRIMARY KEY)')
            temp_conn.executemany(
                'INSERT OR IGNORE INTO temp_phones (phone) VALUES (?)',
                [(p,) for p in unique_phones]
            )
            temp_conn.commit()

            temp_conn.execute(f"ATTACH DATABASE '{db_path}' AS main_db")
            temp_conn.execute('CREATE TABLE verified_extracted (phone TEXT PRIMARY KEY)')

            # 分批从主库提取并解析手机号
            main_cursor = temp_conn.execute(
                "SELECT legal_phone FROM main_db.customers WHERE verify_result='pass' AND legal_phone!=''"
            )

            batch = []
            batch_size = 1000
            for row in main_cursor:
                phones = re.findall(r'1\d{10}', row[0] or '')
                for p in phones:
                    batch.append((p,))
                    if len(batch) >= batch_size:
                        temp_conn.executemany(
                            'INSERT OR IGNORE INTO verified_extracted (phone) VALUES (?)',
                            batch
                        )
                        batch = []
            if batch:
                temp_conn.executemany(
                    'INSERT OR IGNORE INTO verified_extracted (phone) VALUES (?)',
                    batch
                )
            temp_conn.commit()
            del batch
            gc.collect()

            # JOIN 找交集
            cursor = temp_conn.execute(
                """
                SELECT t.phone FROM temp_phones t
                INNER JOIN verified_extracted v ON t.phone = v.phone
                """
            )
            verified_set = set(row[0] for row in cursor.fetchall())
            temp_conn.execute("DETACH DATABASE main_db")
            return verified_set, len(verified_set)
        finally:
            _close_temp_db(temp_conn, temp_path)
    else:
        # MySQL / PostgreSQL 模式：
        # 先从主库提取所有已核验手机号，再和本次数据比对
        # （已优化为只存手机号到临时表，不存完整记录）
        main_conn = _get_main_conn(db_path, db_type)
        cursor = main_conn.cursor()
        cursor.execute(
            "SELECT legal_phone FROM customers WHERE verify_result='pass' AND legal_phone!=''"
        )

        # 流式提取所有已核验手机号（只保留在本次数据范围内的）
        target_phones_set = set(unique_phones)
        verified_set = set()

        for row in cursor.fetchall():
            phone_text = row['legal_phone'] if hasattr(row, 'keys') else row[0]
            phones = re.findall(r'1\d{10}', phone_text or '')
            for p in phones:
                if p in target_phones_set:
                    verified_set.add(p)
                    # 全部找到了就提前退出
                    if len(verified_set) == len(target_phones_set):
                        cursor.close()
                        return verified_set, len(verified_set)

        cursor.close()
        return verified_set, len(verified_set)
