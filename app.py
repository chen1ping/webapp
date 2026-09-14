"""
陈平安的资料库 - 后端服务

功能:
- 用户注册/登录
- 客户公海（企业数据浏览、搜索、领取）
- 三步核验流程（上传数据 -> 生成核验文件 -> 处理结果）
- 管理员后台（用户管理、数据统计）
"""

import os
import sys
import re
import gc
import zlib
import zipfile
import sqlite3
import hashlib
import json
import uuid
import logging
import traceback
import time
import threading
from datetime import datetime
from functools import wraps


# ========== 应用级缓存 ==========
_cache = {}
CACHE_TTL = 300  # 5分钟缓存

def cache_get(key):
    entry = _cache.get(key)
    if entry and time.time() - entry[1] < CACHE_TTL:
        return entry[0]
    return None

def cache_set(key, value):
    _cache[key] = (value, time.time())

def cache_invalidate():
    _cache.clear()

from flask import (Flask, render_template, request, redirect, url_for,
                   session, jsonify, send_file, flash)

# 确保能找到处理模块
# frozen模式：__file__指向临时解压目录，数据文件需放在exe同级目录
if getattr(sys, 'frozen', False):
    # PyInstaller打包后：exe所在目录用于数据持久化
    APP_DIR = os.path.dirname(sys.executable)
    # 临时解压目录用于模板和静态文件
    BUNDLE_DIR = sys._MEIPASS if hasattr(sys, '_MEIPASS') else APP_DIR
    sys.path.insert(0, BUNDLE_DIR)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = APP_DIR
    sys.path.insert(0, BUNDLE_DIR)

SCRIPT_DIR = BUNDLE_DIR

from step1_generator import generate_alipay_files
import step23_processor
import importlib

# Flask模板和静态文件路径
if getattr(sys, 'frozen', False):
    app = Flask(__name__,
                template_folder=os.path.join(BUNDLE_DIR, 'templates'),
                static_folder=os.path.join(BUNDLE_DIR, 'static'))
else:
    app = Flask(__name__)
app.secret_key = 'customer_pool_system_2026'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200MB

# Session安全配置
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
    PERMANENT_SESSION_LIFETIME=86400 * 7,  # 7天
    TEMPLATES_AUTO_RELOAD=True,
    SEND_FILE_MAX_AGE_DEFAULT=0,
)

# ========== 日志配置 ==========
_data_dir = os.path.join(APP_DIR, 'data')
os.makedirs(_data_dir, exist_ok=True)

# 加载或生成安全密钥（持久化到data目录，重启后session不失效）
_secret_file = os.path.join(_data_dir, 'secret.key')
if os.path.exists(_secret_file):
    with open(_secret_file, 'r') as f:
        app.secret_key = f.read().strip()
else:
    import secrets
    app.secret_key = secrets.token_hex(32)
    with open(_secret_file, 'w') as f:
        f.write(app.secret_key)
from logging.handlers import RotatingFileHandler
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        RotatingFileHandler(
            os.path.join(_data_dir, 'app.log'),
            maxBytes=10 * 1024 * 1024,  # 10MB
            backupCount=5,               # 保留5个历史文件
            encoding='utf-8'
        ),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ========== 全局错误处理 ==========
@app.errorhandler(500)
def internal_error(error):
    logger.error(f"500错误: {str(error)}\n{traceback.format_exc()}")
    gc.collect()  # 出错后清理内存
    return render_template('error.html', message='服务器内部错误，请稍后重试或联系管理员'), 500

@app.errorhandler(404)
def not_found(error):
    return render_template('error.html', message='页面不存在'), 404

@app.errorhandler(413)
def too_large(error):
    return render_template('error.html', message='上传文件过大，请限制在200MB以内'), 413

@app.errorhandler(zlib.error)
def handle_zlib_error(error):
    """捕获zlib解压错误（损坏的xlsx文件），避免崩溃"""
    logger.warning(f"zlib解压错误（文件可能损坏）: {error}")
    gc.collect()
    from flask import request, redirect, url_for, flash
    try:
        flash('文件损坏或格式不正确，请检查后重新上传', 'error')
        return redirect(request.referrer or url_for('library'))
    except:
        return render_template('error.html', message='文件损坏或格式不正确'), 400

@app.errorhandler(zipfile.BadZipFile)
def handle_bad_zip(error):
    """捕获损坏的zip/xlsx文件错误"""
    logger.warning(f"BadZipFile（文件可能损坏）: {error}")
    gc.collect()
    from flask import request, redirect, url_for, flash
    try:
        flash('文件损坏或不是有效的Excel文件，请检查后重新上传', 'error')
        return redirect(request.referrer or url_for('library'))
    except:
        return render_template('error.html', message='文件损坏或格式不正确'), 400


# ========== 健康检查 ==========
@app.route('/health')
def health_check():
    """健康检查端点，返回服务状态和基本指标"""
    try:
        conn = get_db()
        db_ok = True
        total = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
        conn.close()
    except:
        db_ok = False
        total = 0
    
    try:
        import psutil
        mem = psutil.Process().memory_info().rss / 1024 / 1024  # MB
        mem_str = f"{mem:.1f}MB"
    except ImportError:
        mem_str = "N/A"
    except:
        mem_str = "N/A"
    
    return jsonify({
        'status': 'ok' if db_ok else 'degraded',
        'database': 'ok' if db_ok else 'error',
        'total_records': total,
        'memory': mem_str,
        'uptime': getattr(app, '_start_time', 'unknown'),
        'version': '2.0'
    }), 200 if db_ok else 503

# ========== 静态文件缓存加速 ==========
@app.after_request
def add_cache_headers(response):
    if request.path.startswith('/static/'):
        if '.css' in request.path or '.js' in request.path:
            response.cache_control.max_age = 300  # CSS/JS缓存5分钟，便于更新
        else:
            response.cache_control.max_age = 86400 * 7  # 图片等缓存7天
    return response

# 目录配置（数据文件使用APP_DIR确保持久化）
BASE_DIR = APP_DIR
DB_PATH = os.path.join(BASE_DIR, 'data', 'system.db')
UPLOAD_DIR = os.path.join(BASE_DIR, 'data', 'uploads')
RESULT_DIR = os.path.join(BASE_DIR, 'data', 'results')
EXPORT_DIR = os.path.join(BASE_DIR, 'data', 'exports')

# B2B平台列表（预定义，用于分类和筛选）
B2B_PLATFORMS = [
    '海智在线',
    'fb',
    '小蓝书',
    '爱采购',
    '阿里国际',
    '中国制造网',
    '马可波罗网',
    '爱企查',
]

for d in [UPLOAD_DIR, RESULT_DIR, EXPORT_DIR, os.path.join(BASE_DIR, 'data')]:
    os.makedirs(d, exist_ok=True)

# ========== 后台任务管理器 ==========
# 解决长时间导入导致HTTP阻塞、看门狗误杀的问题
_tasks = {}
_tasks_lock = threading.Lock()

def create_task(task_type, description=''):
    """创建一个后台任务，返回任务ID"""
    task_id = str(uuid.uuid4())[:8]
    with _tasks_lock:
        _tasks[task_id] = {
            'id': task_id,
            'type': task_type,
            'description': description,
            'status': 'pending',  # pending / running / done / failed
            'progress': 0,        # 0-100
            'message': '等待执行...',
            'result': None,
            'error': None,
            'created_at': time.time(),
            'started_at': None,
            'finished_at': None,
        }
    return task_id

def update_task_progress(task_id, progress=None, message=None, status=None, result=None, error=None):
    """更新任务进度"""
    with _tasks_lock:
        if task_id not in _tasks:
            return
        task = _tasks[task_id]
        if progress is not None:
            task['progress'] = min(100, max(0, progress))
        if message is not None:
            task['message'] = message
        if status is not None:
            task['status'] = status
            if status == 'running':
                task['started_at'] = time.time()
            elif status in ('done', 'failed'):
                task['finished_at'] = time.time()
        if result is not None:
            task['result'] = result
        if error is not None:
            task['error'] = error

def get_task_info(task_id):
    """获取任务信息"""
    with _tasks_lock:
        if task_id not in _tasks:
            return None
        return dict(_tasks[task_id])

def run_background_task(task_id, target_func, *args, **kwargs):
    """在后台线程中运行任务"""
    def worker():
        try:
            update_task_progress(task_id, status='running', message='正在执行...', progress=5)
            result = target_func(*args, **kwargs)
            update_task_progress(task_id, status='done', message='执行完成', progress=100, result=result)
        except Exception as e:
            logger.error(f"后台任务 {task_id} 失败: {e}\n{traceback.format_exc()}")
            update_task_progress(task_id, status='failed', message=f'执行失败: {str(e)}', error=str(e))
        finally:
            gc.collect()

    t = threading.Thread(target=worker, daemon=True)
    t.start()

def safe_read_excel(file_path, max_rows=500000):
    """安全读取Excel，多层异常保护，防止崩溃。自动检测表头行（兼容免责声明行）"""
    import pandas as pd
    df = None
    errors = []

    def _detect_and_fix_header(df_raw):
        """检测表头行是否正确，不正确则自动查找真正的表头"""
        known_header_keywords = ['公司名称', '企业名称', '名称', '法定代表人', '法人', '统一社会信用代码',
                                 '信用代码', '注册资本', '成立日期', '经营状态', '所属行业', '行业',
                                 '注册地址', '地址', '经营范围', '省份', '城市', '区县', '参保人数',
                                 '社保人数', '手机', '电话', '官网', '股东']
        cols = [str(c).strip() for c in df_raw.columns]
        score = sum(1 for kw in known_header_keywords if any(kw in c for c in cols))
        # 如果当前表头得分>=3，认为是正确的
        if score >= 3:
            return df_raw

        # 否则在前10行中找真正的表头
        # 注意：当前df是header=0读取的，所以df.iloc[N]对应文件第N+1行
        df_raw_reset = df_raw.reset_index(drop=True)
        for idx in range(min(10, len(df_raw_reset))):
            row_vals = [str(v).strip() for v in df_raw_reset.iloc[idx].tolist()]
            row_score = sum(1 for kw in known_header_keywords if any(kw in str(v) for v in row_vals))
            if row_score >= 3:
                # 重新读取，以这一行为表头（文件行号 = idx + 1）
                # 根据文件类型选择引擎，xls用xlrd，xlsx用openpyxl
                is_xls = file_path.lower().endswith('.xls')
                re_read_engine = 'xlrd' if is_xls else 'openpyxl'
                try:
                    new_df = pd.read_excel(file_path, dtype=str, engine=re_read_engine, header=idx + 1).fillna('')
                    if len(new_df) > max_rows:
                        new_df = new_df.iloc[:max_rows]
                    return new_df
                except:
                    # 备用引擎
                    try:
                        backup_engine = 'openpyxl' if is_xls else 'xlrd'
                        new_df = pd.read_excel(file_path, dtype=str, engine=backup_engine, header=idx + 1).fillna('')
                        if len(new_df) > max_rows:
                            new_df = new_df.iloc[:max_rows]
                        return new_df
                    except:
                        pass
                break

        # 找不到就返回原始的
        return df_raw

    # 策略1: openpyxl
    try:
        df = pd.read_excel(file_path, dtype=str, engine='openpyxl').fillna('')
        if len(df) > max_rows:
            df = df.iloc[:max_rows]
        df = _detect_and_fix_header(df)
        return df, None
    except Exception as e:
        errors.append(f'openpyxl: {e}')

    # 策略2: xlrd
    try:
        df = pd.read_excel(file_path, dtype=str, engine='xlrd').fillna('')
        if len(df) > max_rows:
            df = df.iloc[:max_rows]
        df = _detect_and_fix_header(df)
        return df, None
    except Exception as e:
        errors.append(f'xlrd: {e}')

    # 策略3: xlrd 直接读取（忽略损坏）
    try:
        import xlrd
        book = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
        sh = book.sheet_by_index(0)
        # 在前10行找表头
        header_row = 0
        known_header_keywords = ['公司名称', '企业名称', '名称', '法定代表人', '法人', '统一社会信用代码',
                                 '注册资本', '成立日期', '经营状态', '行业', '地址', '电话']
        best_score = 0
        for r in range(min(10, sh.nrows)):
            row_vals = [str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)]
            score = sum(1 for kw in known_header_keywords if any(kw in str(v) for v in row_vals))
            if score > best_score:
                best_score = score
                header_row = r

        headers = [str(sh.cell_value(header_row, c)).strip() for c in range(sh.ncols)]
        rows = []
        max_r = min(sh.nrows - 1, max_rows + header_row)
        for r in range(header_row + 1, max_r + 1):
            try:
                rows.append({headers[c]: str(sh.cell_value(r, c)).strip() for c in range(sh.ncols)})
            except:
                continue
        df = pd.DataFrame(rows)
        return df, None
    except Exception as e:
        errors.append(f'xlrd_raw: {e}')

    # 策略4: 尝试用zipfile检查并修复zip结构
    try:
        import zipfile
        with zipfile.ZipFile(file_path, 'r') as _zf:
            pass
        try:
            df = pd.read_excel(file_path, dtype=str, engine='openpyxl', read_only=True).fillna('')
            if len(df) > max_rows:
                df = df.iloc[:max_rows]
            df = _detect_and_fix_header(df)
            return df, None
        except Exception as e:
            errors.append(f'openpyxl_ro: {e}')
    except zipfile.BadZipFile as e:
        errors.append(f'BadZipFile: {e}')
    except Exception as e:
        errors.append(f'zip_check: {e}')

    return None, '; '.join(errors)

# ========== 数据库自动备份 ==========
BACKUP_DIR = os.path.join(BASE_DIR, 'data', 'backups')
os.makedirs(BACKUP_DIR, exist_ok=True)
MAX_BACKUPS = 20  # 最多保留20个备份

def backup_database(reason='manual'):
    """备份数据库，返回备份文件路径。失败返回None。"""
    try:
        import shutil
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = os.path.join(BACKUP_DIR, f'system_{timestamp}_{reason}.db')
        # 使用SQLite在线备份方式更安全
        src_conn = sqlite3.connect(DB_PATH)
        dst_conn = sqlite3.connect(backup_file)
        src_conn.backup(dst_conn)
        dst_conn.close()
        src_conn.close()
        
        # 清理旧备份，只保留最新的MAX_BACKUPS个
        backups = sorted([f for f in os.listdir(BACKUP_DIR) if f.endswith('.db')])
        if len(backups) > MAX_BACKUPS:
            for old in backups[:len(backups) - MAX_BACKUPS]:
                try:
                    os.remove(os.path.join(BACKUP_DIR, old))
                except:
                    pass
        
        logger.info(f"数据库备份成功: {backup_file} ({reason})")
        return backup_file
    except Exception as e:
        logger.error(f"数据库备份失败: {e}")
        return None


# ========== 数据库工具 ==========

def _parse_insured_number(val):
    """解析参保人数，支持'10人'、'10-20人'、'50人以上'等格式"""
    if val is None:
        return 0
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'none', 'null', ''):
        return 0
    s = s.replace('人', '').replace('以上', '').replace('以下', '').strip()
    if '-' in s:
        parts = s.split('-')
        try:
            return int(float(parts[0].strip()))
        except:
            pass
    try:
        return int(float(s))
    except:
        return 0

def get_db():
    """获取数据库连接，启用WAL模式和超时机制防止锁死"""
    try:
        conn = sqlite3.connect(DB_PATH, timeout=30, isolation_level=None)
        conn.row_factory = sqlite3.Row
        conn.create_function('parse_insured', 1, _parse_insured_number)
        conn.execute('PRAGMA journal_mode=WAL')
        conn.execute('PRAGMA busy_timeout=30000')
        conn.execute('PRAGMA wal_autocheckpoint=1000')
        conn.execute('PRAGMA synchronous=NORMAL')
        conn.execute('PRAGMA cache_size=-64000')  # 64MB内存缓存
        conn.execute('PRAGMA temp_store=MEMORY')
        conn.execute('PRAGMA mmap_size=268435456')  # 256MB mmap
        return conn
    except sqlite3.Error as e:
        logger.error(f"数据库连接失败: {e}")
        raise


def init_db():
    """初始化数据库"""
    conn = get_db()
    c = conn.cursor()

    # 用户表
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT DEFAULT 'user',
        real_name TEXT,
        phone TEXT,
        company TEXT,
        created_at TEXT,
        last_login TEXT
    )''')

    # 客户公海表
    c.execute('''CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        company_name TEXT,
        legal_person TEXT,
        legal_phone TEXT,
        registered_capital TEXT,
        established_date TEXT,
        business_status TEXT,
        industry TEXT,
        address TEXT,
        business_scope TEXT,
        credit_code TEXT,
        province TEXT,
        city TEXT,
        district TEXT,
        insured_count TEXT,
        website TEXT,
        source TEXT,
        batch_name TEXT,
        status TEXT DEFAULT 'public',
        owner_id INTEGER,
        claimed_at TEXT,
        created_at TEXT,
        remark TEXT,
        verify_result TEXT DEFAULT 'pass',
        fail_reason TEXT
    )''')

    # 兼容旧数据库，添加新字段
    try:
        c.execute("ALTER TABLE customers ADD COLUMN verify_result TEXT DEFAULT 'pass'")
    except:
        pass
    try:
        c.execute("ALTER TABLE customers ADD COLUMN fail_reason TEXT")
    except:
        pass
    try:
        c.execute("ALTER TABLE customers ADD COLUMN website TEXT")
    except:
        pass
    try:
        c.execute("ALTER TABLE customers ADD COLUMN shareholder_info TEXT")
    except:
        pass

    # 批次表
    c.execute('''CREATE TABLE IF NOT EXISTS batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        source TEXT,
        total_companies INTEGER DEFAULT 0,
        total_phones INTEGER DEFAULT 0,
        verified INTEGER DEFAULT 0,
        no_exception INTEGER DEFAULT 0,
        status TEXT DEFAULT 'step1',
        upload_file TEXT,
        map_file TEXT,
        step1_files TEXT,
        step3_file TEXT,
        failed_file TEXT,
        created_by INTEGER,
        created_at TEXT,
        remark TEXT,
        shareholder_file TEXT
    )''')

    # 操作日志
    c.execute('''CREATE TABLE IF NOT EXISTS operation_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        action TEXT,
        target TEXT,
        detail TEXT,
        created_at TEXT
    )''')

    # 兼容旧数据库：batches表添加shareholder_file字段
    try:
        c.execute("ALTER TABLE batches ADD COLUMN shareholder_file TEXT")
    except:
        pass

    # 创建索引（加速50万+数据查询）
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_company ON customers(company_name)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_phone ON customers(legal_phone)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_credit ON customers(credit_code)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_batch ON customers(batch_name)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_verify ON customers(verify_result)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_province ON customers(province)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_source ON customers(source)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_customers_company_phone ON customers(company_name, legal_phone)')

    # 创建默认管理员
    admin_pwd = hashlib.md5('chenpingan123'.encode()).hexdigest()
    c.execute("SELECT id FROM users WHERE username='chenpingan'")
    if not c.fetchone():
        c.execute('''INSERT INTO users (username, password, role, real_name, created_at)
                     VALUES (?, ?, 'admin', '陈平安', ?)''',
                  ('chenpingan', admin_pwd, datetime.now().strftime('%Y-%m-%d %H:%M:%S')))

    conn.commit()
    conn.close()
    
    # ===== 创建索引（性能优化）=====
    conn = get_db()
    c = conn.cursor()
    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_customers_status ON customers(status)",
        "CREATE INDEX IF NOT EXISTS idx_customers_owner_id ON customers(owner_id)",
        "CREATE INDEX IF NOT EXISTS idx_customers_verify_result ON customers(verify_result)",
        "CREATE INDEX IF NOT EXISTS idx_customers_province ON customers(province)",
        "CREATE INDEX IF NOT EXISTS idx_customers_city ON customers(city)",
        "CREATE INDEX IF NOT EXISTS idx_customers_district ON customers(district)",
        "CREATE INDEX IF NOT EXISTS idx_customers_batch_name ON customers(batch_name)",
        "CREATE INDEX IF NOT EXISTS idx_customers_source ON customers(source)",
        "CREATE INDEX IF NOT EXISTS idx_customers_credit_code ON customers(credit_code)",
        "CREATE INDEX IF NOT EXISTS idx_customers_company_name ON customers(company_name)",
        "CREATE INDEX IF NOT EXISTS idx_customers_legal_phone ON customers(legal_phone)",
        "CREATE INDEX IF NOT EXISTS idx_customers_cname_phone_batch ON customers(company_name, legal_phone, batch_name)",
        "CREATE INDEX IF NOT EXISTS idx_customers_status_verify ON customers(status, verify_result)",
        "CREATE INDEX IF NOT EXISTS idx_customers_id_desc ON customers(id DESC)",
        "CREATE INDEX IF NOT EXISTS idx_customers_owner_status ON customers(owner_id, status)",
        "CREATE INDEX IF NOT EXISTS idx_customers_pass_id ON customers(verify_result, id)",
        "CREATE INDEX IF NOT EXISTS idx_batches_created_by ON batches(created_by)",
        "CREATE INDEX IF NOT EXISTS idx_batches_status ON batches(status)",
        "CREATE INDEX IF NOT EXISTS idx_logs_user_id ON operation_logs(user_id)",
        "CREATE INDEX IF NOT EXISTS idx_logs_created_at ON operation_logs(created_at)",
    ]
    for idx_sql in indexes:
        try:
            c.execute(idx_sql)
        except Exception:
            pass
    conn.commit()
    conn.close()
    
    print("数据库初始化完成")


_pwd_salt_key = 'cpa_salt_2026!'

def hash_password(pwd):
    """SHA256 + 盐值 哈希（兼容旧MD5，登录时自动升级）"""
    return hashlib.sha256((pwd + _pwd_salt_key).encode()).hexdigest()

def verify_password(pwd, stored_hash):
    """验证密码，支持SHA256和旧MD5两种格式"""
    if len(stored_hash) == 32:  # 旧MD5
        return hashlib.md5(pwd.encode()).hexdigest() == stored_hash
    return hash_password(pwd) == stored_hash

def needs_hash_upgrade(stored_hash):
    """判断是否需要升级到SHA256"""
    return len(stored_hash) == 32


# ========== 登录防爆破 ==========
_login_attempts = {}  # {ip: [timestamps...]}
_MAX_FAILS = 10       # 最大失败次数
_LOCK_WINDOW = 300    # 锁定时间窗口（秒）

def _get_client_ip():
    """获取客户端真实IP"""
    if request.headers.get('X-Forwarded-For'):
        return request.headers.get('X-Forwarded-For').split(',')[0].strip()
    return request.remote_addr or 'unknown'

def _is_login_locked(ip):
    """检查IP是否被锁定"""
    now = time.time()
    if ip in _login_attempts:
        # 清理过期记录
        _login_attempts[ip] = [t for t in _login_attempts[ip] if now - t < _LOCK_WINDOW]
        if len(_login_attempts[ip]) >= _MAX_FAILS:
            return True
    return False

def _record_login_fail(ip):
    """记录登录失败"""
    now = time.time()
    if ip not in _login_attempts:
        _login_attempts[ip] = []
    _login_attempts[ip].append(now)
    # 清理过期
    _login_attempts[ip] = [t for t in _login_attempts[ip] if now - t < _LOCK_WINDOW]

def _clear_login_attempts(ip):
    """登录成功后清除记录"""
    if ip in _login_attempts:
        del _login_attempts[ip]


def log_action(user_id, action, target='', detail=''):
    conn = get_db()
    conn.execute('''INSERT INTO operation_logs (user_id, action, target, detail, created_at)
                    VALUES (?, ?, ?, ?, ?)''',
                 (user_id, action, target, detail,
                  datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    conn.commit()
    conn.close()


# ========== 登录验证装饰器 ==========

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('需要管理员权限', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated


# ========== 路由 ==========

@app.route('/')
def index():
    if 'user_id' not in session:
        return redirect(url_for('login'))
    return redirect(url_for('customer_pool'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        ip = _get_client_ip()

        # 防爆破检查
        if _is_login_locked(ip):
            flash('登录失败次数过多，请5分钟后再试', 'error')
            return render_template('login.html')

        conn = get_db()
        user = conn.execute(
            "SELECT * FROM users WHERE username=?", (username,)
        ).fetchone()
        conn.close()

        if user and verify_password(password, user['password']):
            _clear_login_attempts(ip)
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['role'] = user['role']
            session['real_name'] = user['real_name'] or user['username']
            session.permanent = True

            # 更新最后登录时间 + 自动升级旧MD5密码到SHA256
            conn = get_db()
            if needs_hash_upgrade(user['password']):
                conn.execute("UPDATE users SET password=?, last_login=? WHERE id=?",
                             (hash_password(password),
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S'), user['id']))
            else:
                conn.execute("UPDATE users SET last_login=? WHERE id=?",
                             (datetime.now().strftime('%Y-%m-%d %H:%M:%S'), user['id']))
            conn.commit()
            conn.close()

            log_action(user['id'], 'login', '', '登录成功')
            return redirect(url_for('customer_pool'))
        else:
            _record_login_fail(ip)
            flash('用户名或密码错误', 'error')

    return render_template('login.html')


@app.route('/logout')
def logout():
    if 'user_id' in session:
        log_action(session['user_id'], 'logout', '', '退出登录')
    session.clear()
    return redirect(url_for('login'))


@app.route('/register', methods=['GET', 'POST'])
@admin_required
def register():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        real_name = request.form.get('real_name', '').strip()
        phone = request.form.get('phone', '').strip()

        if len(username) < 3 or len(password) < 4:
            flash('用户名至少3位，密码至少4位', 'error')
        else:
            conn = get_db()
            try:
                conn.execute('''INSERT INTO users (username, password, real_name, phone, role, created_at)
                                VALUES (?, ?, ?, ?, 'user', ?)''',
                             (username, hash_password(password), real_name, phone,
                              datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                conn.commit()
                flash('注册成功，请登录', 'success')
                return redirect(url_for('login'))
            except sqlite3.IntegrityError:
                flash('用户名已存在', 'error')
            finally:
                conn.close()

    return render_template('register.html')


# ---------- 客户公海 ----------

@app.route('/pool')
@login_required
def customer_pool():
    page = int(request.args.get('page', 1))
    per_page = 20
    keyword = request.args.get('keyword', '').strip()
    industry = request.args.get('industry', '').strip()
    province = request.args.get('province', '').strip()
    batch = request.args.get('batch', '').strip()
    verify = request.args.get('verify', 'all')  # all/pass/failed/unverified
    status = request.args.get('status', 'public')
    has_shareholder = request.args.get('has_shareholder', '')  # ''/'yes'/'no'

    conn = get_db()

    # 构建查询
    conditions = []
    params = []

    if status == 'public':
        conditions.append("status='public'")
    elif status == 'my':
        conditions.append(f"owner_id={session['user_id']}")
    elif status == 'all':
        pass

    if keyword:
        conditions.append("(company_name LIKE ? OR legal_person LIKE ? OR legal_phone LIKE ?)")
        params.extend([f'%{keyword}%'] * 3)

    if industry:
        conditions.append("industry LIKE ?")
        params.append(f'%{industry}%')

    if province:
        conditions.append("province=?")
        params.append(province)

    if batch:
        conditions.append("batch_name=?")
        params.append(batch)

    if verify == 'pass':
        conditions.append("verify_result='pass'")
    elif verify == 'failed':
        conditions.append("verify_result='failed'")
    elif verify == 'unverified':
        conditions.append("(verify_result IS NULL OR verify_result='')")

    if has_shareholder == 'yes':
        conditions.append("(shareholder_info IS NOT NULL AND shareholder_info!='')")
    elif has_shareholder == 'no':
        conditions.append("(shareholder_info IS NULL OR shareholder_info='')")

    where_sql = ' AND '.join(conditions) if conditions else '1=1'

    # 总数
    total = conn.execute(
        f"SELECT COUNT(*) FROM customers WHERE {where_sql}", params
    ).fetchone()[0]

    # 分页数据
    offset = (page - 1) * per_page
    customers = conn.execute(
        f'''SELECT c.*, u.real_name as owner_name 
            FROM customers c 
            LEFT JOIN users u ON c.owner_id = u.id
            WHERE {where_sql}
            ORDER BY c.id DESC
            LIMIT ? OFFSET ?''',
        params + [per_page, offset]
    ).fetchall()

    # 省份列表（缓存）
    provinces = cache_get('library_provinces')
    if provinces is None:
        provinces = conn.execute(
            "SELECT DISTINCT province FROM customers WHERE province!='' ORDER BY province"
        ).fetchall()
        cache_set('library_provinces', provinces)

    # 批次列表（缓存）
    batches = cache_get('library_batches')
    if batches is None:
        batches = conn.execute(
            "SELECT DISTINCT batch_name FROM customers WHERE batch_name!='' ORDER BY batch_name"
        ).fetchall()
        cache_set('library_batches', batches)

    # 按批次统计（缓存）
    batch_stats = cache_get('pool_batch_stats')
    if batch_stats is None:
        batch_stats = conn.execute(
            '''SELECT batch_name,
                      SUM(CASE WHEN status='public' THEN 1 ELSE 0 END) as public_count,
                      COUNT(*) as total
               FROM customers
               WHERE batch_name!=''
               GROUP BY batch_name
               ORDER BY total DESC'''
        ).fetchall()
        cache_set('pool_batch_stats', batch_stats)

    # 核验统计+总统计（单次聚合查询）
    counts = conn.execute('''SELECT
        COUNT(*) as total_all,
        SUM(CASE WHEN verify_result='pass' THEN 1 ELSE 0 END) as total_pass,
        SUM(CASE WHEN verify_result='failed' THEN 1 ELSE 0 END) as total_failed,
        SUM(CASE WHEN verify_result IS NULL OR verify_result='' THEN 1 ELSE 0 END) as total_unverified,
        SUM(CASE WHEN status='public' THEN 1 ELSE 0 END) as total_public,
        SUM(CASE WHEN owner_id=? THEN 1 ELSE 0 END) as total_my
        FROM customers''', (session['user_id'],)).fetchone()
    verify_stats = {
        'total_pass': counts['total_pass'] or 0,
        'total_failed': counts['total_failed'] or 0,
        'total_unverified': counts['total_unverified'] or 0,
    }
    stats = {
        'total_public': counts['total_public'] or 0,
        'total_my': counts['total_my'] or 0,
        'total_all': counts['total_all'] or 0,
    }

    conn.close()

    total_pages = (total + per_page - 1) // per_page

    return render_template('pool.html',
                           customers=customers,
                           page=page, total_pages=total_pages, total=total,
                           keyword=keyword, industry=industry, province=province,
                           batch=batch, verify=verify, status=status,
                           has_shareholder=has_shareholder,
                           provinces=provinces, batches=batches,
                           batch_stats=batch_stats, verify_stats=verify_stats,
                           stats=stats)


@app.route('/customer/<int:cid>')
@login_required
def customer_detail(cid):
    conn = get_db()
    customer = conn.execute(
        '''SELECT c.*, u.real_name as owner_name 
           FROM customers c 
           LEFT JOIN users u ON c.owner_id = u.id
           WHERE c.id=?''', (cid,)
    ).fetchone()
    conn.close()

    if not customer:
        flash('客户不存在', 'error')
        return redirect(url_for('customer_pool'))

    return render_template('customer_detail.html', customer=customer)


@app.route('/customer/<int:cid>/claim', methods=['POST'])
@login_required
def claim_customer(cid):
    conn = get_db()
    customer = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()

    if not customer:
        conn.close()
        return jsonify({'success': False, 'msg': '客户不存在'})

    if customer['status'] != 'public':
        conn.close()
        return jsonify({'success': False, 'msg': '该客户已被领取'})

    conn.execute('''UPDATE customers SET status='claimed', owner_id=?, claimed_at=?
                    WHERE id=? AND status='public' ''',
                 (session['user_id'],
                  datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                  cid))
    conn.commit()
    conn.close()

    log_action(session['user_id'], 'claim_customer', str(cid), customer['company_name'])
    return jsonify({'success': True, 'msg': '领取成功'})


@app.route('/customer/<int:cid>/release', methods=['POST'])
@login_required
def release_customer(cid):
    conn = get_db()
    customer = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()

    if not customer:
        conn.close()
        return jsonify({'success': False, 'msg': '客户不存在'})

    if customer['owner_id'] != session['user_id'] and session.get('role') != 'admin':
        conn.close()
        return jsonify({'success': False, 'msg': '无权操作'})

    conn.execute('''UPDATE customers SET status='public', owner_id=NULL, claimed_at=NULL
                    WHERE id=?''', (cid,))
    conn.commit()
    conn.close()

    log_action(session['user_id'], 'release_customer', str(cid), customer['company_name'])
    return jsonify({'success': True, 'msg': '已放回公海'})


# ---------- 数据处理（三步流程） ----------

@app.route('/upload')
@login_required
def upload_page():
    conn = get_db()
    batches = conn.execute(
        '''SELECT b.*, u.real_name as creator_name 
           FROM batches b LEFT JOIN users u ON b.created_by=u.id
           ORDER BY b.id DESC LIMIT 20'''
    ).fetchall()
    conn.close()

    return render_template('upload.html', batches=batches)


@app.route('/upload/step1', methods=['POST'])
@login_required
def upload_step1():
    if 'file' not in request.files:
        flash('请选择文件', 'error')
        return redirect(url_for('upload_page'))

    file = request.files['file']
    sh_file = request.files.get('shareholder_file')
    batch_name = request.form.get('batch_name', '').strip()
    source = request.form.get('source', '爱企查')
    max_rows = int(request.form.get('max_rows', 3000))
    header_row = int(request.form.get('header_row', 3))
    filter_industrial = 'filter_industrial' in request.form

    if not file.filename:
        flash('请选择文件', 'error')
        return redirect(url_for('upload_page'))

    if not batch_name:
        batch_name = f"批次{datetime.now().strftime('%Y%m%d%H%M%S')}"

    # 保存上传文件
    ext = os.path.splitext(file.filename)[1]
    batch_id = str(uuid.uuid4())[:8]
    upload_path = os.path.join(UPLOAD_DIR, f'{batch_id}_{file.filename}')
    file.save(upload_path)

    # 保存股东信息文件（可选）
    shareholder_path = None
    if sh_file and sh_file.filename:
        sh_ext = os.path.splitext(sh_file.filename)[1]
        shareholder_path = os.path.join(UPLOAD_DIR, f'{batch_id}_股东_{sh_file.filename}')
        sh_file.save(shareholder_path)
        logger.info(f"股东信息文件已保存: {shareholder_path}")

    try:
        # 使用数据库模式（不占内存），不再加载全部已核验数据到内存集合
        # verified_phones 和 verified_companies 传入 None，
        # generate_alipay_files 内部通过 db_path 按需查询

        # 执行Step 1（数据库模式，低内存占用）
        result = generate_alipay_files(
            data_file=upload_path,
            output_prefix=batch_name,
            header_row=header_row,
            max_rows_per_file=max_rows,
            random_seed=100,
            output_dir=RESULT_DIR,
            shareholder_file=shareholder_path,
            verified_phones=None,
            filter_industrial=filter_industrial,
            verified_companies=None,
            db_path=DB_PATH,
        )

        # 保存批次记录
        conn = get_db()
        cursor = conn.execute(
            '''INSERT INTO batches (name, source, total_companies, total_phones,
               status, upload_file, map_file, step1_files, created_by, created_at, shareholder_file)
               VALUES (?, ?, ?, ?, 'step1_done', ?, ?, ?, ?, ?, ?)''',
            (batch_name, source, result['total_companies'], result['total_phones'],
             upload_path, result['map_file'],
             json.dumps(result['output_files'], ensure_ascii=False),
             session['user_id'],
             datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
             shareholder_path)
        )
        conn.commit()
        batch_db_id = cursor.lastrowid
        conn.close()

        log_action(session['user_id'], 'step1_create', str(batch_db_id),
                   f'{batch_name}: {result["total_phones"]}条')

        industrial_msg = f'，工业品筛选过滤 {result.get("industrial_filtered", 0)} 家非工业类企业' if filter_industrial and result.get('industrial_filtered', 0) > 0 else ''
        company_msg = f'，资料库匹配过滤 {result.get("company_filtered", 0)} 家已核验企业' if result.get('company_filtered', 0) > 0 else ''
        flash(f'Step 1 完成！共 {result["total_phones"]} 条手机号，拆分为 {result["num_files"]} 份{industrial_msg}{company_msg}', 'success')

    except Exception as e:
        import traceback
        error_detail = traceback.format_exc()
        logger.error(f"Step 1 失败: {e}\n{error_detail}")
        flash(f'Step 1 失败: {str(e)}', 'error')
        return redirect(url_for('upload_page'))

    return redirect(url_for('batch_detail', bid=batch_db_id))


@app.route('/batch/<int:bid>')
@login_required
def batch_detail(bid):
    conn = get_db()
    batch = conn.execute(
        '''SELECT b.*, u.real_name as creator_name 
           FROM batches b LEFT JOIN users u ON b.created_by=u.id
           WHERE b.id=?''', (bid,)
    ).fetchone()
    conn.close()

    if not batch:
        flash('批次不存在', 'error')
        return redirect(url_for('upload_page'))

    step1_files = json.loads(batch['step1_files']) if batch['step1_files'] else []

    return render_template('batch_detail.html', batch=batch, step1_files=step1_files)


@app.route('/batch/<int:bid>/download_step1')
@login_required
def download_step1(bid):
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    conn.close()

    if not batch or not batch['step1_files']:
        flash('文件不存在', 'error')
        return redirect(url_for('batch_detail', bid=bid))

    # 打包下载所有step1文件
    import zipfile
    zip_path = os.path.join(EXPORT_DIR, f'step1_{batch["name"]}.zip')
    files = json.loads(batch['step1_files'])

    with zipfile.ZipFile(zip_path, 'w') as zf:
        for f in files:
            if os.path.exists(f):
                zf.write(f, os.path.basename(f))

    return send_file(zip_path, as_attachment=True)


@app.route('/batch/<int:bid>/step23', methods=['POST'])
@login_required
def batch_step23(bid):
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()

    if not batch:
        conn.close()
        flash('批次不存在', 'error')
        return redirect(url_for('upload_page'))

    # 收集上传的核验结果文件
    result_files = []
    result_dir = os.path.join(RESULT_DIR, f'batch_{bid}')
    os.makedirs(result_dir, exist_ok=True)

    files = request.files.getlist('result_files')
    for f in files:
        if f.filename:
            fpath = os.path.join(result_dir, f.filename)
            f.save(fpath)
            result_files.append(fpath)

    merge_phones = 'merge_phones' in request.form
    remark = request.form.get('remark', '')
    industry_filter = request.form.get('industry_filter', '')
    industry_keywords = [kw.strip() for kw in industry_filter.split(',') if kw.strip()] if industry_filter else None
    shareholder_filter = request.form.get('shareholder_filter', '')
    insured_filter = request.form.get('insured_filter', '')
    filter_industrial = 'filter_industrial' in request.form

    # 处理核验结果前自动备份
    backup_database(f'before_step23_batch{bid}')

    try:
        importlib.reload(step23_processor)
        result = step23_processor.process_results(
            result_dir=result_dir,
            map_file=batch['map_file'],
            output_prefix=batch['name'],
            merge_phones=merge_phones,
            industry_keywords=industry_keywords,
            remark=remark,
            operator_name=session.get('real_name', '陈平安'),
            output_dir=RESULT_DIR,
            shareholder_filter=shareholder_filter,
            insured_filter=insured_filter,
            filter_industrial=filter_industrial
        )

        # 导入到客户公海（先按公司合并，再跨批次去重，最后导入）
        import pandas as pd
        import re as _re
        df = pd.read_excel(result['output_file'], dtype=str).fillna('')
        import_count = 0
        update_count = 0
        skipped_count = 0
        conn2 = get_db()
        now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # Step 1: 先按公司名称合并（同一家公司的多个号码合并到一行）
        df['公司名称'] = df['公司名称'].astype(str).str.strip()
        has_name = df[df['公司名称'] != ''].copy()
        no_name = df[df['公司名称'] == ''].copy()

        merged_rows = []
        if len(has_name) > 0:
            for company, group in has_name.groupby('公司名称', sort=False):
                all_phones = []
                for p in group['法人电话'].astype(str).tolist():
                    for ph in _re.findall(r'1\d{10}', p):
                        if ph not in all_phones:
                            all_phones.append(ph)
                phones_str = '/'.join(all_phones) if all_phones else ''
                first = group.iloc[0]
                merged_rows.append({
                    '公司名称': company,
                    '法人姓名': str(first.get('法人姓名', '')),
                    '法人电话': phones_str,
                    '注册资本': str(first.get('注册资本', '')),
                    '成立日期': str(first.get('成立日期', '')),
                    '经营状态': str(first.get('经营状态', '')),
                    '所属行业': str(first.get('所属行业', '')),
                    '注册地址': str(first.get('注册地址', '')),
                    '经营范围': str(first.get('经营范围', '')),
                    '统一社会信用代码': str(first.get('统一社会信用代码', '')),
                    '所属省份': str(first.get('所属省份', '')),
                    '所属城市': str(first.get('所属城市', '')),
                    '所属区县': str(first.get('所属区县', '')),
                    '参保人数': str(first.get('参保人数', '')),
                    '官网': str(first.get('官网', '')),
                    '备注': str(first.get('备注', '')),
                })

        if len(no_name) > 0:
            for _, row in no_name.iterrows():
                merged_rows.append({
                    '公司名称': '',
                    '法人姓名': str(row.get('法人姓名', '')),
                    '法人电话': str(row.get('法人电话', '')),
                    '注册资本': str(row.get('注册资本', '')),
                    '成立日期': str(row.get('成立日期', '')),
                    '经营状态': str(row.get('经营状态', '')),
                    '所属行业': str(row.get('所属行业', '')),
                    '注册地址': str(row.get('注册地址', '')),
                    '经营范围': str(row.get('经营范围', '')),
                    '统一社会信用代码': str(row.get('统一社会信用代码', '')),
                    '所属省份': str(row.get('所属省份', '')),
                    '所属城市': str(row.get('所属城市', '')),
                    '所属区县': str(row.get('所属区县', '')),
                    '参保人数': str(row.get('参保人数', '')),
                    '官网': str(row.get('官网', '')),
                    '备注': str(row.get('备注', '')),
                })

        df_merged = pd.DataFrame(merged_rows)
        logger.info(f"资料库导入: 原始 {len(df)} 条 -> 合并后 {len(df_merged)} 条 (公司合并减少 {len(df) - len(df_merged)} 条)")

        # Step 2: 查询资料库中已存在的记录（跨所有批次，按手机号和公司名查询）
        all_phones = set()
        all_companies = set()
        for _, row in df_merged.iterrows():
            for ph in _re.findall(r'1\d{10}', str(row.get('法人电话', ''))):
                all_phones.add(ph)
            company = str(row.get('公司名称', '')).strip()
            if company:
                all_companies.add(company)

        existing_phones = set()
        existing_companies = set()

        if all_phones:
            phones_list = list(all_phones)
            batch_size = 500
            for i in range(0, len(phones_list), batch_size):
                bp = phones_list[i:i+batch_size]
                placeholders = ','.join(['?'] * len(bp))
                rows = conn2.execute(
                    f"SELECT legal_phone FROM customers WHERE verify_result='pass' AND legal_phone IN ({placeholders})",
                    bp
                ).fetchall()
                for r in rows:
                    for ph in _re.findall(r'1\d{10}', r['legal_phone'] or ''):
                        existing_phones.add(ph)

        if all_companies:
            companies_list = list(all_companies)
            batch_size = 500
            for i in range(0, len(companies_list), batch_size):
                bc = companies_list[i:i+batch_size]
                placeholders = ','.join(['?'] * len(bc))
                rows = conn2.execute(
                    f"SELECT company_name FROM customers WHERE verify_result='pass' AND company_name IN ({placeholders})",
                    bc
                ).fetchall()
                for r in rows:
                    existing_companies.add(r['company_name'].strip() if r['company_name'] else '')

        logger.info(f"资料库导入: 已有 {len(existing_phones)} 个号码, {len(existing_companies)} 家公司已核验通过")

        # Step 3: 去重 - 保留当前批次，删除之前批次的重复记录
        insert_rows = []
        seen_phones = set()
        seen_companies = set()
        delete_phone_ids = []
        delete_company_ids = []

        # 查询需要删除的旧记录ID
        if existing_phones:
            phones_list = list(existing_phones)
            batch_size = 500
            for i in range(0, len(phones_list), batch_size):
                bp = phones_list[i:i+batch_size]
                placeholders = ','.join(['?'] * len(bp))
                rows = conn2.execute(
                    f"SELECT id, legal_phone FROM customers WHERE verify_result='pass' AND legal_phone IN ({placeholders})",
                    bp
                ).fetchall()
                for r in rows:
                    delete_phone_ids.append(r['id'])

        if existing_companies:
            companies_list = list(existing_companies)
            batch_size = 500
            for i in range(0, len(companies_list), batch_size):
                bc = companies_list[i:i+batch_size]
                placeholders = ','.join(['?'] * len(bc))
                rows = conn2.execute(
                    f"SELECT id, company_name FROM customers WHERE verify_result='pass' AND company_name IN ({placeholders})",
                    bc
                ).fetchall()
                for r in rows:
                    delete_company_ids.append(r['id'])

        all_delete_ids = set(delete_phone_ids) | set(delete_company_ids)
        skipped_count = len(all_delete_ids)

        for _, row in df_merged.iterrows():
            company = str(row.get('公司名称', '')).strip()
            phone_str = str(row.get('法人电话', ''))
            phones = _re.findall(r'1\d{10}', phone_str)

            # 本批次内去重
            if company and company in seen_companies:
                continue
            if company:
                seen_companies.add(company)
            if phones:
                if any(p in seen_phones for p in phones):
                    continue
                for p in phones:
                    seen_phones.add(p)

            insert_rows.append((
                company,
                str(row.get('法人姓名', '')),
                phone_str,
                str(row.get('注册资本', '')),
                str(row.get('成立日期', '')),
                str(row.get('经营状态', '')),
                str(row.get('所属行业', '')),
                str(row.get('注册地址', '')),
                str(row.get('经营范围', '')),
                str(row.get('统一社会信用代码', '')),
                str(row.get('所属省份', '')),
                str(row.get('所属城市', '')),
                str(row.get('所属区县', '')),
                str(row.get('参保人数', '')),
                str(row.get('官网', '')),
                batch['source'], batch['name'], now_str, str(row.get('备注', ''))
            ))

        logger.info(f"资料库导入: 删除 {skipped_count} 条旧记录, 新增 {len(insert_rows)} 条")

        # Step 4: 先删除旧记录，再批量导入新记录（每2000条一批）
        conn2.execute('BEGIN')
        if all_delete_ids:
            ids_list = list(all_delete_ids)
            batch_size = 2000
            for i in range(0, len(ids_list), batch_size):
                bi = ids_list[i:i+batch_size]
                placeholders = ','.join(['?'] * len(bi))
                conn2.execute(f"DELETE FROM customers WHERE id IN ({placeholders})", bi)

        if insert_rows:
            batch_size = 2000
            for i in range(0, len(insert_rows), batch_size):
                batch_rows = insert_rows[i:i+batch_size]
                conn2.executemany(
                    '''INSERT INTO customers (company_name, legal_person, legal_phone,
                       registered_capital, established_date, business_status, industry,
                       address, business_scope, credit_code, province, city, district,
                       insured_count, website, source, batch_name, status, created_at, remark,
                       verify_result)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'public', ?, ?, 'pass')''',
                    batch_rows
                )
                import_count += len(batch_rows)
        conn2.execute('COMMIT')
        conn2.close()

        # 已切换为数据库模式，不再需要内存缓存失效
        # （所有已核验数据按需从数据库查询，不占内存）

        # 更新批次状态
        conn.execute(
            '''UPDATE batches SET status='completed', verified=?, no_exception=?,
               step3_file=?, failed_file=? WHERE id=?''',
            (result['total_verified'], result['no_exception'],
             result['output_file'], result.get('failed_file', ''), bid)
        )
        conn.commit()
        conn.close()

        log_action(session['user_id'], 'step23_complete', str(bid),
                   f'{batch["name"]}: 删除{skipped_count}条旧记录, 新增{import_count}条')

        flash(f'Step 2+3 完成！删除 {skipped_count} 条旧记录，新增 {import_count} 条到公海', 'success')

    except Exception as e:
        try:
            conn.close()
        except:
            pass
        try:
            conn2.close()
        except:
            pass
        logger.error(f"Step 2+3 失败: {e}\n{traceback.format_exc()}")
        flash(f'Step 2+3 失败: {str(e)}', 'error')

    return redirect(url_for('batch_detail', bid=bid))


@app.route('/batch/<int:bid>/download_result')
@login_required
def download_result(bid):
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    conn.close()

    if not batch or not batch['step3_file'] or not os.path.exists(batch['step3_file']):
        flash('文件不存在', 'error')
        return redirect(url_for('batch_detail', bid=bid))

    return send_file(batch['step3_file'], as_attachment=True)


@app.route('/batch/<int:bid>/download_failed')
@login_required
def download_failed(bid):
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    conn.close()

    if not batch or not batch['failed_file'] or not os.path.exists(batch['failed_file']):
        flash('文件不存在', 'error')
        return redirect(url_for('batch_detail', bid=bid))

    return send_file(batch['failed_file'], as_attachment=True)


@app.route('/batch/<int:bid>/import_pool', methods=['POST'])
@login_required
def import_to_pool(bid):
    # 从已完成的批次导入到公海（重复调用）
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    conn.close()

    if not batch or not batch['step3_file']:
        flash('无数据可导入', 'error')
        return redirect(url_for('batch_detail', bid=bid))

    import pandas as pd
    df = pd.read_excel(batch['step3_file'], dtype=str).fillna('')
    import_count = 0

    conn2 = get_db()
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    INSERT_POOL_SQL = '''INSERT INTO customers (company_name, legal_person, legal_phone,
               registered_capital, established_date, business_status, industry,
               address, business_scope, credit_code, province, city, district,
               insured_count, source, batch_name, status, created_at, remark,
               verify_result, shareholder_info)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'public', ?, ?, 'pass', ?)'''
    POOL_BATCH = 2000
    pool_buf = []
    conn2.execute('BEGIN')
    for _, row in df.iterrows():
        shareholder = ''
        for col in ['股东信息', '大股东', '股份信息', '股东名称', '主要股东', '股东']:
            if col in row and str(row.get(col, '')).strip():
                shareholder = str(row.get(col, '')).strip()
                break

        pool_buf.append((
            str(row.get('公司名称', '')),
            str(row.get('法人姓名', '')),
            str(row.get('法人电话', '')),
            str(row.get('注册资本', '')),
            str(row.get('成立日期', '')),
            str(row.get('经营状态', '')),
            str(row.get('所属行业', '')),
            str(row.get('注册地址', '')),
            str(row.get('经营范围', '')),
            str(row.get('统一社会信用代码', '')),
            str(row.get('所属省份', '')),
            str(row.get('所属城市', '')),
            str(row.get('所属区县', '')),
            str(row.get('参保人数', '')),
            batch['source'],
            batch['name'],
            now_str,
            str(row.get('备注', '')),
            shareholder
        ))

        if len(pool_buf) >= POOL_BATCH:
            conn2.executemany(INSERT_POOL_SQL, pool_buf)
            conn2.execute('COMMIT')
            conn2.execute('BEGIN')
            import_count += len(pool_buf)
            pool_buf = []

    if pool_buf:
        conn2.executemany(INSERT_POOL_SQL, pool_buf)
        import_count += len(pool_buf)
    conn2.execute('COMMIT')
    conn2.close()

    del df, pool_buf
    gc.collect()

    log_action(session['user_id'], 'import_pool', str(bid),
               f'{batch["name"]}: 导入{import_count}条')

    flash(f'已导入 {import_count} 条客户到公海', 'success')
    return redirect(url_for('batch_detail', bid=bid))


@app.route('/batch/<int:bid>/delete', methods=['POST'])
@login_required
def batch_delete(bid):
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    if not batch:
        conn.close()
        flash('批次不存在', 'error')
        return redirect(url_for('upload_page'))

    # 管理员或创建者本人可删
    if session.get('role') != 'admin' and batch['created_by'] != session['user_id']:
        conn.close()
        flash('没有权限删除', 'error')
        return redirect(url_for('upload_page'))

    # 删除关联文件
    import shutil
    result_dir = os.path.join(RESULT_DIR, f'batch_{bid}')
    if os.path.exists(result_dir):
        shutil.rmtree(result_dir, ignore_errors=True)

    # 删除批次记录（不删除customers里的资料，保留入库的数据）
    conn.execute("DELETE FROM batches WHERE id=?", (bid,))
    conn.commit()
    conn.close()

    log_action(session['user_id'], 'batch_delete', str(bid), f'删除批次: {batch["name"]}')
    flash(f'已删除批次「{batch["name"]}」', 'success')
    return redirect(url_for('upload_page'))


# ---------- 电话系统模板 ----------

@app.route('/phone_template')
@login_required
def phone_template_list():
    conn = get_db()
    # 已完成Step1、等待核验的批次（status=step1_done）
    pending = conn.execute(
        '''SELECT b.*, u.real_name as creator_name
           FROM batches b LEFT JOIN users u ON b.created_by=u.id
           WHERE b.status='step1_done'
           ORDER BY b.id DESC'''
    ).fetchall()
    # 已完成Step2+3的批次（status=completed，有step3_file）
    completed = conn.execute(
        '''SELECT b.*, u.real_name as creator_name
           FROM batches b LEFT JOIN users u ON b.created_by=u.id
           WHERE b.status='completed'
           ORDER BY b.id DESC'''
    ).fetchall()
    conn.close()
    return render_template('phone_template.html', pending=pending, completed=completed)


@app.route('/phone_template/<int:bid>')
@login_required
def phone_template_fill(bid):
    conn = get_db()
    batch = conn.execute(
        '''SELECT b.*, u.real_name as creator_name
           FROM batches b LEFT JOIN users u ON b.created_by=u.id
           WHERE b.id=?''', (bid,)
    ).fetchone()
    conn.close()
    if not batch:
        flash('批次不存在', 'error')
        return redirect(url_for('phone_template_list'))
    step1_files = json.loads(batch['step1_files']) if batch['step1_files'] else []
    return render_template('phone_template_fill.html', batch=batch, step1_files=step1_files)


def _do_phone_template_fill(task_id, bid, params, user_id):
    """后台执行电话模板填充和导入（task_id用于更新进度）"""
    try:
        update_task_progress(task_id, progress=10, message='正在初始化...')

        conn = get_db()
        batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
        if not batch:
            conn.close()
            update_task_progress(task_id, status='failed', error='批次不存在')
            return

        result_dir = os.path.join(RESULT_DIR, f'batch_{bid}')

        merge_phones = params.get('merge_phones', False)
        remark = params.get('remark', '')
        industry_keywords = params.get('industry_keywords', None)
        import_pool = params.get('import_pool', False)
        shareholder_filter = params.get('shareholder_filter', '')
        insured_filter = params.get('insured_filter', '')
        filter_industrial = params.get('filter_industrial', False)

        update_task_progress(task_id, progress=20, message='正在处理核验结果...')

        importlib.reload(step23_processor)
        result = step23_processor.process_results(
            result_dir=result_dir,
            map_file=batch['map_file'],
            output_prefix=batch['name'],
            merge_phones=merge_phones,
            industry_keywords=industry_keywords,
            remark=remark,
            operator_name='陈平安',
            output_dir=RESULT_DIR,
            shareholder_filter=shareholder_filter,
            insured_filter=insured_filter,
            filter_industrial=filter_industrial
        )

        update_task_progress(task_id, progress=50, message='核验处理完成，正在导入资料库...')

        import_count = 0
        update_count = 0
        failed_import_count = 0
        failed_update_count = 0

        if import_pool:
            import pandas as pd
            import re as _re2

            # 使用安全读取
            df, _ = safe_read_excel(result['output_file'])
            if df is None:
                df = pd.DataFrame()

            total_in_file = len(df)
            logger.info(f"[导入资料库] 文件记录数: {total_in_file}")

            conn2 = get_db()
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # Step 1: 先按公司名称合并（同一家公司的多个号码合并到一行）
            df['公司名称'] = df['公司名称'].astype(str).str.strip()
            has_name = df[df['公司名称'] != ''].copy()
            no_name = df[df['公司名称'] == ''].copy()

            merged_rows = []
            if len(has_name) > 0:
                for company, group in has_name.groupby('公司名称', sort=False):
                    all_phones = []
                    for p in group['法人电话'].astype(str).tolist():
                        for ph in _re2.findall(r'1\d{10}', p):
                            if ph not in all_phones:
                                all_phones.append(ph)
                    phones_str = '/'.join(all_phones) if all_phones else ''
                    first = group.iloc[0]
                    shareholder = str(first.get('股份信息', '') or first.get('股东信息', '') or first.get('大股东', '') or '')
                    website_val = str(first.get('官网', '') or first.get('企业官网', '') or first.get('网站', '') or first.get('网址', ''))
                    merged_rows.append({
                        '公司名称': company,
                        '法人姓名': str(first.get('法人姓名', '')),
                        '法人电话': phones_str,
                        '注册资本': str(first.get('注册资本', '')),
                        '成立日期': str(first.get('成立日期', '')),
                        '经营状态': str(first.get('经营状态', '')),
                        '所属行业': str(first.get('所属行业', '')),
                        '注册地址': str(first.get('注册地址', '')),
                        '经营范围': str(first.get('经营范围', '')),
                        '统一社会信用代码': str(first.get('统一社会信用代码', '')),
                        '所属省份': str(first.get('所属省份', '')),
                        '所属城市': str(first.get('所属城市', '')),
                        '所属区县': str(first.get('所属区县', '')),
                        '参保人数': str(first.get('参保人数', '')),
                        '官网': website_val,
                        '备注': str(first.get('备注', '')),
                        '股东信息': shareholder,
                    })

            if len(no_name) > 0:
                for _, row in no_name.iterrows():
                    shareholder = str(row.get('股份信息', '') or row.get('股东信息', '') or row.get('大股东', '') or '')
                    website_val = str(row.get('官网', '') or row.get('企业官网', '') or row.get('网站', '') or row.get('网址', ''))
                    merged_rows.append({
                        '公司名称': '',
                        '法人姓名': str(row.get('法人姓名', '')),
                        '法人电话': str(row.get('法人电话', '')),
                        '注册资本': str(row.get('注册资本', '')),
                        '成立日期': str(row.get('成立日期', '')),
                        '经营状态': str(row.get('经营状态', '')),
                        '所属行业': str(row.get('所属行业', '')),
                        '注册地址': str(row.get('注册地址', '')),
                        '经营范围': str(row.get('经营范围', '')),
                        '统一社会信用代码': str(row.get('统一社会信用代码', '')),
                        '所属省份': str(row.get('所属省份', '')),
                        '所属城市': str(row.get('所属城市', '')),
                        '所属区县': str(row.get('所属区县', '')),
                        '参保人数': str(row.get('参保人数', '')),
                        '官网': website_val,
                        '备注': str(row.get('备注', '')),
                        '股东信息': shareholder,
                    })

            df_merged = pd.DataFrame(merged_rows) if merged_rows else pd.DataFrame()
            logger.info(f"[导入资料库] 原始 {total_in_file} 条 -> 合并后 {len(df_merged)} 条")

            if len(df_merged) == 0:
                logger.warning("[导入资料库] 没有有效记录可导入")
            else:
                # Step 2: 查询资料库中已存在的记录（跨所有批次，按手机号和公司名查询）
                all_phones = set()
                all_companies = set()
                for _, row in df_merged.iterrows():
                    for ph in _re2.findall(r'1\d{10}', str(row.get('法人电话', ''))):
                        all_phones.add(ph)
                    company = str(row.get('公司名称', '')).strip()
                    if company:
                        all_companies.add(company)

                existing_phones = set()
                existing_companies = set()

                if all_phones:
                    phones_list = list(all_phones)
                    batch_size_q = 500
                    for i in range(0, len(phones_list), batch_size_q):
                        bp = phones_list[i:i+batch_size_q]
                        placeholders = ','.join(['?'] * len(bp))
                        rows = conn2.execute(
                            f"SELECT legal_phone FROM customers WHERE verify_result='pass' AND legal_phone IN ({placeholders})",
                            bp
                        ).fetchall()
                        for r in rows:
                            for ph in _re2.findall(r'1\d{10}', r['legal_phone'] or ''):
                                existing_phones.add(ph)

                if all_companies:
                    companies_list = list(all_companies)
                    batch_size_q = 500
                    for i in range(0, len(companies_list), batch_size_q):
                        bc = companies_list[i:i+batch_size_q]
                        placeholders = ','.join(['?'] * len(bc))
                        rows = conn2.execute(
                            f"SELECT company_name FROM customers WHERE verify_result='pass' AND company_name IN ({placeholders})",
                            bc
                        ).fetchall()
                        for r in rows:
                            existing_companies.add(r['company_name'].strip() if r['company_name'] else '')

                logger.info(f"[导入资料库] 已有 {len(existing_phones)} 个号码, {len(existing_companies)} 家公司已核验通过")

                # Step 3: 查询需要删除的旧记录ID
                delete_ids = set()
                if existing_phones:
                    phones_list = list(existing_phones)
                    batch_size_q = 500
                    for i in range(0, len(phones_list), batch_size_q):
                        bp = phones_list[i:i+batch_size_q]
                        placeholders = ','.join(['?'] * len(bp))
                        rows = conn2.execute(
                            f"SELECT id FROM customers WHERE verify_result='pass' AND legal_phone IN ({placeholders})",
                            bp
                        ).fetchall()
                        for r in rows:
                            delete_ids.add(r['id'])

                if existing_companies:
                    companies_list = list(existing_companies)
                    batch_size_q = 500
                    for i in range(0, len(companies_list), batch_size_q):
                        bc = companies_list[i:i+batch_size_q]
                        placeholders = ','.join(['?'] * len(bc))
                        rows = conn2.execute(
                            f"SELECT id FROM customers WHERE verify_result='pass' AND company_name IN ({placeholders})",
                            bc
                        ).fetchall()
                        for r in rows:
                            delete_ids.add(r['id'])

                skipped_count = len(delete_ids)

                # Step 4: 准备插入数据（本批次内去重）
                insert_rows = []
                seen_phones = set()
                seen_companies = set()
                total_rows = len(df_merged)

                for _, row in df_merged.iterrows():
                    company = str(row.get('公司名称', '')).strip()
                    phone_str = str(row.get('法人电话', ''))
                    phones = _re2.findall(r'1\d{10}', phone_str)

                    # 本批次内去重
                    if company and company in seen_companies:
                        continue
                    if company:
                        seen_companies.add(company)
                    if phones:
                        if any(p in seen_phones for p in phones):
                            continue
                        for p in phones:
                            seen_phones.add(p)

                    shareholder = str(row.get('股东信息', ''))
                    website_val = str(row.get('官网', ''))
                    insert_rows.append((
                        company,
                        str(row.get('法人姓名', '')),
                        phone_str,
                        str(row.get('注册资本', '')),
                        str(row.get('成立日期', '')),
                        str(row.get('经营状态', '')),
                        str(row.get('所属行业', '')),
                        str(row.get('注册地址', '')),
                        str(row.get('经营范围', '')),
                        str(row.get('统一社会信用代码', '')),
                        str(row.get('所属省份', '')),
                        str(row.get('所属城市', '')),
                        str(row.get('所属区县', '')),
                        str(row.get('参保人数', '')),
                        website_val,
                        batch['source'], batch['name'], now_str, str(row.get('备注', '')),
                        'pass', shareholder
                    ))

                logger.info(f"[导入资料库] 删除 {skipped_count} 条旧记录, 新增 {len(insert_rows)} 条")

                # Step 5: 先删除旧记录，再批量导入新记录
                INSERT_SQL2 = '''INSERT INTO customers (company_name, legal_person, legal_phone,
                       registered_capital, established_date, business_status, industry,
                       address, business_scope, credit_code, province, city, district,
                       insured_count, website, source, batch_name, status, created_at, remark,
                       verify_result, shareholder_info)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'public', ?, ?, ?, ?)'''

                conn2.execute('BEGIN')
                if delete_ids:
                    ids_list = list(delete_ids)
                    batch_size_d = 2000
                    for i in range(0, len(ids_list), batch_size_d):
                        bi = ids_list[i:i+batch_size_d]
                        placeholders = ','.join(['?'] * len(bi))
                        conn2.execute(f"DELETE FROM customers WHERE id IN ({placeholders})", bi)

                if insert_rows:
                    batch_size_i = 2000
                    for i in range(0, len(insert_rows), batch_size_i):
                        batch_rows = insert_rows[i:i+batch_size_i]
                        conn2.executemany(INSERT_SQL2, batch_rows)
                        import_count += len(batch_rows)

                        pct = 50 + int((i + len(batch_rows)) / max(total_rows, 1) * 35)
                        update_task_progress(task_id, progress=min(pct, 85),
                                           message=f'正在导入资料库 ({i + len(batch_rows)}/{total_rows})')

                conn2.execute('COMMIT')

                del df_merged, insert_rows
                del existing_phones, existing_companies
                gc.collect()

                logger.info(f"[导入资料库] 完成: 删除 {skipped_count} 条旧记录, 新增 {import_count} 条")

            del df
            gc.collect()

            # 导入失败记录到资料库（分批优化）
            if result.get('failed_file') and os.path.exists(result['failed_file']):
                update_task_progress(task_id, progress=88, message='正在导入失败记录...')
                df_failed, _ = safe_read_excel(result['failed_file'])
                if df_failed is not None:
                    logger.info(f"[导入失败记录] 文件记录数: {len(df_failed)}")

                    # 用电话去重
                    f_phones = df_failed['收款方账号'].astype(str).str.strip().tolist() if '收款方账号' in df_failed.columns else []
                    f_existing_by_phone = {}

                    phones_with_val = [p for p in f_phones if p and p.lower() != 'nan']
                    if phones_with_val:
                        phones_unique = list(set(phones_with_val))
                        batch_size = 500
                        for i in range(0, len(phones_unique), batch_size):
                            bp = phones_unique[i:i+batch_size]
                            placeholders = ','.join(['?'] * len(bp))
                            erows = conn2.execute(
                                f"SELECT id, company_name, legal_phone FROM customers WHERE legal_phone IN ({placeholders}) AND batch_name=?",
                                bp + [batch['name']]
                            ).fetchall()
                            for r in erows:
                                f_existing_by_phone[r['legal_phone']] = r['id']

                    logger.info(f"[导入失败记录] 已存在记录(按电话匹配): {len(f_existing_by_phone)} 条")

                    f_insert_rows = []
                    f_update_rows = []
                    F_UPDATE_SQL = '''UPDATE customers SET
                                legal_person=?, registered_capital=?, established_date=?,
                                business_status=?, industry=?, address=?, business_scope=?,
                                credit_code=?, province=?, city=?, district=?, insured_count=?,
                                source=?, verify_result='failed', fail_reason=?
                                WHERE id=?'''
                    F_INSERT_SQL = '''INSERT INTO customers (company_name, legal_person, legal_phone,
                               registered_capital, established_date, business_status, industry,
                               address, business_scope, credit_code, province, city, district,
                               insured_count, source, batch_name, status, created_at, remark,
                               verify_result, fail_reason)
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'public', ?, ?, 'failed', ?)'''
                    f_batch_size = 2000
                    conn2.execute('BEGIN')
                    for _, row in df_failed.iterrows():
                        company = str(row.get('企业名称', '')).strip()
                        if company.lower() == 'nan':
                            company = ''
                        phone = str(row.get('收款方账号', '')).strip()
                        if phone.lower() == 'nan':
                            phone = ''
                        fail_reason = str(row.get('账户异常原因', ''))

                        # 公司名和电话都为空的跳过
                        if not company and not phone:
                            continue

                        f_data = (
                            str(row.get('法定代表人', '')),
                            str(row.get('注册资本', '')),
                            str(row.get('成立日期', '')),
                            str(row.get('经营状态', '')),
                            str(row.get('所属行业', '')),
                            str(row.get('注册地址', '')),
                            str(row.get('经营范围', '')),
                            str(row.get('统一社会信用代码', '')),
                            str(row.get('所属省份', '')),
                            str(row.get('所属城市', '')),
                            str(row.get('所属区县', '')),
                            str(row.get('参保人数', '')),
                            batch['source'],
                            fail_reason,
                        )

                        # 判断是否已存在
                        existing_id = None
                        if phone and phone in f_existing_by_phone:
                            existing_id = f_existing_by_phone[phone]

                        if existing_id is not None:
                            f_update_rows.append(f_data + (existing_id,))
                            if len(f_update_rows) >= f_batch_size:
                                conn2.executemany(F_UPDATE_SQL, f_update_rows)
                                conn2.execute('COMMIT')
                                conn2.execute('BEGIN')
                                failed_update_count += len(f_update_rows)
                                f_update_rows = []
                        else:
                            f_insert_rows.append((
                                company, str(row.get('法定代表人', '')), phone,
                                str(row.get('注册资本', '')), str(row.get('成立日期', '')),
                                str(row.get('经营状态', '')), str(row.get('所属行业', '')),
                                str(row.get('注册地址', '')), str(row.get('经营范围', '')),
                                str(row.get('统一社会信用代码', '')), str(row.get('所属省份', '')),
                                str(row.get('所属城市', '')), str(row.get('所属区县', '')),
                                str(row.get('参保人数', '')),
                                batch['source'], batch['name'], now_str, '', fail_reason
                            ))
                            if phone:
                                f_existing_by_phone[phone] = -1
                            if len(f_insert_rows) >= f_batch_size:
                                conn2.executemany(F_INSERT_SQL, f_insert_rows)
                                conn2.execute('COMMIT')
                                conn2.execute('BEGIN')
                                failed_import_count += len(f_insert_rows)
                                f_insert_rows = []

                    if f_update_rows:
                        conn2.executemany(F_UPDATE_SQL, f_update_rows)
                        failed_update_count += len(f_update_rows)
                    if f_insert_rows:
                        conn2.executemany(F_INSERT_SQL, f_insert_rows)
                        failed_import_count += len(f_insert_rows)
                    conn2.execute('COMMIT')

                    del df_failed, f_insert_rows, f_update_rows, f_existing_by_phone
                    gc.collect()

                    logger.info(f"[导入失败记录] 完成: 新增 {failed_import_count} 条, 更新 {failed_update_count} 条")

            conn2.close()
            cache_invalidate()  # 清除缓存，让下次查询重新加载

        update_task_progress(task_id, progress=95, message='正在更新批次状态...')

        conn.execute(
            '''UPDATE batches SET status='completed', verified=?, no_exception=?,
               step3_file=?, failed_file=? WHERE id=?''',
            (result['total_verified'], result['no_exception'],
             result['output_file'], result.get('failed_file', ''), bid)
        )
        conn.commit()
        conn.close()

        try:
            log_action(user_id, 'phone_fill', str(bid),
                       f'{batch["name"]}: {result["final_count"]}条通过, {failed_import_count+failed_update_count}条失败, 新增{import_count + failed_import_count}条, 更新{update_count + failed_update_count}条入资料库')
        except:
            pass

        result_data = {
            'final_count': result['final_count'],
            'total_verified': result['total_verified'],
            'no_exception': result['no_exception'],
            'matched_count': result.get('matched_count', 0),
            'industrial_filtered': result.get('industrial_filtered', 0),
            'failed_count': failed_import_count + failed_update_count,
            'import_count': import_count + failed_import_count,
            'update_count': update_count + failed_update_count,
            'pass_import_count': import_count,
            'pass_update_count': update_count,
            'failed_import_count': failed_import_count,
            'failed_update_count': failed_update_count,
            'import_pool': import_pool,
            'output_file': result['output_file'],
            'failed_file': result.get('failed_file', ''),
        }

        done_msg = f'完成！核验{result["total_verified"]}条，通过{result["final_count"]}条'
        if import_pool:
            done_msg += f'，新增{import_count + failed_import_count}条，更新{update_count + failed_update_count}条'

        update_task_progress(task_id, status='done', progress=100,
                           message=done_msg,
                           result=result_data)

    except Exception as e:
        logger.error(f"电话模板填充任务 {task_id} 失败: {e}\n{traceback.format_exc()}")
        update_task_progress(task_id, status='failed', error=str(e),
                           message=f'处理失败: {str(e)}')
        try:
            conn.close()
        except:
            pass
        try:
            conn2.close()
        except:
            pass


@app.route('/phone_template/<int:bid>/fill', methods=['POST'])
@login_required
def phone_template_do_fill(bid):
    """电话系统模板填充（后台异步执行，防止HTTP阻塞）"""
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    if not batch:
        conn.close()
        flash('批次不存在', 'error')
        return redirect(url_for('phone_template_list'))

    # 收集核验结果文件
    result_dir = os.path.join(RESULT_DIR, f'batch_{bid}')
    os.makedirs(result_dir, exist_ok=True)
    files = request.files.getlist('result_files')
    for f in files:
        if f.filename:
            f.save(os.path.join(result_dir, f.filename))
    conn.close()

    merge_phones = 'merge_phones' in request.form
    remark = request.form.get('remark', '')
    industry_filter = request.form.get('industry_filter', '')
    industry_keywords = [kw.strip() for kw in industry_filter.split(',') if kw.strip()] if industry_filter else None
    import_pool = 'import_pool' in request.form
    shareholder_filter = request.form.get('shareholder_filter', '')
    insured_filter = request.form.get('insured_filter', '')
    filter_industrial = 'filter_industrial' in request.form

    params = {
        'merge_phones': merge_phones,
        'remark': remark,
        'industry_keywords': industry_keywords,
        'import_pool': import_pool,
        'shareholder_filter': shareholder_filter,
        'insured_filter': insured_filter,
        'filter_industrial': filter_industrial,
    }

    # 创建后台任务
    task_id = create_task('phone_fill', f'电话模板填充 - {batch["name"]}')
    run_background_task(task_id, _do_phone_template_fill,
                       task_id, bid, params, session['user_id'])

    return jsonify({
        'task_id': task_id,
        'redirect_url': url_for('phone_template_fill', bid=bid, task_id=task_id)
    })


@app.route('/phone_template/<int:bid>/download')
@login_required
def phone_template_download(bid):
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    conn.close()
    if not batch or not batch['step3_file'] or not os.path.exists(batch['step3_file']):
        flash('文件不存在', 'error')
        return redirect(url_for('phone_template_list'))
    return send_file(batch['step3_file'], as_attachment=True)


@app.route('/phone_template/<int:bid>/download_failed')
@login_required
def phone_template_download_failed(bid):
    conn = get_db()
    batch = conn.execute("SELECT * FROM batches WHERE id=?", (bid,)).fetchone()
    conn.close()
    if not batch or not batch['failed_file'] or not os.path.exists(batch['failed_file']):
        flash('文件不存在', 'error')
        return redirect(url_for('phone_template_list'))
    return send_file(batch['failed_file'], as_attachment=True)


# ---------- 资料库 ----------

@app.route('/library')
@login_required
def library():
    page = int(request.args.get('page', 1))
    per_page = 20
    keyword = request.args.get('keyword', '').strip()
    # 省份支持多选
    selected_provinces = request.args.getlist('provinces')
    if not selected_provinces:
        single = request.args.get('province', '').strip()
        if single:
            selected_provinces = [single]
    # 城市支持多选
    selected_cities = request.args.getlist('cities')
    if not selected_cities:
        single_city = request.args.get('city', '').strip()
        if single_city:
            selected_cities = [single_city]
    # 区县支持多选
    selected_districts = request.args.getlist('districts')
    if not selected_districts:
        single_district = request.args.get('district', '').strip()
        if single_district:
            selected_districts = [single_district]
    source = request.args.get('source', '').strip()
    batch = request.args.get('batch', '').strip()
    verify = request.args.get('verify', 'all')  # all/pass/failed
    category = request.args.get('category', '').strip()
    min_insured = request.args.get('min_insured', '').strip()
    max_insured = request.args.get('max_insured', '').strip()
    max_insured_dup = request.args.get('max_insured_dup', '') == '1'

    conn = get_db()

    conditions = []
    params = []

    if verify == 'pass':
        conditions.append("verify_result='pass'")
    elif verify == 'failed':
        conditions.append("verify_result='failed'")

    if keyword:
        conditions.append("(company_name LIKE ? OR legal_person LIKE ? OR legal_phone LIKE ?)")
        params.extend([f'%{keyword}%'] * 3)

    # 品类筛选：匹配所属行业和经营范围
    if category:
        cat_keywords = [kw.strip() for kw in category.split(',') if kw.strip()]
        if cat_keywords:
            cat_conditions = []
            for kw in cat_keywords:
                cat_conditions.append("(industry LIKE ? OR business_scope LIKE ?)")
                params.extend([f'%{kw}%', f'%{kw}%'])
            # 多个关键词之间是"或"的关系（匹配任意一个即可）
            conditions.append('(' + ' OR '.join(cat_conditions) + ')')

    # 参保人数筛选
    if min_insured:
        try:
            threshold = int(float(min_insured))
            conditions.append("parse_insured(insured_count) >= ?")
            params.append(threshold)
        except (ValueError, TypeError):
            pass
    if max_insured:
        try:
            threshold = int(float(max_insured))
            conditions.append("parse_insured(insured_count) <= ?")
            params.append(threshold)
        except (ValueError, TypeError):
            pass

    if selected_provinces:
        placeholders = ','.join(['?'] * len(selected_provinces))
        conditions.append(f"province IN ({placeholders})")
        params.extend(selected_provinces)

    if selected_cities:
        placeholders = ','.join(['?'] * len(selected_cities))
        conditions.append(f"city IN ({placeholders})")
        params.extend(selected_cities)

    if selected_districts:
        placeholders = ','.join(['?'] * len(selected_districts))
        conditions.append(f"district IN ({placeholders})")
        params.extend(selected_districts)

    if source:
        conditions.append("source LIKE ?")
        params.append(f'%{source}%')

    if batch:
        conditions.append("batch_name=?")
        params.append(batch)

    where_sql = ' AND '.join(conditions) if conditions else '1=1'

    total = conn.execute(
        f"SELECT COUNT(*) FROM customers WHERE {where_sql}", params
    ).fetchone()[0]

    offset = (page - 1) * per_page
    records = conn.execute(
        f'''SELECT * FROM customers WHERE {where_sql}
            ORDER BY id DESC LIMIT ? OFFSET ?''',
        params + [per_page, offset]
    ).fetchall()

    # 省份列表（缓存）
    provinces = cache_get('library_provinces')
    if provinces is None:
        provinces = conn.execute(
            "SELECT DISTINCT province FROM customers WHERE province!='' ORDER BY province"
        ).fetchall()
        cache_set('library_provinces', provinces)

    # 城市列表（按选中省份过滤，如果有选省份的话）
    city_cache_key = f'library_cities_{"_".join(sorted(selected_provinces))}'
    cities = cache_get(city_cache_key)
    if cities is None:
        city_where = "WHERE city!=''"
        city_params = []
        if selected_provinces:
            placeholders = ','.join(['?'] * len(selected_provinces))
            city_where += f" AND province IN ({placeholders})"
            city_params.extend(selected_provinces)
        cities = conn.execute(
            f"SELECT DISTINCT city FROM customers {city_where} ORDER BY city",
            city_params
        ).fetchall()
        cache_set(city_cache_key, cities)

    # 区县列表（按选中省份和城市过滤）
    district_cache_key = f'library_districts_{"_".join(sorted(selected_provinces))}_{"_".join(sorted(selected_cities))}'
    districts = cache_get(district_cache_key)
    if districts is None:
        district_where = "WHERE district!=''"
        district_params = []
        if selected_provinces:
            placeholders = ','.join(['?'] * len(selected_provinces))
            district_where += f" AND province IN ({placeholders})"
            district_params.extend(selected_provinces)
        if selected_cities:
            placeholders = ','.join(['?'] * len(selected_cities))
            district_where += f" AND city IN ({placeholders})"
            district_params.extend(selected_cities)
        districts = conn.execute(
            f"SELECT DISTINCT district FROM customers {district_where} ORDER BY district",
            district_params
        ).fetchall()
        cache_set(district_cache_key, districts)

    # 来源列表（缓存）
    sources = cache_get('library_sources')
    if sources is None:
        db_sources = [r['source'] for r in conn.execute(
            "SELECT DISTINCT source FROM customers WHERE source!='' ORDER BY source"
        ).fetchall()]
        all_sources = list(dict.fromkeys(db_sources + B2B_PLATFORMS))
        sources = [{'source': s} for s in all_sources]
        cache_set('library_sources', sources)

    # 统计（缓存，单次聚合）
    stats = cache_get('library_stats')
    if stats is None:
        stats_row = conn.execute('''SELECT 
            COUNT(*) as total_all,
            SUM(CASE WHEN verify_result='pass' THEN 1 ELSE 0 END) as total_pass,
            SUM(CASE WHEN verify_result='failed' THEN 1 ELSE 0 END) as total_failed
            FROM customers''').fetchone()
        stats = {
            'total_all': stats_row['total_all'] or 0,
            'total_pass': stats_row['total_pass'] or 0,
            'total_failed': stats_row['total_failed'] or 0,
        }
        cache_set('library_stats', stats)

    # 按省份统计（缓存）
    province_stats = cache_get('library_province_stats')
    if province_stats is None:
        province_stats = conn.execute(
            '''SELECT province, 
                      SUM(CASE WHEN verify_result='pass' THEN 1 ELSE 0 END) as pass_count,
                      SUM(CASE WHEN verify_result='failed' THEN 1 ELSE 0 END) as failed_count,
                      COUNT(*) as total
               FROM customers 
               WHERE province!='' 
               GROUP BY province 
               ORDER BY total DESC'''
        ).fetchall()
        cache_set('library_province_stats', province_stats)

    # 按平台统计（缓存）
    source_stats = cache_get('library_source_stats')
    if source_stats is None:
        db_source_stats = conn.execute(
            '''SELECT source,
                      SUM(CASE WHEN verify_result='pass' THEN 1 ELSE 0 END) as pass_count,
                      SUM(CASE WHEN verify_result='failed' THEN 1 ELSE 0 END) as failed_count,
                      COUNT(*) as total
               FROM customers
               WHERE source!=''
               GROUP BY source
               ORDER BY total DESC'''
        ).fetchall()
        
        source_stats_map = {}
        for row in db_source_stats:
            s_name = row['source']
            matched = False
            for platform in B2B_PLATFORMS:
                if platform in s_name or s_name in platform:
                    matched = True
                    if platform in source_stats_map:
                        source_stats_map[platform]['pass_count'] += row['pass_count']
                        source_stats_map[platform]['failed_count'] += row['failed_count']
                        source_stats_map[platform]['total'] += row['total']
                    else:
                        source_stats_map[platform] = {
                            'source': platform,
                            'pass_count': row['pass_count'],
                            'failed_count': row['failed_count'],
                            'total': row['total']
                        }
                    break
        if not matched:
            source_stats_map[s_name] = dict(row)
        
        for platform in B2B_PLATFORMS:
            if platform not in source_stats_map:
                source_stats_map[platform] = {
                    'source': platform,
                    'pass_count': 0,
                    'failed_count': 0,
                    'total': 0
                }
        
        source_stats = sorted(source_stats_map.values(), key=lambda x: x['total'], reverse=True)
        cache_set('library_source_stats', source_stats)

    # 按任务统计（缓存）
    batch_stats = cache_get('library_batch_stats')
    if batch_stats is None:
        batch_stats = conn.execute(
            '''SELECT batch_name,
                      SUM(CASE WHEN verify_result='pass' THEN 1 ELSE 0 END) as pass_count,
                      SUM(CASE WHEN verify_result='failed' THEN 1 ELSE 0 END) as failed_count,
                      COUNT(*) as total
               FROM customers
               WHERE batch_name!=''
               GROUP BY batch_name
               ORDER BY total DESC'''
        ).fetchall()
        cache_set('library_batch_stats', batch_stats)

    # 批次列表（缓存）
    batches = cache_get('library_batches')
    if batches is None:
        batches = conn.execute(
            "SELECT DISTINCT batch_name FROM customers WHERE batch_name!='' ORDER BY batch_name"
        ).fetchall()
        cache_set('library_batches', batches)

    conn.close()

    total_pages = (total + per_page - 1) // per_page

    # 构建省份查询字符串（用于模板URL拼接）
    province_query = ''.join([f'&provinces={p}' for p in selected_provinces])
    city_query = ''.join([f'&cities={c}' for c in selected_cities])
    district_query = ''.join([f'&districts={d}' for d in selected_districts])
    category_query = f'&category={category}' if category else ''
    min_insured_query = f'&min_insured={min_insured}' if min_insured else ''
    max_insured_query = f'&max_insured={max_insured}' if max_insured else ''

    # 构建导出链接（后端直接生成，避免Jinja2转义 & 为 &amp;）
    from urllib.parse import quote
    export_parts = [f'verify={quote(verify, safe="")}', f'source={quote(source, safe="")}', f'batch={quote(batch, safe="")}']
    if category:
        export_parts.append(f'category={quote(category, safe="")}')
    if min_insured:
        export_parts.append(f'min_insured={quote(min_insured, safe="")}')
    if max_insured:
        export_parts.append(f'max_insured={quote(max_insured, safe="")}')
    for p in selected_provinces:
        export_parts.append(f'provinces={quote(p, safe="")}')
    for c in selected_cities:
        export_parts.append(f'cities={quote(c, safe="")}')
    for d in selected_districts:
        export_parts.append(f'districts={quote(d, safe="")}')
    export_url = '/library/export?' + '&'.join(export_parts)

    return render_template('library.html',
                           records=records,
                           page=page, total_pages=total_pages, total=total,
                           keyword=keyword, category=category,
                           selected_provinces=selected_provinces, selected_cities=selected_cities,
                           selected_districts=selected_districts,
                           province_query=province_query, city_query=city_query,
                           district_query=district_query,
                           category_query=category_query,
                           min_insured=min_insured, min_insured_query=min_insured_query,
                           max_insured=max_insured, max_insured_query=max_insured_query,
                           export_url=export_url,
                           source=source, batch=batch, verify=verify,
                           provinces=provinces, cities=cities, districts=districts,
                           sources=sources, batches=batches,
                           stats=stats, province_stats=province_stats,
                           source_stats=source_stats, batch_stats=batch_stats)


@app.route('/library/<int:cid>')
@login_required
def library_detail(cid):
    conn = get_db()
    record = conn.execute("SELECT * FROM customers WHERE id=?", (cid,)).fetchone()
    conn.close()
    if not record:
        flash('记录不存在', 'error')
        return redirect(url_for('library'))
    return render_template('library_detail.html', record=record)


def _do_library_import(task_id, upload_paths, source, filter_industrial, remark, user_id):
    """后台执行资料库导入（task_id用于更新进度）"""
    import pandas as pd

    # 列名兼容映射
    col_aliases = {
        '法人姓名': ['法人姓名', '法定代表人', '法人'],
        '法人电话': ['法人电话', '电话', '手机号码', '联系电话', '手机', '手机号'],
        '注册资本': ['注册资本'],
        '成立日期': ['成立日期', '注册时间'],
        '经营状态': ['经营状态', '企业状态'],
        '所属行业': ['所属行业', '行业'],
        '注册地址': ['注册地址', '地址'],
        '经营范围': ['经营范围'],
        '统一社会信用代码': ['统一社会信用代码', '信用代码'],
        '所属省份': ['所属省份', '省份'],
        '所属城市': ['所属城市', '城市'],
        '所属区县': ['所属区县', '区县'],
        '参保人数': ['参保人数', '社保人数', '员工人数', '人员规模'],
        '官网': ['官网', '企业官网', '网站', '网址'],
        '股份信息': ['股份信息', '股东信息', '大股东', '股东名称', '主要股东', '股东'],
    }

    def get_val(row, key):
        for alias in col_aliases.get(key, [key]):
            # 模糊匹配：列名包含关键词即可
            for col in row.index:
                col_str = str(col)
                if alias in col_str:
                    val = str(row[col]).strip()
                    if val and val.lower() != 'nan':
                        return val
        return ''

    try:
        # 导入前自动备份
        update_task_progress(task_id, progress=8, message='正在备份数据库...')
        backup_database('before_import')
    except Exception as e:
        logger.warning(f"导入前备份失败: {e}")

    conn = get_db()
    import_count = 0
    file_count = 0
    errors = []
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    total_files = len(upload_paths)

    INSERT_SQL = '''INSERT INTO customers (company_name, legal_person, legal_phone,
                   registered_capital, established_date, business_status, industry,
                   address, business_scope, credit_code, province, city, district,
                   insured_count, website, source, batch_name, status, created_at, remark,
                   verify_result, shareholder_info)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'public', ?, ?, 'pass', ?)'''
    BATCH_SIZE = 2000  # 每批插入2000条

    for fi, (upload_path, original_name) in enumerate(upload_paths):
        file_count += 1
        base_progress = 10 + int((fi / total_files) * 85)  # 10% ~ 95%

        try:
            update_task_progress(task_id, progress=base_progress,
                               message=f'正在读取文件 {fi+1}/{total_files}: {original_name}')

            # 使用安全读取Excel
            df, read_err = safe_read_excel(upload_path)
            if df is None:
                errors.append(f'{original_name}: 无法读取 ({read_err})')
                continue

            # 列名标准化：企业名称 -> 公司名称
            for alias in ['企业名称', '公司名称', '名称', '单位名称', '机构名称', '商户名称', '店铺名称']:
                if alias in df.columns and alias != '公司名称':
                    df = df.rename(columns={alias: '公司名称'})
                    logger.info(f'{original_name}: 列名标准化 {alias} -> 公司名称')
                    break

            if '公司名称' not in df.columns:
                errors.append(f'{original_name}: 未找到"公司名称"列')
                del df
                gc.collect()
                continue

            # 强制工业品筛选
            if filter_industrial:
                try:
                    from step23_processor import is_excluded_industry
                    industry_col = None
                    for col in ['所属行业', '行业', '国民经济行业', '行业分类', '行业类型']:
                        if col in df.columns:
                            industry_col = col
                            break
                    if industry_col:
                        before_count = len(df)
                        df = df[~df[industry_col].apply(is_excluded_industry)].reset_index(drop=True)
                        logger.info(f'{original_name}: 工业品筛选过滤 {before_count - len(df)} 家，剩余 {len(df)} 家')
                except Exception as e:
                    logger.warning(f"工业品筛选失败: {e}")

            # 分批处理：每2000条提交一次，避免内存堆积
            file_import = 0
            batch_rows = []
            total_rows = len(df)
            conn.execute('BEGIN')
            for idx, row in df.iterrows():
                company_name = str(row.get('公司名称', '')).strip()
                if not company_name or company_name.lower() == 'nan':
                    continue

                batch_rows.append((
                    company_name,
                    get_val(row, '法人姓名'),
                    get_val(row, '法人电话'),
                    get_val(row, '注册资本'),
                    get_val(row, '成立日期'),
                    get_val(row, '经营状态'),
                    get_val(row, '所属行业'),
                    get_val(row, '注册地址'),
                    get_val(row, '经营范围'),
                    get_val(row, '统一社会信用代码'),
                    get_val(row, '所属省份'),
                    get_val(row, '所属城市'),
                    get_val(row, '所属区县'),
                    get_val(row, '参保人数'),
                    get_val(row, '官网'),
                    source,
                    '直接导入',
                    now,
                    remark,
                    get_val(row, '股份信息')
                ))
                file_import += 1

                if len(batch_rows) >= BATCH_SIZE:
                    conn.executemany(INSERT_SQL, batch_rows)
                    conn.execute('COMMIT')
                    conn.execute('BEGIN')
                    batch_rows = []
                    # 更新进度
                    if total_rows > 0:
                        row_pct = (idx + 1) / total_rows
                        file_pct = base_progress + int(row_pct * 85 / total_files)
                        update_task_progress(task_id, progress=file_pct,
                                           message=f'导入中 {fi+1}/{total_files}: {original_name} ({idx+1}/{total_rows})')

            # 插入剩余记录
            if batch_rows:
                conn.executemany(INSERT_SQL, batch_rows)
            conn.execute('COMMIT')

            import_count += file_import
            del df, batch_rows
            gc.collect()
            logger.info(f'{original_name}: 导入完成，共 {file_import} 条')

        except Exception as e:
            logger.error(f"导入文件 {original_name} 失败: {e}\n{traceback.format_exc()}")
            errors.append(f'{original_name}: 导入失败 ({e})')
            try:
                conn.execute('ROLLBACK')
            except:
                pass
            continue

    try:
        conn.close()
    except:
        pass
    cache_invalidate()

    try:
        log_action(user_id, 'library_import', f'{file_count}个文件',
                   f'导入{import_count}条, 来源: {source}')
    except:
        pass

    result = {
        'import_count': import_count,
        'file_count': file_count,
        'errors': errors,
    }

    if errors:
        update_task_progress(task_id, status='done', progress=100,
                           message=f'导入完成：成功 {import_count} 条，{len(errors)} 个文件失败',
                           result=result)
    else:
        update_task_progress(task_id, status='done', progress=100,
                           message=f'导入完成：成功 {import_count} 条（{file_count} 个文件）',
                           result=result)


@app.route('/library/import', methods=['POST'])
@login_required
def library_import():
    """直接导入Excel文件到资料库（后台异步执行，防止HTTP阻塞）"""
    files = request.files.getlist('files')
    if not files or all(not f.filename for f in files):
        flash('请选择文件', 'error')
        return redirect(url_for('library'))

    source = request.form.get('source', '').strip()
    filter_industrial = 'filter_industrial' in request.form
    remark = request.form.get('remark', '').strip()

    # 先保存所有上传文件
    upload_paths = []
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    for i, file in enumerate(files):
        if not file.filename:
            continue
        upload_path = os.path.join(UPLOAD_DIR, f'library_import_{timestamp}_{i}_{file.filename}')
        file.save(upload_path)
        upload_paths.append((upload_path, file.filename))

    if not upload_paths:
        flash('请选择文件', 'error')
        return redirect(url_for('library'))

    # 创建后台任务
    task_id = create_task('library_import', f'资料库导入 ({len(upload_paths)}个文件)')
    run_background_task(task_id, _do_library_import,
                       task_id, upload_paths, source, filter_industrial, remark, session['user_id'])

    # 重定向到资料库页面，前端轮询进度
    return redirect(url_for('library', task_id=task_id))


@app.route('/api/task/<task_id>')
@login_required
def api_task_status(task_id):
    """查询后台任务进度"""
    task = get_task_info(task_id)
    if not task:
        return jsonify({'error': '任务不存在'}), 404
    return jsonify(task)


@app.route('/match')
@login_required
def match_page():
    """公司名匹配页面"""
    conn = get_db()
    stats = {
        'total_pass': conn.execute("SELECT COUNT(*) FROM customers WHERE verify_result='pass'").fetchone()[0],
        'total_all': conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0],
    }
    conn.close()
    return render_template('match.html', stats=stats)


@app.route('/library/match', methods=['POST'])
@login_required
def library_match():
    """根据公司名称匹配资料库中核验成功的企业，导出完整信息"""
    import pandas as pd

    # 获取公司名称
    company_names_text = request.form.get('company_names', '').strip()
    name_file = request.files.get('name_file')

    names = []
    if company_names_text:
        names = [n.strip() for n in company_names_text.split('\n') if n.strip()]

    if name_file and name_file.filename:
        upload_path = os.path.join(UPLOAD_DIR, f'match_names_{datetime.now().strftime("%Y%m%d%H%M%S")}_{name_file.filename}')
        name_file.save(upload_path)
        ext = os.path.splitext(name_file.filename)[1].lower()
        if ext == '.txt':
            with open(upload_path, 'r', encoding='utf-8') as f:
                names.extend([n.strip() for n in f.readlines() if n.strip()])
        elif ext in ('.xlsx', '.xls'):
            try:
                df_probe = pd.read_excel(upload_path, dtype=str, header=None, nrows=5)
                header_row = None
                company_col = None
                company_col_names = ['公司名称', '企业名称', '名称', '公司', '企业', '单位名称', '机构名称', '商户名称']
                for row_idx in range(min(5, len(df_probe))):
                    for col_idx in range(df_probe.shape[1]):
                        val = str(df_probe.iloc[row_idx, col_idx]).strip()
                        if val in company_col_names:
                            header_row = row_idx
                            company_col = col_idx
                            break
                    if header_row is not None:
                        break
                if header_row is not None:
                    df_names = pd.read_excel(upload_path, dtype=str, header=header_row)
                    col_name = df_names.columns[company_col]
                    for val in df_names[col_name].dropna():
                        v = str(val).strip()
                        if v and v.lower() != 'nan':
                            names.append(v)
                else:
                    df_names = pd.read_excel(upload_path, dtype=str, header=None)
                    for col in df_names.columns:
                        for val in df_names[col].dropna():
                            v = str(val).strip()
                            if v and v.lower() != 'nan':
                                names.append(v)
            except Exception:
                try:
                    df_names = pd.read_excel(upload_path, dtype=str)
                    for col in df_names.columns:
                        if '公司' in str(col) or '名称' in str(col) or '企业' in str(col):
                            for val in df_names[col].dropna():
                                v = str(val).strip()
                                if v and v.lower() != 'nan':
                                    names.append(v)
                except Exception:
                    pass

    # 去重
    names = list(dict.fromkeys(names))

    if not names:
        flash('请输入公司名称或上传文件', 'error')
        return redirect(url_for('library'))

    conn = None
    try:
        conn = get_db()

        import unicodedata
        def normalize_name(s):
            """全角转半角，统一括号等字符"""
            if not s:
                return ''
            s = str(s).strip()
            try:
                s = unicodedata.normalize('NFKC', s)
            except Exception:
                pass
            return s

        # 归一化后的输入名 -> 原始名映射
        normalized_names = {}
        for name in names:
            nn = normalize_name(name)
            if nn and nn not in normalized_names:
                normalized_names[nn] = name

        # 优化：构建按首字符的倒排索引，加速子串匹配
        from collections import defaultdict
        name_by_first_char = defaultdict(list)
        for nn in normalized_names:
            if nn:
                name_by_first_char[nn[0]].append(nn)

        logger.info(f"公司匹配: 输入{len(names)}个公司名, 归一化后{len(normalized_names)}个, 开始匹配")

        BATCH_LOAD = 5000
        all_records = []
        matched_company_names = set()
        preview_records = []
        last_id = 0

        while True:
            rows = conn.execute(
                "SELECT * FROM customers WHERE verify_result='pass' AND id > ? ORDER BY id LIMIT ?",
                (last_id, BATCH_LOAD)
            ).fetchall()
            if not rows:
                break
            for r in rows:
                last_id = r['id']
                company = r['company_name'] or ''
                if not company:
                    continue
                company_norm = normalize_name(company)
                if not company_norm:
                    continue
                company_chars = set(company_norm)
                matched = False
                for char in company_chars:
                    if char in name_by_first_char:
                        for nn in name_by_first_char[char]:
                            if nn in company_norm:
                                matched = True
                                break
                    if matched:
                        break
                if matched:
                    all_records.append(r)
                    matched_company_names.add(company)
                    if len(preview_records) < 100:
                        preview_records.append({
                            'company_name': company,
                            'legal_person': r['legal_person'] or '',
                            'legal_phone': r['legal_phone'] or '',
                            'industry': r['industry'] or '',
                            'province': r['province'] or '',
                            'city': r['city'] or '',
                            'insured_count': r['insured_count'] or '',
                        })
            del rows
            gc.collect()

        logger.info(f"公司匹配: 匹配到{len(all_records)}条记录")

        # 用openpyxl流式写入，避免内存溢出
        from openpyxl import Workbook
        export_path = os.path.join(EXPORT_DIR, f'公司名匹配_核验成功_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')
        wb = Workbook(write_only=True)
        ws1 = wb.create_sheet('匹配结果')

        headers_match = ['公司名称', '法人姓名', '法人电话', '注册资本', '成立日期',
                         '经营状态', '所属行业', '注册地址', '经营范围', '统一社会信用代码',
                         '所属省份', '所属城市', '所属区县', '参保人数', '官网',
                         '来源', '批次', '核验结果', '资料人员', '备注']
        ws1.append(headers_match)

        match_count = 0
        for r in all_records:
            ws1.append([
                r['company_name'] or '',
                r['legal_person'] or '',
                r['legal_phone'] or '',
                r['registered_capital'] or '',
                r['established_date'] or '',
                r['business_status'] or '',
                r['industry'] or '',
                r['address'] or '',
                r['business_scope'] or '',
                r['credit_code'] or '',
                r['province'] or '',
                r['city'] or '',
                r['district'] or '',
                r['insured_count'] or '',
                r['website'] or '',
                r['source'] or '',
                r['batch_name'] or '',
                '通过',
                '陈平安',
                r['remark'] or '',
            ])
            match_count += 1

        # 统计未匹配的公司名 - 优化：使用倒排索引避免O(N*M)遍历
        matched_norm_names = set()
        for mname in matched_company_names:
            mn = normalize_name(mname)
            if mn:
                matched_norm_names.add(mn)

        matched_by_first_char = defaultdict(list)
        for mn in matched_norm_names:
            if mn:
                matched_by_first_char[mn[0]].append(mn)

        unmatched = []
        for name in names:
            nn = normalize_name(name)
            if not nn:
                continue
            found = False
            nn_chars = set(nn)
            for char in nn_chars:
                if char in matched_by_first_char:
                    for mn in matched_by_first_char[char]:
                        if nn in mn or mn in nn:
                            found = True
                            break
                if found:
                    break
            if not found:
                unmatched.append(name)

        if unmatched:
            ws2 = wb.create_sheet('未匹配')
            ws2.append(['未匹配的公司名称'])
            for name in unmatched:
                ws2.append([name])

        wb.save(export_path)
        del wb, all_records
        gc.collect()

        log_action(session['user_id'], 'library_match', f'输入{len(names)}个公司名',
                   f'匹配{match_count}条, 未匹配{len(unmatched)}条')

        # 返回JSON结果，前端展示后手动下载
        import os as _os
        download_filename = _os.path.basename(export_path)
        result_data = {
            'success': True,
            'total_input': len(names),
            'matched_count': match_count,
            'unmatched_count': len(unmatched),
            'unmatched_names': unmatched[:200],
            'download_url': f'/match/download/{download_filename}',
        }

        # 预览记录已在匹配时收集，无需再次查库

        result_data['preview_records'] = preview_records

        # 存储导出文件路径供下载接口使用
        session['match_export_path'] = export_path

        from flask import jsonify
        return jsonify(result_data)
    except Exception as e:
        logger.error(f"公司名匹配失败: {e}\n{traceback.format_exc()}")
        flash(f'匹配失败: {str(e)}', 'error')
        return redirect(url_for('match_page'))
    finally:
        if conn:
            conn.close()


@app.route('/match/download/<path:filename>')
@login_required
def match_download(filename):
    """下载匹配结果文件"""
    export_path = session.get('match_export_path', '')
    if export_path and os.path.exists(export_path):
        return send_file(export_path, as_attachment=True)
    # 尝试从文件名查找
    full_path = os.path.join(EXPORT_DIR, filename)
    if os.path.exists(full_path):
        return send_file(full_path, as_attachment=True)
    flash('文件不存在或已过期，请重新匹配', 'error')
    return redirect(url_for('match_page'))


@app.route('/library/export')
@login_required
def library_export():
    verify = request.args.get('verify', 'all')
    # 省份多选
    selected_provinces = request.args.getlist('provinces')
    if not selected_provinces:
        single = request.args.get('province', '')
        if single:
            selected_provinces = [single]
    # 城市多选
    selected_cities = request.args.getlist('cities')
    if not selected_cities:
        single_city = request.args.get('city', '')
        if single_city:
            selected_cities = [single_city]
    # 区县多选
    selected_districts = request.args.getlist('districts')
    if not selected_districts:
        single_district = request.args.get('district', '')
        if single_district:
            selected_districts = [single_district]
    source = request.args.get('source', '')
    batch = request.args.get('batch', '')
    category = request.args.get('category', '').strip()
    min_insured = request.args.get('min_insured', '').strip()
    max_insured = request.args.get('max_insured', '').strip()
    max_insured_dup = request.args.get('max_insured_dup', '') == '1'
    export_count = request.args.get('export_count', '').strip()
    export_start = request.args.get('export_start', '').strip()

    conn = None
    try:
        conn = get_db()
        conditions = []
        params = []

        if verify == 'pass':
            conditions.append("verify_result='pass'")
        elif verify == 'failed':
            conditions.append("verify_result='failed'")
        if selected_provinces:
            placeholders = ','.join(['?'] * len(selected_provinces))
            conditions.append(f"province IN ({placeholders})")
            params.extend(selected_provinces)
        if selected_cities:
            placeholders = ','.join(['?'] * len(selected_cities))
            conditions.append(f"city IN ({placeholders})")
            params.extend(selected_cities)
        if selected_districts:
            placeholders = ','.join(['?'] * len(selected_districts))
            conditions.append(f"district IN ({placeholders})")
            params.extend(selected_districts)
        if source:
            conditions.append("source LIKE ?")
            params.append(f'%{source}%')
        if batch:
            conditions.append("batch_name=?")
            params.append(batch)
        # 品类筛选
        if category:
            cat_keywords = [kw.strip() for kw in category.split(',') if kw.strip()]
            if cat_keywords:
                cat_conditions = []
                for kw in cat_keywords:
                    cat_conditions.append("(industry LIKE ? OR business_scope LIKE ?)")
                    params.extend([f'%{kw}%', f'%{kw}%'])
                conditions.append('(' + ' OR '.join(cat_conditions) + ')')
        # 参保人数筛选
        if min_insured:
            try:
                threshold = int(float(min_insured))
                conditions.append("parse_insured(insured_count) >= ?")
                params.append(threshold)
            except (ValueError, TypeError):
                pass
        if max_insured:
            try:
                threshold = int(float(max_insured))
                conditions.append("parse_insured(insured_count) <= ?")
                params.append(threshold)
            except (ValueError, TypeError):
                pass

        where_sql = ' AND '.join(conditions) if conditions else '1=1'

        # 先计算总数
        total_filtered = conn.execute(
            f"SELECT COUNT(*) FROM customers WHERE {where_sql}", params
        ).fetchone()[0]

        # SQLite语法: LIMIT count OFFSET offset（顺序不能反）
        sql_limit = ''
        sql_params = list(params)
        if export_count:
            try:
                limit = int(export_count)
                if limit > 0:
                    if export_start:
                        try:
                            start = int(export_start)
                            if start > 1:
                                sql_limit = f' LIMIT ? OFFSET ?'
                                sql_params.extend([limit, start - 1])
                            else:
                                sql_limit = f' LIMIT ?'
                                sql_params.append(limit)
                        except (ValueError, TypeError):
                            sql_limit = f' LIMIT ?'
                            sql_params.append(limit)
                    else:
                        sql_limit = f' LIMIT ?'
                        sql_params.append(limit)
            except (ValueError, TypeError):
                pass
        elif export_start:
            try:
                start = int(export_start)
                if start > 1:
                    sql_limit = f' LIMIT -1 OFFSET ?'
                    sql_params.append(start - 1)
            except (ValueError, TypeError):
                pass

        from openpyxl import Workbook
        cat_keywords = [kw.strip() for kw in category.split(',') if kw.strip()] if category else []

        # 如果需要同公司去重(max_insured_dup)，先获取要保留的ID集合
        keep_ids = None
        if max_insured_dup and total_filtered > 0:
            # 分批获取所有记录的ID和公司名和参保人数
            keep_ids = set()
            company_best = {}  # company -> (insured_num, id)
            offset_q = 0
            BATCH_Q = 5000
            while True:
                q_rows = conn.execute(
                    f"SELECT id, company_name, insured_count FROM customers WHERE {where_sql} ORDER BY id DESC LIMIT ? OFFSET ?",
                    params + [BATCH_Q, offset_q]
                ).fetchall()
                if not q_rows:
                    break
                for r in q_rows:
                    comp = str(r['company_name'] or '').strip()
                    if comp == '':
                        keep_ids.add(r['id'])
                    else:
                        num = _parse_insured_number(r['insured_count'])
                        if comp not in company_best or num > company_best[comp][0]:
                            if comp in company_best:
                                keep_ids.discard(company_best[comp][1])
                            company_best[comp] = (num, r['id'])
                            keep_ids.add(r['id'])
                offset_q += BATCH_Q
                del q_rows
            # 加入有公司名的最佳记录
            for _, (_, rid) in company_best.items():
                keep_ids.add(rid)
            del company_best
            gc.collect()
            logger.info(f"max_insured_dup: 保留 {len(keep_ids)} 条 (总 {total_filtered} 条)")

        # 分批查询 + openpyxl write_only流式写入
        export_path = os.path.join(EXPORT_DIR, f'资料库导出_{verify}_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')
        wb = Workbook(write_only=True)
        ws = wb.create_sheet('Sheet1')

        headers_export = ['公司名称', '法人姓名', '法人电话', '注册资本', '成立日期',
                           '经营状态', '所属行业', '注册地址', '经营范围', '统一社会信用代码',
                           '所属省份', '所属城市', '所属区县', '参保人数', '官网', '匹配品类',
                           '来源', '批次', '核验结果', '失败原因', '资料人员', '备注']
        ws.append(headers_export)

        # 分批读取并写入
        BATCH_SIZE = 2000
        offset = 0
        export_count_actual = 0
        while True:
            batch_params = list(params) + [BATCH_SIZE, offset]
            batch_rows = conn.execute(
                f"SELECT * FROM customers WHERE {where_sql} ORDER BY id DESC LIMIT ? OFFSET ?",
                batch_params
            ).fetchall()
            if not batch_rows:
                break
            for r in batch_rows:
                # 如果有去重，跳过不需要的记录
                if keep_ids is not None and r['id'] not in keep_ids:
                    continue
                # 计算匹配的品类
                matched_cats = []
                if cat_keywords:
                    industry = str(r['industry'] or '')
                    scope = str(r['business_scope'] or '')
                    combined = industry + scope
                    for kw in cat_keywords:
                        if kw in combined:
                            matched_cats.append(kw)
                ws.append([
                    r['company_name'] or '',
                    r['legal_person'] or '',
                    r['legal_phone'] or '',
                    r['registered_capital'] or '',
                    r['established_date'] or '',
                    r['business_status'] or '',
                    r['industry'] or '',
                    r['address'] or '',
                    r['business_scope'] or '',
                    r['credit_code'] or '',
                    r['province'] or '',
                    r['city'] or '',
                    r['district'] or '',
                    r['insured_count'] or '',
                    r['website'] or '',
                    '/'.join(matched_cats),
                    r['source'] or '',
                    r['batch_name'] or '',
                    '通过' if r['verify_result'] == 'pass' else '失败',
                    r['fail_reason'] or '',
                    '陈平安',
                    r['remark'] or '',
                ])
                export_count_actual += 1
            offset += BATCH_SIZE
            del batch_rows
            gc.collect()

        wb.save(export_path)
        del wb, ws, keep_ids
        gc.collect()

        logger.info(f"导出 {export_count_actual} 条 (总筛选 {total_filtered} 条)")

        logger.info(f"用户 {session.get('username')} 导出资料库: {export_count_actual} 条, 条件: verify={verify}, source={source}, batch={batch}, min_insured={min_insured}, max_insured={max_insured}")
        return send_file(export_path, as_attachment=True)
    except Exception as e:
        logger.error(f"资料库导出失败: {e}\n{traceback.format_exc()}")
        return render_template('error.html', message=f'导出失败: {str(e)}'), 500
    finally:
        if conn:
            conn.close()


# ---------- 管理员后台 ----------

@app.route('/admin')
@admin_required
def admin_dashboard():
    conn = get_db()

    stats = {
        'total_users': conn.execute("SELECT COUNT(*) FROM users").fetchone()[0],
        'total_batches': conn.execute("SELECT COUNT(*) FROM batches").fetchone()[0],
    }
    cust_stats = conn.execute('''SELECT 
        COUNT(*) as total,
        SUM(CASE WHEN status='public' THEN 1 ELSE 0 END) as public_count,
        SUM(CASE WHEN status='claimed' THEN 1 ELSE 0 END) as claimed_count
        FROM customers''').fetchone()
    stats['total_customers'] = cust_stats['total'] or 0
    stats['public_customers'] = cust_stats['public_count'] or 0
    stats['claimed_customers'] = cust_stats['claimed_count'] or 0

    # 用户统计
    user_stats = conn.execute(
        '''SELECT u.id, u.username, u.real_name, u.role, u.created_at, u.last_login,
           COUNT(c.id) as claimed_count
           FROM users u LEFT JOIN customers c ON u.id = c.owner_id
           GROUP BY u.id ORDER BY u.id DESC'''
    ).fetchall()

    # 最近操作日志
    recent_logs = conn.execute(
        '''SELECT l.*, u.username 
           FROM operation_logs l LEFT JOIN users u ON l.user_id=u.id
           ORDER BY l.id DESC LIMIT 20'''
    ).fetchall()

    # 批次统计
    batches = conn.execute(
        '''SELECT b.*, u.real_name as creator_name 
           FROM batches b LEFT JOIN users u ON b.created_by=u.id
           ORDER BY b.id DESC LIMIT 10'''
    ).fetchall()

    conn.close()

    return render_template('admin.html',
                           stats=stats, user_stats=user_stats,
                           recent_logs=recent_logs, batches=batches)


@app.route('/admin/users')
@admin_required
def admin_users():
    conn = get_db()
    users = conn.execute("SELECT * FROM users ORDER BY id DESC").fetchall()
    conn.close()
    return render_template('admin_users.html', users=users)


@app.route('/admin/users/add', methods=['POST'])
@admin_required
def add_user():
    username = request.form.get('username', '').strip()
    password = request.form.get('password', '').strip()
    real_name = request.form.get('real_name', '').strip()
    phone = request.form.get('phone', '').strip()
    role = request.form.get('role', 'user')

    if not username or not password:
        flash('用户名和密码不能为空', 'error')
        return redirect(url_for('admin_users'))

    conn = get_db()
    # 检查用户名是否已存在
    existing = conn.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone()
    if existing:
        conn.close()
        flash('用户名已存在', 'error')
        return redirect(url_for('admin_users'))

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn.execute(
        '''INSERT INTO users (username, password, real_name, phone, role, created_at, last_login)
           VALUES (?, ?, ?, ?, ?, ?, ?)''',
        (username, hash_password(password), real_name, phone, role, now, None)
    )
    conn.commit()
    conn.close()

    log_action(session['user_id'], 'add_user', username, f'角色: {role}')
    flash(f'用户 {username} 添加成功', 'success')
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/reset', methods=['POST'])
@admin_required
def reset_password(uid):
    new_pwd = request.form.get('password', '123456')
    conn = get_db()
    conn.execute("UPDATE users SET password=? WHERE id=?",
                 (hash_password(new_pwd), uid))
    conn.commit()
    conn.close()
    log_action(session['user_id'], 'reset_password', str(uid), '')
    flash('密码已重置', 'success')
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/role', methods=['POST'])
@admin_required
def change_role(uid):
    role = request.form.get('role', 'user')
    conn = get_db()
    conn.execute("UPDATE users SET role=? WHERE id=?", (role, uid))
    conn.commit()
    conn.close()
    log_action(session['user_id'], 'change_role', str(uid), role)
    flash('角色已更新', 'success')
    return redirect(url_for('admin_users'))


@app.route('/admin/users/<int:uid>/delete', methods=['POST'])
@admin_required
def delete_user(uid):
    if uid == session['user_id']:
        flash('不能删除自己', 'error')
        return redirect(url_for('admin_users'))

    conn = get_db()
    conn.execute("DELETE FROM users WHERE id=?", (uid,))
    conn.commit()
    conn.close()
    log_action(session['user_id'], 'delete_user', str(uid), '')
    flash('用户已删除', 'success')
    return redirect(url_for('admin_users'))


# ---------- 客户导出 ----------

@app.route('/export/customers')
@login_required
def export_customers():
    status = request.args.get('status', 'public')
    keyword = request.args.get('keyword', '')
    batch = request.args.get('batch', '')
    verify = request.args.get('verify', 'all')
    has_shareholder = request.args.get('has_shareholder', '')

    conn = get_db()

    conditions = []
    params = []

    if status == 'public':
        conditions.append("status='public'")
    elif status == 'my':
        conditions.append(f"owner_id={session['user_id']}")

    if keyword:
        conditions.append("(company_name LIKE ? OR legal_person LIKE ?)")
        params.extend([f'%{keyword}%'] * 2)

    if batch:
        conditions.append("batch_name=?")
        params.append(batch)

    if verify == 'pass':
        conditions.append("verify_result='pass'")
    elif verify == 'failed':
        conditions.append("verify_result='failed'")
    elif verify == 'unverified':
        conditions.append("(verify_result IS NULL OR verify_result='')")

    if has_shareholder == 'yes':
        conditions.append("(shareholder_info IS NOT NULL AND shareholder_info!='')")
    elif has_shareholder == 'no':
        conditions.append("(shareholder_info IS NULL OR shareholder_info='')")

    where_sql = ' AND '.join(conditions) if conditions else '1=1'

    # 分批查询防止内存溢出
    BATCH_SIZE = 2000
    offset = 0
    total_count = conn.execute(
        f"SELECT COUNT(*) FROM customers WHERE {where_sql}", params
    ).fetchone()[0]

    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

    export_path = os.path.join(EXPORT_DIR, f'客户导出_{status}_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx')
    wb = Workbook()
    ws = wb.active
    ws.title = '客户列表'

    headers = ['公司名称', '法人姓名', '法人电话', '注册资本', '成立日期',
               '经营状态', '所属行业', '注册地址', '经营范围', '统一社会信用代码',
               '所属省份', '所属城市', '所属区县', '参保人数', '来源', '批次',
               '核验状态', '失败原因', '资料人员', '备注']
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)

    row_idx = 2
    while True:
        batch_rows = conn.execute(
            f"SELECT * FROM customers WHERE {where_sql} ORDER BY id DESC LIMIT ? OFFSET ?",
            params + [BATCH_SIZE, offset]
        ).fetchall()
        if not batch_rows:
            break
        for c in batch_rows:
            ws.cell(row_idx, 1, c['company_name'] or '')
            ws.cell(row_idx, 2, c['legal_person'] or '')
            ws.cell(row_idx, 3, c['legal_phone'] or '')
            ws.cell(row_idx, 4, c['registered_capital'] or '')
            ws.cell(row_idx, 5, c['established_date'] or '')
            ws.cell(row_idx, 6, c['business_status'] or '')
            ws.cell(row_idx, 7, c['industry'] or '')
            ws.cell(row_idx, 8, c['address'] or '')
            ws.cell(row_idx, 9, c['business_scope'] or '')
            ws.cell(row_idx, 10, c['credit_code'] or '')
            ws.cell(row_idx, 11, c['province'] or '')
            ws.cell(row_idx, 12, c['city'] or '')
            ws.cell(row_idx, 13, c['district'] or '')
            ws.cell(row_idx, 14, c['insured_count'] or '')
            ws.cell(row_idx, 15, c['source'] or '')
            ws.cell(row_idx, 16, c['batch_name'] or '')
            verify_text = '通过' if c['verify_result'] == 'pass' else ('失败' if c['verify_result'] == 'failed' else '待核验')
            ws.cell(row_idx, 17, verify_text)
            ws.cell(row_idx, 18, c['fail_reason'] or '')
            ws.cell(row_idx, 19, '陈平安')
            ws.cell(row_idx, 20, c['remark'] or '')
            row_idx += 1
        offset += BATCH_SIZE
        del batch_rows
        gc.collect()

    conn.close()
    wb.save(export_path)
    del wb, ws
    gc.collect()

    logger.info(f"用户 {session.get('username')} 导出客户: {total_count} 条")
    return send_file(export_path, as_attachment=True)


# ---------- 去重模块 ----------

@app.route('/dedup')
@login_required
def dedup_page():
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]

    MAX_SHOW = 500  # 最多显示500组，防止内存溢出

    # 1. 按统一社会信用代码去重分析（credit_code不为空）
    dup_by_credit = conn.execute(
        f'''SELECT credit_code, company_name, COUNT(*) as cnt,
                  GROUP_CONCAT(id, ',') as ids,
                  GROUP_CONCAT(legal_phone, ',') as phones
           FROM customers
           WHERE credit_code != '' AND credit_code IS NOT NULL
           GROUP BY credit_code
           HAVING cnt > 1
           ORDER BY cnt DESC LIMIT {MAX_SHOW}'''
    ).fetchall()

    # 2. 按公司名称+法人电话去重分析（credit_code为空）
    dup_by_name_phone = conn.execute(
        f'''SELECT company_name, legal_phone, COUNT(*) as cnt,
                  GROUP_CONCAT(id, ',') as ids
           FROM customers
           WHERE (credit_code = '' OR credit_code IS NULL)
             AND company_name != '' AND legal_phone != ''
           GROUP BY company_name, legal_phone
           HAVING cnt > 1
           ORDER BY cnt DESC LIMIT {MAX_SHOW}'''
    ).fetchall()

    # 3. 仅按公司名称去重（跨表重复检查标准，电话不同时）
    dup_by_company = conn.execute(
        f'''SELECT company_name, COUNT(*) as cnt,
                  GROUP_CONCAT(id, ',') as ids,
                  COUNT(DISTINCT legal_phone) as phone_count
           FROM customers
           WHERE company_name != ''
           GROUP BY company_name
           HAVING cnt > 1 AND phone_count > 1
           ORDER BY cnt DESC LIMIT {MAX_SHOW}'''
    ).fetchall()

    # 4. 按法人电话去重（同一号码出现在多条记录中）
    dup_by_phone = conn.execute(
        f'''SELECT legal_phone, COUNT(*) as cnt,
                  GROUP_CONCAT(id, ',') as ids,
                  COUNT(DISTINCT company_name) as company_count,
                  GROUP_CONCAT(company_name, ' | ') as companies
           FROM customers
           WHERE legal_phone != '' AND legal_phone IS NOT NULL
           GROUP BY legal_phone
           HAVING cnt > 1
           ORDER BY cnt DESC LIMIT {MAX_SHOW}'''
    ).fetchall()

    # 统计总数（用SQL COUNT，不加载全部数据）
    credit_dup_records = conn.execute(
        "SELECT COALESCE(SUM(cnt), 0) FROM (SELECT COUNT(*) as cnt FROM customers WHERE credit_code != '' AND credit_code IS NOT NULL GROUP BY credit_code HAVING cnt > 1)"
    ).fetchone()[0]
    name_phone_dup_records = conn.execute(
        "SELECT COALESCE(SUM(cnt), 0) FROM (SELECT COUNT(*) as cnt FROM customers WHERE (credit_code = '' OR credit_code IS NULL) AND company_name != '' AND legal_phone != '' GROUP BY company_name, legal_phone HAVING cnt > 1)"
    ).fetchone()[0]
    company_dup_records = conn.execute(
        "SELECT COALESCE(SUM(cnt), 0) FROM (SELECT COUNT(*) as cnt FROM customers WHERE company_name != '' GROUP BY company_name HAVING cnt > 1 AND COUNT(DISTINCT legal_phone) > 1)"
    ).fetchone()[0]
    phone_dup_records = conn.execute(
        "SELECT COALESCE(SUM(cnt), 0) FROM (SELECT COUNT(*) as cnt FROM customers WHERE legal_phone != '' AND legal_phone IS NOT NULL GROUP BY legal_phone HAVING cnt > 1)"
    ).fetchone()[0]

    conn.close()
    gc.collect()

    return render_template('dedup.html',
                           total=total,
                           dup_by_credit=dup_by_credit,
                           dup_by_name_phone=dup_by_name_phone,
                           dup_by_company=dup_by_company,
                           dup_by_phone=dup_by_phone,
                           credit_dup_records=credit_dup_records,
                           name_phone_dup_records=name_phone_dup_records,
                           company_dup_records=company_dup_records,
                           phone_dup_records=phone_dup_records)


@app.route('/dedup/execute', methods=['POST'])
@login_required
def dedup_execute():
    strategy = request.form.get('strategy', 'keep_newest')
    scope = request.form.getlist('scope')  # credit, name_phone, company

    # 去重前自动备份
    backup_database('before_dedup')

    conn = get_db()
    deleted_count = 0
    kept_count = 0

    def resolve_keep_id(ids_str):
        ids = [int(x) for x in ids_str.split(',') if x.strip()]
        if not ids:
            return None
        if strategy == 'keep_newest':
            return max(ids)
        elif strategy == 'keep_oldest':
            return min(ids)
        elif strategy == 'keep_complete':
            # 批量查询所有记录，避免N+1
            placeholders = ','.join(['?'] * len(ids))
            rows = conn.execute(
                f"SELECT id, * FROM customers WHERE id IN ({placeholders})", ids
            ).fetchall()
            if not rows:
                return max(ids)
            best_id = rows[0]['id']
            best_score = -1
            for row in rows:
                score = sum(1 for k in row.keys() if row[k] not in (None, ''))
                if score > best_score:
                    best_score = score
                    best_id = row['id']
            return best_id
        return max(ids)

    def batch_delete(ids_list):
        """批量删除，避免逐条DELETE"""
        nonlocal deleted_count
        if not ids_list:
            return
        # SQLite参数限制999，分批处理
        batch_size = 900
        for i in range(0, len(ids_list), batch_size):
            batch = ids_list[i:i+batch_size]
            placeholders = ','.join(['?'] * len(batch))
            conn.execute(f"DELETE FROM customers WHERE id IN ({placeholders})", batch)
            deleted_count += len(batch)

    try:
        # 1. 按统一社会信用代码去重
        if 'credit' in scope:
            groups = conn.execute(
                '''SELECT credit_code, GROUP_CONCAT(id, ',') as ids, COUNT(*) as cnt
                   FROM customers
                   WHERE credit_code != '' AND credit_code IS NOT NULL
                   GROUP BY credit_code HAVING cnt > 1'''
            ).fetchall()
            for g in groups:
                keep_id = resolve_keep_id(g['ids'])
                if keep_id:
                    ids = [int(x) for x in g['ids'].split(',') if x.strip()]
                    to_delete = [x for x in ids if x != keep_id]
                    batch_delete(to_delete)
                    kept_count += 1

        # 2. 按公司名称+法人电话去重（credit_code为空）
        if 'name_phone' in scope:
            groups = conn.execute(
                '''SELECT company_name, legal_phone, GROUP_CONCAT(id, ',') as ids, COUNT(*) as cnt
                   FROM customers
                   WHERE (credit_code = '' OR credit_code IS NULL)
                     AND company_name != '' AND legal_phone != ''
                   GROUP BY company_name, legal_phone HAVING cnt > 1'''
            ).fetchall()
            for g in groups:
                keep_id = resolve_keep_id(g['ids'])
                if keep_id:
                    ids = [int(x) for x in g['ids'].split(',') if x.strip()]
                    to_delete = [x for x in ids if x != keep_id]
                    batch_delete(to_delete)
                    kept_count += 1

        # 3. 仅按公司名称去重（跨表重复，保留电话最全的）
        if 'company' in scope:
            groups = conn.execute(
                '''SELECT company_name, GROUP_CONCAT(id, ',') as ids, COUNT(*) as cnt,
                          COUNT(DISTINCT legal_phone) as phone_count
                   FROM customers
                   WHERE company_name != ''
                   GROUP BY company_name HAVING cnt > 1 AND phone_count > 1'''
            ).fetchall()
            for g in groups:
                keep_id = resolve_keep_id(g['ids'])
                if keep_id:
                    ids = [int(x) for x in g['ids'].split(',') if x.strip()]
                    to_delete = [x for x in ids if x != keep_id]
                    batch_delete(to_delete)
                    kept_count += 1

        # 4. 按法人电话去重（同一号码出现在不同公司中）
        if 'phone' in scope:
            groups = conn.execute(
                '''SELECT legal_phone, GROUP_CONCAT(id, ',') as ids, COUNT(*) as cnt
                   FROM customers
                   WHERE legal_phone != '' AND legal_phone IS NOT NULL
                   GROUP BY legal_phone HAVING cnt > 1'''
            ).fetchall()
            for g in groups:
                keep_id = resolve_keep_id(g['ids'])
                if keep_id:
                    ids = [int(x) for x in g['ids'].split(',') if x.strip()]
                    to_delete = [x for x in ids if x != keep_id]
                    batch_delete(to_delete)
                    kept_count += 1

        conn.commit()
        conn.close()
        cache_invalidate()  # 清除缓存

        log_action(session['user_id'], 'dedup', '',
                   f'删除{deleted_count}条, 保留{kept_count}组, 策略:{strategy}, 范围:{",".join(scope)}')
        flash(f'去重完成！删除 {deleted_count} 条重复记录，保留 {kept_count} 条唯一记录', 'success')

    except Exception as e:
        conn.rollback()
        conn.close()
        flash(f'去重失败: {str(e)}', 'error')
        import traceback
        traceback.print_exc()

    return redirect(url_for('dedup_page'))


# ========== 启动 ==========

if __name__ == '__main__':
    import ssl as _ssl
    import threading
    import time as _time

    init_db()
    app._start_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    cert_path = os.path.join(APP_DIR, 'cert.pem')
    key_path = os.path.join(APP_DIR, 'key.pem')

    ctx = _ssl.SSLContext(_ssl.PROTOCOL_TLS_SERVER)
    ctx.load_cert_chain(certfile=cert_path, keyfile=key_path)

    # HTTP→HTTPS 重定向线程
    def run_http_redirect():
        from http.server import HTTPServer, BaseHTTPRequestHandler

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self): self._do()
            def do_POST(self): self._do()
            def do_HEAD(self): self._do()
            def do_PUT(self): self._do()
            def do_DELETE(self): self._do()
            def _do(self):
                host = self.headers.get('Host', '192.168.1.107').split(':')[0]
                self.send_response(301)
                self.send_header('Location', f'https://{host}{self.path}')
                self.send_header('Connection', 'close')
                self.end_headers()
            def log_message(self, *a): pass

        try:
            HTTPServer(('0.0.0.0', 80), Handler).serve_forever()
        except Exception as e:
            print(f" * HTTP redirect failed: {e}")

    t = threading.Thread(target=run_http_redirect, daemon=True)
    t.start()
    _time.sleep(2)  # 等待80端口绑定

    print("\n" + "=" * 50)
    print("  陈平安的资料库 (HTTPS 加密)")
    print("  访问地址: https://chenpinganyyds")
    print("  也可通过 https://localhost 访问")
    print("  局域网: https://192.168.1.107")
    print("  HTTP(80) 自动跳转 HTTPS(443)")
    print("=" * 50 + "\n")

    app.run(host='0.0.0.0', port=443, debug=False, threaded=True, ssl_context=ctx)
