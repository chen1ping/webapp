"""
Gunicorn 配置文件（生产环境用）
用法: gunicorn -c gunicorn_config.py app:app
"""

import os
import multiprocessing

# 监听地址
bind = "127.0.0.1:8000"

# worker 数量：CPU核数 * 2 + 1
workers = int(os.environ.get('GUNICORN_WORKERS', multiprocessing.cpu_count() * 2 + 1))

# worker 类型（gevent支持异步，性能更好）
worker_class = "gevent"

# 每个 worker 的最大并发连接数
worker_connections = 1000

# 超时时间（秒）
timeout = 120

# 优雅重启超时
graceful_timeout = 30

# 保持连接
keepalive = 5

# 日志
accesslog = "logs/access.log"
errorlog = "logs/error.log"
loglevel = "info"

# 进程名称
proc_name = "webapp"

# 后台运行
daemon = False

# PID 文件
pidfile = "logs/gunicorn.pid"

# 预加载应用（减少内存占用）
preload_app = True

# worker 最大请求数（防止内存泄漏，超过后重启）
max_requests = 1000
max_requests_jitter = 100
