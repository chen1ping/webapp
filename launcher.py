# -*- coding: utf-8 -*-
"""
陈平安资料库 - 桌面应用程序
系统托盘 + Flask服务器 + 内网穿透远程访问
"""

import os
import sys
import threading
import subprocess
import socket
import time
import webbrowser
import json
import urllib.request
import re

# 设置工作目录
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
    BASE_DIR = os.path.join(os.path.dirname(sys.executable), '_internal')
    if not os.path.exists(BASE_DIR):
        BASE_DIR = APP_DIR
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR = APP_DIR

os.chdir(BASE_DIR)
sys.path.insert(0, BASE_DIR)

# 目录
DB_PATH = os.path.join(BASE_DIR, 'database.db')
RESULT_DIR = os.path.join(BASE_DIR, 'results')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
os.makedirs(RESULT_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)


def get_local_ip():
    """获取本机局域网IP"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "127.0.0.1"


def check_port(port):
    """检查端口是否可用"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(('0.0.0.0', port))
        s.close()
        return True
    except:
        return False


def find_available_port():
    """找到可用端口"""
    for port in [80, 8080, 8000, 8888, 5000]:
        if check_port(port):
            return port
    return 5000


def download_cloudflared():
    """下载cloudflared内网穿透工具"""
    exe_path = os.path.join(APP_DIR, 'cloudflared.exe')
    if os.path.exists(exe_path):
        return exe_path
    try:
        url = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"
        print("正在下载内网穿透工具（仅需一次）...")
        urllib.request.urlretrieve(url, exe_path)
        return exe_path
    except Exception as e:
        print(f"下载失败: {e}")
        return None


def start_tunnel(port):
    """启动cloudflare内网穿透隧道，返回公网URL"""
    exe_path = download_cloudflared()
    if not exe_path:
        return None, None

    try:
        proc = subprocess.Popen(
            [exe_path, 'tunnel', '--url', f'http://localhost:{port}', '--no-autoupdate'],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        tunnel_url = None
        start_time = time.time()
        while time.time() - start_time < 40:
            line = proc.stdout.readline().decode('utf-8', errors='ignore')
            if line:
                print(f"[隧道] {line.strip()}")
            if 'trycloudflare.com' in line:
                match = re.search(r'https://[a-z0-9-]+\.trycloudflare\.com', line)
                if match:
                    tunnel_url = match.group(0)
                    break
            if proc.poll() is not None:
                break

        return tunnel_url, proc
    except Exception as e:
        print(f"隧道启动失败: {e}")
        return None, None


class ServerApp:
    """服务器应用主类"""

    def __init__(self):
        self.port = find_available_port()
        self.local_ip = get_local_ip()
        self.tunnel_url = None
        self.tunnel_proc = None
        self.server_thread = None
        self.running = True

    def start_server(self):
        """启动Flask服务器"""
        try:
            import app as flask_app
            flask_app.app.run(host='0.0.0.0', port=self.port, debug=False, threaded=True)
        except Exception as e:
            print(f"服务器错误: {e}")

    def start(self):
        """启动所有服务"""
        print("=" * 55)
        print("  陈平安资料库 - 正在启动...")
        print("=" * 55)

        # 1. 启动Flask服务器
        print("\n[1/3] 启动服务器...")
        self.server_thread = threading.Thread(target=self.start_server, daemon=True)
        self.server_thread.start()
        time.sleep(3)
        print(f"  服务器已启动 (端口 {self.port})")

        # 2. 启动内网穿透
        print("\n[2/3] 启动远程访问隧道...")
        self.tunnel_url, self.tunnel_proc = start_tunnel(self.port)
        if self.tunnel_url:
            print(f"  远程访问已开启")
        else:
            print(f"  远程隧道未启动（仅局域网可用）")

        # 3. 保存配置并打开浏览器
        print("\n[3/3] 准备就绪")
        self.save_config()

        # 显示信息
        self.show_info()

        # 自动打开浏览器
        webbrowser.open(f'http://127.0.0.1:{self.port}')

        # 保持运行
        print("\n服务器运行中，按 Ctrl+C 停止...")
        try:
            while self.running:
                time.sleep(1)
        except KeyboardInterrupt:
            self.cleanup()

    def show_info(self):
        """显示访问信息"""
        print("\n" + "=" * 55)
        print("  ★ 陈平安资料库 已启动 ★")
        print("=" * 55)
        print(f"\n  【电脑访问】")
        print(f"    http://127.0.0.1:{self.port}")
        print(f"    http://{self.local_ip}:{self.port}")
        print(f"\n  【手机访问】")
        if self.tunnel_url:
            print(f"    ★ 远程访问（任何网络）:")
            print(f"    {self.tunnel_url}")
            print(f"\n    → 手机浏览器打开上方链接")
            print(f"    → 可添加到桌面当APP使用")
        else:
            print(f"    局域网访问（同一WiFi）:")
            print(f"    http://{self.local_ip}:{self.port}")
        print(f"\n  【端口】{self.port}")
        print("=" * 55)

        # 生成手机扫码链接页面
        self.create_access_page()

    def create_access_page(self):
        """生成扫码访问页面"""
        html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>陈平安资料库 - 访问入口</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:"KaiTi","STKaiti","楷体",serif; background:#f7f4ee; color:#1a1a1a; display:flex; justify-content:center; align-items:center; min-height:100vh; }}
.box {{ background:#fdfcfa; border:1px solid #d4d0c8; border-top:3px solid #8b0000; padding:40px; text-align:center; width:90%; max-width:420px; box-shadow:0 4px 16px rgba(0,0,0,0.1); }}
h1 {{ font-size:28px; letter-spacing:6px; margin-bottom:8px; font-weight:400; }}
.sub {{ color:#777; font-size:13px; letter-spacing:3px; margin-bottom:24px; }}
.link-box {{ background:#f0ebe4; border:1px solid #d4d0c8; padding:16px; margin:12px 0; border-radius:2px; }}
.link-box .label {{ font-size:11px; color:#999; letter-spacing:2px; margin-bottom:6px; }}
.link-box a {{ display:block; font-size:14px; color:#8b0000; text-decoration:none; word-break:break-all; padding:8px; background:#fff; border:1px solid #e8e5de; }}
.link-box a:hover {{ background:#faf4f4; }}
.remote {{ border-color:#8b0000; background:#faf4f4; }}
.remote a {{ color:#8b0000; font-weight:bold; }}
.tip {{ font-size:12px; color:#999; margin-top:20px; line-height:1.8; }}
.seal {{ display:inline-block; border:2px solid #8b0000; color:#8b0000; padding:8px 12px; font-size:14px; transform:rotate(-5deg); opacity:0.7; margin-bottom:16px; }}
</style>
</head>
<body>
<div class="box">
<div class="seal">平安<br>之印</div>
<h1>陈平安资料库</h1>
<div class="sub">访问入口</div>
{"<div class='link-box remote'><div class='label'>★ 手机远程访问（任何网络）</div><a href='" + self.tunnel_url + "'>" + self.tunnel_url + "</a></div>" if self.tunnel_url else ""}
<div class="link-box"><div class="label">电脑本机访问</div><a href="http://127.0.0.1:{self.port}">http://127.0.0.1:{self.port}</a></div>
<div class="link-box"><div class="label">局域网访问（同WiFi）</div><a href="http://{self.local_ip}:{self.port}">http://{self.local_ip}:{self.port}</a></div>
<div class="tip">
提示：手机用浏览器打开远程访问链接<br>
然后「添加到主屏幕」即可当APP使用<br>
<br>
端口：{self.port} ｜ 启动时间：{time.strftime('%Y-%m-%d %H:%M')}
</div>
</div>
</body>
</html>'''
        path = os.path.join(APP_DIR, 'access.html')
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)

    def save_config(self):
        """保存配置"""
        config = {
            'local_ip': self.local_ip,
            'port': self.port,
            'tunnel_url': self.tunnel_url,
            'local_url': f'http://127.0.0.1:{self.port}',
            'lan_url': f'http://{self.local_ip}:{self.port}',
        }
        config_path = os.path.join(APP_DIR, 'app_config.json')
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def cleanup(self):
        """清理"""
        self.running = False
        if self.tunnel_proc:
            self.tunnel_proc.terminate()
            print("远程隧道已关闭")
        print("服务器已停止，再见！")


def main():
    app = ServerApp()
    app.start()


if __name__ == '__main__':
    main()
