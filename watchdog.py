"""独立看门狗 - 后台运行，子进程方式监控服务器"""
import subprocess
import time
import os
import sys
import socket
import threading
import ctypes
import ctypes.wintypes as wintypes

PYTHON = sys.executable
APP = r"C:\WebApp\app.py"
LOG = r"C:\WebApp\server_log.txt"
PID_FILE = r"C:\WebApp\watchdog.pid"

kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
psapi = ctypes.WinDLL('psapi', use_last_error=True)
PROCESS_TERMINATE = 0x0001
PROCESS_QUERY_INFORMATION = 0x0400
MAX_MEMORY_MB = 1500  # 内存超过1500MB自动重启（后台任务需要更多内存）

class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("PageFaultCount", wintypes.DWORD),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
    ]

def get_process_memory_mb(pid):
    """获取进程内存占用（MB）"""
    try:
        h = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
        if not h:
            return 0
        pmc = PROCESS_MEMORY_COUNTERS()
        pmc.cb = ctypes.sizeof(pmc)
        if psapi.GetProcessMemoryInfo(h, ctypes.byref(pmc), ctypes.sizeof(pmc)):
            mem_mb = pmc.WorkingSetSize / (1024 * 1024)
            kernel32.CloseHandle(h)
            return mem_mb
        kernel32.CloseHandle(h)
    except:
        pass
    return 0

def kill_process_by_pid(pid):
    """用SeDebugPrivilege杀进程"""
    h = kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
    if h:
        result = kernel32.TerminateProcess(h, 1)
        kernel32.CloseHandle(h)
        return result
    # 回退到taskkill
    os.system(f'taskkill /F /PID {pid} >nul 2>&1')
    return True

def check_port(port):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(3)
        result = s.connect_ex(("127.0.0.1", port))
        s.close()
        return result == 0
    except:
        return False

def kill_port(port):
    try:
        lines = os.popen(f'netstat -ano | findstr ":{port} " | findstr LISTENING').read()
        for line in lines.strip().split('\n'):
            parts = line.split()
            if len(parts) >= 5 and parts[-1].isdigit():
                pid = parts[-1]
                kill_process_by_pid(int(pid))
    except:
        pass

def check_http_health():
    """检查HTTP是否能正常响应（不仅仅是端口开放）"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(5)
        s.connect(("127.0.0.1", 80))
        s.send(b"GET / HTTP/1.0\r\nHost: localhost\r\n\r\n")
        data = s.recv(1024)
        s.close()
        return b"200" in data or b"301" in data or b"302" in data
    except:
        return False

def main():
    with open(PID_FILE, "w") as f:
        f.write(str(os.getpid()))

    restart_count = 0
    server_pid = None
    last_health_ok = time.time()

    while True:
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n[{ts}] Watchdog: starting (attempt {restart_count + 1})", flush=True)

        if check_port(443):
            print("[Watchdog] port 443 busy, killing...", flush=True)
            kill_port(443)
            time.sleep(2)
        if check_port(80):
            print("[Watchdog] port 80 busy, killing...", flush=True)
            kill_port(80)
            time.sleep(2)

        try:
            log_f = open(LOG, "w", encoding="utf-8", errors="replace")
        except PermissionError:
            log_f = open(LOG, "w")

        proc = subprocess.Popen(
            [PYTHON, APP],
            stdout=log_f,
            stderr=subprocess.STDOUT,
            cwd=r"C:\WebApp",
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | 0x00000008
        )

        server_pid = proc.pid
        print(f"[Watchdog] server PID={server_pid}", flush=True)
        last_health_ok = time.time()

        # 等待进程结束 + 心跳检测
        while True:
            ret = proc.poll()
            if ret is not None:
                ts2 = time.strftime("%Y-%m-%d %H:%M:%S")
                print(f"[{ts2}] Watchdog: exited code={ret}", flush=True)
                break

            # 每30秒检查健康状态
            time.sleep(30)

            # 内存检查：超过阈值则重启
            mem_mb = get_process_memory_mb(server_pid)
            if mem_mb > MAX_MEMORY_MB:
                print(f"[Watchdog] Memory {int(mem_mb)}MB exceeds {MAX_MEMORY_MB}MB, restarting...", flush=True)
                kill_process_by_pid(server_pid)
                break

            if check_port(443):
                if check_http_health():
                    last_health_ok = time.time()
                else:
                    elapsed = time.time() - last_health_ok
                    if elapsed > 300:  # 5分钟无响应才重启（后台任务不阻塞HTTP）
                        print(f"[Watchdog] HTTP unresponsive for {int(elapsed)}s, killing PID={server_pid}", flush=True)
                        kill_process_by_pid(server_pid)
                        break
            else:
                # 端口未开放，可能是正在启动
                if time.time() - last_health_ok > 120:  # 2分钟启动时间
                    print(f"[Watchdog] port 443 down for 120s, killing PID={server_pid}", flush=True)
                    kill_process_by_pid(server_pid)
                    break

        try:
            log_f.close()
        except:
            pass

        restart_count += 1
        print("[Watchdog] restart in 3s...", flush=True)
        time.sleep(3)

if __name__ == "__main__":
    main()
