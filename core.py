# -*- coding: utf-8 -*-

import os
import sys
import json
import time
import atexit
import socket
import ctypes
import subprocess
from ctypes import wintypes
from datetime import datetime
from urllib import request as urlrequest

from constants import (
    APP_NAME,
    STARTUP_SHORTCUT_NAME,
    RUN_KEY_PATH,
    RUN_VALUE_NAME,
    ERROR_ALREADY_EXISTS,
)
from winapi import kernel32, user32, shell32, wininet, HANDLE_T, HWND_T

# -----------------------------
# Helpers
# -----------------------------
def _w(s: str) -> str:
    return s if isinstance(s, str) else str(s)

def ensure_dirs(path: str):
    os.makedirs(path, exist_ok=True)

def get_appdata_dir() -> str:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(appdata, APP_NAME)

APPDATA_DIR = get_appdata_dir()
CONFIG_PATH = os.path.join(APPDATA_DIR, "config.json")
ERROR_LOG = os.path.join(APPDATA_DIR, "error.log")

def log_error(e: Exception):
    try:
        ensure_dirs(APPDATA_DIR)
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat(timespec='seconds')} {repr(e)}\n")
    except Exception:
        pass

def get_entry_script_path() -> str:
    if getattr(sys, "frozen", False):
        return os.path.abspath(sys.executable)

    path = os.path.abspath(sys.argv[0])
    if os.path.isfile(path):
        return path

    main_mod = sys.modules.get("__main__")
    if main_mod and getattr(main_mod, "__file__", None):
        return os.path.abspath(main_mod.__file__)

    base = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(base, "auto_shutdown_win32_native.py")
    if os.path.exists(candidate):
        return candidate

    return os.path.abspath(__file__)

def resource_path(filename: str) -> str:
    try:
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(get_entry_script_path())
    except Exception:
        base = os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.join(base, filename)

# -----------------------------
# Config
# -----------------------------
DEFAULT_REMIND_TEMPLATE = "{base_info}\n\n建议：确认是否需要关机/休眠/合盖。"

DEFAULT_CONFIG = {
    "pushplus_token": "",
    "pushplus_topic": "",
    "pushplus_api": "https://www.pushplus.plus/send",
    "remind_template": DEFAULT_REMIND_TEMPLATE,
    "online_remind_times": 0,  # 0=仅提醒；N=提醒N次后休眠

    "uptime_hours": 2,
    "idle_minutes": 60,
    "pre_hibernate_countdown_sec": 60,
    "resume_grace_sec": 120,
    "tray_balloon_enabled": False,
    "last_hibernate_time": "",
    "last_hibernate_notice_time": "",

    "net_check_url": "https://baidu.com",
    "net_check_timeout_sec": 2,

    "autostart_enabled": False,
}

def migrate_config_if_needed():
    try:
        ensure_dirs(APPDATA_DIR)
        if os.path.exists(CONFIG_PATH):
            return
        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        old_path = os.path.join(base_dir, "config.json")
        if os.path.exists(old_path):
            with open(old_path, "r", encoding="utf-8") as f:
                data = f.read()
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                f.write(data)
            try:
                os.replace(old_path, old_path + ".migrated.bak")
            except Exception:
                pass
    except Exception as e:
        log_error(e)

def load_config() -> dict:
    migrate_config_if_needed()
    ensure_dirs(APPDATA_DIR)
    cfg = dict(DEFAULT_CONFIG)
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                user_cfg = json.load(f)
            if isinstance(user_cfg, dict):
                cfg.update(user_cfg)
                if "online_remind_times" not in user_cfg and "online_hibernate_policy" in user_cfg:
                    try:
                        cfg["online_remind_times"] = int(user_cfg.get("online_hibernate_policy", 0))
                    except Exception:
                        pass
                if "online_hibernate_policy" in cfg:
                    cfg.pop("online_hibernate_policy", None)
    except Exception as e:
        log_error(e)
    return cfg

def save_config(cfg: dict):
    try:
        ensure_dirs(APPDATA_DIR)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_error(e)

# -----------------------------
# Single instance (Mutex)
# -----------------------------
_single_instance_handle = None

def ensure_single_instance(mutex_name: str = r"Local\AutoShutdownReminder_SingleInstance") -> bool:
    global _single_instance_handle
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = HANDLE_T
    kernel32.CloseHandle.argtypes = [HANDLE_T]
    kernel32.CloseHandle.restype = wintypes.BOOL

    h = kernel32.CreateMutexW(None, True, _w(mutex_name))
    if not h:
        return True
    last_err = ctypes.get_last_error()
    if last_err == ERROR_ALREADY_EXISTS:
        kernel32.CloseHandle(h)
        return False
    _single_instance_handle = h

    def _release():
        try:
            if _single_instance_handle:
                kernel32.CloseHandle(_single_instance_handle)
        except Exception:
            pass
    atexit.register(_release)
    return True

# -----------------------------
# Idle/Uptime (WinAPI)
# -----------------------------
class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.UINT), ("dwTime", wintypes.DWORD)]

user32.GetLastInputInfo.argtypes = [ctypes.POINTER(LASTINPUTINFO)]
user32.GetLastInputInfo.restype = wintypes.BOOL
kernel32.GetTickCount64.argtypes = []
kernel32.GetTickCount64.restype = ctypes.c_ulonglong
kernel32.GetTickCount.argtypes = []
kernel32.GetTickCount.restype = wintypes.DWORD

def get_idle_seconds() -> int:
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if not user32.GetLastInputInfo(ctypes.byref(lii)):
        return 0
    tick = kernel32.GetTickCount()
    idle_ms = int(tick) - int(lii.dwTime)
    if idle_ms < 0:
        idle_ms = 0
    return idle_ms // 1000

def get_uptime_seconds() -> int:
    return int(kernel32.GetTickCount64() // 1000)

# -----------------------------
# Network checks
# -----------------------------
def is_online_winapi() -> bool:
    try:
        flags = wintypes.DWORD(0)
        return bool(wininet.InternetGetConnectedState(ctypes.byref(flags), 0))
    except Exception:
        return False

def _parse_host(url: str) -> str:
    u = (url or "").strip()
    if "://" in u:
        u = u.split("://", 1)[1]
    return u.split("/", 1)[0].strip()

def is_online_two_level(cfg: dict) -> bool:
    # 1) WinAPI
    if not is_online_winapi():
        return False

    # 2) DNS + lightweight HTTP GET
    url = str(cfg.get("net_check_url", "https://baidu.com")).strip()
    timeout = int(cfg.get("net_check_timeout_sec", 2))
    try:
        host = _parse_host(url)
        if not host:
            return False
        socket.gethostbyname(host)
    except Exception:
        return False
    try:
        req = urlrequest.Request(url, method="GET", headers={"User-Agent": "AutoShutdown/1.0"})
        with urlrequest.urlopen(req, timeout=timeout) as resp:
            code = getattr(resp, "status", None) or resp.getcode()
        return int(code) // 100 in (2, 3)
    except Exception:
        return False

# -----------------------------
# Pushplus
# -----------------------------
def pushplus_send(cfg: dict, title: str, content: str):
    api = str(cfg.get("pushplus_api", "https://www.pushplus.plus/send")).strip()
    payload = {
        "token": cfg.get("pushplus_token", ""),
        "title": title,
        "content": content,
        "topic": cfg.get("pushplus_topic", ""),
        "template": "txt",
        "channel": "wechat"
    }
    try:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urlrequest.Request(
            api,
            data=data,
            method="POST",
            headers={"Content-Type": "application/json; charset=utf-8", "User-Agent": "AutoShutdown/1.0"},
        )
        with urlrequest.urlopen(req, timeout=8) as resp:
            body = resp.read().decode("utf-8", errors="replace")
            code = getattr(resp, "status", None) or resp.getcode()
        if int(code) // 100 != 2:
            return False, f"HTTP {code}\n{body}"
        return True, body
    except Exception as e:
        return False, repr(e)

# -----------------------------
# Hibernate / powercfg
# -----------------------------

def _run_subprocess_hidden(args, capture_output=False, text=False, timeout=None, check=False):
    """Run subprocess without showing a console window (Windows)."""
    try:
        creationflags = 0
        startupinfo = None
        if os.name == "nt":
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
        return subprocess.run(
            args,
            capture_output=capture_output,
            text=text,
            timeout=timeout,
            check=check,
            shell=False,
            startupinfo=startupinfo,
            creationflags=creationflags,
        )
    except Exception:
        # last resort fallback
        return subprocess.run(args, capture_output=capture_output, text=text, timeout=timeout, check=check, shell=False)

# 修改休眠函数，增加备用方法
def go_hibernate():
    """
    尝试多种方法使系统进入休眠状态：
    1. 首先尝试使用shutdown /h命令
    2. 如果失败，则尝试使用 rundll32 powrprof.dll,SetSuspendState
    """
    try:
        # 方法1: 使用 shutdown /h 命令
        result = _run_subprocess_hidden(["shutdown", "/h"], check=False, timeout=10)
        if result.returncode == 0:
            return True
    except Exception:
        pass
    
    try:
        # 方法2: 使用 rundll32 调用电源管理函数
        # SetSuspendState(Hibernate, ForceCritical, DisableWakeEvent)
        result = _run_subprocess_hidden([
            "rundll32.exe", 
            "powrprof.dll,SetSuspendState", 
            "1,1,0"
        ], check=False, timeout=10)
        if result.returncode == 0:
            return True
    except Exception:
        pass
    
    # 如果以上方法都失败，则返回False
    return False

def run_powercfg(args):
    try:
        r = _run_subprocess_hidden(["powercfg"] + list(args), capture_output=True, text=True, timeout=10, check=False)
        return r.returncode, r.stdout or "", r.stderr or ""
    except Exception as e:
        return 1, "", repr(e)

def check_hibernate_available_from_powercfg_a(output: str):
    text = output or ""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    section = None
    hib_in_avail = False
    hib_in_not = False
    for ln in lines:
        low = ln.lower()
        if ("此系统上有以下睡眠状态" in ln) or ("the following sleep states are available" in low):
            section = "avail"; continue
        if ("此系统上没有以下睡眠状态" in ln) or ("the following sleep states are not available" in low):
            section = "not"; continue

        is_hib = ("hibernate" in low) or ("休眠" in ln)
        # 排除“混合睡眠”
        if ("混合睡眠" in ln) or ("hybrid sleep" in low):
            is_hib = False
        if section == "avail" and is_hib:
            hib_in_avail = True
        if section == "not" and is_hib:
            hib_in_not = True

    if hib_in_avail:
        return True, "休眠在当前系统上可用。"
    if hib_in_not:
        return False, "休眠在当前系统上不可用。"
    return False, "未能从 powercfg 输出中确认休眠状态。"

def elevate_enable_hibernate_via_uac() -> bool:
    try:
        shell32.ShellExecuteW.argtypes = [HWND_T, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_int]
        shell32.ShellExecuteW.restype = HANDLE_T
        ret = shell32.ShellExecuteW(None, "runas", "cmd.exe", "/c powercfg /h on", None, 1)
        # per ShellExecute docs: return > 32 means success
        return int(ctypes.cast(ret, ctypes.c_ssize_t).value) > 32
    except Exception:
        return False

# -----------------------------
# Autostart (HKCU Run)
# -----------------------------

def get_running_exe_path() -> str:
    """Get current process executable path using WinAPI (works for Nuitka-built exe)."""
    try:
        kernel32.GetModuleFileNameW.argtypes = [wintypes.HMODULE, wintypes.LPWSTR, wintypes.DWORD]
        kernel32.GetModuleFileNameW.restype = wintypes.DWORD
    except Exception:
        pass
    try:
        kernel32.GetLongPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        kernel32.GetLongPathNameW.restype = wintypes.DWORD
    except Exception:
        pass

    buf_len = 2048
    path = ""
    while True:
        buf = ctypes.create_unicode_buffer(buf_len)
        n = 0
        try:
            n = int(kernel32.GetModuleFileNameW(None, buf, buf_len))
        except Exception:
            n = 0
        if n == 0:
            path = os.path.abspath(sys.executable)
            break
        if n < buf_len - 1:
            path = buf.value
            break
        buf_len = min(buf_len * 2, 32768)
        if buf_len >= 32768:
            path = buf.value
            break

    # Normalize to long path (avoid 8.3 short paths causing mismatches)
    try:
        lb = ctypes.create_unicode_buffer(32768)
        m = int(kernel32.GetLongPathNameW(path, lb, 32768))
        if m:
            path = lb.value
    except Exception:
        pass
    return os.path.abspath(path)

def is_compiled_exe() -> bool:
    exe = os.path.abspath(get_running_exe_path())
    base = os.path.basename(exe).lower()
    if not exe.lower().endswith(".exe"):
        return False
    if base in ("python.exe", "pythonw.exe"):
        return False
    return True

def get_pythonw_path() -> str:
    exe = os.path.abspath(sys.executable)
    if exe.lower().endswith("python.exe"):
        pyw = exe[:-10] + "pythonw.exe"
        if os.path.exists(pyw):
            return pyw
    return exe

def build_expected_autostart_command() -> str:
    if is_compiled_exe():
        exe = os.path.abspath(get_running_exe_path())
        return f"\"{exe}\""
    pyw = get_pythonw_path()
    script = os.path.abspath(get_entry_script_path())
    return f"\"{pyw}\" \"{script}\""

def get_startup_folder() -> str:
    # Current user startup folder
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(appdata, r"Microsoft\Windows\Start Menu\Programs\Startup")


def get_startup_shortcut_path() -> str:
    return os.path.join(get_startup_folder(), STARTUP_SHORTCUT_NAME)


def _ps_escape(s: str) -> str:
    # Escape for embedding in single-quoted PowerShell string
    return (s or "").replace("'", "''")


def run_powershell(ps: str, timeout_sec: int = 10):
    """Run a PowerShell snippet (UTF-8 output). Returns (rc, stdout, stderr)."""
    try:
        p = _run_subprocess_hidden(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=timeout_sec,
            check=False,
        )
        return int(p.returncode), (p.stdout or ""), (p.stderr or "")
    except Exception as e:
        return 1, "", repr(e)


def build_expected_shortcut_spec():
    """
    Returns (target_path, arguments, icon_location, working_dir).
    For exe: target=exe, args=''
    For script: target=pythonw, args="script"
    """
    ico = resource_path("AutoShutdown.ico")
    if is_compiled_exe():
        exe = os.path.abspath(get_running_exe_path())
        wd = os.path.dirname(exe)
        icon_loc = ico if os.path.exists(ico) else exe
        return exe, "", icon_loc, wd

    pyw = get_pythonw_path()
    script = os.path.abspath(get_entry_script_path())
    wd = os.path.dirname(script)
    icon_loc = ico if os.path.exists(ico) else pyw
    # Keep quotes to survive spaces
    return pyw, f"\"{script}\"", icon_loc, wd


def create_startup_shortcut() -> bool:
    try:
        ensure_dirs(get_startup_folder())
        lnk = get_startup_shortcut_path()
        target, args, icon_loc, wd = build_expected_shortcut_spec()

        ps = f"""
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('{_ps_escape(lnk)}')
$sc.TargetPath = '{_ps_escape(target)}'
$sc.Arguments = '{_ps_escape(args)}'
$sc.WorkingDirectory = '{_ps_escape(wd)}'
$sc.IconLocation = '{_ps_escape(icon_loc)},0'
# WindowStyle: 7 = Minimized, 1 = Normal. Use 7 to reduce focus stealing.
$sc.WindowStyle = 7
$sc.Save()
"""
        rc, out, err = run_powershell(ps, timeout_sec=10)
        return rc == 0 and os.path.exists(lnk)
    except Exception as e:
        log_error(e)
        return False


def delete_startup_shortcut() -> bool:
    try:
        lnk = get_startup_shortcut_path()
        if os.path.exists(lnk):
            os.remove(lnk)
        return True
    except Exception as e:
        log_error(e)
        return False


def read_startup_shortcut_spec():
    """
    Returns dict: {TargetPath, Arguments, IconLocation, WorkingDirectory} or None if missing/unreadable.
    """
    lnk = get_startup_shortcut_path()
    if not os.path.exists(lnk):
        return None
    ps = f"""
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut('{_ps_escape(lnk)}')
@{{
  TargetPath=$sc.TargetPath;
  Arguments=$sc.Arguments;
  IconLocation=$sc.IconLocation;
  WorkingDirectory=$sc.WorkingDirectory
}} | ConvertTo-Json -Compress
"""
    rc, out, err = run_powershell(ps, timeout_sec=10)
    if rc != 0:
        return None
    try:
        return json.loads(out.strip())
    except Exception:
        return None


def _norm_path(p: str) -> str:
    try:
        return os.path.normcase(os.path.normpath(p or "")).strip()
    except Exception:
        return (p or "").lower().strip()


def _norm_args(a: str) -> str:
    return " ".join((a or "").strip().split()).lower()


def cleanup_old_registry_run_entry():
    """
    Optional: remove legacy HKCU\\...\\Run entry if present, to avoid duplicate/invalid startup entries.
    """
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
            try:
                winreg.DeleteValue(key, RUN_VALUE_NAME)
            except FileNotFoundError:
                pass
    except Exception:
        pass

def normalize_cmd(s: str) -> str:
    s = (s or "").strip()
    return " ".join(s.split()).lower()

def read_autostart_value():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_READ) as key:
            val, _typ = winreg.QueryValueEx(key, RUN_VALUE_NAME)
        return str(val)
    except Exception:
        return None

def write_autostart_value(cmd: str) -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, RUN_VALUE_NAME, 0, winreg.REG_SZ, cmd)
        return True
    except Exception as e:
        log_error(e)
        return False

def delete_autostart_value() -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
            try:
                winreg.DeleteValue(key, RUN_VALUE_NAME)
            except FileNotFoundError:
                pass
        return True
    except Exception as e:
        log_error(e)
        return False
