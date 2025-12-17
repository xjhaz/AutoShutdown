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
from datetime import datetime, timedelta
from urllib import request as urlrequest

# -----------------------------
# Subprocess helpers (hide console windows for child console apps)
# -----------------------------
def _run_subprocess_hidden(args, **kwargs):
    """Run subprocess without showing a console window (Windows)."""
    if os.name == "nt":
        # CREATE_NO_WINDOW avoids a new console window for console programs (e.g., powershell/powercfg).
        creationflags = int(kwargs.pop("creationflags", 0)) | int(getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000))
        kwargs["creationflags"] = creationflags

        # Also request hidden window style for additional robustness.
        try:
            si = kwargs.pop("startupinfo", None)
            if si is None:
                si = subprocess.STARTUPINFO()
                si.dwFlags |= int(getattr(subprocess, "STARTF_USESHOWWINDOW", 1))
                si.wShowWindow = 0  # SW_HIDE
            kwargs["startupinfo"] = si
        except Exception:
            pass

        # Avoid opening a shell (which may spawn cmd.exe).
        kwargs.setdefault("shell", False)

    return subprocess.run(args, **kwargs)


APP_NAME = "AutoShutdown"

# Startup folder shortcut (no admin, robust)
STARTUP_SHORTCUT_NAME = "AutoShutdownReminder.lnk"
RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_VALUE_NAME = "AutoShutdownReminder"

# -----------------------------
# Pointer-sized types (avoid missing wintypes.* on some builds)
# -----------------------------
IS_64 = ctypes.sizeof(ctypes.c_void_p) == 8
WPARAM_T = ctypes.c_uint64 if IS_64 else ctypes.c_uint32
LPARAM_T = ctypes.c_int64 if IS_64 else ctypes.c_int32
LRESULT_T = ctypes.c_int64 if IS_64 else ctypes.c_int32
UINT_PTR_T = ctypes.c_uint64 if IS_64 else ctypes.c_uint32

HANDLE_T = getattr(wintypes, "HANDLE", ctypes.c_void_p)
HWND_T = getattr(wintypes, "HWND", HANDLE_T)
HMENU_T = getattr(wintypes, "HMENU", HANDLE_T)
HINSTANCE_T = getattr(wintypes, "HINSTANCE", HANDLE_T)
HICON_T = getattr(wintypes, "HICON", HANDLE_T)
HCURSOR_T = getattr(wintypes, "HCURSOR", HANDLE_T)
HBRUSH_T = getattr(wintypes, "HBRUSH", HANDLE_T)
ATOM_T = getattr(wintypes, "ATOM", wintypes.WORD)

# -----------------------------
# DLLs
# -----------------------------
user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
wininet = ctypes.WinDLL("wininet", use_last_error=True)
comctl32 = ctypes.WinDLL("comctl32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

# -----------------------------
# WinAPI constants
# -----------------------------
WM_DESTROY = 0x0002
WM_CLOSE = 0x0010
WM_COMMAND = 0x0111
WM_TIMER = 0x0113
WM_APP = 0x8000
WM_LBUTTONDBLCLK = 0x0203
WM_RBUTTONUP = 0x0205
WM_SETFONT = 0x0030

CS_HREDRAW = 0x0002
CS_VREDRAW = 0x0001

WS_OVERLAPPED = 0x00000000
WS_CAPTION = 0x00C00000
WS_SYSMENU = 0x00080000
WS_VISIBLE = 0x10000000
WS_MINIMIZEBOX = 0x00020000
WS_CLIPCHILDREN = 0x02000000
WS_CHILD = 0x40000000
WS_TABSTOP = 0x00010000
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_TOPMOST = 0x00000008
WS_EX_DLGMODALFRAME = 0x00000001

ES_AUTOHSCROLL = 0x0080
ES_NUMBER = 0x2000
BS_PUSHBUTTON = 0x00000000
SS_LEFT = 0x00000000

SW_SHOW = 5

MB_OK = 0x00000000
MB_YESNO = 0x00000004
MB_DEFBUTTON2 = 0x00000100
MB_ICONINFORMATION = 0x00000040
MB_ICONWARNING = 0x00000030
MB_ICONERROR = 0x00000010

MF_STRING = 0x0000
MF_SEPARATOR = 0x0800
MF_CHECKED = 0x0008
MF_UNCHECKED = 0x0000
MF_BYCOMMAND = 0x0000

TPM_RIGHTBUTTON = 0x0002
TPM_RETURNCMD = 0x0100
TPM_NONOTIFY = 0x0080

IMAGE_ICON = 1
LR_LOADFROMFILE = 0x0010
LR_DEFAULTSIZE = 0x0040

NIM_ADD = 0x00000000
NIM_MODIFY = 0x00000001
NIM_DELETE = 0x00000002

NIF_MESSAGE = 0x00000001
NIF_ICON = 0x00000002
NIF_TIP = 0x00000004
NIF_INFO = 0x00000010

NIIF_INFO = 0x00000001
NIIF_WARNING = 0x00000002
NIIF_ERROR = 0x00000003

ERROR_ALREADY_EXISTS = 183

# Timers
TIMER_MAIN = 1
TIMER_SETTINGS_DELAYCHECK = 2
TIMER_COUNTDOWN = 3

TRAY_CALLBACK_MSG = WM_APP + 1

# Menu IDs
MID_ONCE_NO_REMIND = 1001
MID_ONCE_NO_HIBERNATE = 1002
MID_AUTOSTART = 1003
MID_SETTINGS = 1004
MID_RESET_ONCE = 1005
MID_QUIT = 1006

# Settings controls
SID_TOKEN = 2001
SID_TOPIC = 2002
SID_API = 2003
SID_UPTIME_H = 2004
SID_IDLE_M = 2005
SID_CHECK_SEC = 2006
SID_COOLDOWN_M = 2007
SID_COUNTDOWN_S = 2008
SID_NET_URL = 2009
SID_NET_TIMEOUT = 2010

SID_BTN_CHECK_HIB = 2101
SID_BTN_ENABLE_HIB = 2102
SID_BTN_TEST_MSG = 2103
SID_BTN_TEST_HIB = 2104
SID_BTN_SAVE = 2105
SID_BTN_CANCEL = 2106

# Countdown controls
CID_LABEL = 3001
CID_BTN_CANCEL = 3002
CID_BTN_NOW = 3003

# -----------------------------
# WinAPI prototypes (critical for 64-bit safety)
# -----------------------------
# Window proc
WNDPROC = ctypes.WINFUNCTYPE(LRESULT_T, HWND_T, wintypes.UINT, WPARAM_T, LPARAM_T)

# DefWindowProc
user32.DefWindowProcW.argtypes = [HWND_T, wintypes.UINT, WPARAM_T, LPARAM_T]
user32.DefWindowProcW.restype = LRESULT_T

# RegisterClass
user32.RegisterClassW.argtypes = [ctypes.c_void_p]
user32.RegisterClassW.restype = ATOM_T

# CreateWindowEx
user32.CreateWindowExW.argtypes = [
    wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    HWND_T, HMENU_T, HINSTANCE_T, wintypes.LPVOID
]
user32.CreateWindowExW.restype = HWND_T

# Message loop
user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), HWND_T, wintypes.UINT, wintypes.UINT]
user32.GetMessageW.restype = ctypes.c_int
user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.TranslateMessage.restype = wintypes.BOOL
user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
user32.DispatchMessageW.restype = LRESULT_T

# Posting
user32.PostMessageW.argtypes = [HWND_T, wintypes.UINT, WPARAM_T, LPARAM_T]
user32.PostMessageW.restype = wintypes.BOOL

# Timers
user32.SetTimer.argtypes = [HWND_T, UINT_PTR_T, wintypes.UINT, wintypes.LPVOID]
user32.SetTimer.restype = UINT_PTR_T
user32.KillTimer.argtypes = [HWND_T, UINT_PTR_T]
user32.KillTimer.restype = wintypes.BOOL

# Misc
user32.MessageBoxW.argtypes = [HWND_T, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.UINT]
user32.MessageBoxW.restype = ctypes.c_int
user32.GetWindowTextLengthW.argtypes = [HWND_T]
user32.GetWindowTextLengthW.restype = ctypes.c_int
user32.GetWindowTextW.argtypes = [HWND_T, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowTextW.restype = ctypes.c_int
user32.SetWindowTextW.argtypes = [HWND_T, wintypes.LPCWSTR]
user32.SetWindowTextW.restype = wintypes.BOOL
user32.SendMessageW.argtypes = [HWND_T, wintypes.UINT, WPARAM_T, LPARAM_T]
user32.SendMessageW.restype = LRESULT_T
user32.DestroyWindow.argtypes = [HWND_T]
user32.DestroyWindow.restype = wintypes.BOOL
user32.EnableWindow.argtypes = [HWND_T, wintypes.BOOL]
user32.EnableWindow.restype = wintypes.BOOL
user32.SetForegroundWindow.argtypes = [HWND_T]
user32.SetForegroundWindow.restype = wintypes.BOOL
user32.ShowWindow.argtypes = [HWND_T, ctypes.c_int]
user32.ShowWindow.restype = wintypes.BOOL
user32.PostQuitMessage.argtypes = [ctypes.c_int]
user32.PostQuitMessage.restype = None
user32.GetCursorPos.argtypes = [ctypes.c_void_p]
user32.GetCursorPos.restype = wintypes.BOOL

# Menus
user32.CreatePopupMenu.argtypes = []
user32.CreatePopupMenu.restype = HMENU_T
user32.AppendMenuW.argtypes = [HMENU_T, wintypes.UINT, UINT_PTR_T, wintypes.LPCWSTR]
user32.AppendMenuW.restype = wintypes.BOOL
user32.CheckMenuItem.argtypes = [HMENU_T, wintypes.UINT, wintypes.UINT]
user32.CheckMenuItem.restype = wintypes.DWORD
user32.TrackPopupMenu.argtypes = [HMENU_T, wintypes.UINT, ctypes.c_int, ctypes.c_int, ctypes.c_int, HWND_T, wintypes.LPVOID]
user32.TrackPopupMenu.restype = wintypes.UINT

# Tray
shell32.Shell_NotifyIconW.argtypes = [wintypes.DWORD, ctypes.c_void_p]
shell32.Shell_NotifyIconW.restype = wintypes.BOOL

# InternetGetConnectedState
wininet.InternetGetConnectedState.argtypes = [ctypes.POINTER(wintypes.DWORD), wintypes.DWORD]
wininet.InternetGetConnectedState.restype = wintypes.BOOL

# LoadImage
user32.LoadImageW.argtypes = [HINSTANCE_T, wintypes.LPCWSTR, wintypes.UINT, ctypes.c_int, ctypes.c_int, wintypes.UINT]
user32.LoadImageW.restype = HANDLE_T

# Cursor
user32.LoadCursorW.argtypes = [HINSTANCE_T, HANDLE_T]
user32.LoadCursorW.restype = HCURSOR_T

# EnumChildWindows
EnumProc = ctypes.WINFUNCTYPE(wintypes.BOOL, HWND_T, LPARAM_T)
user32.EnumChildWindows.argtypes = [HWND_T, EnumProc, LPARAM_T]
user32.EnumChildWindows.restype = wintypes.BOOL

# IsWindow
user32.IsWindow.argtypes = [HWND_T]
user32.IsWindow.restype = wintypes.BOOL

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

def resource_path(filename: str) -> str:
    try:
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base = os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.join(base, filename)

# -----------------------------
# DPI awareness / scaling / fonts
# -----------------------------
def set_dpi_awareness():
    # Per-monitor v2 preferred
    try:
        fn = getattr(user32, "SetProcessDpiAwarenessContext", None)
        if fn:
            fn.argtypes = [HANDLE_T]
            fn.restype = wintypes.BOOL
            fn(ctypes.c_void_p(-4))  # PER_MONITOR_AWARE_V2
            return
    except Exception:
        pass
    # shcore fallback
    try:
        shcore = ctypes.WinDLL("shcore", use_last_error=True)
        shcore.SetProcessDpiAwareness.argtypes = [ctypes.c_int]
        shcore.SetProcessDpiAwareness.restype = ctypes.c_int
        shcore.SetProcessDpiAwareness(2)  # PER_MONITOR
        return
    except Exception:
        pass
    # system aware fallback
    try:
        fn2 = getattr(user32, "SetProcessDPIAware", None)
        if fn2:
            fn2.argtypes = []
            fn2.restype = wintypes.BOOL
            fn2()
    except Exception:
        pass

def get_system_dpi() -> int:
    try:
        fn = getattr(user32, "GetDpiForSystem", None)
        if fn:
            fn.argtypes = []
            fn.restype = wintypes.UINT
            return int(fn())
    except Exception:
        pass
    return 96

def scale_by_dpi(v: int, dpi: int) -> int:
    return int(round(int(v) * float(dpi) / 96.0))

class LOGFONTW(ctypes.Structure):
    _fields_ = [
        ("lfHeight", ctypes.c_long),
        ("lfWidth", ctypes.c_long),
        ("lfEscapement", ctypes.c_long),
        ("lfOrientation", ctypes.c_long),
        ("lfWeight", ctypes.c_long),
        ("lfItalic", ctypes.c_byte),
        ("lfUnderline", ctypes.c_byte),
        ("lfStrikeOut", ctypes.c_byte),
        ("lfCharSet", ctypes.c_byte),
        ("lfOutPrecision", ctypes.c_byte),
        ("lfClipPrecision", ctypes.c_byte),
        ("lfQuality", ctypes.c_byte),
        ("lfPitchAndFamily", ctypes.c_byte),
        ("lfFaceName", wintypes.WCHAR * 32),
    ]

def create_ui_font(dpi: int, point_size: int = 10, face: str = "Segoe UI"):
    try:
        kernel32.MulDiv.argtypes = [ctypes.c_int, ctypes.c_int, ctypes.c_int]
        kernel32.MulDiv.restype = ctypes.c_int
        height = -int(kernel32.MulDiv(int(point_size), int(dpi), 72))
        lf = LOGFONTW()
        lf.lfHeight = height
        lf.lfWeight = 400
        lf.lfCharSet = 1  # DEFAULT_CHARSET
        lf.lfQuality = 5  # CLEARTYPE_QUALITY
        lf.lfFaceName = face
        gdi32.CreateFontIndirectW.argtypes = [ctypes.POINTER(LOGFONTW)]
        gdi32.CreateFontIndirectW.restype = HANDLE_T
        return gdi32.CreateFontIndirectW(ctypes.byref(lf))
    except Exception:
        return None

def apply_font(hwnd: HWND_T, hfont):
    if not hwnd or not hfont:
        return
    try:
        user32.SendMessageW(hwnd, WM_SETFONT, WPARAM_T(int(hfont)), LPARAM_T(1))
    except Exception:
        pass

def apply_font_to_all_children(hwnd_parent: HWND_T, hfont):
    if not hwnd_parent or not hfont:
        return
    @EnumProc
    def _cb(hwnd_child, lparam):
        apply_font(hwnd_child, hfont)
        return True
    try:
        user32.EnumChildWindows(hwnd_parent, _cb, LPARAM_T(0))
    except Exception:
        pass

def delete_gdi_object(hobj):
    try:
        if hobj:
            gdi32.DeleteObject.argtypes = [HANDLE_T]
            gdi32.DeleteObject.restype = wintypes.BOOL
            gdi32.DeleteObject(hobj)
    except Exception:
        pass

# -----------------------------
# Config
# -----------------------------
DEFAULT_CONFIG = {
    "pushplus_token": "",
    "pushplus_topic": "",
    "pushplus_api": "https://www.pushplus.plus/send",

    "uptime_hours": 2,
    "idle_minutes": 60,
    "check_interval_sec": 120,
    "cooldown_minutes": 60,
    "pre_hibernate_countdown_sec": 60,

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
def go_hibernate():
    _run_subprocess_hidden(["shutdown", "/h"], check=False)

def run_powercfg(args):
    try:
        r = _run_subprocess_hidden(["powercfg"] + list(args), capture_output=True, text=True, shell=False)
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
        return True, "休眠（Hibernate）在当前系统上可用。"
    if hib_in_not:
        return False, "休眠（Hibernate）在当前系统上不可用（可查看详细输出原因）。"
    return False, "未能从 powercfg 输出中确认休眠状态（可查看详细输出）。"

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

def get_self_executable_path_winapi() -> str:
    """获取当前进程可执行文件的完整路径（WinAPI：GetModuleFileNameW）。"""
    try:
        kernel32.GetModuleFileNameW.argtypes = [HANDLE_T, wintypes.LPWSTR, wintypes.DWORD]
        kernel32.GetModuleFileNameW.restype = wintypes.DWORD
        buf = ctypes.create_unicode_buffer(32768)
        n = kernel32.GetModuleFileNameW(None, buf, len(buf))
        if n == 0:
            raise OSError(ctypes.get_last_error())
        return buf.value
    except Exception:
        # 兜底：尽量返回当前解释器/可执行文件
        try:
            return os.path.abspath(sys.executable)
        except Exception:
            return os.path.abspath(sys.argv[0])

def get_current_program_path() -> str:
    """程序当前路径 + 程序名称（即：当前进程的 exe 完整路径）。"""
    p = get_self_executable_path_winapi()
    # 尽量转换为长路径，避免 8.3 短路径导致的误判/异常
    p = to_long_path_winapi(p)
    try:
        return os.path.abspath(p)
    except Exception:
        return p

def get_current_program_dir() -> str:
    try:
        return os.path.dirname(get_current_program_path())
    except Exception:
        return os.getcwd()

def to_long_path_winapi(p: str) -> str:
    """将路径尽量转换为长路径（用于比较时消除 8.3 短路径差异）。"""
    s = (p or "").strip().strip('"')
    if not s:
        return ""
    try:
        s = os.path.abspath(s)
    except Exception:
        pass
    try:
        kernel32.GetLongPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        kernel32.GetLongPathNameW.restype = wintypes.DWORD
        buf = ctypes.create_unicode_buffer(32768)
        n = kernel32.GetLongPathNameW(s, buf, len(buf))
        if n and n < len(buf):
            return buf.value
    except Exception:
        pass
    return s

def is_compiled_exe() -> bool:
    exe = get_current_program_path()
    base = os.path.basename(exe).lower()
    if not exe.lower().endswith(".exe"):
        return False
    # PyInstaller / Nuitka / cx_Freeze 等常见打包会设置 sys.frozen
    if getattr(sys, "frozen", False):
        return True
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
        exe = get_current_program_path()
        return f"\"{exe}\""
    pyw = get_pythonw_path()
    script = os.path.abspath(__file__)
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
    """
    Run a PowerShell snippet (UTF-8 output). Returns (rc, stdout, stderr).
    """
    try:
        p = _run_subprocess_hidden(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True,
            text=True,
            timeout=timeout_sec
        )
        return p.returncode, p.stdout or "", p.stderr or ""
    except Exception as e:
        return 1, "", repr(e)


def build_expected_shortcut_spec():
    """
    Returns (target_path, arguments, icon_location, working_dir).
    规则：
      - 打包 exe：target = 程序当前路径 + 程序名称（WinAPI 获取），args = ''
      - 直接运行脚本：target = pythonw，args = "script"
    """
    ico = resource_path("AutoShutdown.ico")
    if is_compiled_exe():
        exe_full = get_current_program_path()
        exe_dir = os.path.dirname(exe_full)
        exe_name = os.path.basename(exe_full)
        target = os.path.join(exe_dir, exe_name)
        wd = exe_dir
        icon_loc = ico if os.path.exists(ico) else target
        return target, "", icon_loc, wd

    pyw = get_pythonw_path()
    script = os.path.abspath(__file__)
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
        p2 = to_long_path_winapi(p or "")
        return os.path.normcase(os.path.normpath(p2 or "")).strip()
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

# -----------------------------
# TaskDialog (3 choices) helper
# -----------------------------
class TASKDIALOG_BUTTON(ctypes.Structure):
    _fields_ = [("nButtonID", ctypes.c_int), ("pszButtonText", wintypes.LPCWSTR)]

class TASKDIALOGCONFIG(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("hwndParent", HWND_T),
        ("hInstance", HINSTANCE_T),
        ("dwFlags", wintypes.DWORD),
        ("dwCommonButtons", wintypes.DWORD),
        ("pszWindowTitle", wintypes.LPCWSTR),
        ("MainIcon", ctypes.c_void_p),
        ("pszMainInstruction", wintypes.LPCWSTR),
        ("pszContent", wintypes.LPCWSTR),
        ("cButtons", wintypes.UINT),
        ("pButtons", ctypes.POINTER(TASKDIALOG_BUTTON)),
        ("nDefaultButton", ctypes.c_int),
        ("cRadioButtons", wintypes.UINT),
        ("pRadioButtons", ctypes.c_void_p),
        ("nDefaultRadioButton", ctypes.c_int),
        ("pszVerificationText", wintypes.LPCWSTR),
        ("pszExpandedInformation", wintypes.LPCWSTR),
        ("pszExpandedControlText", wintypes.LPCWSTR),
        ("pszCollapsedControlText", wintypes.LPCWSTR),
        ("FooterIcon", ctypes.c_void_p),
        ("pszFooter", wintypes.LPCWSTR),
        ("pfCallback", ctypes.c_void_p),
        ("lpCallbackData", LPARAM_T),
        ("cxWidth", wintypes.UINT),
    ]

TDF_POSITION_RELATIVE_TO_WINDOW = 0x00001000

def task_dialog_3choice(hwnd_parent, title: str, instruction: str, content: str,
                        b1_text="修复", b2_text="关闭", b3_text="忽略",
                        default_id=101):
    try:
        TaskDialogIndirect = comctl32.TaskDialogIndirect
    except Exception:
        return None

    TaskDialogIndirect.argtypes = [
        ctypes.POINTER(TASKDIALOGCONFIG),
        ctypes.POINTER(ctypes.c_int),
        ctypes.POINTER(ctypes.c_int),
        ctypes.POINTER(ctypes.c_bool),
    ]
    TaskDialogIndirect.restype = ctypes.c_int

    buttons = (TASKDIALOG_BUTTON * 3)()
    buttons[0].nButtonID = 101; buttons[0].pszButtonText = _w(b1_text)
    buttons[1].nButtonID = 102; buttons[1].pszButtonText = _w(b2_text)
    buttons[2].nButtonID = 103; buttons[2].pszButtonText = _w(b3_text)

    cfg = TASKDIALOGCONFIG()
    cfg.cbSize = ctypes.sizeof(TASKDIALOGCONFIG)
    cfg.hwndParent = hwnd_parent
    cfg.hInstance = kernel32.GetModuleHandleW(None)
    cfg.dwFlags = TDF_POSITION_RELATIVE_TO_WINDOW
    cfg.dwCommonButtons = 0
    cfg.pszWindowTitle = _w(title)
    cfg.MainIcon = None
    cfg.pszMainInstruction = _w(instruction)
    cfg.pszContent = _w(content)
    cfg.cButtons = 3
    cfg.pButtons = ctypes.cast(buttons, ctypes.POINTER(TASKDIALOG_BUTTON))
    cfg.nDefaultButton = int(default_id)

    pressed = ctypes.c_int(0)
    hr = TaskDialogIndirect(ctypes.byref(cfg), ctypes.byref(pressed), None, None)
    if hr != 0:
        return None
    return int(pressed.value)

# -----------------------------
# UI basics
# -----------------------------
def message_box(hwnd, text, title="AutoShutdown", flags=MB_OK | MB_ICONINFORMATION):
    return user32.MessageBoxW(hwnd, _w(text), _w(title), flags)

def get_text(hwnd_ctrl) -> str:
    length = user32.GetWindowTextLengthW(hwnd_ctrl)
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd_ctrl, buf, length + 1)
    return buf.value

def set_text(hwnd_ctrl, text: str):
    user32.SetWindowTextW(hwnd_ctrl, _w(text))

def create_ctrl(parent, cls, text, x, y, w, h, cid, style=0, exstyle=0):
    return user32.CreateWindowExW(
        exstyle,
        _w(cls),
        _w(text),
        style | WS_CHILD | WS_VISIBLE,
        int(x), int(y), int(w), int(h),
        parent,
        HMENU_T(int(cid)) if cid else HMENU_T(0),
        kernel32.GetModuleHandleW(None),
        None
    )

# -----------------------------
# Tray icon struct
# -----------------------------
class NOTIFYICONDATA(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", HWND_T),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", HICON_T),
        ("szTip", wintypes.WCHAR * 128),
        ("dwState", wintypes.DWORD),
        ("dwStateMask", wintypes.DWORD),
        ("szInfo", wintypes.WCHAR * 256),
        ("uTimeoutOrVersion", wintypes.UINT),
        ("szInfoTitle", wintypes.WCHAR * 64),
        ("dwInfoFlags", wintypes.DWORD),
        ("guidItem", ctypes.c_byte * 16),
        ("hBalloonIcon", HICON_T),
    ]

def load_icon_handle(path: str) -> HICON_T:
    return ctypes.cast(
        user32.LoadImageW(None, _w(path), IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE),
        HICON_T
    )

def tray_add(hwnd, hicon, tip="AutoShutdown"):
    nid = NOTIFYICONDATA()
    nid.cbSize = ctypes.sizeof(NOTIFYICONDATA)
    nid.hWnd = hwnd
    nid.uID = 1
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP
    nid.uCallbackMessage = TRAY_CALLBACK_MSG
    nid.hIcon = hicon
    nid.szTip = _w(tip)
    shell32.Shell_NotifyIconW(NIM_ADD, ctypes.byref(nid))
    return nid

def tray_delete(nid: NOTIFYICONDATA):
    shell32.Shell_NotifyIconW(NIM_DELETE, ctypes.byref(nid))

def tray_balloon(nid: NOTIFYICONDATA, title: str, msg: str, level: int = NIIF_INFO, timeout_ms: int = 5000):
    nid.uFlags = NIF_INFO | NIF_MESSAGE | NIF_ICON | NIF_TIP
    nid.szInfoTitle = _w(title)[:63]
    nid.szInfo = _w(msg)[:255]
    nid.dwInfoFlags = level
    nid.uTimeoutOrVersion = max(1000, min(int(timeout_ms), 30000))
    shell32.Shell_NotifyIconW(NIM_MODIFY, ctypes.byref(nid))

# -----------------------------
# Window class registration
# -----------------------------
_hwnd_to_obj = {}

def _register_window_class(class_name: str, wndproc):
    class WNDCLASSW(ctypes.Structure):
        _fields_ = [
            ("style", wintypes.UINT),
            ("lpfnWndProc", ctypes.c_void_p),
            ("cbClsExtra", ctypes.c_int),
            ("cbWndExtra", ctypes.c_int),
            ("hInstance", HINSTANCE_T),
            ("hIcon", HICON_T),
            ("hCursor", HCURSOR_T),
            ("hbrBackground", HBRUSH_T),
            ("lpszMenuName", wintypes.LPCWSTR),
            ("lpszClassName", wintypes.LPCWSTR),
        ]
    hinst = kernel32.GetModuleHandleW(None)
    wc = WNDCLASSW()
    wc.style = CS_HREDRAW | CS_VREDRAW
    wc.lpfnWndProc = ctypes.cast(wndproc, ctypes.c_void_p)
    wc.cbClsExtra = 0
    wc.cbWndExtra = 0
    wc.hInstance = hinst
    wc.hIcon = None
    wc.hCursor = user32.LoadCursorW(None, HANDLE_T(32512))  # IDC_ARROW
    wc.hbrBackground = HBRUSH_T(5)  # COLOR_WINDOW+1
    wc.lpszMenuName = None
    wc.lpszClassName = _w(class_name)
    user32.RegisterClassW(ctypes.byref(wc))

# -----------------------------
# Countdown dialog
# -----------------------------
class CountdownDialog:
    CLASS_NAME = "AutoShutdown_CountdownDialog"

    def __init__(self, owner_hwnd, seconds: int, detail_text: str, title: str, app=None):
        self.owner_hwnd = owner_hwnd
        self.app = app
        self.remaining = max(1, int(seconds))
        self.cancelled = False
        self.accepted = False
        self.detail_text = detail_text
        self.title = title
        self.hwnd = None
        self.h_label = None
        self.hfont = None

    def show(self):
        hinst = kernel32.GetModuleHandleW(None)
        dpi = get_system_dpi()
        S = lambda v: scale_by_dpi(v, dpi)

        width, height = S(520), S(270)
        x, y = S(260), S(260)

        self.hwnd = user32.CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_DLGMODALFRAME,
            _w(self.CLASS_NAME),
            _w(self.title),
            WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_VISIBLE | WS_CLIPCHILDREN,
            x, y, width, height,
            self.owner_hwnd,
            None,
            hinst,
            None
        )
        _hwnd_to_obj[self.hwnd] = self

        self.hfont = create_ui_font(dpi, point_size=10, face="Segoe UI")
        apply_font(self.hwnd, self.hfont)

        create_ctrl(self.hwnd, "STATIC", "将进入休眠（Hibernate）。", S(16), S(12), S(480), S(22), 0, SS_LEFT)
        create_ctrl(self.hwnd, "STATIC", self.detail_text, S(16), S(40), S(480), S(105), 0, SS_LEFT)
        self.h_label = create_ctrl(self.hwnd, "STATIC", "", S(16), S(150), S(480), S(22), CID_LABEL, SS_LEFT)
        create_ctrl(self.hwnd, "BUTTON", "取消本次休眠", S(140), S(185), S(160), S(36), CID_BTN_CANCEL, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "立即休眠", S(320), S(185), S(160), S(36), CID_BTN_NOW, BS_PUSHBUTTON | WS_TABSTOP)

        apply_font_to_all_children(self.hwnd, self.hfont)

        self._update_label()
        user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_COUNTDOWN), 1000, None)

        if self.owner_hwnd:
            user32.EnableWindow(self.owner_hwnd, False)

    def close(self):
        try:
            if self.hwnd:
                user32.KillTimer(self.hwnd, UINT_PTR_T(TIMER_COUNTDOWN))
                user32.DestroyWindow(self.hwnd)
        except Exception:
            pass
        finally:
            delete_gdi_object(self.hfont)
            self.hfont = None
            if self.owner_hwnd:
                user32.EnableWindow(self.owner_hwnd, True)
                user32.SetForegroundWindow(self.owner_hwnd)

    def _update_label(self):
        if self.h_label:
            set_text(self.h_label, f"{self.remaining} 秒后将自动休眠。")

    def on_timer(self):
        self.remaining -= 1
        if self.remaining <= 0:
            self.accepted = True
            self.close()
            return
        self._update_label()

    def on_cancel(self):
        self.cancelled = True
        self.accepted = False
        self.close()

    def on_now(self):
        self.cancelled = False
        self.accepted = True
        self.close()

# -----------------------------
# Settings window
# -----------------------------
class SettingsWindow:
    CLASS_NAME = "AutoShutdown_SettingsWindow"

    def __init__(self, app, owner_hwnd):
        self.app = app
        self.owner_hwnd = owner_hwnd
        self.hwnd = None
        self.controls = {}
        self.hfont = None
        self._delaycheck_pending = False

    def is_open(self):
        return bool(self.hwnd and user32.IsWindow(self.hwnd))

    def show(self):
        if self.is_open():
            user32.ShowWindow(self.hwnd, SW_SHOW)
            user32.SetForegroundWindow(self.hwnd)
            return

        hinst = kernel32.GetModuleHandleW(None)
        dpi = get_system_dpi()
        S = lambda v: scale_by_dpi(v, dpi)

        width, height = S(720), S(600)
        x, y = S(240), S(180)

        self.hwnd = user32.CreateWindowExW(
            WS_EX_DLGMODALFRAME,
            _w(self.CLASS_NAME),
            _w("设置"),
            WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_VISIBLE | WS_CLIPCHILDREN,
            x, y, width, height,
            self.owner_hwnd,
            None,
            hinst,
            None
        )
        _hwnd_to_obj[self.hwnd] = self

        self.hfont = create_ui_font(dpi, point_size=10, face="Segoe UI")
        apply_font(self.hwnd, self.hfont)

        left_label_x = S(20)
        left_edit_x = S(240)
        y0 = S(20)
        row_h = S(32)
        edit_w = S(440)
        label_w = S(210)

        def add_row(label, cid, value, is_number=False):
            nonlocal y0
            create_ctrl(self.hwnd, "STATIC", label, left_label_x, y0 + S(6), label_w, S(22), 0, SS_LEFT)
            style = ES_AUTOHSCROLL | WS_TABSTOP
            if is_number:
                style |= ES_NUMBER
            h = create_ctrl(self.hwnd, "EDIT", value, left_edit_x, y0, edit_w, S(26), cid, style)
            self.controls[cid] = h
            y0 += row_h

        cfg = self.app.cfg
        add_row("pushplus token：", SID_TOKEN, str(cfg.get("pushplus_token", "")))
        add_row("群组 topic：", SID_TOPIC, str(cfg.get("pushplus_topic", "")))
        add_row("pushplus API：", SID_API, str(cfg.get("pushplus_api", "https://www.pushplus.plus/send")))

        y0 += S(10)
        add_row("开机时长阈值（小时）：", SID_UPTIME_H, str(int(cfg.get("uptime_hours", 2))), True)
        add_row("空闲阈值（分钟）：", SID_IDLE_M, str(int(cfg.get("idle_minutes", 60))), True)
        add_row("检查间隔（秒）：", SID_CHECK_SEC, str(int(cfg.get("check_interval_sec", 120))), True)
        add_row("触发冷却（分钟）：", SID_COOLDOWN_M, str(int(cfg.get("cooldown_minutes", 60))), True)
        add_row("休眠倒计时（秒）：", SID_COUNTDOWN_S, str(int(cfg.get("pre_hibernate_countdown_sec", 60))), True)
        add_row("二级网络校验 URL：", SID_NET_URL, str(cfg.get("net_check_url", "https://baidu.com")))
        add_row("二级网络校验超时（秒）：", SID_NET_TIMEOUT, str(int(cfg.get("net_check_timeout_sec", 2))), True)

        by = y0 + S(14)
        create_ctrl(self.hwnd, "BUTTON", "检查休眠是否可用（powercfg /a）", S(20), by, S(330), S(34), SID_BTN_CHECK_HIB, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "开启休眠（powercfg /h on）", S(370), by, S(310), S(34), SID_BTN_ENABLE_HIB, BS_PUSHBUTTON | WS_TABSTOP)
        by += S(48)
        create_ctrl(self.hwnd, "BUTTON", "测试消息发送", S(20), by, S(330), S(34), SID_BTN_TEST_MSG, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "测试休眠（60秒可取消）", S(370), by, S(310), S(34), SID_BTN_TEST_HIB, BS_PUSHBUTTON | WS_TABSTOP)
        by += S(64)
        create_ctrl(self.hwnd, "BUTTON", "保存", S(410), by, S(130), S(36), SID_BTN_SAVE, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "取消", S(550), by, S(130), S(36), SID_BTN_CANCEL, BS_PUSHBUTTON | WS_TABSTOP)
        by += S(52)
        create_ctrl(self.hwnd, "STATIC", "提示：关闭本窗口不会退出程序；请在托盘菜单选择“退出”。",
                    S(20), by, S(660), S(22), 0, SS_LEFT)

        apply_font_to_all_children(self.hwnd, self.hfont)

    def close(self):
        try:
            if self.hwnd:
                user32.DestroyWindow(self.hwnd)
        except Exception:
            pass
        delete_gdi_object(self.hfont)
        self.hfont = None
        self.hwnd = None

    def _get_cfg_from_ui(self) -> dict:
        cfg = dict(self.app.cfg)

        def s(cid):
            return get_text(self.controls[cid]).strip()

        def i(cid, default, minv=None, maxv=None):
            try:
                v = int(s(cid))
            except Exception:
                v = int(default)
            if minv is not None:
                v = max(minv, v)
            if maxv is not None:
                v = min(maxv, v)
            return v

        cfg["pushplus_token"] = s(SID_TOKEN)
        cfg["pushplus_topic"] = s(SID_TOPIC)
        cfg["pushplus_api"] = s(SID_API) or "https://www.pushplus.plus/send"

        cfg["uptime_hours"] = i(SID_UPTIME_H, cfg.get("uptime_hours", 2), 0, 168)
        cfg["idle_minutes"] = i(SID_IDLE_M, cfg.get("idle_minutes", 60), 0, 24*60)
        cfg["check_interval_sec"] = i(SID_CHECK_SEC, cfg.get("check_interval_sec", 120), 10, 3600)
        cfg["cooldown_minutes"] = i(SID_COOLDOWN_M, cfg.get("cooldown_minutes", 60), 1, 24*60)
        cfg["pre_hibernate_countdown_sec"] = i(SID_COUNTDOWN_S, cfg.get("pre_hibernate_countdown_sec", 60), 5, 3600)
        cfg["net_check_url"] = s(SID_NET_URL) or "https://baidu.com"
        cfg["net_check_timeout_sec"] = i(SID_NET_TIMEOUT, cfg.get("net_check_timeout_sec", 2), 1, 10)

        # autostart enabled is controlled by tray; keep current
        cfg["autostart_enabled"] = bool(self.app.cfg.get("autostart_enabled", False))
        return cfg

    def on_check_hibernate(self):
        code, out, err = run_powercfg(["/a"])
        if code != 0:
            message_box(self.hwnd, "执行 powercfg /a 失败。\n\n" + (err or out or ""), "休眠可用性检查", MB_OK | MB_ICONERROR)
            return
        ok, reason = check_hibernate_available_from_powercfg_a(out)
        try:
            ensure_dirs(APPDATA_DIR)
            with open(os.path.join(APPDATA_DIR, "powercfg_a.txt"), "w", encoding="utf-8") as f:
                f.write(out)
        except Exception:
            pass
        text = ("休眠（Hibernate）可用：是\n" if ok else "休眠（Hibernate）可用：否\n") + reason + "\n\n详细输出已写入：%APPDATA%\\AutoShutdown\\powercfg_a.txt"
        message_box(self.hwnd, text, "休眠可用性检查", MB_OK | (MB_ICONINFORMATION if ok else MB_ICONWARNING))

    def on_enable_hibernate(self):
        code, out, err = run_powercfg(["/h", "on"])
        if code == 0:
            message_box(self.hwnd, "已执行开启休眠命令，将在 2 秒后自动检查结果。", "开启休眠", MB_OK | MB_ICONINFORMATION)
            self._delaycheck_pending = True
            user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_SETTINGS_DELAYCHECK), 2000, None)
            return

        launched = elevate_enable_hibernate_via_uac()
        if launched:
            message_box(self.hwnd, "已发起管理员权限请求（UAC）。\n如果你选择“是”，系统将开启休眠。\n程序将在 4 秒后自动检查一次结果。", "开启休眠", MB_OK | MB_ICONINFORMATION)
            self._delaycheck_pending = True
            user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_SETTINGS_DELAYCHECK), 4000, None)
        else:
            message_box(self.hwnd, "无法发起管理员命令。\n你可以手动以管理员身份运行：powercfg /h on\n\n" + (err or out or ""), "开启休眠失败", MB_OK | MB_ICONERROR)

    def on_test_msg(self):
        cfg = self._get_cfg_from_ui()
        if not cfg.get("pushplus_token"):
            message_box(self.hwnd, "pushplus token 不能为空。", "测试消息发送", MB_OK | MB_ICONWARNING)
            return
        if not cfg.get("pushplus_topic"):
            message_box(self.hwnd, "群组 topic 不能为空。", "测试消息发送", MB_OK | MB_ICONWARNING)
            return
        title = "AutoShutdown 测试消息"
        content = "这是一条测试群组消息。\n时间：%s\ntopic：%s" % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), cfg.get("pushplus_topic"))
        ok, detail = pushplus_send(cfg, title, content)
        message_box(self.hwnd, ("发送成功。\n\n" if ok else "发送失败。\n\n") + (detail or ""), "测试消息发送",
                    MB_OK | (MB_ICONINFORMATION if ok else MB_ICONERROR))

    def on_test_hib(self):
        detail = (
            "测试休眠：60 秒后将进入休眠（Hibernate）。\n"
            "你可以点击“取消本次休眠”或“立即休眠”。\n"
            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        dlg = CountdownDialog(self.hwnd, 60, detail, "测试休眠（60秒可取消）", app=self.app)
        dlg.show()
        self.app.active_dialog = dlg

    def on_save(self):
        cfg = self._get_cfg_from_ui()
        if not cfg.get("pushplus_token"):
            message_box(self.hwnd, "pushplus token 不能为空。", "提示", MB_OK | MB_ICONWARNING); return
        if not cfg.get("pushplus_topic"):
            message_box(self.hwnd, "群组 topic 不能为空。", "提示", MB_OK | MB_ICONWARNING); return

        self.app.cfg = cfg
        save_config(cfg)
        self.app.apply_main_timer()
        self.app.tray_info("AutoShutdown", "设置已保存并生效。", NIIF_INFO)
        self.close()

    def on_cancel(self):
        self.close()

    def on_timer(self, timer_id: int):
        if timer_id == TIMER_SETTINGS_DELAYCHECK and self._delaycheck_pending:
            self._delaycheck_pending = False
            user32.KillTimer(self.hwnd, UINT_PTR_T(TIMER_SETTINGS_DELAYCHECK))
            self.on_check_hibernate()

# -----------------------------
# App
# -----------------------------
class App:
    MAIN_CLASS = "AutoShutdown_HiddenMain"

    def __init__(self):
        self.cfg = load_config()
        self.hwnd = None
        self.nid = None
        self.menu = None
        self.hicon = None

        self.suppress_once_remind = False
        self.suppress_once_hibernate = False
        self.last_trigger_time = datetime.min

        self.settings = None
        self.active_dialog = None

    def tray_info(self, title: str, msg: str, level=NIIF_INFO):
        try:
            tray_balloon(self.nid, title, msg, level=level, timeout_ms=5000)
        except Exception as e:
            log_error(e)

    def create_tray(self):
        icon_path = resource_path("AutoShutdown.ico")
        try:
            hicon = load_icon_handle(icon_path) if os.path.exists(icon_path) else None
        except Exception:
            hicon = None
        if not hicon:
            # IDI_APPLICATION
            user32.LoadIconW.argtypes = [HINSTANCE_T, HANDLE_T]
            user32.LoadIconW.restype = HICON_T
            hicon = user32.LoadIconW(None, HANDLE_T(32512))
        self.hicon = hicon
        self.nid = tray_add(self.hwnd, self.hicon, tip="AutoShutdown")

    def destroy_tray(self):
        try:
            if self.nid:
                tray_delete(self.nid)
        except Exception:
            pass
        self.nid = None

    def build_menu(self):
        self.menu = user32.CreatePopupMenu()

        def append(mid, text):
            user32.AppendMenuW(self.menu, MF_STRING, UINT_PTR_T(mid), _w(text))

        def sep():
            user32.AppendMenuW(self.menu, MF_SEPARATOR, UINT_PTR_T(0), None)

        append(MID_ONCE_NO_REMIND, "本次不提醒（仅影响下一次触发）")
        append(MID_ONCE_NO_HIBERNATE, "本次不关机（仅影响下一次触发）")
        sep()
        append(MID_AUTOSTART, "开机自启")
        append(MID_SETTINGS, "设置...")
        append(MID_RESET_ONCE, "恢复默认（取消本次抑制）")
        sep()
        append(MID_QUIT, "退出")
        self._update_menu_checks()

    def _update_menu_checks(self):
        def check(mid, on):
            user32.CheckMenuItem(self.menu, mid, MF_BYCOMMAND | (MF_CHECKED if on else MF_UNCHECKED))
        check(MID_ONCE_NO_REMIND, self.suppress_once_remind)
        check(MID_ONCE_NO_HIBERNATE, self.suppress_once_hibernate)
        check(MID_AUTOSTART, bool(self.cfg.get("autostart_enabled", False)))

    def show_menu(self):
        self._update_menu_checks()

        class POINT(ctypes.Structure):
            _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

        pt = POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        user32.SetForegroundWindow(self.hwnd)
        cmd = user32.TrackPopupMenu(self.menu, TPM_RIGHTBUTTON | TPM_RETURNCMD | TPM_NONOTIFY, pt.x, pt.y, 0, self.hwnd, None)
        if cmd:
            user32.PostMessageW(self.hwnd, WM_COMMAND, WPARAM_T(cmd), LPARAM_T(0))

    def apply_main_timer(self):
        try:
            interval = int(self.cfg.get("check_interval_sec", 120))
            interval = max(10, interval)
            user32.KillTimer(self.hwnd, UINT_PTR_T(TIMER_MAIN))
            user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_MAIN), interval * 1000, None)
        except Exception as e:
            log_error(e)
            user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_MAIN), 120 * 1000, None)

    def _format_td(self, td: timedelta) -> str:
        sec = int(td.total_seconds())
        h = sec // 3600
        m = (sec % 3600) // 60
        return f"{h}小时{m}分钟"

    def get_thresholds(self):
        uptime_th = timedelta(hours=int(self.cfg.get("uptime_hours", 2)))
        idle_th = timedelta(minutes=int(self.cfg.get("idle_minutes", 60)))
        cooldown = timedelta(minutes=int(self.cfg.get("cooldown_minutes", 60)))
        countdown = int(self.cfg.get("pre_hibernate_countdown_sec", 60))
        return uptime_th, idle_th, cooldown, countdown

    def should_trigger(self, uptime: timedelta, idle: timedelta) -> bool:
        uptime_th, idle_th, cooldown, _ = self.get_thresholds()
        if uptime < uptime_th:
            return False
        if idle < idle_th:
            return False
        if datetime.now() - self.last_trigger_time < cooldown:
            return False
        return True

    def consume_once_flags(self):
        self.suppress_once_remind = False
        self.suppress_once_hibernate = False

    def prepare_hibernate_flow(self, base_info: str):
        _, _, _, countdown = self.get_thresholds()

        if self.suppress_once_hibernate:
            self.tray_info("AutoShutdown", "本次已设置不关机：跳过休眠。", NIIF_INFO)
            self.last_trigger_time = datetime.now()
            self.consume_once_flags()
            return

        self.tray_info("AutoShutdown", "无网络/发送失败：将弹出可取消休眠提示。", NIIF_WARNING)
        detail = base_info + f"\n\n{countdown} 秒后自动休眠（可取消）。"
        dlg = CountdownDialog(self.hwnd, countdown, detail, "即将进入休眠（可取消）", app=self)
        dlg.show()
        self.active_dialog = dlg
        self.last_trigger_time = datetime.now()
        self.consume_once_flags()

    def tick(self):
        uptime = timedelta(seconds=get_uptime_seconds())
        idle = timedelta(seconds=get_idle_seconds())
        if not self.should_trigger(uptime, idle):
            return

        title = "电脑长时间未关机提醒"
        base_info = (
            f"电脑已运行：{self._format_td(uptime)}\n"
            f"空闲时间：{int(idle.total_seconds() / 60)} 分钟\n"
            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )

        online = is_online_two_level(self.cfg)
        if online:
            if self.suppress_once_remind:
                self.tray_info("AutoShutdown", "本次已设置不提醒：跳过微信通知。", NIIF_INFO)
                self.last_trigger_time = datetime.now()
                self.consume_once_flags()
                return

            ok, _detail = pushplus_send(self.cfg, title, base_info + "\n\n建议：确认是否需要关机/休眠/合盖。")
            if ok:
                self.tray_info("AutoShutdown", "已发送群组微信通知。", NIIF_INFO)
                self.last_trigger_time = datetime.now()
                self.consume_once_flags()
                return

            # send failed -> treat as offline
            self.prepare_hibernate_flow(base_info)
            return

        self.prepare_hibernate_flow(base_info)

    def autostart_integrity_check(self):
        try:
            # 配置为“关闭自启”时，确保没有残留（包括旧注册表方式）
            if not bool(self.cfg.get("autostart_enabled", False)):
                delete_startup_shortcut()
                cleanup_old_registry_run_entry()
                return

            expected_target, expected_args, expected_icon, expected_wd = build_expected_shortcut_spec()
            current = read_startup_shortcut_spec()

            mismatch = True
            if isinstance(current, dict):
                ct = _norm_path(current.get("TargetPath", ""))
                ca = _norm_args(current.get("Arguments", ""))
                et = _norm_path(expected_target)
                ea = _norm_args(expected_args)
                mismatch = not (ct == et and ca == ea)


            if not mismatch:
                return

            # 启用自启时：每次启动都检查 lnk 是否存在且路径是否正确；异常则优先自动修复
            try:
                cleanup_old_registry_run_entry()
                ok = create_startup_shortcut()
                if ok:
                    # 确保配置一致（防止外部删除 lnk 后仍显示“已开启”但无法自启）
                    self.cfg["autostart_enabled"] = True
                    save_config(self.cfg)
                    return
            except Exception as e:
                log_error(e)

            # 自动修复失败：再让用户选择 修复 / 关闭 / 忽略
            # 修复 / 关闭 / 忽略 三选一
            cur_desc = "(无)" if not current else (
                f"Target={current.get('TargetPath','')}\nArgs={current.get('Arguments','')}"
            )
            exp_desc = f"Target={expected_target}\nArgs={expected_args}"

            pressed = task_dialog_3choice(
                self.hwnd,
                "检测到自启项异常",
                "开机自启快捷方式可能被修改或缺失",
                "为安全起见，建议修复为期望命令（防止路径被篡改）。\n\n"
                f"当前 Startup 快捷方式：{cur_desc}\n\n"
                f"期望：{exp_desc}\n",
                b1_text="修复",
                b2_text="关闭",
                b3_text="忽略",
                default_id=101
            )

            if pressed is None:
                # 降级：两步 MessageBox
                r = message_box(self.hwnd,
                                "检测到开机自启快捷方式可能被修改或缺失。\n\n"
                                f"当前：{cur_desc}\n\n"
                                f"期望：{exp_desc}\n\n"
                                "是否修复为期望值？",
                                "检测到自启项异常",
                                MB_YESNO | MB_ICONWARNING)
                if r == 6:
                    pressed = 101
                else:
                    r2 = message_box(self.hwnd,
                                     "是否关闭开机自启？\n\n选择“否”将忽略本次提示。",
                                     "检测到自启项异常",
                                     MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2)
                    pressed = 102 if r2 == 6 else 103

            if pressed == 101:
                cleanup_old_registry_run_entry()
                ok = create_startup_shortcut()
                self.cfg["autostart_enabled"] = True
                save_config(self.cfg)
                self.tray_info("AutoShutdown", "已修复开机自启快捷方式。" if ok else "修复失败，请查看 error.log。", NIIF_INFO if ok else NIIF_ERROR)
            elif pressed == 102:
                delete_startup_shortcut()
                cleanup_old_registry_run_entry()
                self.cfg["autostart_enabled"] = False
                save_config(self.cfg)
                self.tray_info("AutoShutdown", "已关闭开机自启。", NIIF_INFO)
            else:
                # ignore
                pass
        except Exception as e:
            log_error(e)

    def on_menu(self, mid: int):
        if mid == MID_ONCE_NO_REMIND:
            self.suppress_once_remind = not self.suppress_once_remind
        elif mid == MID_ONCE_NO_HIBERNATE:
            self.suppress_once_hibernate = not self.suppress_once_hibernate
        elif mid == MID_RESET_ONCE:
            self.consume_once_flags()
            self.tray_info("AutoShutdown", "已恢复默认（取消本次抑制）。", NIIF_INFO)
        elif mid == MID_AUTOSTART:
            enabled = bool(self.cfg.get("autostart_enabled", False))
            if not enabled:
                cleanup_old_registry_run_entry()
                ok = create_startup_shortcut()
                if ok:
                    self.cfg["autostart_enabled"] = True
                    save_config(self.cfg)
                    self.tray_info("AutoShutdown", "已开启开机自启（Startup 快捷方式）。", NIIF_INFO)
                else:
                    self.tray_info("AutoShutdown", "开启自启失败（无法创建 Startup 快捷方式）。", NIIF_ERROR)
            else:
                delete_startup_shortcut()
                cleanup_old_registry_run_entry()
                self.cfg["autostart_enabled"] = False
                save_config(self.cfg)
                self.tray_info("AutoShutdown", "已关闭开机自启。", NIIF_INFO)
        elif mid == MID_SETTINGS:
            if self.settings is None:
                self.settings = SettingsWindow(self, self.hwnd)
            self.settings.show()
        elif mid == MID_QUIT:
            self.quit()

        self._update_menu_checks()

    def quit(self):
        try:
            self.destroy_tray()
        except Exception:
            pass
        user32.PostQuitMessage(0)

    def create_main_window(self):
        _register_window_class(self.MAIN_CLASS, _main_wndproc)
        _register_window_class(SettingsWindow.CLASS_NAME, _settings_wndproc)
        _register_window_class(CountdownDialog.CLASS_NAME, _countdown_wndproc)

        hinst = kernel32.GetModuleHandleW(None)
        self.hwnd = user32.CreateWindowExW(
            WS_EX_TOOLWINDOW,
            _w(self.MAIN_CLASS),
            _w("AutoShutdownHidden"),
            WS_OVERLAPPED,
            0, 0, 0, 0,
            None, None, hinst, None
        )
        _hwnd_to_obj[self.hwnd] = self

        self.create_tray()
        self.build_menu()
        self.apply_main_timer()
        self.autostart_integrity_check()

    def run(self):
        self.create_main_window()
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

# -----------------------------
# Window procedures
# -----------------------------
@WNDPROC
def _main_wndproc(hwnd, msg, wparam, lparam):
    app = _hwnd_to_obj.get(hwnd)
    try:
        if msg == TRAY_CALLBACK_MSG:
            if int(lparam) == WM_RBUTTONUP:
                app.show_menu()
            elif int(lparam) == WM_LBUTTONDBLCLK:
                app.on_menu(MID_SETTINGS)
            return 0

        if msg == WM_COMMAND:
            mid = int(wparam) & 0xFFFF
            app.on_menu(mid)
            return 0

        if msg == WM_TIMER:
            if int(wparam) == TIMER_MAIN:
                if app.active_dialog is None:
                    app.tick()
                return 0

        if msg == WM_DESTROY:
            try:
                app.destroy_tray()
            except Exception:
                pass
            user32.PostQuitMessage(0)
            return 0
    except Exception as e:
        log_error(e)

    return int(user32.DefWindowProcW(hwnd, msg, wparam, lparam))

@WNDPROC
def _settings_wndproc(hwnd, msg, wparam, lparam):
    win = _hwnd_to_obj.get(hwnd)
    try:
        if msg == WM_COMMAND:
            cid = int(wparam) & 0xFFFF
            if cid == SID_BTN_CHECK_HIB:
                win.on_check_hibernate()
            elif cid == SID_BTN_ENABLE_HIB:
                win.on_enable_hibernate()
            elif cid == SID_BTN_TEST_MSG:
                win.on_test_msg()
            elif cid == SID_BTN_TEST_HIB:
                win.on_test_hib()
            elif cid == SID_BTN_SAVE:
                win.on_save()
            elif cid == SID_BTN_CANCEL:
                win.on_cancel()
            return 0

        if msg == WM_TIMER:
            win.on_timer(int(wparam))
            return 0

        if msg == WM_CLOSE:
            user32.DestroyWindow(hwnd)
            return 0

        if msg == WM_DESTROY:
            _hwnd_to_obj.pop(hwnd, None)
            return 0
    except Exception as e:
        log_error(e)

    return int(user32.DefWindowProcW(hwnd, msg, wparam, lparam))

@WNDPROC
def _countdown_wndproc(hwnd, msg, wparam, lparam):
    dlg = _hwnd_to_obj.get(hwnd)
    try:
        if msg == WM_COMMAND:
            cid = int(wparam) & 0xFFFF
            if cid == CID_BTN_CANCEL:
                dlg.on_cancel()
            elif cid == CID_BTN_NOW:
                dlg.on_now()
            return 0

        if msg == WM_TIMER and int(wparam) == TIMER_COUNTDOWN:
            dlg.on_timer()
            return 0

        if msg == WM_CLOSE:
            dlg.on_cancel()
            return 0

        if msg == WM_DESTROY:
            dlg_obj = dlg
            _hwnd_to_obj.pop(hwnd, None)

            try:
                if dlg_obj and getattr(dlg_obj, "app", None) is not None:
                    dlg_obj.app.active_dialog = None
                accepted = bool(dlg_obj and dlg_obj.accepted and (not dlg_obj.cancelled))
                if accepted:
                    go_hibernate()
            except Exception as e:
                log_error(e)
            return 0
    except Exception as e:
        log_error(e)

    return int(user32.DefWindowProcW(hwnd, msg, wparam, lparam))

# -----------------------------
# main
# -----------------------------
def main():
    if os.name != "nt":
        print("This script is intended for Windows.")
        return

    set_dpi_awareness()
    ensure_dirs(APPDATA_DIR)

    if not ensure_single_instance():
        return

    app = App()
    app.run()

if __name__ == "__main__":
    main()
