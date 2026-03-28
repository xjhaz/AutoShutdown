# -*- coding: utf-8 -*-

import os
import ctypes
from ctypes import wintypes
from datetime import datetime

from constants import (
    WM_DESTROY,
    WM_CLOSE,
    WM_COMMAND,
    WM_TIMER,
    WM_LBUTTONDBLCLK,
    WM_RBUTTONUP,
    WM_SETFONT,
    WM_POWERBROADCAST,
    PBT_APMRESUMECRITICAL,
    PBT_APMRESUMESUSPEND,
    PBT_APMRESUMEAUTOMATIC,
    CS_HREDRAW,
    CS_VREDRAW,
    WS_OVERLAPPED,
    WS_CAPTION,
    WS_SYSMENU,
    WS_VISIBLE,
    WS_MINIMIZEBOX,
    WS_CLIPCHILDREN,
    WS_CHILD,
    WS_TABSTOP,
    WS_VSCROLL,
    WS_EX_TOPMOST,
    WS_EX_DLGMODALFRAME,
    WS_EX_CLIENTEDGE,
    ES_AUTOHSCROLL,
    ES_NUMBER,
    ES_MULTILINE,
    ES_AUTOVSCROLL,
    ES_WANTRETURN,
    BS_PUSHBUTTON,
    CBS_DROPDOWNLIST,
    CBS_HASSTRINGS,
    CB_ADDSTRING,
    CB_GETCURSEL,
    CB_SETCURSEL,
    CBN_SELCHANGE,
    SS_LEFT,
    SW_SHOW,
    MB_OK,
    MB_YESNO,
    MB_DEFBUTTON2,
    MB_ICONINFORMATION,
    MB_ICONWARNING,
    MB_ICONERROR,
    IMAGE_ICON,
    LR_LOADFROMFILE,
    LR_DEFAULTSIZE,
    NIM_ADD,
    NIM_MODIFY,
    NIM_DELETE,
    NIF_MESSAGE,
    NIF_ICON,
    NIF_TIP,
    NIF_INFO,
    NIIF_INFO,
    NIIF_WARNING,
    NIIF_ERROR,
    TIMER_MAIN,
    TIMER_SETTINGS_DELAYCHECK,
    TIMER_COUNTDOWN,
    TRAY_CALLBACK_MSG,
    SID_TOKEN,
    SID_TOPIC,
    SID_API,
    SID_UPTIME_H,
    SID_IDLE_M,
    SID_COUNTDOWN_S,
    SID_NET_URL,
    SID_NET_TIMEOUT,
    SID_ONLINE_POLICY,
    SID_REMIND_TEMPLATE,
    SID_PUSHPLUS_CHANNEL,
    SID_TIMEOUT_ACTION,
    SID_BTN_CHECK_HIB,
    SID_BTN_ENABLE_HIB,
    SID_BTN_TEST_MSG,
    SID_BTN_TEST_HIB,
    SID_BTN_TEST_LOCK,
    SID_BTN_SAVE,
    SID_BTN_CANCEL,
    CID_LABEL,
    CID_BTN_CANCEL,
    CID_BTN_NOW,
)
from winapi import (
    user32,
    kernel32,
    shell32,
    comctl32,
    gdi32,
    HANDLE_T,
    HWND_T,
    HMENU_T,
    HINSTANCE_T,
    HICON_T,
    HCURSOR_T,
    HBRUSH_T,
    WPARAM_T,
    LPARAM_T,
    UINT_PTR_T,
    WNDPROC,
    EnumProc,
)
from core import (
    _w,
    ensure_dirs,
    APPDATA_DIR,
    log_error,
    DEFAULT_REMIND_TEMPLATE_HIBERNATE,
    DEFAULT_REMIND_TEMPLATE_LOCK,
    pushplus_send,
    get_pushplus_channel_name,
    normalize_pushplus_channel,
    normalize_timeout_action,
    get_timeout_action_meta,
    run_powercfg,
    check_hibernate_available_from_powercfg_a,
    elevate_enable_hibernate_via_uac,
    save_config,
    resource_path,
    perform_timeout_action,
)

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


def combo_add_string(hwnd_ctrl, text: str):
    buf = ctypes.create_unicode_buffer(_w(text))
    user32.SendMessageW(hwnd_ctrl, CB_ADDSTRING, WPARAM_T(0), LPARAM_T(ctypes.addressof(buf)))

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
    ]

def load_icon_handle(path: str) -> HICON_T:
    return user32.LoadImageW(None, _w(path), IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)

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

    def __init__(self, owner_hwnd, seconds: int, detail_text: str, title: str, app=None, action: str = "hibernate"):
        self.owner_hwnd = owner_hwnd
        self.app = app
        self.remaining = max(1, int(seconds))
        self.cancelled = False
        self.accepted = False
        self.detail_text = detail_text
        self.title = title
        self.action = normalize_timeout_action(action)
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
        action_meta = get_timeout_action_meta(self.action)

        create_ctrl(self.hwnd, "STATIC", action_meta["intro"], S(16), S(12), S(480), S(22), 0, SS_LEFT)
        create_ctrl(self.hwnd, "STATIC", self.detail_text, S(16), S(40), S(480), S(105), 0, SS_LEFT)
        self.h_label = create_ctrl(self.hwnd, "STATIC", "", S(16), S(150), S(480), S(22), CID_LABEL, SS_LEFT)
        create_ctrl(self.hwnd, "BUTTON", action_meta["cancel_button"], S(140), S(185), S(160), S(36), CID_BTN_CANCEL, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", action_meta["now_button"], S(320), S(185), S(160), S(36), CID_BTN_NOW, BS_PUSHBUTTON | WS_TABSTOP)

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
        try:
            set_text(self.h_label, f"剩余 {self.remaining} 秒...")
        except Exception:
            pass

    def on_timer(self):
        if self.remaining <= 1:
            self.accepted = True
            self.close()
            return
        self.remaining -= 1
        self._update_label()

    def on_cancel(self):
        self.cancelled = True
        self.close()

    def on_now(self):
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
        self.combo_options = {}
        self.remind_templates = {}
        self.current_template_action = "hibernate"
        self.current_pushplus_channel = "app"
        self.h_topic_label = None
        self.h_topic_edit = None
        self.h_template_label = None
        self.h_template_edit = None
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

        width, height = S(720), S(850)
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
        right_margin = S(20)
        col_gap = S(20)
        content_w = width - left_label_x - right_margin
        label_w = max(S(280), int(content_w * 0.38))
        label_w = min(label_w, S(340))
        left_edit_x = left_label_x + label_w + col_gap
        y0 = S(20)
        row_h = S(32)
        edit_w = max(S(260), content_w - label_w - col_gap)

        def add_row(label, cid, value, is_number=False):
            nonlocal y0
            create_ctrl(self.hwnd, "STATIC", label, left_label_x, y0 + S(6), label_w, S(22), 0, SS_LEFT)
            style = ES_AUTOHSCROLL | WS_TABSTOP
            if is_number:
                style |= ES_NUMBER
            h = create_ctrl(self.hwnd, "EDIT", value, left_edit_x, y0, edit_w, S(26), cid, style)
            self.controls[cid] = h
            y0 += row_h

        def add_combo_row(label, cid, options, current_value):
            nonlocal y0
            create_ctrl(self.hwnd, "STATIC", label, left_label_x, y0 + S(6), label_w, S(22), 0, SS_LEFT)
            h = create_ctrl(
                self.hwnd,
                "COMBOBOX",
                "",
                left_edit_x,
                y0,
                edit_w,
                S(220),
                cid,
                CBS_DROPDOWNLIST | CBS_HASSTRINGS | WS_TABSTOP,
            )
            self.controls[cid] = h
            self.combo_options[cid] = list(options)

            selected = 0
            for idx, (text, value) in enumerate(options):
                combo_add_string(h, text)
                if value == current_value:
                    selected = idx
            user32.SendMessageW(h, CB_SETCURSEL, WPARAM_T(selected), LPARAM_T(0))
            y0 += row_h


        cfg = self.app.cfg
        self.remind_templates = {
            "hibernate": str(cfg.get("remind_template_hibernate", DEFAULT_REMIND_TEMPLATE_HIBERNATE)),
            "lock": str(cfg.get("remind_template_lock", DEFAULT_REMIND_TEMPLATE_LOCK)),
        }
        add_row("pushplus token：", SID_TOKEN, str(cfg.get("pushplus_token", "")))
        self.h_topic_label = create_ctrl(self.hwnd, "STATIC", "", left_label_x, y0 + S(6), label_w, S(22), 0, SS_LEFT)
        self.h_topic_edit = create_ctrl(self.hwnd, "EDIT", str(cfg.get("pushplus_topic", "")), left_edit_x, y0, edit_w, S(26), SID_TOPIC, ES_AUTOHSCROLL | WS_TABSTOP)
        self.controls[SID_TOPIC] = self.h_topic_edit
        y0 += row_h
        add_row("pushplus API：", SID_API, str(cfg.get("pushplus_api", "https://www.pushplus.plus/send")))
        add_combo_row(
            "pushplus 通知渠道：",
            SID_PUSHPLUS_CHANNEL,
            [("微信渠道", "wechat"), ("App 渠道", "app")],
            normalize_pushplus_channel(cfg.get("pushplus_channel", "app")),
        )

        y0 += S(10)
        add_row("开机时长阈值（小时）：", SID_UPTIME_H, str(int(cfg.get("uptime_hours", 1))), True)
        add_row("空闲阈值（分钟）：", SID_IDLE_M, str(int(cfg.get("idle_minutes", 1))), True)
        add_row("超时动作倒计时（秒）：", SID_COUNTDOWN_S, str(int(cfg.get("pre_hibernate_countdown_sec", 60))), True)
        add_combo_row(
            "超时后执行动作：",
            SID_TIMEOUT_ACTION,
            [("休眠", "hibernate"), ("锁屏", "lock")],
            normalize_timeout_action(cfg.get("timeout_action", "hibernate")),
        )
        add_row("二级网络校验 URL：", SID_NET_URL, str(cfg.get("net_check_url", "https://baidu.com")))
        add_row("二级网络校验超时（秒）：", SID_NET_TIMEOUT, str(int(cfg.get("net_check_timeout_sec", 2))), True)
        add_row("联网提醒后执行动作次数（0=仅提醒）：", SID_ONLINE_POLICY, str(int(cfg.get("online_remind_times", 0))), True)

        # 自定义提醒内容（支持 {base_info} 占位符）
        self.h_template_label = create_ctrl(self.hwnd, "STATIC", "", left_label_x, y0 + S(6), content_w, S(22), 0, SS_LEFT)
        y0 += S(28)
        self.h_template_edit = create_ctrl(
            self.hwnd,
            "EDIT",
            "",
            left_label_x,
            y0,
            content_w,
            S(140),
            SID_REMIND_TEMPLATE,
            WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | ES_WANTRETURN | WS_VSCROLL,
            WS_EX_CLIENTEDGE,
        )
        self.controls[SID_REMIND_TEMPLATE] = self.h_template_edit
        y0 += S(150)

        by = y0 + S(14)
        two_btn_gap = S(20)
        two_btn_w = (content_w - two_btn_gap) // 2
        create_ctrl(self.hwnd, "BUTTON", "检查休眠是否可用", left_label_x, by, two_btn_w, S(34), SID_BTN_CHECK_HIB, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "开启休眠", left_label_x + two_btn_w + two_btn_gap, by, two_btn_w, S(34), SID_BTN_ENABLE_HIB, BS_PUSHBUTTON | WS_TABSTOP)
        by += S(48)
        three_btn_gap = S(20)
        three_btn_w = (content_w - three_btn_gap * 2) // 3
        create_ctrl(self.hwnd, "BUTTON", "测试消息发送", left_label_x, by, three_btn_w, S(34), SID_BTN_TEST_MSG, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "测试休眠（60秒可取消）", left_label_x + three_btn_w + three_btn_gap, by, three_btn_w, S(34), SID_BTN_TEST_HIB, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "测试锁屏（60秒可取消）", left_label_x + (three_btn_w + three_btn_gap) * 2, by, three_btn_w, S(34), SID_BTN_TEST_LOCK, BS_PUSHBUTTON | WS_TABSTOP)
        by += S(64)
        action_btn_gap = S(16)
        action_btn_w = S(130)
        action_right = left_label_x + content_w
        cancel_x = action_right - action_btn_w
        save_x = cancel_x - action_btn_gap - action_btn_w
        create_ctrl(self.hwnd, "BUTTON", "保存", save_x, by, action_btn_w, S(36), SID_BTN_SAVE, BS_PUSHBUTTON | WS_TABSTOP)
        create_ctrl(self.hwnd, "BUTTON", "取消", cancel_x, by, action_btn_w, S(36), SID_BTN_CANCEL, BS_PUSHBUTTON | WS_TABSTOP)
        by += S(52)
        create_ctrl(self.hwnd, "STATIC", "提示：关闭本窗口不会退出程序；请在托盘菜单选择“退出”。",
                    left_label_x, by, content_w, S(22), 0, SS_LEFT)

        apply_font_to_all_children(self.hwnd, self.hfont)
        self.on_pushplus_channel_changed(force=True)
        self.on_timeout_action_changed(force=True)

    def close(self):
        try:
            if self.hwnd:
                user32.DestroyWindow(self.hwnd)
        except Exception:
            pass
        delete_gdi_object(self.hfont)
        self.hfont = None
        self.hwnd = None

    def _get_timeout_action_from_ui(self) -> str:
        hwnd = self.controls.get(SID_TIMEOUT_ACTION)
        options = self.combo_options.get(SID_TIMEOUT_ACTION, [])
        idx = int(user32.SendMessageW(hwnd, CB_GETCURSEL, WPARAM_T(0), LPARAM_T(0)))
        if 0 <= idx < len(options):
            return options[idx][1]
        return "hibernate"

    def _get_pushplus_channel_from_ui(self) -> str:
        hwnd = self.controls.get(SID_PUSHPLUS_CHANNEL)
        options = self.combo_options.get(SID_PUSHPLUS_CHANNEL, [])
        idx = int(user32.SendMessageW(hwnd, CB_GETCURSEL, WPARAM_T(0), LPARAM_T(0)))
        if 0 <= idx < len(options):
            return options[idx][1]
        return "app"

    def on_pushplus_channel_changed(self, force: bool = False):
        next_channel = normalize_pushplus_channel(self._get_pushplus_channel_from_ui())
        if not force and next_channel == self.current_pushplus_channel:
            return
        self.current_pushplus_channel = next_channel
        if next_channel == "wechat":
            set_text(self.h_topic_label, "群组 topic（微信群发可选）：")
            user32.EnableWindow(self.h_topic_edit, True)
        else:
            set_text(self.h_topic_label, "topic（仅微信渠道群发使用）：")
            user32.EnableWindow(self.h_topic_edit, False)

    def _sync_template_editor(self):
        if not self.h_template_edit or not self.current_template_action:
            return
        self.remind_templates[self.current_template_action] = get_text(self.h_template_edit)

    def on_timeout_action_changed(self, force: bool = False):
        next_action = normalize_timeout_action(self._get_timeout_action_from_ui())
        if not force and next_action == self.current_template_action:
            return
        if not force:
            self._sync_template_editor()
        self.current_template_action = next_action
        action_meta = get_timeout_action_meta(next_action)
        label = f"{action_meta['name']}提醒文案（可用 {{base_info}}加入电脑已运行时间、空闲时间和当前时间）："
        set_text(self.h_template_label, label)
        set_text(self.h_template_edit, self.remind_templates.get(next_action, ""))

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

        def combo_value(cid, default):
            hwnd = self.controls.get(cid)
            options = self.combo_options.get(cid, [])
            idx = int(user32.SendMessageW(hwnd, CB_GETCURSEL, WPARAM_T(0), LPARAM_T(0)))
            if 0 <= idx < len(options):
                return options[idx][1]
            return default

        cfg["pushplus_token"] = s(SID_TOKEN)
        cfg["pushplus_topic"] = s(SID_TOPIC)
        cfg["pushplus_api"] = s(SID_API) or "https://www.pushplus.plus/send"
        cfg["pushplus_channel"] = combo_value(SID_PUSHPLUS_CHANNEL, "app")

        cfg["uptime_hours"] = i(SID_UPTIME_H, cfg.get("uptime_hours", 2), 0, 168)
        cfg["idle_minutes"] = i(SID_IDLE_M, cfg.get("idle_minutes", 60), 0, 24*60)
        cfg["pre_hibernate_countdown_sec"] = i(SID_COUNTDOWN_S, cfg.get("pre_hibernate_countdown_sec", 60), 1, 3600)
        cfg["timeout_action"] = combo_value(SID_TIMEOUT_ACTION, "hibernate")
        cfg["net_check_url"] = s(SID_NET_URL) or "https://baidu.com"
        cfg["net_check_timeout_sec"] = i(SID_NET_TIMEOUT, cfg.get("net_check_timeout_sec", 2), 1, 10)

        cfg["online_remind_times"] = i(SID_ONLINE_POLICY, cfg.get("online_remind_times", 0), 0, 99)

        # 自定义提醒内容模板
        self._sync_template_editor()
        tpl_hibernate = (self.remind_templates.get("hibernate", "") or "").strip()
        tpl_lock = (self.remind_templates.get("lock", "") or "").strip()
        cfg["remind_template_hibernate"] = tpl_hibernate if tpl_hibernate else DEFAULT_REMIND_TEMPLATE_HIBERNATE
        cfg["remind_template_lock"] = tpl_lock if tpl_lock else DEFAULT_REMIND_TEMPLATE_LOCK
        cfg["remind_template"] = cfg["remind_template_hibernate"]


        # autostart enabled is controlled by tray; keep current
        cfg["autostart_enabled"] = bool(self.app.cfg.get("autostart_enabled", False))
        return cfg

    def on_check_hibernate(self):
        code, out, err = run_powercfg(["/a"])
        if code != 0:
            message_box(self.hwnd, "检查失败。\n\n" + (err or out or ""), "休眠可用性检查", MB_OK | MB_ICONERROR)
            return
        ok, reason = check_hibernate_available_from_powercfg_a(out)
        try:
            ensure_dirs(APPDATA_DIR)
            with open(os.path.join(APPDATA_DIR, "powercfg_a.txt"), "w", encoding="utf-8") as f:
                f.write(out)
        except Exception:
            pass
        text = ("休眠可用：是\n" if ok else "休眠可用：否\n") + reason + "\n\n详细输出已写入：%APPDATA%\\AutoShutdown\\powercfg_a.txt"
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
            message_box(self.hwnd, "已发起管理员权限请求。\n如果你选择“是”，系统将开启休眠。\n程序将在 4 秒后自动检查一次结果。", "开启休眠", MB_OK | MB_ICONINFORMATION)
            self._delaycheck_pending = True
            user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_SETTINGS_DELAYCHECK), 4000, None)
        else:
            message_box(self.hwnd, "无法发起管理员命令。\n你可以手动以管理员身份运行：powercfg /h on\n\n" + (err or out or ""), "开启休眠失败", MB_OK | MB_ICONERROR)

    def on_test_msg(self):
        cfg = self._get_cfg_from_ui()
        if not cfg.get("pushplus_token"):
            message_box(self.hwnd, "pushplus token 不能为空。", "测试消息发送", MB_OK | MB_ICONWARNING)
            return
        channel_name = get_pushplus_channel_name(cfg.get("pushplus_channel", "app"))
        title = "AutoShutdown 测试消息"
        content = "这是一条测试消息。\n渠道：%s\n时间：%s" % (
            channel_name,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )
        if cfg.get("pushplus_channel") == "wechat" and cfg.get("pushplus_topic"):
            content += "\ntopic：%s" % cfg.get("pushplus_topic")
        ok, detail = pushplus_send(cfg, title, content)
        message_box(self.hwnd, ("发送成功。\n\n" if ok else "发送失败。\n\n") + (detail or ""), "测试消息发送",
                    MB_OK | (MB_ICONINFORMATION if ok else MB_ICONERROR))

    def _show_test_action_dialog(self, action: str):
        action_meta = get_timeout_action_meta(action)
        detail = (
            f"测试{action_meta['name']}：60 秒后将自动{action_meta['verb']}。\n"
            f"你可以点击“{action_meta['cancel_button']}”或“{action_meta['now_button']}”。\n"
            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        dlg = CountdownDialog(self.hwnd, 60, detail, f"测试{action_meta['name']}（60秒可取消）", app=self.app, action=action_meta["value"])
        dlg.show()
        self.app.active_dialog = dlg

    def on_test_hib(self):
        self._show_test_action_dialog("hibernate")

    def on_test_lock(self):
        self._show_test_action_dialog("lock")

    def on_save(self):
        cfg = self._get_cfg_from_ui()
        if not cfg.get("pushplus_token"):
            message_box(self.hwnd, "pushplus token 不能为空。", "提示", MB_OK | MB_ICONWARNING); return

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

        if msg == WM_POWERBROADCAST:
            if int(wparam) in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND, PBT_APMRESUMECRITICAL):
                if app:
                    app.on_resume_event()
                return 1

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
            notify = (int(wparam) >> 16) & 0xFFFF
            if cid == SID_PUSHPLUS_CHANNEL and notify == CBN_SELCHANGE:
                win.on_pushplus_channel_changed()
                return 0
            if cid == SID_TIMEOUT_ACTION and notify == CBN_SELCHANGE:
                win.on_timeout_action_changed()
                return 0
            if cid == SID_BTN_CHECK_HIB:
                win.on_check_hibernate()
            elif cid == SID_BTN_ENABLE_HIB:
                win.on_enable_hibernate()
            elif cid == SID_BTN_TEST_MSG:
                win.on_test_msg()
            elif cid == SID_BTN_TEST_HIB:
                win.on_test_hib()
            elif cid == SID_BTN_TEST_LOCK:
                win.on_test_lock()
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
                    action_meta = get_timeout_action_meta(getattr(dlg_obj, "action", "hibernate"))
                    marked = None
                    if action_meta["value"] == "hibernate" and dlg_obj and getattr(dlg_obj, "app", None) is not None:
                        try:
                            marked = dlg_obj.app.mark_hibernate_time()
                        except Exception:
                            marked = None
                    ok, err_msg = perform_timeout_action(action_meta["value"])
                    if not ok:
                        if marked and dlg_obj and getattr(dlg_obj, "app", None) is not None:
                            try:
                                dlg_obj.app.revert_hibernate_time(*marked)
                            except Exception:
                                pass
                        try:
                            user32.MessageBoxW(
                                None,
                                _w(err_msg or action_meta["error_message"]),
                                _w(action_meta["error_title"]),
                                MB_OK | MB_ICONERROR)
                        except:
                            pass
            except Exception as e:
                log_error(e)
            return 0
    except Exception as e:
        log_error(e)

    return int(user32.DefWindowProcW(hwnd, msg, wparam, lparam))
