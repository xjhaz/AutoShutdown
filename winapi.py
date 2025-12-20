# -*- coding: utf-8 -*-

import ctypes
from ctypes import wintypes

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
