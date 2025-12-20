import os
import sys
import json
import ctypes
import socket
import subprocess
import atexit
import ctypes
from ctypes import wintypes
from datetime import datetime, timedelta

import requests
from PyQt6.QtCore import QTimer, Qt
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import (
    QApplication, QSystemTrayIcon, QMenu,
    QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout,
    QFormLayout, QLineEdit, QSpinBox, QDialogButtonBox, QMessageBox, QTextEdit
)
ERROR_ALREADY_EXISTS = 183
_single_instance_handle = None

def ensure_single_instance(mutex_name: str = r"Local\AutoShutdownReminder_SingleInstance") -> bool:
    """
    返回 True 表示首次实例（允许继续运行）
    返回 False 表示已有实例在运行（应退出）
    """
    global _single_instance_handle

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL

    h = kernel32.CreateMutexW(None, True, mutex_name)
    if not h:
        # 创建失败：保守起见允许启动（也可改为 False）
        return True

    last_err = ctypes.get_last_error()
    if last_err == ERROR_ALREADY_EXISTS:
        # 已有实例：释放句柄并拒绝启动
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

def resource_path(filename: str) -> str:
    """兼容：源码运行 / Nuitka standalone / Nuitka onefile"""
    try:
        if getattr(sys, "frozen", False):   # Nuitka 打包态通常为 True
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base = os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.join(base, filename)


DEFAULT_CONFIG = {
    "pushplus_token": "037219ad39c447b2849c898c3e585887",
    "pushplus_topic": "520",
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

APP_NAME = "AutoShutdown"
RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_VALUE_NAME = "AutoShutdownReminder"


def get_appdata_dir() -> str:
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    return os.path.join(appdata, APP_NAME)


APPDATA_DIR = get_appdata_dir()
CONFIG_PATH = os.path.join(APPDATA_DIR, "config.json")
ERROR_LOG = os.path.join(APPDATA_DIR, "error.log")


def ensure_dirs():
    os.makedirs(APPDATA_DIR, exist_ok=True)


def log_error(e: Exception):
    try:
        ensure_dirs()
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().isoformat(timespec='seconds')} {repr(e)}\n")
    except Exception:
        pass


def migrate_config_if_needed():
    try:
        ensure_dirs()
        if os.path.exists(CONFIG_PATH):
            return

        base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        old_path = os.path.join(base_dir, "config.json")
        if not os.path.exists(old_path):
            return

        with open(old_path, "r", encoding="utf-8") as f:
            data = f.read()
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            f.write(data)

        bak_path = old_path + ".migrated.bak"
        try:
            os.replace(old_path, bak_path)
        except Exception:
            pass
    except Exception as e:
        log_error(e)


def load_config() -> dict:
    migrate_config_if_needed()
    ensure_dirs()
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
        ensure_dirs()
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log_error(e)


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]


def get_idle_seconds() -> int:
    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
    if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)) == 0:
        return 0
    tick = ctypes.windll.kernel32.GetTickCount()
    idle_ms = tick - lii.dwTime
    if idle_ms < 0:
        idle_ms = 0
    return int(idle_ms / 1000)


def get_uptime_seconds() -> int:
    GetTickCount64 = ctypes.windll.kernel32.GetTickCount64
    GetTickCount64.restype = ctypes.c_ulonglong
    return int(GetTickCount64() / 1000)


def is_online_winapi() -> bool:
    try:
        flags = ctypes.c_ulong(0)
        res = ctypes.windll.wininet.InternetGetConnectedState(ctypes.byref(flags), 0)
        return bool(res)
    except Exception:
        return False


def is_online_two_level(cfg: dict) -> bool:
    if not is_online_winapi():
        return False

    url = str(cfg.get("net_check_url", "https://baidu.com")).strip()
    timeout = int(cfg.get("net_check_timeout_sec", 2))

    try:
        host = url.split("://", 1)[-1].split("/", 1)[0].strip()
        if not host:
            return False
        socket.gethostbyname(host)
    except Exception:
        return False

    try:
        r = requests.get(url, timeout=timeout)
        return (r.status_code // 100 in (2, 3))
    except Exception:
        return False


def pushplus_send(cfg: dict, title: str, content: str):
    """
    返回 (ok: bool, detail: str)
    """
    try:
        payload = {
            "token": cfg.get("pushplus_token", ""),
            "title": title,
            "content": content,
            "topic": cfg.get("pushplus_topic", ""),
            "template": "txt",
            "channel": "wechat"
        }
        api = cfg.get("pushplus_api", "https://www.pushplus.plus/send")
        r = requests.post(api, json=payload, timeout=8)
        if r.status_code // 100 != 2:
            return False, f"HTTP {r.status_code}\n{r.text}"
        # pushplus 返回 JSON，尽量展示但不强依赖结构
        return True, r.text
    except Exception as e:
        return False, repr(e)


def go_hibernate():
    subprocess.run(["shutdown", "/h"], check=False)


def format_td(td: timedelta) -> str:
    sec = int(td.total_seconds())
    h = sec // 3600
    m = (sec % 3600) // 60
    return f"{h}小时{m}分钟"


def run_powercfg(args):
    try:
        r = subprocess.run(
            ["powercfg"] + list(args),
            capture_output=True,
            text=True,
            shell=False
        )
        return r.returncode, r.stdout or "", r.stderr or ""
    except Exception as e:
        return 1, "", repr(e)


def check_hibernate_available_from_powercfg_a(output: str):
    """
    更稳健的解析 powercfg /a 输出（兼容中英文、不同措辞）。
    返回 (available: bool, reason: str)
    """
    text = output or ""
    lines = [ln.strip() for ln in text.splitlines()]

    section = None  # "avail" / "not_avail"
    hib_in_avail = False
    hib_in_not_avail = False

    for ln in lines:
        if not ln:
            continue
        low = ln.lower()

        # 可用段落标题
        if (
            "the following sleep states are available" in low
            or "available on this system" in low
            or "此系统上有以下睡眠状态" in ln
            or "在此系统上有以下睡眠状态" in ln
            or "在此系统上可用" in ln
            or "可用的睡眠状态" in ln
        ):
            section = "avail"
            continue

        # 不可用段落标题
        if (
            "the following sleep states are not available" in low
            or "not available on this system" in low
            or "此系统上没有以下睡眠状态" in ln
            or "在此系统上没有以下睡眠状态" in ln
            or "在此系统上不可用" in ln
            or "不可用的睡眠状态" in ln
        ):
            section = "not_avail"
            continue

        is_hibernate_line = ("hibernate" in low) or ("休眠" in ln)

        # 避免把“混合睡眠/Hybrid Sleep”当成休眠
        if "混合睡眠" in ln or "混合休眠" in ln or "hybrid sleep" in low:
            is_hibernate_line = False

        if section == "avail" and is_hibernate_line:
            hib_in_avail = True
        elif section == "not_avail" and is_hibernate_line:
            hib_in_not_avail = True

    if hib_in_avail:
        return True, "休眠（Hibernate）在当前系统上可用。"
    if hib_in_not_avail:
        return False, "休眠（Hibernate）在当前系统上不可用（可查看详细输出原因）。"
    return False, "未能从 powercfg 输出中确认休眠状态（可查看详细输出）。"


def elevate_enable_hibernate_via_uac() -> bool:
    try:
        ret = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            "cmd.exe",
            "/c powercfg /h on",
            None,
            1
        )
        return ret > 32
    except Exception:
        return False


def is_compiled_exe() -> bool:
    exe = os.path.abspath(sys.executable)
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
        exe = os.path.abspath(sys.executable)
        return f"\"{exe}\""
    pyw = get_pythonw_path()
    script = os.path.abspath(__file__)
    return f"\"{pyw}\" \"{script}\""


def normalize_cmd(s: str) -> str:
    s = (s or "").strip()
    parts = s.split()
    return " ".join(parts).lower()


def read_autostart_value():
    if os.name != "nt":
        return None
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_READ) as key:
            val, _typ = winreg.QueryValueEx(key, RUN_VALUE_NAME)
        return str(val)
    except Exception:
        return None


def write_autostart_value(cmd: str) -> bool:
    if os.name != "nt":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, RUN_VALUE_NAME, 0, winreg.REG_SZ, cmd)
        return True
    except Exception as e:
        log_error(e)
        return False


def delete_autostart_value() -> bool:
    if os.name != "nt":
        return False
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


def get_autostart_enabled() -> bool:
    return read_autostart_value() is not None


class AutostartIntegrityDialog(QDialog):
    def __init__(self, current_value: str, expected_value: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("检测到自启项异常")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setWindowIcon(QApplication.windowIcon())
        self.choice = "ignore"

        layout = QVBoxLayout()
        layout.addWidget(QLabel("检测到开机自启项可能被修改或缺失。请选择处理方式："))

        info = QTextEdit()
        info.setReadOnly(True)
        info.setPlainText(
            "当前注册表 Run 值：\n"
            f"{current_value or '(无)'}\n\n"
            "期望值（本程序）：\n"
            f"{expected_value}\n"
        )
        info.setMinimumHeight(180)
        layout.addWidget(info)

        btns = QHBoxLayout()
        b_repair = QPushButton("修复（写回期望值）")
        b_disable = QPushButton("关闭开机自启")
        b_ignore = QPushButton("忽略")

        b_repair.clicked.connect(self._choose_repair)
        b_disable.clicked.connect(self._choose_disable)
        b_ignore.clicked.connect(self._choose_ignore)

        btns.addWidget(b_repair)
        btns.addWidget(b_disable)
        btns.addWidget(b_ignore)
        layout.addLayout(btns)
        self.setLayout(layout)

    def _choose_repair(self):
        self.choice = "repair"
        self.accept()

    def _choose_disable(self):
        self.choice = "disable"
        self.accept()

    def _choose_ignore(self):
        self.choice = "ignore"
        self.accept()


class PreHibernateDialog(QDialog):
    def __init__(self, seconds: int, detail_text: str):
        super().__init__()
        self.setWindowTitle("即将进入休眠（可取消）")
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        self.remaining = max(1, int(seconds))
        self.cancelled = False

        layout = QVBoxLayout()
        lbl_main = QLabel("将进入休眠（Hibernate）。")
        lbl_main.setWordWrap(True)
        layout.addWidget(lbl_main)

        lbl_detail = QLabel(detail_text)
        lbl_detail.setWordWrap(True)
        layout.addWidget(lbl_detail)

        self.lbl_count = QLabel("")
        layout.addWidget(self.lbl_count)

        btn_layout = QHBoxLayout()
        btn_cancel = QPushButton("取消本次休眠")
        btn_cancel.clicked.connect(self.on_cancel)
        btn_layout.addWidget(btn_cancel)

        btn_now = QPushButton("立即休眠")
        btn_now.clicked.connect(self.on_now)
        btn_layout.addWidget(btn_now)

        layout.addLayout(btn_layout)
        self.setLayout(layout)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(1000)
        self.update_label()

    def update_label(self):
        self.lbl_count.setText(f"{self.remaining} 秒后将自动休眠。")

    def tick(self):
        self.remaining -= 1
        if self.remaining <= 0:
            self.timer.stop()
            self.accept()
            return
        self.update_label()

    def on_cancel(self):
        self.cancelled = True
        self.timer.stop()
        self.reject()

    def on_now(self):
        self.timer.stop()
        self.accept()


class SettingsDialog(QDialog):
    def __init__(self, cfg: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.cfg = dict(cfg)

        self._tip_box = None

        layout = QVBoxLayout()
        form = QFormLayout()

        self.ed_token = QLineEdit(self.cfg.get("pushplus_token", ""))
        self.ed_topic = QLineEdit(self.cfg.get("pushplus_topic", ""))
        self.ed_api = QLineEdit(self.cfg.get("pushplus_api", "https://www.pushplus.plus/send"))

        form.addRow("pushplus token：", self.ed_token)
        form.addRow("群组 topic：", self.ed_topic)
        form.addRow("pushplus API：", self.ed_api)

        self.sb_uptime_h = QSpinBox(); self.sb_uptime_h.setRange(0, 168)
        self.sb_uptime_h.setValue(int(self.cfg.get("uptime_hours", 2)))
        form.addRow("开机时长阈值（小时）：", self.sb_uptime_h)

        self.sb_idle_m = QSpinBox(); self.sb_idle_m.setRange(0, 24 * 60)
        self.sb_idle_m.setValue(int(self.cfg.get("idle_minutes", 60)))
        form.addRow("空闲阈值（分钟）：", self.sb_idle_m)

        self.sb_check = QSpinBox(); self.sb_check.setRange(10, 3600)
        self.sb_check.setValue(int(self.cfg.get("check_interval_sec", 120)))
        form.addRow("检查间隔（秒）：", self.sb_check)

        self.sb_cooldown = QSpinBox(); self.sb_cooldown.setRange(1, 24 * 60)
        self.sb_cooldown.setValue(int(self.cfg.get("cooldown_minutes", 60)))
        form.addRow("触发冷却（分钟）：", self.sb_cooldown)

        self.sb_countdown = QSpinBox(); self.sb_countdown.setRange(5, 3600)
        self.sb_countdown.setValue(int(self.cfg.get("pre_hibernate_countdown_sec", 60)))
        form.addRow("休眠倒计时（秒）：", self.sb_countdown)

        self.ed_net_url = QLineEdit(str(self.cfg.get("net_check_url", "https://baidu.com")))
        form.addRow("二级网络校验 URL：", self.ed_net_url)

        self.sb_net_to = QSpinBox(); self.sb_net_to.setRange(1, 10)
        self.sb_net_to.setValue(int(self.cfg.get("net_check_timeout_sec", 2)))
        form.addRow("二级网络校验超时（秒）：", self.sb_net_to)

        layout.addLayout(form)

        # 第一排：休眠检查/开启
        hib_btns = QHBoxLayout()
        self.btn_check_hib = QPushButton("检查休眠是否可用（powercfg /a）")
        self.btn_enable_hib = QPushButton("开启休眠（powercfg /h on）")
        self.btn_check_hib.clicked.connect(self.on_check_hibernate)
        self.btn_enable_hib.clicked.connect(self.on_enable_hibernate)
        hib_btns.addWidget(self.btn_check_hib)
        hib_btns.addWidget(self.btn_enable_hib)
        layout.addLayout(hib_btns)

        # 第二排：测试消息 / 测试休眠（新增）
        test_btns = QHBoxLayout()
        self.btn_test_msg = QPushButton("测试消息发送")
        self.btn_test_hib = QPushButton("测试休眠（60秒可取消）")
        self.btn_test_msg.clicked.connect(self.on_test_pushplus)
        self.btn_test_hib.clicked.connect(self.on_test_hibernate)
        test_btns.addWidget(self.btn_test_msg)
        test_btns.addWidget(self.btn_test_hib)
        layout.addLayout(test_btns)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

        self.setLayout(layout)

    def _close_tip_box_if_any(self):
        try:
            if self._tip_box is not None:
                self._tip_box.close()
                self._tip_box.deleteLater()
                self._tip_box = None
        except Exception:
            self._tip_box = None

    def _show_nonblocking_tip(self, title: str, text: str, ms: int = 1800):
        self._close_tip_box_if_any()

        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setIcon(QMessageBox.Icon.Information)
        box.setText(text)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.setModal(False)
        box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        box.show()
        self._tip_box = box

        def _auto_close():
            try:
                if self._tip_box is box:
                    box.accept()
                    box.deleteLater()
                    self._tip_box = None
            except Exception:
                pass

        QTimer.singleShot(max(300, int(ms)), _auto_close)

    def _temp_cfg_from_ui(self) -> dict:
        """
        用当前 UI 文本构造临时 cfg（不保存），用于测试发送等。
        """
        cfg = self.get_config()
        return cfg

    def on_check_hibernate(self):
        self._close_tip_box_if_any()

        code, out, err = run_powercfg(["/a"])
        msg = QMessageBox(self)
        msg.setWindowTitle("休眠可用性检查")
        if code != 0:
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setText("执行 powercfg /a 失败。")
            msg.setInformativeText("可能原因：系统策略限制或 powercfg 不可用。")
            msg.setDetailedText((err or "") + "\n" + (out or ""))
            msg.exec()
            return

        available, reason = check_hibernate_available_from_powercfg_a(out)
        if available:
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setText("休眠（Hibernate）可用：是")
            msg.setInformativeText(reason)
        else:
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText("休眠（Hibernate）可用：否")
            msg.setInformativeText(reason)

        msg.setDetailedText(out)
        msg.exec()

    def on_enable_hibernate(self):
        code, out, err = run_powercfg(["/h", "on"])
        if code == 0:
            self._show_nonblocking_tip("开启休眠", "已执行开启休眠命令，将在 2 秒后自动检查结果。", ms=1600)
            QTimer.singleShot(2000, self.on_check_hibernate)
            return

        launched = elevate_enable_hibernate_via_uac()
        if launched:
            self._show_nonblocking_tip(
                "开启休眠",
                "已发起管理员权限请求（UAC）。\n"
                "如果你选择“是”，系统将开启休眠。\n"
                "程序将在 4 秒后自动检查一次结果。",
                ms=2200
            )
            QTimer.singleShot(4000, self.on_check_hibernate)
        else:
            msg = QMessageBox(self)
            msg.setWindowTitle("开启休眠失败")
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setText("无法发起管理员命令。")
            msg.setInformativeText("你可以手动以管理员身份运行：powercfg /h on")
            msg.setDetailedText((err or "") + "\n" + (out or ""))
            msg.exec()

    # ===== 新增：测试 pushplus =====
    def on_test_pushplus(self):
        self._close_tip_box_if_any()
        cfg = self._temp_cfg_from_ui()

        token = (cfg.get("pushplus_token") or "").strip()
        topic = (cfg.get("pushplus_topic") or "").strip()
        api = (cfg.get("pushplus_api") or "").strip()

        if not token:
            QMessageBox.warning(self, "测试消息发送", "pushplus token 不能为空。")
            return
        if not topic:
            QMessageBox.warning(self, "测试消息发送", "群组 topic 不能为空。")
            return
        if not api:
            QMessageBox.warning(self, "测试消息发送", "pushplus API 不能为空。")
            return

        title = "AutoShutdown 测试消息"
        content = (
            "这是一条测试群组消息。\n"
            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"topic：{topic}"
        )

        self._show_nonblocking_tip("测试消息发送", "正在发送测试消息…", ms=1200)

        ok, detail = pushplus_send(cfg, title, content)
        msg = QMessageBox(self)
        msg.setWindowTitle("测试消息发送")
        if ok:
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setText("发送成功。")
        else:
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setText("发送失败。")
            msg.setInformativeText("请检查网络、token/topic 是否正确，或 pushplus 服务状态。")
        msg.setDetailedText(detail or "")
        msg.exec()

    # ===== 新增：测试休眠 =====
    def on_test_hibernate(self):
        self._close_tip_box_if_any()

        detail = (
            "测试休眠：60 秒后将进入休眠（Hibernate）。\n"
            "你可以点击“取消本次休眠”或“立即休眠”。\n"
            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )

        dlg = PreHibernateDialog(60, detail)
        result = dlg.exec()

        if result == QDialog.DialogCode.Accepted and not dlg.cancelled:
            go_hibernate()
        else:
            QMessageBox.information(self, "测试休眠", "已取消本次测试休眠。")

    def get_config(self) -> dict:
        cfg = dict(self.cfg)
        cfg["pushplus_token"] = self.ed_token.text().strip()
        cfg["pushplus_topic"] = self.ed_topic.text().strip()
        cfg["pushplus_api"] = self.ed_api.text().strip()

        cfg["uptime_hours"] = int(self.sb_uptime_h.value())
        cfg["idle_minutes"] = int(self.sb_idle_m.value())
        cfg["check_interval_sec"] = int(self.sb_check.value())
        cfg["cooldown_minutes"] = int(self.sb_cooldown.value())
        cfg["pre_hibernate_countdown_sec"] = int(self.sb_countdown.value())

        cfg["net_check_url"] = self.ed_net_url.text().strip()
        cfg["net_check_timeout_sec"] = int(self.sb_net_to.value())
        return cfg


class TrayApp:
    def __init__(self):
        self.cfg = load_config()

        self.last_trigger_time = datetime.min
        self.suppress_once_remind = False
        self.suppress_once_hibernate = False

        self.tray = QSystemTrayIcon()
        icon_path = resource_path("AutoShutdown.ico")
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
        else:
            icon = QIcon.fromTheme("computer")
            if icon.isNull():
                icon = QIcon()

        self.tray.setIcon(icon)

        self.tray.setToolTip("AutoShutdown")

        self.menu = QMenu()

        self.act_no_remind = QAction("本次不提醒（仅影响下一次触发）", self.menu)
        self.act_no_remind.setCheckable(True)
        self.act_no_remind.triggered.connect(self.toggle_no_remind)
        self.menu.addAction(self.act_no_remind)

        self.act_no_hibernate = QAction("本次不关机（仅影响下一次触发）", self.menu)
        self.act_no_hibernate.setCheckable(True)
        self.act_no_hibernate.triggered.connect(self.toggle_no_hibernate)
        self.menu.addAction(self.act_no_hibernate)

        self.menu.addSeparator()

        self.act_autostart = QAction("开机自启", self.menu)
        self.act_autostart.setCheckable(True)
        self.act_autostart.setChecked(get_autostart_enabled())
        self.act_autostart.triggered.connect(self.on_toggle_autostart)
        self.menu.addAction(self.act_autostart)

        self.act_settings = QAction("设置...", self.menu)
        self.act_settings.triggered.connect(self.open_settings)
        self.menu.addAction(self.act_settings)

        self.act_reset = QAction("恢复默认（取消本次抑制）", self.menu)
        self.act_reset.triggered.connect(self.reset_once_flags)
        self.menu.addAction(self.act_reset)

        self.menu.addSeparator()

        self.act_quit = QAction("退出", self.menu)
        self.act_quit.triggered.connect(QApplication.quit)
        self.menu.addAction(self.act_quit)

        self.tray.setContextMenu(self.menu)
        self.tray.show()

        self.timer = QTimer()
        self.timer.timeout.connect(self.tick)
        self.apply_timer_interval()

        self.check_autostart_integrity_on_start()

    def apply_timer_interval(self):
        try:
            interval = int(self.cfg.get("check_interval_sec", 120))
            interval = max(10, interval)
            self.timer.stop()
            self.timer.start(interval * 1000)
        except Exception as e:
            log_error(e)
            self.timer.start(120 * 1000)

    def get_thresholds(self):
        uptime_th = timedelta(hours=int(self.cfg.get("uptime_hours", 2)))
        idle_th = timedelta(minutes=int(self.cfg.get("idle_minutes", 60)))
        cooldown = timedelta(minutes=int(self.cfg.get("cooldown_minutes", 60)))
        countdown = int(self.cfg.get("pre_hibernate_countdown_sec", 60))
        return uptime_th, idle_th, cooldown, countdown

    def check_autostart_integrity_on_start(self):
        try:
            if not bool(self.cfg.get("autostart_enabled", False)):
                return

            expected = build_expected_autostart_command()
            current = read_autostart_value()
            mismatch = (current is None) or (normalize_cmd(current) != normalize_cmd(expected))
            if not mismatch:
                return

            dlg = AutostartIntegrityDialog(current_value=current or "", expected_value=expected)
            dlg.exec()

            if dlg.choice == "repair":
                ok = write_autostart_value(expected)
                self.act_autostart.setChecked(get_autostart_enabled())
                self.cfg["autostart_enabled"] = True
                save_config(self.cfg)
                self.tray.showMessage(
                    "AutoShutdown",
                    "已修复开机自启项。" if ok else "修复失败，请查看 error.log。",
                    QSystemTrayIcon.MessageIcon.Information if ok else QSystemTrayIcon.MessageIcon.Critical
                )

            elif dlg.choice == "disable":
                delete_autostart_value()
                self.act_autostart.setChecked(False)
                self.cfg["autostart_enabled"] = False
                save_config(self.cfg)
                self.tray.showMessage("AutoShutdown", "已关闭开机自启。", QSystemTrayIcon.MessageIcon.Information)

            else:
                self.act_autostart.setChecked(get_autostart_enabled())

        except Exception as e:
            log_error(e)

    def toggle_no_remind(self):
        self.suppress_once_remind = self.act_no_remind.isChecked()

    def toggle_no_hibernate(self):
        self.suppress_once_hibernate = self.act_no_hibernate.isChecked()

    def reset_once_flags(self):
        self.suppress_once_remind = False
        self.suppress_once_hibernate = False
        self.act_no_remind.setChecked(False)
        self.act_no_hibernate.setChecked(False)

    def consume_once_flags(self):
        self.reset_once_flags()

    def on_toggle_autostart(self):
        enabled = self.act_autostart.isChecked()
        try:
            if enabled:
                cmd = build_expected_autostart_command()
                ok = write_autostart_value(cmd)
                if not ok:
                    self.act_autostart.setChecked(get_autostart_enabled())
                    self.tray.showMessage("AutoShutdown", "开启自启失败（请查看 error.log）。", QSystemTrayIcon.MessageIcon.Critical)
                    return
                self.cfg["autostart_enabled"] = True
                save_config(self.cfg)
                self.tray.showMessage("AutoShutdown", "已开启开机自启。", QSystemTrayIcon.MessageIcon.Information)
            else:
                delete_autostart_value()
                self.cfg["autostart_enabled"] = False
                save_config(self.cfg)
                self.tray.showMessage("AutoShutdown", "已关闭开机自启。", QSystemTrayIcon.MessageIcon.Information)

        except Exception as e:
            log_error(e)
            self.act_autostart.setChecked(get_autostart_enabled())

    def open_settings(self):
        dlg = SettingsDialog(self.cfg)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_cfg = dlg.get_config()

            if not new_cfg.get("pushplus_token"):
                QMessageBox.warning(None, "提示", "pushplus token 不能为空。")
                return
            if not new_cfg.get("pushplus_topic"):
                QMessageBox.warning(None, "提示", "群组 topic 不能为空。")
                return

            new_cfg["autostart_enabled"] = bool(self.cfg.get("autostart_enabled", False))
            self.cfg = new_cfg
            save_config(self.cfg)
            self.apply_timer_interval()
            self.tray.showMessage("AutoShutdown", "设置已保存并生效。", QSystemTrayIcon.MessageIcon.Information)

    def should_trigger(self, uptime: timedelta, idle: timedelta) -> bool:
        uptime_th, idle_th, cooldown, _ = self.get_thresholds()
        if uptime < uptime_th:
            return False
        if idle < idle_th:
            return False
        if datetime.now() - self.last_trigger_time < cooldown:
            return False
        return True

    def prepare_hibernate_flow(self, base_info: str):
        _, _, _, countdown = self.get_thresholds()

        if self.suppress_once_hibernate:
            self.tray.showMessage("AutoShutdown", "本次已设置不关机：跳过休眠。", QSystemTrayIcon.MessageIcon.Information)
            self.last_trigger_time = datetime.now()
            self.consume_once_flags()
            return

        self.tray.showMessage("AutoShutdown", "无网络/发送失败：将弹出可取消休眠提示。", QSystemTrayIcon.MessageIcon.Warning)
        detail = base_info + f"\n\n{countdown} 秒后自动休眠（可取消）。"

        dlg = PreHibernateDialog(countdown, detail)
        result = dlg.exec()

        self.last_trigger_time = datetime.now()
        self.consume_once_flags()

        if result == QDialog.DialogCode.Accepted and not dlg.cancelled:
            go_hibernate()
        else:
            self.tray.showMessage("AutoShutdown", "已取消本次休眠。", QSystemTrayIcon.MessageIcon.Information)

    def tick(self):
        try:
            uptime = timedelta(seconds=get_uptime_seconds())
            idle = timedelta(seconds=get_idle_seconds())

            if not self.should_trigger(uptime, idle):
                return

            title = "电脑长时间未关机提醒"
            base_info = (
                f"电脑已运行：{format_td(uptime)}\n"
                f"空闲时间：{int(idle.total_seconds() / 60)} 分钟\n"
                f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            )

            online = is_online_two_level(self.cfg)

            if online:
                if self.suppress_once_remind:
                    self.tray.showMessage("AutoShutdown", "本次已设置不提醒：跳过微信通知。", QSystemTrayIcon.MessageIcon.Information)
                    self.last_trigger_time = datetime.now()
                    self.consume_once_flags()
                    return

                ok, _detail = pushplus_send(self.cfg, title, base_info + "\n\n建议：确认是否需要关机/休眠/合盖。")
                if ok:
                    self.tray.showMessage("AutoShutdown", "已发送群组微信通知。", QSystemTrayIcon.MessageIcon.Information)
                    self.last_trigger_time = datetime.now()
                    self.consume_once_flags()
                    return

                self.prepare_hibernate_flow(base_info)
                return

            self.prepare_hibernate_flow(base_info)

        except Exception as e:
            log_error(e)


def main():
    if os.name != "nt":
        print("This script is intended for Windows 11.")
        return

    ensure_dirs()
    app = QApplication(sys.argv)

    # 单实例：如果已经运行，直接退出（防止开一堆）
    if not ensure_single_instance():
        return

    # 关键：关闭设置窗口等最后窗口时，不退出整个托盘程序
    app.setQuitOnLastWindowClosed(False)

    icon_path = resource_path("AutoShutdown.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    _tray = TrayApp()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
    
