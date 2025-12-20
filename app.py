# -*- coding: utf-8 -*-

import os
import ctypes
from ctypes import wintypes
from datetime import datetime, timedelta

from constants import (
    APP_NAME,
    APP_VERSION,
    GITHUB_URL,
    WS_EX_TOOLWINDOW,
    WS_OVERLAPPED,
    MF_STRING,
    MF_SEPARATOR,
    MF_CHECKED,
    MF_UNCHECKED,
    MF_BYCOMMAND,
    TPM_RIGHTBUTTON,
    TPM_RETURNCMD,
    TPM_NONOTIFY,
    NIIF_INFO,
    NIIF_WARNING,
    NIIF_ERROR,
    TIMER_MAIN,
    WM_COMMAND,
    MID_ONCE_NO_REMIND,
    MID_ONCE_NO_HIBERNATE,
    MID_RESET_ONCE,
    MID_AUTOSTART,
    MID_TRAY_BALLOON,
    MID_SETTINGS,
    MID_ABOUT,
    MID_QUIT,
    MB_OK,
    MB_YESNO,
    MB_DEFBUTTON2,
    MB_ICONINFORMATION,
    MB_ICONWARNING,
)
from winapi import (
    user32,
    kernel32,
    HINSTANCE_T,
    HANDLE_T,
    HICON_T,
    WPARAM_T,
    LPARAM_T,
    UINT_PTR_T,
)
from core import (
    _w,
    ensure_dirs,
    APPDATA_DIR,
    ensure_single_instance,
    log_error,
    resource_path,
    load_config,
    save_config,
    DEFAULT_REMIND_TEMPLATE,
    get_idle_seconds,
    get_uptime_seconds,
    is_online_two_level,
    pushplus_send,
    build_expected_shortcut_spec,
    read_startup_shortcut_spec,
    create_startup_shortcut,
    delete_startup_shortcut,
    cleanup_old_registry_run_entry,
    _norm_path,
    _norm_args,
)
from ui import (
    set_dpi_awareness,
    tray_add,
    tray_delete,
    tray_balloon,
    load_icon_handle,
    message_box,
    task_dialog_3choice,
    SettingsWindow,
    CountdownDialog,
    _register_window_class,
    _hwnd_to_obj,
    _main_wndproc,
    _settings_wndproc,
    _countdown_wndproc,
)

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
        self._online_ok_count = 0
        self.last_trigger_time = datetime.min

        self.settings = None
        self.active_dialog = None
        
        # 添加新属性来跟踪联网提醒状态
        self.online_remind_count = 0
        self.last_online_remind_time = None
        self.resume_grace_until = datetime.min

    def tray_info(self, title: str, msg: str, level=NIIF_INFO):
        try:
            if not bool(self.cfg.get("tray_balloon_enabled", False)):
                return
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

        append(MID_ONCE_NO_REMIND, "本次不提醒")
        append(MID_ONCE_NO_HIBERNATE, "本次不关机")
        sep()
        append(MID_AUTOSTART, "开机自启")
        append(MID_TRAY_BALLOON, "托盘气泡通知")
        append(MID_SETTINGS, "设置...")
        append(MID_ABOUT, "关于...")
        append(MID_RESET_ONCE, "恢复默认提醒/休眠设置")
        sep()
        append(MID_QUIT, "退出")
        self._update_menu_checks()

    def _update_menu_checks(self):
        def check(mid, on):
            user32.CheckMenuItem(self.menu, mid, MF_BYCOMMAND | (MF_CHECKED if on else MF_UNCHECKED))
        check(MID_ONCE_NO_REMIND, self.suppress_once_remind)
        check(MID_ONCE_NO_HIBERNATE, self.suppress_once_hibernate)
        check(MID_AUTOSTART, bool(self.cfg.get("autostart_enabled", False)))
        check(MID_TRAY_BALLOON, bool(self.cfg.get("tray_balloon_enabled", False)))

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
            # 获取当前空闲时间和阈值
            idle_seconds = get_idle_seconds()
            idle_threshold_minutes = int(self.cfg.get("idle_minutes", 60))
            idle_threshold_seconds = idle_threshold_minutes * 60

            remind_times = 0
            try:
                remind_times = int(self.cfg.get("online_remind_times", 0))
            except Exception:
                remind_times = 0
            remind_times = 0 if remind_times < 0 else remind_times

            if (remind_times > 0 and self.online_remind_count > 0 and self.last_online_remind_time
                    and self.online_remind_count < remind_times):
                target = self.last_online_remind_time + timedelta(seconds=idle_threshold_seconds)
                remaining = int((target - datetime.now()).total_seconds())
                interval = remaining if remaining > 0 else 1
            else:
                if idle_seconds < idle_threshold_seconds:
                    interval = max(1, idle_threshold_seconds - idle_seconds)
                else:
                    # 已达到阈值，避免 1 秒内重复触发，按阈值间隔检查
                    interval = max(1, idle_threshold_seconds)

            now = datetime.now()
            if self.resume_grace_until > now:
                grace_remaining = int((self.resume_grace_until - now).total_seconds())
                if grace_remaining > 0:
                    interval = min(interval, grace_remaining)
            interval = max(1, interval)
            
            user32.KillTimer(self.hwnd, UINT_PTR_T(TIMER_MAIN))
            user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_MAIN), interval * 1000, None)
            
            # 添加调试信息，打印下一次检查的时间
            next_check = now + timedelta(seconds=interval)
            # print(f"[DEBUG] 当前空闲秒数: {idle_seconds}, 阈值: {idle_threshold_seconds}")
            # print(f"[DEBUG] 下次检查间隔: {interval} 秒, 预计检查时间: {next_check.strftime('%Y-%m-%d %H:%M:%S')}")
        except Exception as e:
            log_error(e)
            user32.SetTimer(self.hwnd, UINT_PTR_T(TIMER_MAIN), 60 * 1000, None)  # 默认每分钟检查一次

    def _format_td(self, td: timedelta) -> str:
        sec = int(td.total_seconds())
        h = sec // 3600
        m = (sec % 3600) // 60
        return f"{h}小时{m}分钟"

    def _parse_dt(self, value: str):
        if not value:
            return None
        try:
            return datetime.fromisoformat(value)
        except Exception:
            try:
                return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
            except Exception:
                return None

    def _format_dt(self, dt: datetime) -> str:
        return dt.strftime("%Y-%m-%d %H:%M:%S")

    def _set_resume_grace(self):
        try:
            grace_sec = int(self.cfg.get("resume_grace_sec", 120))
        except Exception:
            grace_sec = 120
        if grace_sec <= 0:
            self.resume_grace_until = datetime.min
            return
        self.resume_grace_until = datetime.now() + timedelta(seconds=grace_sec)

    def _maybe_show_last_hibernate_notice(self, prefix: str = ""):
        last_str = str(self.cfg.get("last_hibernate_time", "")).strip()
        if not last_str:
            return
        notice_str = str(self.cfg.get("last_hibernate_notice_time", "")).strip()
        if notice_str == last_str:
            return
        last_dt = self._parse_dt(last_str)
        if not last_dt:
            return
        msg = f"上次休眠时间：{self._format_dt(last_dt)}"
        if prefix:
            msg = f"{prefix}\n{msg}"
        self.tray_info("AutoShutdown", msg, NIIF_INFO)
        self.cfg["last_hibernate_notice_time"] = last_str
        save_config(self.cfg)

    def mark_hibernate_time(self):
        prev_time = str(self.cfg.get("last_hibernate_time", ""))
        prev_notice = str(self.cfg.get("last_hibernate_notice_time", ""))
        ts = datetime.now().isoformat(timespec="seconds")
        self.cfg["last_hibernate_time"] = ts
        self.cfg["last_hibernate_notice_time"] = ""
        save_config(self.cfg)
        return ts, prev_time, prev_notice

    def revert_hibernate_time(self, ts: str, prev_time: str, prev_notice: str):
        if str(self.cfg.get("last_hibernate_time", "")) != ts:
            return
        self.cfg["last_hibernate_time"] = prev_time
        self.cfg["last_hibernate_notice_time"] = prev_notice
        save_config(self.cfg)

    def on_resume_event(self):
        self._online_ok_count = 0
        self.online_remind_count = 0
        self.last_online_remind_time = None
        self._set_resume_grace()
        self._maybe_show_last_hibernate_notice("检测到系统从休眠恢复")
        self.apply_main_timer()

    def get_thresholds(self):
        uptime_th = timedelta(hours=int(self.cfg.get("uptime_hours", 2)))
        idle_th = timedelta(minutes=int(self.cfg.get("idle_minutes", 60)))
        countdown = int(self.cfg.get("pre_hibernate_countdown_sec", 60))
        return uptime_th, idle_th, countdown

    def should_trigger(self, uptime: timedelta, idle: timedelta) -> bool:
        if datetime.now() < self.resume_grace_until:
            return False
        uptime_th, idle_th, _ = self.get_thresholds()
        if uptime < uptime_th:
            return False
        if idle < idle_th:
            return False
        return True

    def consume_once_flags(self):
        self.suppress_once_remind = False
        self.suppress_once_hibernate = False

    def prepare_hibernate_flow(self, base_info: str, reason: str = "无网络/发送失败"):
        _, _, countdown = self.get_thresholds()

        if self.suppress_once_hibernate:
            self.tray_info("AutoShutdown", "本次已设置不关机：跳过休眠。", NIIF_INFO)
            self.last_trigger_time = datetime.now()
            self.consume_once_flags()
            return

        self.tray_info("AutoShutdown", f"{reason}：将弹出可取消休眠提示。", NIIF_WARNING)
        detail = base_info + f"\n\n{countdown} 秒后自动休眠。"
        dlg = CountdownDialog(self.hwnd, countdown, detail, "即将进入休眠", app=self)
        dlg.show()
        self.active_dialog = dlg
        self.last_trigger_time = datetime.now()
        self.consume_once_flags()
        
    def tick(self):
        uptime = timedelta(seconds=get_uptime_seconds())
        idle = timedelta(seconds=get_idle_seconds())

        # 若不满足阈值，清零"联网成功提醒计数"（用于"两次提醒后休眠"）
        try:
            uptime_th, idle_th, _ = self.get_thresholds()
        except Exception:
            uptime_th = timedelta(hours=int(self.cfg.get("uptime_hours", 2)))
            idle_th = timedelta(minutes=int(self.cfg.get("idle_minutes", 60)))
        try:
            if uptime < uptime_th or idle < idle_th:
                self._online_ok_count = 0
                # 重置联网提醒状态
                self.online_remind_count = 0
                self.last_online_remind_time = None
        except Exception:
            pass

        # 重新应用定时器 - 这样可以根据当前空闲状态调整下次检查时间
        self.apply_main_timer()

        if not self.should_trigger(uptime, idle):
            return

        title = "电脑长时间未关机提醒"
        base_info = (
            f"电脑已运行：{self._format_td(uptime)}\n"
            f"空闲时间：{int(idle.total_seconds() / 60)} 分钟\n"
            f"时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )

        # 渲染提醒内容（支持 {base_info}）
        tpl = str(self.cfg.get("remind_template", DEFAULT_REMIND_TEMPLATE) or "").strip()
        if not tpl:
            tpl = DEFAULT_REMIND_TEMPLATE
        if "{base_info}" in tpl:
            content = tpl.replace("{base_info}", base_info)
        else:
            content = base_info + "\n\n" + tpl

        online = is_online_two_level(self.cfg)
        if online:
            if self.suppress_once_remind:
                self.tray_info("AutoShutdown", "本次已设置不提醒：跳过联网消息发送。", NIIF_INFO)
                self.last_trigger_time = datetime.now()
                self.consume_once_flags()
                return

            remind_times = 0
            try:
                remind_times = int(self.cfg.get("online_remind_times", 0))
            except Exception:
                remind_times = 0
            remind_times = 0 if remind_times < 0 else remind_times

            if (remind_times > 0 and self.online_remind_count > 0 and self.last_online_remind_time
                    and self.online_remind_count < remind_times):
                time_since_first_remind = datetime.now() - self.last_online_remind_time
                if time_since_first_remind < idle_th:
                    return

            ok, _detail = pushplus_send(self.cfg, title, content)
            if ok:
                if remind_times <= 0:
                    self.tray_info("AutoShutdown", "已发送群组微信通知。", NIIF_INFO)
                    self.last_trigger_time = datetime.now()
                    self.consume_once_flags()
                    return

                next_count = self.online_remind_count + 1
                self.online_remind_count = next_count
                self.last_online_remind_time = datetime.now()

                if next_count < remind_times:
                    self.tray_info(
                        "AutoShutdown",
                        f"已发送群组微信通知（{next_count}/{remind_times}）：继续满足条件将再次提醒。",
                        NIIF_INFO
                    )
                    self.last_trigger_time = datetime.now()
                    self.consume_once_flags()
                    self.apply_main_timer()
                    return

                self.tray_info(
                    "AutoShutdown",
                    f"已发送群组微信通知（{next_count}/{remind_times}）：将弹出可取消休眠提示。",
                    NIIF_INFO
                )
                self.prepare_hibernate_flow(base_info, reason=f"联网提醒已发送（{next_count}/{remind_times}）")
                self.online_remind_count = 0
                self.last_online_remind_time = None
                return

            # send failed -> treat as offline
            self.online_remind_count = 0
            self.last_online_remind_time = None
            self.prepare_hibernate_flow(base_info, reason="发送失败")
            return

        self.online_remind_count = 0
        self.last_online_remind_time = None
        self.prepare_hibernate_flow(base_info, reason="无网络")

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

    
    def show_about(self):
        info = f"{APP_NAME}\n版本：{APP_VERSION}\nGitHub：{GITHUB_URL}"
        message_box(self.hwnd, info, "关于", MB_OK | MB_ICONINFORMATION)

    def on_menu(self, mid: int):
        if mid == MID_ONCE_NO_REMIND:
            self.suppress_once_remind = not self.suppress_once_remind
        elif mid == MID_ONCE_NO_HIBERNATE:
            self.suppress_once_hibernate = not self.suppress_once_hibernate
        elif mid == MID_RESET_ONCE:
            self.consume_once_flags()
            self.tray_info("AutoShutdown", "已恢复默认。", NIIF_INFO)
        elif mid == MID_AUTOSTART:
            enabled = bool(self.cfg.get("autostart_enabled", False))
            if not enabled:
                cleanup_old_registry_run_entry()
                ok = create_startup_shortcut()
                if ok:
                    self.cfg["autostart_enabled"] = True
                    save_config(self.cfg)
                    self.tray_info("AutoShutdown", "已开启开机自启。", NIIF_INFO)
                else:
                    self.tray_info("AutoShutdown", "开启自启失败。", NIIF_ERROR)
            else:
                delete_startup_shortcut()
                cleanup_old_registry_run_entry()
                self.cfg["autostart_enabled"] = False
                save_config(self.cfg)
                self.tray_info("AutoShutdown", "已关闭开机自启。", NIIF_INFO)
        elif mid == MID_TRAY_BALLOON:
            enabled = bool(self.cfg.get("tray_balloon_enabled", False))
            self.cfg["tray_balloon_enabled"] = not enabled
            save_config(self.cfg)
        elif mid == MID_SETTINGS:
            if self.settings is None:
                self.settings = SettingsWindow(self, self.hwnd)
            self.settings.show()
        elif mid == MID_ABOUT:
            self.show_about()
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
        self._maybe_show_last_hibernate_notice()
        self.apply_main_timer()
        self.autostart_integrity_check()

    def run(self):
        self.create_main_window()
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

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
