# AutoShutdown（Windows 原生托盘版）

一个 **纯 Python（仅标准库）** 实现的 Windows 托盘常驻工具：  
当满足 **电脑运行时间 + 用户空闲时间** 条件后，根据 **联网状态与策略**，自动 **发送群组微信通知**，并可 **进入可取消的休眠流程**。

本项目不依赖第三方 GUI 框架，完全基于 **Win32 API（ctypes）** 实现，适合追求轻量、可控与可打包分发的场景。

---


## 主要功能

### 1️⃣ 托盘常驻 + 右键菜单

- 程序启动后 **无主窗口**，仅在系统托盘显示图标  
- **右键菜单** 支持：
  - 本次不提醒（仅对下一次生效）
  - 本次不关机（仅对下一次生效）
  - 开机自启（带完整性校验）
  - 设置...
  - 关于...
  - 恢复默认提醒/休眠设置
  - 退出
- **双击托盘图标**：打开设置窗口

---

### 2️⃣ 触发条件

触发需 **同时满足**：

- 开机运行时间 ≥ `uptime_hours`
- 用户空闲时间 ≥ `idle_minutes`

---

### 3️⃣ 联网判断（两级校验）

1. **WinAPI**：`InternetGetConnectedState`
2. **二级校验**：  
   - DNS 解析  
   - 轻量 HTTP GET（URL 与超时可配置）

---

### 4️⃣ 联网状态下的行为（可配置）

当检测到联网且消息发送成功时，支持按“提醒次数后休眠”配置：

| 配置值 | 行为                                   |
| ------ | -------------------------------------- |
| 0      | 仅提醒：发送群组微信通知，不休眠       |
| 1      | 提醒后休眠：发送通知后弹出休眠倒计时   |
| N>1    | 提醒 N 次后休眠：第 N 次触发进入休眠流程 |

---

### 5️⃣ 离线 / 发送失败：可取消休眠倒计时

- 无网络或发送失败 → **弹出置顶倒计时窗口**

- 默认 60 秒（可配置）：

  - 取消本次休眠
  - 立即休眠

- 倒计时结束自动执行：  

  ```
  shutdown /h
  ```

---

### 6️⃣ 设置窗口（原生 Win32）

支持配置：

- pushplus Token / Topic / API
- 开机运行时长阈值（小时）
- 空闲时间阈值（分钟）
- 休眠倒计时（秒）
- 二级网络校验 URL / 超时
- **联网提醒后休眠次数（0=仅提醒）**
- **自定义提醒内容模板**（支持 `{base_info}`）

附带功能：

- 测试消息发送
- 测试休眠（60 秒可取消）
- 休眠可用性检查（`powercfg /a`）
- 一键开启休眠（必要时自动 UAC 提权）

---

### 7️⃣ 自定义提醒内容

提醒内容支持模板，占位符：

- `{base_info}` 自动展开为：
  - 电脑已运行时间
  - 空闲时间
  - 当前时间

示例：

```
{base_info}

建议：确认是否需要关机、休眠或合盖。
```

---

### 8️⃣ 开机自启

- 使用 **Startup 目录 .lnk 快捷方式**
- **WinAPI 获取当前运行 EXE 路径**（避免 8.3 短路径问题）
- 每次启动自动执行：
  - 快捷方式存在性检查
  - 目标路径 / 参数一致性校验
- 异常时弹出三选一对话框：
  - 修复
  - 关闭自启
  - 忽略

> 已自动清理旧版 Registry Run 自启方式，避免冲突

---

### 9️⃣ 配置与日志

- 配置文件：  

  ```
  %APPDATA%\AutoShutdown\config.json
  ```

- 错误日志：  

  ```
  %APPDATA%\AutoShutdown\error.log
  ```

- 自动迁移旧版本配置文件

- 单实例运行（Win32 Mutex）

---

### 10️⃣ 托盘气泡通知
- 托盘右键菜单可切换“托盘气泡通知”（默认关闭）
- 开启后在关键节点弹出系统气泡，如联网提醒成功/失败、进入可取消休眠提示等
- 也可通过配置项 `tray_balloon_enabled` 控制

---

## 快速开始

### 方式一：直接运行（Release）

前往 GitHub Releases 下载已打包版本，安装后运行。

### 方式二：脚本运行（开发 / 调试）

1. Windows 10 / 11 + Python 3.10+

2. 同目录放置：

   - `auto_shutdown_win32_native.py`
   - `AutoShutdown.ico`（可选）

3. 启动：

   ```bash
   python auto_shutdown_win32_native.py
   ```

---

## 配置文件示例（config.json）

```json
{
  "pushplus_token": "",
  "pushplus_topic": "",
  "pushplus_api": "https://www.pushplus.plus/send",
  "remind_template": "{base_info}\n\n建议：确认是否需要关机或休眠。",
  "online_remind_times": 0,
  "uptime_hours": 2,
  "idle_minutes": 60,
  "pre_hibernate_countdown_sec": 60,
  "net_check_url": "https://baidu.com",
  "net_check_timeout_sec": 2,
  "tray_balloon_enabled": false,
  "autostart_enabled": false
}
```

---

## 打包指南（Nuitka，推荐）

```powershell
python -m nuitka `
  --standalone `
  --follow-imports `
  --windows-disable-console `
  --windows-icon-from-ico=AutoShutdown.ico `
  --include-data-file=AutoShutdown.ico=AutoShutdown.ico `
  --include-module=app `
  --include-module=auto_shutdown `
  --include-module=constants `
  --include-module=core `
  --include-module=ui `
  --include-module=winapi `
  auto_shutdown_win32_native.py
```

- 不弹控制台黑窗
- 子进程统一使用隐藏方式启动（`CREATE_NO_WINDOW`）

---

## 项目信息

- 项目名称：AutoShutdown
- 当前版本：0.0.4
- GitHub： https://github.com/xjhaz/AutoShutdown

---

## License

MIT License
