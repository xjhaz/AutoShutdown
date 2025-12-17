# AutoShutdown
一个纯 Python（标准库）实现的 Windows 托盘常驻小工具：在满足“电脑已运行时间 + 空闲时间”阈值后，根据联网情况选择 **发送群组微信通知** 或 **弹出可取消的休眠倒计时**。
release中有打包好的安装包。

> 项目核心逻辑：定时检测 **开机运行时长**（GetTickCount64）与 **用户空闲时长**（GetLastInputInfo），满足阈值且超过冷却时间后触发通知/休眠流程。

---

## 主要功能

### 1) 托盘常驻 + 右键菜单

- 运行后不显示主窗口，托盘图标常驻；右键弹出菜单，双击打开“设置”。
- 菜单项（部分）：
  - **本次不提醒**（仅影响下一次触发）
  - **本次不关机**（仅影响下一次触发）
  - **开机自启**
  - **设置...**
  - **恢复默认（取消本次抑制）**
  - **退出**

### 2) 触发条件与冷却时间

- 触发条件同时满足：
  - 电脑运行时间 ≥ `uptime_hours`
  - 用户空闲时间 ≥ `idle_minutes`
  - 距上次触发 ≥ `cooldown_minutes`（防刷屏/反复休眠）fileciteturn4file0L30-L45

### 3) 联网时：发送群组微信通知（pushplus）

- 两级联网判断：
  1. WinAPI：`InternetGetConnectedState`
  2. DNS 解析 + 轻量 HTTP GET（可配置 URL 与超时）
- 在线且发送成功：弹托盘提示“已发送群组微信通知”。

### 4) 离线/发送失败：弹出可取消休眠倒计时

- 无网络或发送失败时，弹出置顶倒计时窗口（默认 60 秒，可配置），可“取消本次休眠”或“立即休眠”。
- 倒计时结束或点击“立即休眠”后调用 `shutdown /h` 进入休眠。

### 5) 设置窗口

- 支持配置阈值/周期：`uptime_hours`、`idle_minutes`、`check_interval_sec`、`cooldown_minutes`、`pre_hibernate_countdown_sec` 等。
- 支持“测试消息发送”“测试休眠”。
- 支持检查/开启休眠（`powercfg /a`、`powercfg /h on`），并把 `powercfg /a` 结果写入 `%APPDATA%\AutoShutdown\powercfg_a.txt` 便于排查。

### 6) 开机自启：Startup 目录快捷方式 + 完整性检查

- 使用 **Startup 文件夹 .lnk 快捷方式**实现自启，目标路径规则：
  - 打包 EXE：`程序当前路径 + 程序名称`（WinAPI 获取当前进程 exe 完整路径）
  - 脚本运行：`pythonw.exe "脚本路径"`fileciteturn1file3L1-L23
- 每次启动时做 **快捷方式存在性与路径一致性**检查：
  - 若缺失/被篡改，优先自动修复；自动修复失败再提示用户“修复/关闭/忽略”。

### 7) 配置与日志

- 配置文件：`%APPDATA%\AutoShutdown\config.json`（首次运行自动创建；支持从旧目录迁移）。
- 错误日志：`%APPDATA%\AutoShutdown\error.log`。
- 单实例：使用 Mutex `Local\AutoShutdownReminder_SingleInstance` 防止重复运行。

---

## 快速开始

### 运行脚本（开发/调试）

1. 准备环境：Windows 10/11 + Python 3.x（推荐 3.10+）。

2. 同目录放置：

   - `auto_shutdown_win32_native.py`
   - `AutoShutdown.ico`（可选，没有也能运行）

3. 启动：

   ```bash
   python auto_shutdown_win32_native.py
   ```

4. 运行后在系统托盘找到 **AutoShutdown** 图标，右键进入“设置”。

---

## 配置项说明（config.json）

默认配置如下（示意）：

```json
{
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
  "autostart_enabled": false
}
```

字段含义：

- `pushplus_token` / `pushplus_topic`：pushplus 群组推送参数（在线且推送成功时走通知路径）。
- `uptime_hours`：开机运行时长阈值（小时）
- `idle_minutes`：空闲阈值（分钟）
- `check_interval_sec`：检测周期（秒，最低 10 秒）
- `cooldown_minutes`：触发冷却时间（分钟）
- `pre_hibernate_countdown_sec`：离线/发送失败时，休眠倒计时（秒）
- `net_check_url` / `net_check_timeout_sec`：二级联网检测 URL 与超时（秒）
- `autostart_enabled`：是否启用开机自启（由托盘菜单控制保存）。

---

## 打包指南（Nuitka）

> 目标：生成可直接分发的 Windows GUI 程序（无控制台窗口），并把图标资源一起带上。

### 1) 安装 Nuitka

```bash
python -m pip install -U nuitka
```

### 2) 推荐打包命令（standalone）

在项目根目录执行（按你的文件名调整）：

```powershell
python -m nuitka .\auto_shutdown_win32_native.py `
  --standalone `
  --windows-console-mode=disable `
  --windows-icon-from-ico=AutoShutdown.ico `
  --include-data-files=AutoShutdown.ico=AutoShutdown.ico `
  --output-dir=diststand
```

说明：

- `--standalone`：输出包含运行依赖（适合直接分发）
- `--windows-console-mode=disable`：不弹控制台黑窗
- `--include-data-files`：把 `AutoShutdown.ico` 复制到输出目录，确保托盘图标正常加载

打包完成后，在 `diststand\auto_shutdown_win32_native.dist\` 目录中找到 exe 并运行。

### 3) “仍出现黑窗”的常见原因

- 不是主程序控制台，而是 **子进程**（例如 `powercfg`、`shutdown`、`powershell`）弹窗。
  - 本项目已对这些命令优先采用隐藏方式启动（子进程默认不应弹黑窗）。
- 若你的环境仍异常：优先查看 `%APPDATA%\AutoShutdown\error.log` 排查调用失败原因。

---

##（可选）制作安装包

如果你需要“选择安装目录 → 解压/复制文件 → 创建桌面快捷方式 → 安装完成后运行”，可考虑：

- Inno Setup
- NSIS

这些工具适合把 `*.dist` 整个目录作为安装内容进行打包。

---

## 常见问题（FAQ）

### 1) 为什么不休眠？

- 请先在“设置”里点击“休眠可用性检查（powercfg /a）”，确认 Hibernate 可用；必要时点击“开启休眠（powercfg /h on）”。

### 2) 开机自启为什么失效？

- 本项目使用 Startup 目录快捷方式，并在每次启动做完整性检查；若被删除/篡改会尝试自动修复，不行再提示你选择处理方式。

---

## TODO

- 自定义发送消息内容（目前通知文本为固定模板）。
- 添加“关于/版本信息”窗口（或菜单项）。
- 提供选项：在联网状态下 **通知发送成功后仍进入休眠**（当前逻辑是：在线且发送成功则仅通知，不休眠；离线/发送失败才休眠）。
