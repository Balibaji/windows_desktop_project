# Windows 窗口枚举与自动化操作（PowerShell + Python 两套实现）

查找桌面上打开的窗口（**包含但不限于标题**），并对它们执行输入、按键、点击、读写控件文本、关闭等操作。
两套实现功能对齐，都是**零第三方依赖**：

| 文件 | 说明 |
| --- | --- |
| `WindowOps.ps1` | PowerShell 版工具（.NET + Win32 API，Windows 自带 PowerShell 5.1 即可） |
| `window_ops.py` | Python 版工具（纯 `ctypes`，只用标准库，无需 pywin32 / uiautomation） |
| `demo.ps1` | PowerShell 版端到端演示 |
| `demo.py` | Python 版端到端演示 |

## 环境要求

- Windows 10 / 11
- PowerShell 5.1+（系统自带）
- Python 3.8+（`window_ops.py` 只用标准库）

## 快速开始

### 1. PowerShell 版

```powershell
# 端到端演示：枚举窗口 → 找窗口 → 激活 → 真实键盘输入 → 读写控件 → 关闭
powershell -ExecutionPolicy Bypass -File .\demo.ps1

# 列出所有可见窗口
.\WindowOps.ps1 list

# 按标题查找，输出 JSON
.\WindowOps.ps1 list -Pattern "记事本" -Json

# 查看窗口详细信息 + 子控件
.\WindowOps.ps1 info -Pattern "记事本"

# 激活窗口，并输入中文
.\WindowOps.ps1 type -Pattern "记事本" -Text "Hello，世界！"

# 发送组合键
.\WindowOps.ps1 keys -Pattern "记事本" -Keys "ctrl+a","ctrl+c"

# 在窗口内点击（坐标相对窗口左上角，不传则点中心）
.\WindowOps.ps1 click -Pattern "记事本" -X 200 -Y 150

# 直接读写控件文本（不需要窗口在前台）
.\WindowOps.ps1 settext -Pattern "记事本" -Text "直接写入" -ControlClass "Edit"
.\WindowOps.ps1 gettext -Pattern "记事本" -ControlClass "Edit"

# 关闭窗口
.\WindowOps.ps1 close -Pattern "记事本"
```

### 2. Python 版

```bash
# 端到端演示
python demo.py

# 列出所有可见窗口
python window_ops.py list

# 按标题查找，输出 JSON
python window_ops.py list --pattern 记事本 --json

# 查看窗口详细信息 + 子控件
python window_ops.py info --pattern 记事本

# 激活窗口，并输入中文
python window_ops.py type --pattern 记事本 --text "Hello，世界！"

# 发送组合键
python window_ops.py keys --pattern 记事本 --keys ctrl+a ctrl+c

# 在窗口内点击（坐标相对窗口左上角，不传则点中心）
python window_ops.py click --pattern 记事本 --x 200 --y 150

# 直接读写控件文本（不需要窗口在前台）
python window_ops.py settext --pattern 记事本 --text "直接写入" --control-class "Edit"
python window_ops.py gettext --pattern 记事本 --control-class "Edit"

# 关闭窗口
python window_ops.py close --pattern 记事本
```

## 命令 / 动作一览

| 动作 | PowerShell | Python | 作用 |
| --- | --- | --- | --- |
| 列出窗口 | `list` | `list` | 枚举顶层窗口 |
| 详细信息 | `info` | `info` | 窗口属性 + 子控件 |
| 子控件 | `children` | `children` | 只列子控件 |
| 激活 | `activate` | `activate` | 强制带到前台 |
| 文本输入 | `type` | `type` | `SendInput` 真实键盘事件，支持中文 / emoji |
| 组合键 | `keys` | `keys` | `ctrl+a`、`alt+f4`、`enter`、`f5` … |
| 鼠标点击 | `click` | `click` | 左键 / 右键 / 双击 |
| 读控件文本 | `gettext` | `gettext` | `WM_GETTEXT`，Scintilla 自动回退 `SCI_GETTEXT` |
| 写控件文本 | `settext` | `settext` | `WM_SETTEXT`，不需要前台焦点 |
| 关闭 | `close` | `close` | 发送 `WM_CLOSE` |
| 当前前台窗口 | `foreground` | `foreground` | 查看谁在前台 |

### 筛选条件

| PowerShell | Python | 说明 |
| --- | --- | --- |
| `-Pattern` | `--pattern` | 标题包含匹配（加 `-Regex` / `--regex` 用正则） |
| `-Class` | `--class` | 窗口类名，支持通配符 |
| `-ProcessName` | `--process` | 进程名，支持通配符 |
| `-Hwnd` | `--hwnd` | 直接按句柄（支持 `0x1A2B` 写法） |
| `-ProcessId` | `--pid` | 按 PID |
| `-Index` | `--index` | 匹配多个时取第几个（从 0 开始） |
| `-IncludeInvisible` | `--include-invisible` | 包含隐藏窗口 |
| `-Json` | `--json` | JSON 输出 |
| `-MessageMode` | `--message-mode` | `type` 改用 `PostMessage(WM_CHAR)`，不需要焦点 |
| `-ControlClass` | `--control-class` | `gettext` / `settext` 指定控件类名 |

## 能拿到的信息（不止标题）

每个窗口都会返回：

- **Handle**：窗口句柄
- **PID**：进程 ID
- **Process**：进程名
- **Class**：窗口类名（`Notepad`、`Chrome_WidgetWin_1`、`#32770`、`Edit` …）
- **Title**：标题
- **Visible / Minimized / Enabled**：可见 / 最小化 / 启用状态
- **Rect**：屏幕坐标与尺寸
- **Style / ExStyle / Owner**：窗口样式、扩展样式、属主窗口
- **子控件**：`info` / `children` 可以递归列出编辑框、按钮、静态文本等，并读取控件文本

## 工作原理（两套实现一致）

1. **枚举**：`EnumWindows` / `EnumChildWindows`
2. **属性**：`GetWindowTextW`、`GetClassNameW`、`GetWindowThreadProcessId`、
   `GetWindowRect`、`GetWindowLongPtrW`、`IsWindowVisible`、`IsIconic`
3. **激活**：`ShowWindow` + `SetWindowPos` + `AttachThreadInput` + `SetForegroundWindow`，
   失败再回退 `SwitchToThisWindow`（绕过 Windows 的前台锁定）
4. **键盘**：`SendInput` + `KEYEVENTF_UNICODE`，逐字符发送，中文、emoji、任意符号都能输入
5. **鼠标**：`SetCursorPos` + `mouse_event`
6. **控件文本**：`WM_SETTEXT` / `WM_GETTEXT`；Scintilla 控件回退 `SCI_GETTEXT`
7. **无焦点输入**：`PostMessage(WM_CHAR)`（`-MessageMode` / `--message-mode`）

## 注意事项

- **权限**：普通权限进程无法向管理员权限窗口发送输入（UIPI 隔离），需要以相同或更高权限运行脚本。
- **前台锁定**：`SendInput` 依赖真实焦点；脚本已用 `AttachThreadInput` 尝试绕过，但个别系统策略下仍可能失败。
  此时可改用消息模式（`-MessageMode` / `--message-mode`）直接给控件发消息。
- **受保护进程**（任务管理器、部分安全软件）通常拒绝被操作。
- **控制台窗口**（`cmd`、Windows Terminal）不是普通控件，`WM_GETTEXT` 读不到屏幕文本；
  演示里通过“输入命令 → 命令写文件 → 读文件”来验证键盘输入确实生效。
- **UWP / 打包应用**的顶层窗口可能由 `ApplicationFrameHost` 代理，窗口类名和子控件结构与传统 Win32 程序不同。
- 两套脚本启动时都调用了 `SetProcessDPIAware()`，保证 `GetWindowRect` 与鼠标坐标在缩放屏幕上一致。
- PowerShell 脚本保存为 **UTF-8 with BOM**，否则 Windows PowerShell 5.1 会把中文按 ANSI 读取。

## 演示做了什么

`demo.ps1` / `demo.py` 会自动：

1. 枚举当前所有可见窗口并打印前 8 个；
2. 启动一个控制台窗口（标题 `WinOps-PSDemo` / `WinOps-PyDemo`）并按标题找到它；
3. 打印窗口句柄、PID、进程、类名、位置尺寸、样式等；
4. 激活窗口，用 `SendInput` 输入命令并回车，通过结果文件验证输入生效；
5. 用 `Ctrl+L` 清屏、`Enter` 执行命令，演示组合键；
6. 启动“字符映射表”，枚举子控件，用 `WM_SETTEXT` 写入、`WM_GETTEXT` 读回；
7. 鼠标点击文本框 → `Ctrl+End` → 真实键盘追加文本；
8. 关闭所有演示窗口（加 `-KeepOpen` / `--keep-open` 可保留窗口观察）。

---

## RDP / 堡垒机场景

| 文件 | 说明 |
| --- | --- |
| `rdp_ops.py` | RDP 位图自动化 + 本地隧道感知（Python，纯标准库） |
| `RdpOps.ps1` | 同上的 PowerShell 版（mstsc + Win32 + WinRT OCR） |
| `AccessClient_RDP_Tunnel.md` | AccessClient / mstsc 双端口隧道原理与排障手册 |

### 隧道感知（AccessClient 场景）

堡垒机客户端常在本机 `127.0.0.1` 起代理端口，再以
`mstsc /v:127.0.0.1:<port>` 拉起远程桌面。两套脚本都内置了三个子命令：

```bash
python rdp_ops.py tunnel --app AccessClient        # 发现端口/属主/父进程
python rdp_ops.py wait-port --port 52341 --timeout 20
python rdp_ops.py health --any --process mstsc --screen
```

```powershell
.\RdpOps.ps1 tunnel -App "AccessClient*" -Watch
.\RdpOps.ps1 wait-port -Port 52341 -TimeoutSec 20
.\RdpOps.ps1 health -Any -ProcessName mstsc -Screen -Json
```

- `tunnel` 同时解析 mstsc 命令行 `/v:` 与该进程到回环地址的 `ESTABLISHED` 连接，
  所以用 `.rdp` 文件启动也能识别端口；
- `wait-port` 默认只检查 `LISTEN`，不会主动连 RDP 端口（`--connect` / `-Connect` 才会探测）；
- `health` 返回非零退出码，可直接用于流水线；`--screen` 用连续截图哈希判断画面是否卡死。
