
## 先说结论

对本地电脑来说，RDP 窗口就是一张位图，外加一条输入转发通道：

- 本地的 **UI Automation、`WM_GETTEXT`、`WM_SETTEXT`、`PostMessage(WM_CHAR)` 全部无效** —— 远端控件根本不在本机窗口树里。
- 能用的只有**真实输入**：`SendInput` 发键盘、`SetCursorPos`+`mouse_event` 点鼠标。mstsc 会把这些输入通过 RDP 输入通道转发到远端。
- 想“看到结果”，只能**截图 + OCR/模板匹配/视觉模型**。

所以有两条路线，通常优先第二条：

|路线|适用|稳定性|
|---|---|---|
|A. 模拟真人键鼠|远端程序只有图形界面、必须点按钮|一般，受焦点/分辨率/输入法影响|
|B. 直接远程执行命令|只是要跑 cmd/PowerShell/脚本|高，推荐|

---

## 路线 A：用键鼠“操作”RDP 窗口

你项目里的 `window_ops.py` 正好能做，关键点：

### 1. 找到 RDP 窗口

```
# 进程是 mstsc，顶层窗口类名通常是 TscShellContainerClass
python window_ops.py list --class "TscShell*" --json
python window_ops.py info --class TscShellContainerClass
```

### 2. 打开远端 cmd

```
python window_ops.py activate --class TscShellContainerClass
python window_ops.py keys     --class TscShellContainerClass --keys win+r
python window_ops.py type     --class TscShellContainerClass --text "cmd"
python window_ops.py keys     --class TscShellContainerClass --keys enter
```

> 必须用 `type`/`keys`（`SendInput`），不能用 `settext`/`--message-mode`。后者是往本地窗口发消息，mstsc 不会把它翻译成 RDP 输入。

### 3. 在远端执行 PowerShell

短命令直接输入：

```
python window_ops.py keys --class TscShellContainerClass --keys win+r
python window_ops.py type --class TscShellContainerClass --text 'powershell -NoExit -Command "Get-Process | Sort-Object CPU -Descending | Select-Object -First 5"'
python window_ops.py keys --class TscShellContainerClass --keys enter
```

长脚本/中文用**剪贴板粘贴**最稳（需要 mstsc 勾选“本地资源 → 剪贴板”）：

```
$script = @'
$ErrorActionPreference = 'Stop'
Get-Date
Get-Service | Where-Object Status -eq 'Running' | Select-Object -First 10
'@
Set-Clipboard -Value $script

python window_ops.py activate --class TscShellContainerClass
python window_ops.py keys     --class TscShellContainerClass --keys ctrl+v
python window_ops.py keys     --class TscShellContainerClass --keys enter
```

更省事的做法：**开启驱动器重定向**（mstsc → 本地资源 → 详细信息 → 驱动器），远端可以直接运行你本机的脚本：

```
python window_ops.py type --class TscShellContainerClass --text 'powershell -ExecutionPolicy Bypass -File \\tsclient\C\Users\Admin\script.ps1'
```

### 4. 怎么知道执行成功

`SendInput` 是“盲发”，屏幕上的变化不会自动变成返回值。三种反馈方式：

1. **让远端把结果写到文件**，本机读重定向目录/共享目录；
2. **让远端 `Set-Clipboard`**，本机 `Get-Clipboard` 读取；
3. **截图 + OCR**：截取 mstsc 客户区，用 Windows 自带 `Windows.Media.Ocr` 或 Tesseract 识别，定位到文字坐标后再点。

截图要点：

- 不要用 `PrintWindow`，mstsc 走 DirectX/位图渲染，容易全黑；用屏幕区域截取（`Graphics.CopyFromScreen` / Desktop Duplication / PIL `ImageGrab`）。
- 固定窗口大小和远端分辨率，关掉“智能缩放”，坐标才可复现。
- `window_ops.py click --x --y` 是**相对窗口外框**的，RDP 客户区还要加上标题栏高度；如果画面被缩放，远端坐标 = 本地偏移 × (远端宽 / 客户区宽)。

### 5. 必须知道的坑

- **焦点**：`SendInput` 发给当前前台窗口，自动化期间不能有人动鼠标键盘。
- **权限**：mstsc 以管理员运行时，脚本也要管理员，否则 UIPI 拦截。
- **本机锁屏**：本机锁屏后 `SendInput` 失效；远端 RDP 断开后远端会话也会锁屏。
- **Ctrl+Alt+Del** 被本机截获，要用 **Ctrl+Alt+End**。
- **Win 键/Alt+Tab** 在窗口模式下行为取决于 mstsc 的“应用 Windows 键组合”设置；全屏时通常发给远端。
- **输入法**：中文容易出问题，优先剪贴板粘贴，不要在远端敲中文。
- 远端 RDP 会话断开后想继续跑 GUI 自动化，需要保持会话解锁，例如在远端执行 `tscon <会话ID> /dest:console` 把会话切到物理控制台，否则键鼠打不进去。

---

## 路线 B：绕开位图，直接远程执行（推荐）

只要目的不是“点某个图形按钮”，就不要走 RDP 窗口。

### WinRM / PowerShell Remoting（最自然）

远端先启用一次：

```
Enable-PSRemoting -Force
```

本机执行：

```
$cred = Get-Credential
Invoke-Command -ComputerName 192.168.1.50 -Credential $cred -ScriptBlock {
    whoami
    Get-Date
    & 'C:\temp\job.ps1'
}
```

交互式：`Enter-PSSession -ComputerName 192.168.1.50 -Credential $cred`

### 计划任务（走 SMB/RPC，不依赖 WinRM 端口）

```
schtasks /create /s 192.168.1.50 /u DOMAIN\user /p 密码 /tn RdpJob `
  /tr "powershell -ExecutionPolicy Bypass -File C:\temp\job.ps1" /sc once /st 23:59 /f
schtasks /run /s 192.168.1.50 /u DOMAIN\user /p 密码 /tn RdpJob
```

### PsExec / OpenSSH / WMI

```
# PsExec：-i 指定 RDP 会话号（query session 查看），-s 用 SYSTEM
psexec \\192.168.1.50 -u user -p 密码 -i 2 -s cmd /c "powershell -File C:\temp\job.ps1"

# OpenSSH
ssh user@192.168.1.50 "powershell -Command Get-Date"

# WMI 起进程
Invoke-CimMethod -ComputerName 192.168.1.50 -ClassName Win32_Process `
  -MethodName Create -Arguments @{ CommandLine = 'powershell -File C:\temp\job.ps1' }
```

### 重要区别

- WinRM / WMI / PsExec（不带 `-i`）跑在**非交互会话**：cmd、PowerShell 脚本没问题，但 GUI 程序不会出现在 RDP 桌面上。
- 需要 GUI 出现在当前 RDP 桌面时：用**计划任务勾选“只在用户登录时运行”**，或 `psexec -i <RDP会话号>`。
- 双跳（本机 → 远端 → 第三台）才需要 CredSSP，普通场景不用。

---

## 选型建议

- 只是“打开 cmd、跑 PowerShell 脚本、拿结果” → **WinRM / 计划任务 / SSH**，不要碰画面。
- 远端程序只有窗口、没有命令行接口 → **SendInput 键鼠 + 截图 OCR**，并尽量把操作收敛成“快捷键 + 剪贴板 + 固定坐标”。
- 长期无人值守的 GUI 自动化 → 在**远端**装 agent（Python + pyautogui/截图），本机通过 HTTP/socket 下指令，比在本机截 mstsc 稳得多；同时用 `tscon` 保持会话不锁屏。

---

如果你愿意，我可以直接在这个项目里加一个 `rdp_ops.py`，封装：

- `find_rdp_window()`（按 `mstsc` / `TscShellContainerClass` 查找）
- `open_cmd()`、`run_powershell(cmd)`（Win+R + 剪贴板粘贴）
- `screenshot_client()`、`ocr_text()`、`click_text("确定")`
- `wait_for_text(...)` 做执行结果校验
