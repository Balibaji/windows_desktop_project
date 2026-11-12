#Requires -Version 5.1
<#
.SYNOPSIS
    RDP / 位图窗口自动化工具（PowerShell 版，纯 Win32 + .NET + WinRT OCR，无第三方依赖）

.DESCRIPTION
    远程桌面在本地只是一张位图，本地的 UI Automation / WM_GETTEXT / WM_SETTEXT
    看不到远端控件。本脚本只使用两条通道：
      * SendInput 真实键鼠输入（由 mstsc 转发到远端）
      * 截取 mstsc 客户区 + Windows.Media.Ocr 识别文字

    动作：
      find         查找 RDP 窗口（默认 mstsc / msrdc / TscShell*，也支持任意窗口）
      status       窗口 + 客户区在屏幕上的位置、尺寸、DPI
      open-cmd     激活 RDP 窗口 → Win+R → 输入 cmd → 回车
      run-ps       在远端执行 PowerShell（短脚本 EncodedCommand，长脚本剪贴板粘贴）
      type         向远端当前焦点窗口输入文本
      keys         向远端发送组合键
      click        点击 RDP 位图区域（默认客户区坐标）
      screenshot   截取客户区为 PNG / BMP
      ocr          对图片做 OCR
      click-text   截图 → OCR → 找文字 → 点击文字中心
      wait-text    轮询 OCR，直到文字出现或超时
      clip-set / clip-get / paste   剪贴板操作（RDP 默认启用剪贴板重定向）
      tunnel       发现 mstsc → 127.0.0.1:<port> 的本地端口隧道（AccessClient 场景）
      wait-port    等待本机端口进入 LISTEN（隧道就绪检查）
      health       窗口 + 隧道的综合健康检查

.PARAMETER Action
    find | status | open-cmd | run-ps | type | keys | click | screenshot |
    ocr | click-text | wait-text | clip-set | clip-get | paste |
    tunnel | wait-port | health

.PARAMETER Pattern
    按标题筛选（默认不区分大小写的「包含」匹配，加 -Regex 则按正则）。

.PARAMETER Class
    按窗口类名筛选，支持通配符，例如 -Class "TscShell*"。

.PARAMETER ProcessName
    按进程名筛选，支持通配符，例如 -ProcessName "mstsc"。

.PARAMETER Hwnd
    直接按窗口句柄定位，支持十进制或 0x 十六进制。

.PARAMETER ProcessId
    按 PID 定位。

.PARAMETER Any
    不限定 RDP 进程，允许匹配任意窗口（本地测试截图 / OCR 用）。

.PARAMETER Index
    匹配到多个窗口时取第几个（从 0 开始）。

.PARAMETER Text
    type / click-text / wait-text / clip-set / paste 使用的文本。

.PARAMETER Keys
    keys 动作的组合键，例如 "win+r","ctrl+v","enter"。

.PARAMETER X / Y
    click 的坐标。默认相对客户区（位图区域）左上角；-WindowRelative 改为窗口外框。

.PARAMETER Script
    run-ps 直接提供的 PowerShell 脚本内容。

.PARAMETER ScriptFile
    run-ps 读取本机脚本文件内容，再粘贴到远端执行。

.PARAMETER RemoteFile
    run-ps 直接在远端执行文件（可配合 \\tsclient 重定向驱动器）。

.PARAMETER Encoded
    run-ps 对短脚本使用 -EncodedCommand（不依赖剪贴板，受「运行」对话框长度限制）。

.PARAMETER ExitAfter
    脚本执行完后退出远端 PowerShell（默认保留窗口看输出）。

.PARAMETER NoActivate
    截图 / 输入前不自动激活窗口。

.PARAMETER Out
    screenshot / ocr / click-text 的截图输出路径。

.PARAMETER Image
    ocr 直接识别已有图片，不截图。

.PARAMETER Language
    OCR 语言，例如 zh-CN / en-US。默认使用用户配置语言。

.PARAMETER TimeoutSec / IntervalMs
    wait-text / tunnel / wait-port / health 的超时时间与轮询间隔。

.PARAMETER DryRun
    click-text 只输出坐标，不真正点击。

.PARAMETER App
    tunnel 按父进程 / 监听进程名过滤，例如 -App "AccessClient*"。

.PARAMETER Port
    tunnel 只显示指定本地端口；wait-port 指定要等待的端口；
    health 手工指定隧道端口（默认从 mstsc 命令行 / 连接推断）。

.PARAMETER Watch
    tunnel 持续监控端口状态，变化时输出（Ctrl+C 退出）。

.PARAMETER Launch
    tunnel 对已声明但尚未连接的端口拉起 mstsc /v:127.0.0.1:<port>。

.PARAMETER Screen
    health 连续截图比较，用于检测画面是否卡死。

.PARAMETER Connect
    wait-port 额外做一次 TCP 连接探测（会向 RDP 端口发起握手）。

.PARAMETER Json
    以 JSON 输出（find / status / screenshot / ocr / click-text / wait-text /
    clip-get / tunnel / wait-port / health）。

.EXAMPLE
    .\RdpOps.ps1 find
.EXAMPLE
    .\RdpOps.ps1 status -Json
.EXAMPLE
    .\RdpOps.ps1 tunnel
.EXAMPLE
    .\RdpOps.ps1 tunnel -Watch -App "AccessClient*"
.EXAMPLE
    .\RdpOps.ps1 wait-port -Port 52341 -TimeoutSec 20
.EXAMPLE
    .\RdpOps.ps1 health -Any -ProcessName mstsc -Screen -Json
.EXAMPLE
    # 本地测试（不需要 RDP）
    .\RdpOps.ps1 screenshot -Any -ProcessName notepad -Out notepad.png
    .\RdpOps.ps1 ocr -Image notepad.png -Json
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('find', 'status', 'open-cmd', 'run-ps', 'type', 'keys', 'click',
                 'screenshot', 'ocr', 'click-text', 'wait-text',
                 'clip-set', 'clip-get', 'paste',
                 'tunnel', 'wait-port', 'health')]
    [string]$Action = 'find',

    [string]$Pattern,
    [string]$Class,
    [string]$ProcessName,
    [string]$Hwnd = '',
    [int]$ProcessId = 0,
    [switch]$Regex,
    [switch]$IncludeInvisible,
    [switch]$Any,
    [int]$Index = 0,

    [string]$Text,
    [string[]]$Keys,
    [int]$X = -1,
    [int]$Y = -1,
    [switch]$WindowRelative,
    [switch]$DoubleClick,
    [switch]$RightClick,

    [string]$Script,
    [string]$ScriptFile,
    [string]$RemoteFile,
    [switch]$Encoded,
    [switch]$ExitAfter,
    [int]$WaitMs = 900,

    [string]$Out,
    [string]$Image,
    [string]$Language,
    [switch]$NoActivate,
    [switch]$Exact,
    [int]$MatchIndex = 0,
    [int]$TimeoutSec = 15,
    [int]$IntervalMs = 800,
    [switch]$DryRun,

    [string]$App,
    [int]$Port = 0,
    [switch]$Watch,
    [switch]$Launch,
    [switch]$Screen,
    [switch]$Connect,

    [switch]$Json
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    $OutputEncoding = [Text.Encoding]::UTF8
} catch { }

# ---------------------------------------------------------------------------
# Win32 / .NET 封装（C# 5 语法，兼容 Windows PowerShell 5.1）
# ---------------------------------------------------------------------------
if (-not ('RdpNative' -as [type])) {
    Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public class RdpWin
{
    public IntPtr Handle;
    public string Title = "";
    public string ClassName = "";
    public int Pid;
    public string ProcessName = "";
    public int Left;
    public int Top;
    public int Width;
    public int Height;
    public bool Visible;
    public bool Minimized;
}

public static class RdpNative
{
    public const uint WM_CLOSE = 0x0010;

    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct HARDWAREINPUT
    {
        public uint uMsg;
        public ushort wParamL;
        public ushort wParamH;
    }

    [StructLayout(LayoutKind.Explicit)]
    struct INPUTUNION
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
        [FieldOffset(0)] public KEYBDINPUT ki;
        [FieldOffset(0)] public HARDWAREINPUT hi;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct INPUT
    {
        public uint type;
        public INPUTUNION u;
    }

    const uint INPUT_KEYBOARD        = 1;
    const uint KEYEVENTF_KEYUP       = 0x0002;
    const uint KEYEVENTF_UNICODE     = 0x0004;
    const uint MOUSEEVENTF_LEFTDOWN  = 0x0002;
    const uint MOUSEEVENTF_LEFTUP    = 0x0004;
    const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
    const uint MOUSEEVENTF_RIGHTUP   = 0x0010;

    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h, ref POINT p);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);
    [DllImport("user32.dll")] static extern void SwitchToThisWindow(IntPtr h, bool altTab);
    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags, int dx, int dy, uint data, IntPtr extraInfo);
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint count, INPUT[] inputs, int size);
    [DllImport("user32.dll")] static extern short VkKeyScan(char ch);

    static Dictionary<int, string> _procCache = new Dictionary<int, string>();

    static RdpNative()
    {
        try { SetProcessDPIAware(); } catch { }
    }

    static string ProcName(int pid)
    {
        lock (_procCache)
        {
            string name;
            if (_procCache.TryGetValue(pid, out name)) return name;
            try { name = Process.GetProcessById(pid).ProcessName; }
            catch { name = ""; }
            _procCache[pid] = name;
            return name;
        }
    }

    static RdpWin Info(IntPtr h)
    {
        StringBuilder title = new StringBuilder(1024);
        GetWindowText(h, title, title.Capacity);
        StringBuilder cls = new StringBuilder(256);
        GetClassName(h, cls, cls.Capacity);

        uint pid;
        GetWindowThreadProcessId(h, out pid);

        RECT r;
        GetWindowRect(h, out r);

        RdpWin w = new RdpWin();
        w.Handle = h;
        w.Title = title.ToString();
        w.ClassName = cls.ToString();
        w.Pid = (int)pid;
        w.ProcessName = ProcName((int)pid);
        w.Left = r.Left;
        w.Top = r.Top;
        w.Width = r.Right - r.Left;
        w.Height = r.Bottom - r.Top;
        w.Visible = IsWindowVisible(h);
        w.Minimized = IsIconic(h);
        return w;
    }

    public static List<RdpWin> GetWindows(bool visibleOnly)
    {
        List<RdpWin> list = new List<RdpWin>();
        EnumWindows(delegate(IntPtr h, IntPtr l)
        {
            if (!visibleOnly || IsWindowVisible(h)) list.Add(Info(h));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    public static RdpWin GetForegroundInfo()
    {
        return Info(GetForegroundWindow());
    }

    public static IntPtr GetForegroundHandle()
    {
        return GetForegroundWindow();
    }

    public static bool IsWindowHandle(IntPtr h)
    {
        return IsWindow(h);
    }

    public static bool ForceForeground(IntPtr h)
    {
        if (!IsWindow(h)) return false;
        if (IsIconic(h)) ShowWindow(h, 9);  // SW_RESTORE
        ShowWindow(h, 5);                   // SW_SHOW
        SetWindowPos(h, IntPtr.Zero, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040);

        uint fgThread;
        GetWindowThreadProcessId(GetForegroundWindow(), out fgThread);
        uint cur = GetCurrentThreadId();
        bool attached = (fgThread != cur) && AttachThreadInput(cur, fgThread, true);

        bool ok = SetForegroundWindow(h);
        if (!ok)
        {
            SwitchToThisWindow(h, true);
            ok = (GetForegroundWindow() == h);
        }
        if (attached) AttachThreadInput(cur, fgThread, false);
        Thread.Sleep(80);
        return ok;
    }

    // 返回客户区在屏幕上的 { left, top, width, height }
    public static int[] ClientRectScreen(IntPtr h)
    {
        RECT r;
        GetClientRect(h, out r);
        POINT p;
        p.X = 0;
        p.Y = 0;
        ClientToScreen(h, ref p);
        return new int[] { p.X, p.Y, r.Right - r.Left, r.Bottom - r.Top };
    }

    public static double DpiScale(IntPtr h)
    {
        try { return GetDpiForWindow(h) / 96.0; }
        catch { return 1.0; }
    }

    public static string CaptureClient(IntPtr h, string path)
    {
        int[] cr = ClientRectScreen(h);
        int w = cr[2], hgt = cr[3];
        if (w <= 0 || hgt <= 0) throw new ArgumentException("客户区尺寸无效: " + w + "x" + hgt);

        using (Bitmap bmp = new Bitmap(w, hgt, PixelFormat.Format24bppRgb))
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.CopyFromScreen(cr[0], cr[1], 0, 0, new Size(w, hgt), CopyPixelOperation.SourceCopy);
            if (path.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase))
                bmp.Save(path, ImageFormat.Bmp);
            else
                bmp.Save(path, ImageFormat.Png);
        }
        return path;
    }

    static INPUT UniInput(char ch, bool up)
    {
        INPUT i = new INPUT();
        i.type = INPUT_KEYBOARD;
        i.u.ki.wVk = 0;
        i.u.ki.wScan = (ushort)ch;
        i.u.ki.dwFlags = KEYEVENTF_UNICODE | (up ? KEYEVENTF_KEYUP : 0);
        return i;
    }

    static INPUT VkInput(ushort vk, bool up)
    {
        INPUT i = new INPUT();
        i.type = INPUT_KEYBOARD;
        i.u.ki.wVk = vk;
        i.u.ki.wScan = 0;
        i.u.ki.dwFlags = up ? KEYEVENTF_KEYUP : 0;
        return i;
    }

    static void Flush(List<INPUT> batch)
    {
        if (batch.Count == 0) return;
        SendInput((uint)batch.Count, batch.ToArray(), Marshal.SizeOf(typeof(INPUT)));
        batch.Clear();
        Thread.Sleep(12);
    }

    public static void TypeText(string text)
    {
        if (text == null) return;
        List<INPUT> batch = new List<INPUT>();
        foreach (char ch in text)
        {
            if (ch == '\r') continue;
            if (ch == '\n')
            {
                batch.Add(VkInput(0x0D, false));
                batch.Add(VkInput(0x0D, true));
            }
            else if (ch == '\t')
            {
                batch.Add(VkInput(0x09, false));
                batch.Add(VkInput(0x09, true));
            }
            else
            {
                batch.Add(UniInput(ch, false));
                batch.Add(UniInput(ch, true));
            }
            if (batch.Count >= 64) Flush(batch);
        }
        Flush(batch);
    }

    static ushort NameToVk(string name)
    {
        switch (name.ToLowerInvariant())
        {
            case "backspace": case "bs": return 0x08;
            case "tab": return 0x09;
            case "enter": case "return": return 0x0D;
            case "shift": return 0x10;
            case "ctrl": case "control": return 0x11;
            case "alt": return 0x12;
            case "pause": return 0x13;
            case "capslock": return 0x14;
            case "esc": case "escape": return 0x1B;
            case "space": return 0x20;
            case "pageup": case "pgup": return 0x21;
            case "pagedown": case "pgdn": return 0x22;
            case "end": return 0x23;
            case "home": return 0x24;
            case "left": return 0x25;
            case "up": return 0x26;
            case "right": return 0x27;
            case "down": return 0x28;
            case "printscreen": case "prtsc": return 0x2C;
            case "insert": case "ins": return 0x2D;
            case "delete": case "del": return 0x2E;
            case "win": case "lwin": return 0x5B;
            case "apps": return 0x5D;
            case "multiply": return 0x6A;
            case "add": return 0x6B;
            case "subtract": return 0x6D;
            case "decimal": return 0x6E;
            case "divide": return 0x6F;
            default:
                if (name.Length >= 2 && (name[0] == 'f' || name[0] == 'F'))
                {
                    int n;
                    if (int.TryParse(name.Substring(1), out n) && n >= 1 && n <= 24)
                        return (ushort)(0x70 + n - 1);
                }
                return 0;
        }
    }

    public static void SendCombo(string combo)
    {
        if (string.IsNullOrEmpty(combo)) return;
        string[] parts = combo.Split('+');
        bool ctrl = false, alt = false, shift = false, win = false;
        for (int i = 0; i < parts.Length - 1; i++)
        {
            string m = parts[i].Trim().ToLowerInvariant();
            if (m == "ctrl" || m == "control") ctrl = true;
            else if (m == "alt") alt = true;
            else if (m == "shift") shift = true;
            else if (m == "win" || m == "lwin") win = true;
        }
        PressKey(parts[parts.Length - 1].Trim(), ctrl, alt, shift, win);
    }

    public static void PressKey(string key, bool ctrl, bool alt, bool shift, bool win)
    {
        if (string.IsNullOrEmpty(key)) return;

        ushort vk = 0;
        bool needShift = false;

        if (key.Length == 1)
        {
            if (!ctrl && !alt && !win)
            {
                TypeText(key);
                return;
            }
            short scan = VkKeyScan(key[0]);
            if (scan != -1)
            {
                vk = (ushort)(scan & 0xFF);
                needShift = (scan & 0x100) != 0;
            }
        }
        else
        {
            vk = NameToVk(key);
        }

        if (vk == 0) return;

        List<INPUT> seq = new List<INPUT>();
        if (ctrl) seq.Add(VkInput(0x11, false));
        if (alt) seq.Add(VkInput(0x12, false));
        if (shift || needShift) seq.Add(VkInput(0x10, false));
        if (win) seq.Add(VkInput(0x5B, false));

        seq.Add(VkInput(vk, false));
        seq.Add(VkInput(vk, true));

        if (win) seq.Add(VkInput(0x5B, true));
        if (shift || needShift) seq.Add(VkInput(0x10, true));
        if (alt) seq.Add(VkInput(0x12, true));
        if (ctrl) seq.Add(VkInput(0x11, true));

        SendInput((uint)seq.Count, seq.ToArray(), Marshal.SizeOf(typeof(INPUT)));
        Thread.Sleep(30);
    }

    public static void ClickAt(int x, int y, bool doubleClick, bool rightButton)
    {
        SetCursorPos(x, y);
        Thread.Sleep(40);
        uint down = rightButton ? MOUSEEVENTF_RIGHTDOWN : MOUSEEVENTF_LEFTDOWN;
        uint up = rightButton ? MOUSEEVENTF_RIGHTUP : MOUSEEVENTF_LEFTUP;
        mouse_event(down, 0, 0, 0, IntPtr.Zero);
        mouse_event(up, 0, 0, 0, IntPtr.Zero);
        if (doubleClick)
        {
            Thread.Sleep(70);
            mouse_event(down, 0, 0, 0, IntPtr.Zero);
            mouse_event(up, 0, 0, 0, IntPtr.Zero);
        }
        Thread.Sleep(50);
    }
}
'@
}

# ---------------------------------------------------------------------------
# 目标窗口
# ---------------------------------------------------------------------------
$HwndValue = 0L
if ($Hwnd) {
    if ($Hwnd -match '^0[xX]') { $HwndValue = [Convert]::ToInt64($Hwnd.Substring(2), 16) }
    else { $HwndValue = [long]$Hwnd }
}

function Test-RdpWindow {
    param($Window)
    $proc = $Window.ProcessName.ToLowerInvariant()
    $cls = $Window.ClassName.ToLowerInvariant()
    return ($proc -eq 'mstsc' -or $proc -eq 'msrdc' -or
            $cls.StartsWith('tscshell') -or $cls.StartsWith('rail'))
}

function Get-RdpTargets {
    param(
        [string]$TitlePattern,
        [string]$ClassName,
        [string]$ProcName,
        [long]$WindowHandle = 0,
        [int]$PidValue = 0,
        [switch]$UseRegex,
        [switch]$IncludeHidden,
        [switch]$AllowAny
    )

    $result = @([RdpNative]::GetWindows(-not $IncludeHidden.IsPresent))
    $hasFilter = ($TitlePattern -or $ClassName -or $ProcName -or
                  ($WindowHandle -ne 0) -or ($PidValue -ne 0))

    if ($WindowHandle -ne 0) { $result = @($result | Where-Object { $_.Handle.ToInt64() -eq $WindowHandle }) }
    if ($PidValue -ne 0)     { $result = @($result | Where-Object { $_.Pid -eq $PidValue }) }
    if ($ProcName)           { $result = @($result | Where-Object { $_.ProcessName -like $ProcName }) }
    if ($ClassName)          { $result = @($result | Where-Object { $_.ClassName -like $ClassName }) }
    if ($TitlePattern) {
        if ($UseRegex) {
            $result = @($result | Where-Object { $_.Title -match $TitlePattern })
        } else {
            $result = @($result | Where-Object {
                $_.Title.IndexOf($TitlePattern, [StringComparison]::OrdinalIgnoreCase) -ge 0
            })
        }
    }

    if (-not $hasFilter -and -not $AllowAny) {
        $result = @($result | Where-Object { Test-RdpWindow $_ })
    }

    return @($result | Sort-Object Pid, Title)
}

function Get-RdpTarget {
    param(
        [string]$TitlePattern,
        [string]$ClassName,
        [string]$ProcName,
        [long]$WindowHandle = 0,
        [int]$PidValue = 0,
        [switch]$UseRegex,
        [switch]$IncludeHidden,
        [switch]$AllowAny,
        [int]$Pick = 0
    )

    $list = @(Get-RdpTargets -TitlePattern $TitlePattern -ClassName $ClassName `
                             -ProcName $ProcName -WindowHandle $WindowHandle `
                             -PidValue $PidValue -UseRegex:$UseRegex `
                             -IncludeHidden:$IncludeHidden -AllowAny:$AllowAny)
    if ($list.Count -eq 0) {
        throw ("没有找到目标窗口。`n" +
               "  * 连上 RDP 后直接运行（默认找 mstsc / msrdc / TscShell*）；`n" +
               "  * 或用 -Any -Pattern/-Class/-ProcessName/-Hwnd 指定任意窗口。")
    }
    if ($Pick -lt 0 -or $Pick -ge $list.Count) {
        throw "索引 $Pick 越界，共找到 $($list.Count) 个窗口"
    }
    return $list[$Pick]
}

function ConvertTo-RdpWindow {
    param([Parameter(Mandatory)]$Window)

    [pscustomobject][ordered]@{
        Handle      = ('0x{0:X8}' -f $Window.Handle.ToInt64())
        PID         = $Window.Pid
        Process     = $Window.ProcessName
        Class       = $Window.ClassName
        Visible     = $Window.Visible
        Minimized   = $Window.Minimized
        Rect        = ('{0},{1} {2}x{3}' -f $Window.Left, $Window.Top, $Window.Width, $Window.Height)
        Title       = $Window.Title
        IsRdp       = (Test-RdpWindow $Window)
    }
}

function Get-ClientGeometry {
    param([Parameter(Mandatory)]$Window)
    $cr = [RdpNative]::ClientRectScreen($Window.Handle)
    [pscustomobject][ordered]@{
        Left   = $cr[0]
        Top    = $cr[1]
        Width  = $cr[2]
        Height = $cr[3]
    }
}

# ---------------------------------------------------------------------------
# 键鼠 / 窗口操作
# ---------------------------------------------------------------------------
function Assert-RunDialogRemote {
    param([Parameter(Mandatory)]$Target)

    $fg = [RdpNative]::GetForegroundInfo()
    if ($fg.Handle.ToInt64() -eq $Target.Handle.ToInt64()) { return }

    if ($fg.ClassName -eq '#32770' -or $fg.Title -match '(?i)run|运行') {
        [RdpNative]::SendCombo('esc')
        throw ("Win+R 被本机截获，没有发到远端。`n" +
               "请把 mstsc 的「本地资源 → 键盘 → 应用 Windows 键组合」改成「在远程计算机上」，" +
               "或使用全屏连接。")
    }
    throw "Win+R 后前台窗口变成了 [$($fg.Title)]，按键没有送到 RDP 会话"
}

function Invoke-RemoteCommand {
    param(
        [Parameter(Mandatory)]$Target,
        [Parameter(Mandatory)][string]$Command,
        [int]$Wait = 900
    )

    [void][RdpNative]::ForceForeground($Target.Handle)
    Start-Sleep -Milliseconds 150
    [RdpNative]::SendCombo('win+r')
    Start-Sleep -Milliseconds 700
    Assert-RunDialogRemote $Target
    [RdpNative]::SendCombo('ctrl+a')
    [RdpNative]::TypeText($Command)
    [RdpNative]::SendCombo('enter')
    Start-Sleep -Milliseconds $Wait
}

function Invoke-RemoteScriptViaClipboard {
    param(
        [Parameter(Mandatory)]$Target,
        [Parameter(Mandatory)][string]$Body,
        [int]$Wait = 1200,
        [switch]$CloseAfter
    )

    [void][RdpNative]::ForceForeground($Target.Handle)
    Start-Sleep -Milliseconds 150
    [RdpNative]::SendCombo('win+r')
    Start-Sleep -Milliseconds 700
    Assert-RunDialogRemote $Target
    [RdpNative]::SendCombo('ctrl+a')
    [RdpNative]::TypeText('powershell -NoProfile -NoExit')
    [RdpNative]::SendCombo('enter')
    Start-Sleep -Milliseconds 1500

    if ($CloseAfter) { $Body = $Body + "`nexit" }
    Set-Clipboard -Value $Body
    Start-Sleep -Milliseconds 200
    [RdpNative]::SendCombo('ctrl+v')
    Start-Sleep -Milliseconds 400
    [RdpNative]::SendCombo('enter')
    Start-Sleep -Milliseconds $Wait
}

function New-ScreenshotPath {
    Join-Path $env:TEMP ('rdp_shot_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.png')
}

function Get-ClientScreenshot {
    param(
        [Parameter(Mandatory)]$Target,
        [string]$Path,
        [switch]$DoNotActivate
    )

    if (-not $Path) { $Path = New-ScreenshotPath }
    if (-not $DoNotActivate) {
        [void][RdpNative]::ForceForeground($Target.Handle)
        Start-Sleep -Milliseconds 250
    }
    [void][RdpNative]::CaptureClient($Target.Handle, $Path)
    return $Path
}

# ---------------------------------------------------------------------------
# OCR（Windows.Media.Ocr）
# ---------------------------------------------------------------------------
function Wait-WinRtOperation {
    param(
        [Parameter(Mandatory)]$Operation,
        [Parameter(Mandatory)][Type]$ResultType
    )

    $asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    } | Select-Object -First 1

    if (-not $asTask) { throw '当前系统不支持 Windows Runtime 异步调用（AsTask 缺失）' }
    $task = $asTask.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
    [void]$task.Wait(-1)
    return $task.Result
}

function Invoke-Ocr {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$OcrLanguage
    )

    $full = [IO.Path]::GetFullPath($Path)
    if (-not (Test-Path -LiteralPath $full)) { throw "图片不存在: $full" }

    try { Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop } catch { }

    [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
    [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
    [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null
    [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType = WindowsRuntime] | Out-Null

    if ($OcrLanguage) {
        [Windows.Globalization.Language, Windows.Globalization, ContentType = WindowsRuntime] | Out-Null
        $lang = New-Object -TypeName Windows.Globalization.Language -ArgumentList $OcrLanguage
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
    } else {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    }

    if ($null -eq $engine) {
        throw ('OCR 引擎不可用。请在「设置 → 时间和语言 → 语言」中安装对应语言的 OCR 组件，' +
               '或使用 -Language 指定已安装的语言（如 zh-CN / en-US）。')
    }

    $file = Wait-WinRtOperation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($full)) ([Windows.Storage.StorageFile])
    $stream = Wait-WinRtOperation ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Wait-WinRtOperation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Wait-WinRtOperation ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $ocrResult = Wait-WinRtOperation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

    $lines = @()
    foreach ($line in @($ocrResult.Lines)) {
        $words = @()
        $minX = [double]::MaxValue
        $minY = [double]::MaxValue
        $maxX = [double]::MinValue
        $maxY = [double]::MinValue

        foreach ($word in @($line.Words)) {
            $r = $word.BoundingRect
            $words += [pscustomobject][ordered]@{
                Text   = $word.Text
                X      = [int][Math]::Round($r.X)
                Y      = [int][Math]::Round($r.Y)
                Width  = [int][Math]::Round($r.Width)
                Height = [int][Math]::Round($r.Height)
            }
            if ($r.X -lt $minX) { $minX = $r.X }
            if ($r.Y -lt $minY) { $minY = $r.Y }
            if (($r.X + $r.Width) -gt $maxX) { $maxX = $r.X + $r.Width }
            if (($r.Y + $r.Height) -gt $maxY) { $maxY = $r.Y + $r.Height }
        }

        if ($words.Count -eq 0) {
            $minX = 0; $minY = 0; $maxX = 0; $maxY = 0
        }

        $lines += [pscustomobject][ordered]@{
            Text   = $line.Text
            X      = [int][Math]::Round($minX)
            Y      = [int][Math]::Round($minY)
            Width  = [int][Math]::Round($maxX - $minX)
            Height = [int][Math]::Round($maxY - $minY)
            Words  = $words
        }
    }

    [pscustomobject][ordered]@{
        Path   = $full
        Width  = $decoder.PixelWidth
        Height = $decoder.PixelHeight
        Text   = (($lines | ForEach-Object { $_.Text }) -join "`n")
        Lines  = $lines
    }
}

function Find-OcrMatches {
    param(
        [Parameter(Mandatory)]$Ocr,
        [Parameter(Mandatory)][string]$Needle,
        [switch]$UseExact
    )

    $wordHits = @()
    $lineHits = @()

    foreach ($line in @($Ocr.Lines)) {
        foreach ($word in @($line.Words)) {
            $hit = if ($UseExact) {
                $word.Text -ieq $Needle
            } else {
                $word.Text.IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0
            }
            if ($hit) {
                $wordHits += [pscustomobject][ordered]@{
                    Text   = $word.Text
                    X      = $word.X
                    Y      = $word.Y
                    Width  = $word.Width
                    Height = $word.Height
                    Source = 'word'
                    Line   = $line.Text
                }
            }
        }

        $lineHit = if ($UseExact) {
            $line.Text -ieq $Needle
        } else {
            $line.Text.IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0
        }
        if ($lineHit) {
            $lineHits += [pscustomobject][ordered]@{
                Text   = $line.Text
                X      = $line.X
                Y      = $line.Y
                Width  = $line.Width
                Height = $line.Height
                Source = 'line'
                Line   = $line.Text
            }
        }
    }

    if ($wordHits.Count -gt 0) { return $wordHits }
    return $lineHits
}

# ---------------------------------------------------------------------------
# 本地隧道感知（AccessClient / mstsc → 127.0.0.1:<port>）
# ---------------------------------------------------------------------------
function Get-TcpTable {
    if (Get-Command Get-NetTCPConnection -ErrorAction SilentlyContinue) {
        try {
            $rows = Get-NetTCPConnection -ErrorAction Stop
            $mapped = foreach ($row in $rows) {
                [pscustomobject]@{
                    State         = [string]$row.State
                    LocalAddress  = [string]$row.LocalAddress
                    LocalPort     = [int]$row.LocalPort
                    RemoteAddress = [string]$row.RemoteAddress
                    RemotePort    = [int]$row.RemotePort
                    OwningProcess = [int]$row.OwningProcess
                }
            }
            return @($mapped)
        } catch { }
    }

    # 回退：netstat -ano（只解析 TCP，足够定位回环隧道）
    $result = @()
    foreach ($line in (& netstat -ano -p TCP 2>$null)) {
        $parts = @($line -split '\s+' | Where-Object { $_ -ne '' })
        if ($parts.Count -lt 5 -or $parts[0] -ne 'TCP') { continue }
        if ($parts[1] -notmatch '^(?<addr>.+):(?<port>\d+)$') { continue }
        $laddr = $Matches.addr; $lport = [int]$Matches.port
        if ($parts[2] -notmatch '^(?<addr>.+):(?<port>\d+)$') { continue }
        $raddr = $Matches.addr; $rport = [int]$Matches.port
        $stateText = switch ($parts[3].ToUpperInvariant()) {
            'LISTENING'   { 'Listen' }
            'ESTABLISHED' { 'Established' }
            default       { $parts[3] }
        }
        $result += [pscustomobject]@{
            State         = $stateText
            LocalAddress  = $laddr.Trim('[', ']')
            LocalPort     = $lport
            RemoteAddress = $raddr.Trim('[', ']')
            RemotePort    = $rport
            OwningProcess = [int]$parts[4]
        }
    }
    return $result
}

function Parse-LoopbackEndpoint {
    param([string]$CommandLine)
    if ([string]::IsNullOrWhiteSpace($CommandLine)) { return $null }

    $patterns = @(
        '/v:["'']?(?<host>\[[^\]]+\]|[^"''\s:]+)(?::(?<port>\d+))?',
        'full address:s:(?<host>\[[^\]]+\]|[^:\s]+)(?::(?<port>\d+))?'
    )
    foreach ($pattern in $patterns) {
        $m = [regex]::Match($CommandLine, $pattern, 'IgnoreCase')
        if (-not $m.Success) { continue }
        $hostName = $m.Groups['host'].Value.Trim('[', ']').ToLowerInvariant()
        if ($hostName -notin @('127.0.0.1', 'localhost', '::1',
                               '0:0:0:0:0:0:0:1', '::ffff:127.0.0.1')) { continue }
        $portValue = if ($m.Groups['port'].Success) { [int]$m.Groups['port'].Value } else { 3389 }
        return [pscustomobject]@{ Host = $hostName; Port = $portValue }
    }
    return $null
}

function Get-LoopbackTunnels {
    param([int]$Port = 0, [int]$ClientPid = 0)

    $tcp = @(Get-TcpTable)
    $listeners = @($tcp | Where-Object { $_.State -eq 'Listen' })
    $connections = @($tcp | Where-Object { $_.State -eq 'Established' })
    $processes = @(Get-CimInstance Win32_Process `
        -Filter "Name='mstsc.exe' OR Name='msrdc.exe'" -ErrorAction SilentlyContinue |
        Select-Object ProcessId, ParentProcessId, Name, ExecutablePath, CommandLine)
    $windows = @()
    try { $windows = @([RdpNative]::GetWindows($false)) } catch { }

    $loopback = @('127.0.0.1', '::1', '::ffff:127.0.0.1')
    $anyAddress = @('127.0.0.1', '::1', '0.0.0.0', '::')
    $results = @()

    foreach ($proc in $processes) {
        if ($ClientPid -gt 0 -and [int]$proc.ProcessId -ne $ClientPid) { continue }

        $commandLine = [string]$proc.CommandLine
        $declared = Parse-LoopbackEndpoint -CommandLine $commandLine

        $ports = New-Object System.Collections.Generic.List[int]
        if ($declared) { $ports.Add([int]$declared.Port) }
        foreach ($row in @($connections | Where-Object {
                    $_.OwningProcess -eq [int]$proc.ProcessId -and
                    $_.RemoteAddress -in $loopback -and $_.RemotePort -gt 0 })) {
            if (-not $ports.Contains([int]$row.RemotePort)) { $ports.Add([int]$row.RemotePort) }
        }
        if ($Port -gt 0) { $ports = @($ports | Where-Object { $_ -eq $Port }) }
        if ($ports.Count -eq 0) { continue }

        $window = $windows | Where-Object { $_.Pid -eq [int]$proc.ProcessId } | Select-Object -First 1
        $parentName = ''
        try { $parentName = (Get-Process -Id $proc.ParentProcessId -ErrorAction Stop).ProcessName } catch { }

        foreach ($p in $ports) {
            $listenRows = @($listeners | Where-Object {
                    $_.LocalPort -eq $p -and $_.LocalAddress -in $anyAddress })
            $listenerPid = if ($listenRows.Count -gt 0) { [int]$listenRows[0].OwningProcess } else { 0 }
            $listenerName = ''
            if ($listenerPid -gt 0) {
                try { $listenerName = (Get-Process -Id $listenerPid -ErrorAction Stop).ProcessName } catch { }
            }
            $connected = @($connections | Where-Object {
                    $_.OwningProcess -eq [int]$proc.ProcessId -and
                    $_.RemotePort -eq $p -and $_.RemoteAddress -in $loopback }).Count -gt 0

            $results += [pscustomobject][ordered]@{
                Pid             = [int]$proc.ProcessId
                Process         = [string]$proc.Name
                ParentPid       = [int]$proc.ParentProcessId
                ParentProcess   = $parentName
                Port            = [int]$p
                Endpoint        = "127.0.0.1:$p"
                Declared        = [bool]($declared -and [int]$declared.Port -eq [int]$p)
                Listening       = ($listenRows.Count -gt 0)
                ListenerPid     = $listenerPid
                ListenerProcess = $listenerName
                Connected       = $connected
                Ready           = (($listenRows.Count -gt 0) -and $connected)
                WindowHandle    = $(if ($window) { '0x{0:X8}' -f $window.Handle.ToInt64() } else { '' })
                WindowTitle     = $(if ($window) { $window.Title } else { '' })
                CommandLine     = $commandLine
            }
        }
    }

    return @($results | Sort-Object Pid, Port)
}

function Format-RdpTunnel {
    param($Tunnel)
    $state = if ($Tunnel.Listening) {
        "LISTEN:$($Tunnel.ListenerProcess)($($Tunnel.ListenerPid))"
    } else { 'NO-LISTEN' }
    $state += ' / ' + $(if ($Tunnel.Connected) { 'CONNECTED' } else { 'NOT-CONNECTED' })
    $parent = if ($Tunnel.ParentProcess) { $Tunnel.ParentProcess } else { '?' }
    return ('PID {0,-6} {1,-12} -> {2,-21} [{3}]  parent={4}({5})  {6}' -f `
        $Tunnel.Pid, $Tunnel.Process, $Tunnel.Endpoint, $state,
        $parent, $Tunnel.ParentPid, $Tunnel.WindowTitle)
}

function Get-LoopbackPortStatus {
    param([int]$Port)
    $rows = @(Get-TcpTable | Where-Object {
            $_.LocalPort -eq $Port -and $_.State -eq 'Listen' -and
            $_.LocalAddress -in @('127.0.0.1', '::1', '0.0.0.0', '::') })
    $ownerPid = if ($rows.Count -gt 0) { [int]$rows[0].OwningProcess } else { 0 }
    $ownerName = ''
    if ($ownerPid -gt 0) {
        try { $ownerName = (Get-Process -Id $ownerPid -ErrorAction Stop).ProcessName } catch { }
    }
    return [pscustomobject][ordered]@{
        Port      = $Port
        Listening = ($rows.Count -gt 0)
        Addresses = @($rows | ForEach-Object { $_.LocalAddress } | Sort-Object -Unique)
        Pid       = $ownerPid
        Process   = $ownerName
    }
}

function Wait-LoopbackPort {
    param(
        [int]$Port,
        [int]$TimeoutSec = 30,
        [int]$IntervalMs = 500,
        [switch]$Connect
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ($true) {
        $status = Get-LoopbackPortStatus -Port $Port
        if ($status.Listening) {
            if ($Connect) {
                $ok = $false
                $client = New-Object System.Net.Sockets.TcpClient
                try {
                    $ok = $client.ConnectAsync('127.0.0.1', $Port).Wait(2000)
                } catch {
                    $ok = $false
                } finally {
                    $client.Close()
                }
                $status | Add-Member -NotePropertyName TcpConnect -NotePropertyValue $ok -Force
            }
            return $status
        }
        if ((Get-Date) -ge $deadline) { return $status }
        Start-Sleep -Milliseconds ([Math]::Max(50, $IntervalMs))
    }
}

function Get-RdpHealth {
    param(
        $Target,
        [int]$Port = 0,
        [switch]$CheckScreen,
        [int]$ScreenIntervalMs = 600
    )

    $reasons = New-Object System.Collections.Generic.List[string]
    $geom = Get-ClientGeometry $Target

    if (-not $Target.Visible) { $reasons.Add('window_invisible') }
    if ($Target.Minimized) { $reasons.Add('window_minimized') }
    if ($geom.Width -le 0 -or $geom.Height -le 0) { $reasons.Add('client_area_empty') }

    $hung = $false
    try {
        $proc = Get-Process -Id $Target.Pid -ErrorAction Stop
        $hung = -not $proc.Responding
    } catch { }
    if ($hung) { $reasons.Add('window_not_responding') }

    $tunnels = @(Get-LoopbackTunnels -Port $Port -ClientPid $Target.Pid)
    $tunnel = if ($tunnels.Count -gt 0) { $tunnels[0] } else { $null }
    if ($null -eq $tunnel) {
        $reasons.Add('tunnel_not_found')
    } else {
        if (-not $tunnel.Listening) { $reasons.Add('tunnel_port_not_listening') }
        if (-not $tunnel.Connected) { $reasons.Add('tunnel_not_connected') }
    }

    $screen = $null
    if ($CheckScreen) {
        if ($Target.Minimized -or $geom.Width -le 0 -or $geom.Height -le 0) {
            $screen = [pscustomobject][ordered]@{ Checked = $false; Reason = 'window_unavailable' }
        } else {
            $first = Get-ClientScreenshot -Target $Target -DoNotActivate
            $firstHash = (Get-FileHash -LiteralPath $first -Algorithm SHA256).Hash
            Start-Sleep -Milliseconds ([Math]::Max(100, $ScreenIntervalMs))
            $second = Get-ClientScreenshot -Target $Target -DoNotActivate
            $secondHash = (Get-FileHash -LiteralPath $second -Algorithm SHA256).Hash
            $changed = ($firstHash -ne $secondHash)
            if (-not $changed) { $reasons.Add('screen_frozen') }
            $screen = [pscustomobject][ordered]@{
                Checked = $true
                Changed = $changed
                First   = $first
                Second  = $second
            }
        }
    }

    $windowObject = ConvertTo-RdpWindow $Target
    $windowObject | Add-Member -NotePropertyName ClientOrigin -Force `
        -NotePropertyValue ("{0},{1}" -f $geom.Left, $geom.Top)
    $windowObject | Add-Member -NotePropertyName ClientSize -Force `
        -NotePropertyValue ("{0}x{1}" -f $geom.Width, $geom.Height)
    $windowObject | Add-Member -NotePropertyName DpiScale -Force `
        -NotePropertyValue ([Math]::Round([RdpNative]::DpiScale($Target.Handle), 3))
    $windowObject | Add-Member -NotePropertyName Foreground -Force `
        -NotePropertyValue ([RdpNative]::GetForegroundHandle().ToInt64() -eq $Target.Handle.ToInt64())

    return [pscustomobject][ordered]@{
        Healthy    = ($reasons.Count -eq 0)
        Reasons    = @($reasons)
        Window     = $windowObject
        Hung       = $hung
        ClientArea = [pscustomobject][ordered]@{
            Left = $geom.Left; Top = $geom.Top; Width = $geom.Width; Height = $geom.Height
        }
        Tunnel     = $tunnel
        Screen     = $screen
    }
}

# ---------------------------------------------------------------------------
# 动作分发
# ---------------------------------------------------------------------------
$filter = @{
    TitlePattern  = $Pattern
    ClassName     = $Class
    ProcName      = $ProcessName
    WindowHandle  = $HwndValue
    PidValue      = $ProcessId
    UseRegex      = $Regex
    IncludeHidden = $IncludeInvisible
    AllowAny      = $Any
}

switch ($Action) {

    'find' {
        $targets = @(Get-RdpTargets @filter)
        if ($targets.Count -eq 0) {
            if ($Json) { '[]' } else { Write-Warning '没有匹配的窗口' }
            return
        }
        $objects = @($targets | ForEach-Object { ConvertTo-RdpWindow $_ })
        if ($Json) {
            ConvertTo-Json -InputObject $objects -Depth 5
        } else {
            $objects | Format-Table -AutoSize | Out-String -Width 300
        }
    }

    'status' {
        $target = Get-RdpTarget @filter -Pick $Index
        $geom = Get-ClientGeometry $target
        $obj = [pscustomobject][ordered]@{
            Window       = (ConvertTo-RdpWindow $target)
            ClientOrigin = ('{0},{1}' -f $geom.Left, $geom.Top)
            ClientSize   = ('{0}x{1}' -f $geom.Width, $geom.Height)
            DpiScale     = [Math]::Round([RdpNative]::DpiScale($target.Handle), 3)
            Foreground   = ([RdpNative]::GetForegroundHandle().ToInt64() -eq $target.Handle.ToInt64())
        }
        if ($Json) { $obj | ConvertTo-Json -Depth 5 }
        else {
            $obj.Window | Format-List | Out-String -Width 300
            "ClientOrigin = $($obj.ClientOrigin)"
            "ClientSize   = $($obj.ClientSize)"
            "DpiScale     = $($obj.DpiScale)"
            "Foreground   = $($obj.Foreground)"
        }
    }

    'open-cmd' {
        $target = Get-RdpTarget @filter -Pick $Index
        Invoke-RemoteCommand -Target $target -Command 'cmd' -Wait $WaitMs
        "已向 [$($target.Title)] 发送 Win+R → cmd"
    }

    'run-ps' {
        $target = Get-RdpTarget @filter -Pick $Index

        if ($RemoteFile) {
            $command = 'powershell -NoProfile -ExecutionPolicy Bypass -File "' + $RemoteFile + '"'
            Invoke-RemoteCommand -Target $target -Command $command -Wait $WaitMs
            "已在远端执行文件: $RemoteFile"
            return
        }

        $body = $null
        if ($PSBoundParameters.ContainsKey('Script') -and $Script) {
            $body = $Script
        } elseif ($ScriptFile) {
            $body = [IO.File]::ReadAllText([IO.Path]::GetFullPath($ScriptFile))
        } else {
            throw '请用 -Script / -ScriptFile / -RemoteFile 提供要执行的脚本'
        }

        if (-not $body.Trim()) { throw '脚本内容为空' }

        if ($Encoded) {
            $payload = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($body))
            $command = 'powershell -NoProfile '
            if (-not $ExitAfter) { $command += '-NoExit ' }
            $command += '-EncodedCommand ' + $payload

            # 本机「运行」对话框有长度限制，太长自动退回剪贴板方式
            if ($command.Length -le 200) {
                Invoke-RemoteCommand -Target $target -Command $command -Wait $WaitMs
                "已在远端启动 PowerShell（encoded 方式），脚本 $($body.Length) 字符"
                return
            }
        }

        Invoke-RemoteScriptViaClipboard -Target $target -Body $body -Wait $WaitMs -CloseAfter:$ExitAfter
        "已在远端启动 PowerShell（剪贴板方式），脚本 $($body.Length) 字符"
        '提示：剪贴板方式要求 mstsc 已启用「剪贴板」重定向。'
    }

    'type' {
        if ([string]::IsNullOrEmpty($Text)) { throw '请用 -Text 指定要输入的文本' }
        $target = Get-RdpTarget @filter -Pick $Index
        if (-not $NoActivate) { [void][RdpNative]::ForceForeground($target.Handle) }
        [RdpNative]::TypeText($Text)
        "已向 [$($target.Title)] 输入 $($Text.Length) 个字符"
    }

    'keys' {
        if (-not $Keys) { throw '请用 -Keys 指定按键，例如 "ctrl+v","enter"' }
        $target = Get-RdpTarget @filter -Pick $Index
        if (-not $NoActivate) { [void][RdpNative]::ForceForeground($target.Handle) }
        foreach ($k in $Keys) {
            [RdpNative]::SendCombo($k)
            Start-Sleep -Milliseconds 100
        }
        "已向 [$($target.Title)] 发送按键: $($Keys -join ', ')"
    }

    'click' {
        $target = Get-RdpTarget @filter -Pick $Index
        if (-not $NoActivate) { [void][RdpNative]::ForceForeground($target.Handle) }

        if ($WindowRelative) {
            $originX = $target.Left
            $originY = $target.Top
            $width = $target.Width
            $height = $target.Height
        } else {
            $geom = Get-ClientGeometry $target
            $originX = $geom.Left
            $originY = $geom.Top
            $width = $geom.Width
            $height = $geom.Height
        }

        if ($X -ge 0) { $px = $originX + $X } else { $px = $originX + [int]($width / 2) }
        if ($Y -ge 0) { $py = $originY + $Y } else { $py = $originY + [int]($height / 2) }

        [RdpNative]::ClickAt([int]$px, [int]$py, $DoubleClick.IsPresent, $RightClick.IsPresent)
        "已点击屏幕坐标 ($px, $py)"
    }

    'screenshot' {
        $target = Get-RdpTarget @filter -Pick $Index
        $path = Get-ClientScreenshot -Target $target -Path $Out -DoNotActivate:$NoActivate
        $geom = Get-ClientGeometry $target
        $obj = [pscustomobject][ordered]@{
            Path   = $path
            Width  = $geom.Width
            Height = $geom.Height
        }
        if ($Json) { $obj | ConvertTo-Json -Depth 4 }
        else { "已保存截图: $path  ($($geom.Width)x$($geom.Height))" }
    }

    'ocr' {
        if ($Image) {
            $result = Invoke-Ocr -Path $Image -OcrLanguage $Language
        } else {
            $target = Get-RdpTarget @filter -Pick $Index
            $path = Get-ClientScreenshot -Target $target -Path $Out -DoNotActivate:$NoActivate
            $result = Invoke-Ocr -Path $path -OcrLanguage $Language
        }
        if ($Json) { $result | ConvertTo-Json -Depth 8 }
        else { $result.Text }
    }

    'click-text' {
        if (-not $Text) { throw '请用 -Text 指定要查找的文字' }
        $target = Get-RdpTarget @filter -Pick $Index
        $path = Get-ClientScreenshot -Target $target -Path $Out -DoNotActivate:$NoActivate
        $result = Invoke-Ocr -Path $path -OcrLanguage $Language
        $hits = @(Find-OcrMatches -Ocr $result -Needle $Text -UseExact:$Exact)

        if ($hits.Count -eq 0) { throw "未找到文字: $Text（截图: $path）" }
        if ($MatchIndex -lt 0 -or $MatchIndex -ge $hits.Count) {
            throw "匹配索引 $MatchIndex 越界，共找到 $($hits.Count) 个"
        }

        $hit = $hits[$MatchIndex]
        $geom = Get-ClientGeometry $target
        $sx = $geom.Left + $hit.X + [int]($hit.Width / 2)
        $sy = $geom.Top + $hit.Y + [int]($hit.Height / 2)

        $obj = [pscustomobject][ordered]@{
            Found    = $true
            Match    = $hit
            ScreenX  = $sx
            ScreenY  = $sy
            Clicked  = (-not $DryRun)
            Image    = $path
        }

        if ($Json) { $obj | ConvertTo-Json -Depth 6 }
        else {
            "匹配文字 : $($hit.Text)"
            "匹配来源 : $($hit.Source)  (整行: $($hit.Line))"
            "图片坐标 : $($hit.X),$($hit.Y) $($hit.Width)x$($hit.Height)"
            "屏幕坐标 : $sx,$sy"
        }

        if (-not $DryRun) {
            [RdpNative]::ClickAt([int]$sx, [int]$sy, $DoubleClick.IsPresent, $RightClick.IsPresent)
            if (-not $Json) { '已点击' }
        }
    }

    'wait-text' {
        if (-not $Text) { throw '请用 -Text 指定要等待的文字' }
        $target = Get-RdpTarget @filter -Pick $Index
        if (-not $NoActivate) {
            [void][RdpNative]::ForceForeground($target.Handle)
            Start-Sleep -Milliseconds 250
        }

        $deadline = (Get-Date).AddSeconds($TimeoutSec)
        $hit = $null
        $lastPath = $null

        while ($true) {
            $lastPath = Get-ClientScreenshot -Target $target -DoNotActivate
            $result = Invoke-Ocr -Path $lastPath -OcrLanguage $Language
            $hits = @(Find-OcrMatches -Ocr $result -Needle $Text -UseExact:$Exact)
            if ($hits.Count -gt 0) { $hit = $hits[0]; break }
            if ((Get-Date) -ge $deadline) { break }
            Start-Sleep -Milliseconds $IntervalMs
        }

        if ($null -eq $hit) {
            throw "等待文字超时: $Text（最后截图: $lastPath）"
        }

        $geom = Get-ClientGeometry $target
        $sx = $geom.Left + $hit.X + [int]($hit.Width / 2)
        $sy = $geom.Top + $hit.Y + [int]($hit.Height / 2)

        $obj = [pscustomobject][ordered]@{
            Found   = $true
            Match   = $hit
            ScreenX = $sx
            ScreenY = $sy
        }
        if ($Json) { $obj | ConvertTo-Json -Depth 6 }
        else {
            "匹配文字 : $($hit.Text)"
            "匹配来源 : $($hit.Source)  (整行: $($hit.Line))"
            "屏幕坐标 : $sx,$sy"
        }
    }

    'clip-set' {
        if ($null -eq $Text) { throw '请用 -Text 指定要写入剪贴板的文本' }
        Set-Clipboard -Value $Text
        "已写入剪贴板，共 $($Text.Length) 个字符"
    }

    'clip-get' {
        $content = Get-Clipboard -Raw
        if ($Json) {
            [pscustomobject]@{ Text = $content } | ConvertTo-Json -Depth 4
        } else {
            $content
        }
    }

    'paste' {
        if ($null -eq $Text) { throw '请用 -Text 指定要粘贴的文本' }
        $target = Get-RdpTarget @filter -Pick $Index
        Set-Clipboard -Value $Text
        Start-Sleep -Milliseconds 150
        [void][RdpNative]::ForceForeground($target.Handle)
        [RdpNative]::SendCombo('ctrl+v')
        "已粘贴 $($Text.Length) 个字符到 [$($target.Title)]"
    }

    'tunnel' {
        $deadline = (Get-Date).AddSeconds([Math]::Max(0, $TimeoutSec))
        do {
            $tunnels = @(Get-LoopbackTunnels -Port $Port)
            if ($App) {
                $tunnels = @($tunnels | Where-Object {
                        $_.ParentProcess -like $App -or $_.ListenerProcess -like $App })
            }
            if ($tunnels.Count -gt 0 -or (Get-Date) -ge $deadline) { break }
            Start-Sleep -Milliseconds ([Math]::Max(50, $IntervalMs))
        } while ($true)

        if ($Launch) {
            foreach ($item in @($tunnels | Where-Object { -not $_.Connected })) {
                Start-Process mstsc -ArgumentList "/v:127.0.0.1:$($item.Port)"
                "已拉起 mstsc /v:127.0.0.1:$($item.Port)"
            }
        }

        if ($Watch) {
            '监控本地隧道中，Ctrl+C 退出 ...'
            $last = @{}
            while ($true) {
                $tunnels = @(Get-LoopbackTunnels -Port $Port)
                if ($App) {
                    $tunnels = @($tunnels | Where-Object {
                            $_.ParentProcess -like $App -or $_.ListenerProcess -like $App })
                }
                $current = @{}
                foreach ($item in $tunnels) {
                    $current["$($item.Pid):$($item.Port)"] = "$($item.Listening)/$($item.Connected)"
                }
                $stamp = Get-Date -Format 'HH:mm:ss'
                foreach ($item in $tunnels) {
                    $key = "$($item.Pid):$($item.Port)"
                    if (-not $last.ContainsKey($key) -or $last[$key] -ne $current[$key]) {
                        "[$stamp] $(Format-RdpTunnel $item)"
                    }
                }
                foreach ($key in @($last.Keys | Where-Object { -not $current.ContainsKey($_) })) {
                    "[$stamp] 隧道消失: $key"
                }
                $last = $current
                Start-Sleep -Milliseconds ([Math]::Max(200, $IntervalMs))
            }
        }

        if ($Json) {
            ConvertTo-Json -InputObject @($tunnels) -Depth 5
        } elseif ($tunnels.Count -eq 0) {
            Write-Warning '没有发现本地隧道：mstsc 未连接 127.0.0.1，或客户端进程不可见。'
        } else {
            $tunnels | ForEach-Object { Format-RdpTunnel $_ }
        }
        if ($tunnels.Count -eq 0) { exit 1 }
    }

    'wait-port' {
        if ($Port -le 0) { throw '请用 -Port 指定要等待的本地端口' }
        $status = Wait-LoopbackPort -Port $Port -TimeoutSec $TimeoutSec `
            -IntervalMs $IntervalMs -Connect:$Connect
        if ($Json) {
            $status | ConvertTo-Json -Depth 4
        } elseif ($status.Listening) {
            $extra = ''
            if ($Connect) {
                $extra = if ($status.TcpConnect) { ' (TCP 连接成功)' } else { ' (TCP 连接失败)' }
            }
            "端口 $Port 已就绪: $($status.Addresses -join ', ')  " +
            "进程 $($status.Process)($($status.Pid))$extra"
        } else {
            throw "等待端口 $Port 超时（${TimeoutSec}s）"
        }
        if (-not $status.Listening) { exit 1 }
    }

    'health' {
        $targets = @(Get-RdpTargets @filter)
        if ($targets.Count -eq 0) {
            $payload = [pscustomobject][ordered]@{
                Healthy = $false
                Reasons = @('window_not_found')
                Window  = $null
                Tunnel  = $null
            }
            if ($Json) { $payload | ConvertTo-Json -Depth 6 }
            else { Write-Warning '窗口未找到' }
            exit 1
        }
        if ($Index -lt 0 -or $Index -ge $targets.Count) {
            throw "索引 $Index 越界，共找到 $($targets.Count) 个窗口"
        }

        $payload = Get-RdpHealth -Target $targets[$Index] -Port $Port `
            -CheckScreen:$Screen -ScreenIntervalMs $IntervalMs

        if ($Json) {
            $payload | ConvertTo-Json -Depth 8
        } else {
            "健康     : $(if ($payload.Healthy) { 'OK' } else { 'FAIL' })"
            if ($payload.Reasons.Count -gt 0) {
                "问题     : $($payload.Reasons -join ', ')"
            }
            "窗口     : $($payload.Window.Handle) [$($payload.Window.Title)] PID $($payload.Window.PID)"
            "客户区   : $($payload.ClientArea.Left),$($payload.ClientArea.Top) " +
                "$($payload.ClientArea.Width)x$($payload.ClientArea.Height)"
            "状态     : visible=$($payload.Window.Visible) " +
                "minimized=$($payload.Window.Minimized) hung=$($payload.Hung)"
            if ($payload.Tunnel) {
                "隧道     : $($payload.Tunnel.Endpoint)  " +
                "listener=$($payload.Tunnel.ListenerProcess)($($payload.Tunnel.ListenerPid))  " +
                "connected=$($payload.Tunnel.Connected)"
            } else {
                '隧道     : 未发现'
            }
            if ($payload.Screen -and $payload.Screen.Checked) {
                "画面     : $(if ($payload.Screen.Changed) { '有变化' } else { '无变化（可能卡死）' })"
            }
        }

        if (-not $payload.Healthy) { exit 1 }
    }
}
