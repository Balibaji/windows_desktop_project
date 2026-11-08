#Requires -Version 5.1
<#
.SYNOPSIS
    Windows 窗口枚举与自动化操作工具（PowerShell 版，纯 Win32 API + .NET，无第三方依赖）

.DESCRIPTION
    查找 / 枚举当前桌面上打开的窗口，信息包括但不限于标题：
    窗口句柄、PID、进程名、窗口类名、位置与尺寸、可见 / 最小化 / 启用状态、窗口样式。
    还可以枚举窗口内的子控件（编辑框、按钮等）并读取控件文本。

    对找到的窗口可以执行：
      activate   激活到前台
      type       模拟真实键盘输入（SendInput，支持中文 / Unicode）
      keys       发送组合键（ctrl+a、alt+f4、enter ...）
      click      鼠标点击（坐标相对窗口左上角）
      gettext    读取目标控件文本（WM_GETTEXT，Scintilla 自动回退 SCI_GETTEXT）
      settext    直接写入目标控件文本（WM_SETTEXT）
      close      发送 WM_CLOSE 关闭窗口
      foreground 查看当前前台窗口

.PARAMETER Action
    list | info | children | activate | type | keys | click | gettext | settext | close | foreground

.PARAMETER Pattern
    按标题筛选（默认不区分大小写的“包含”匹配，加 -Regex 则按正则匹配）。

.PARAMETER Class
    按窗口类名筛选，支持通配符，例如 -Class "Notepad*"。

.PARAMETER ProcessName
    按进程名筛选，支持通配符，例如 -ProcessName "notepad*"。

.PARAMETER Hwnd
    直接按窗口句柄定位，支持十进制（如 984448）或十六进制（如 0xF0580）。

.PARAMETER ProcessId
    按 PID 定位。

.PARAMETER Index
    匹配到多个窗口时取第几个（从 0 开始，默认 0）。

.PARAMETER IncludeInvisible
    连隐藏窗口一起枚举。

.PARAMETER MessageMode
    type 动作改为向控件 PostMessage(WM_CHAR)，不需要窗口获得前台焦点。

.PARAMETER NoActivate
    执行 type / keys / click 前不自动激活窗口。

.PARAMETER ControlClass
    gettext / settext 指定控件类名（通配符），默认自动挑选编辑类控件。

.PARAMETER Json
    以 JSON 输出（list / info / children / gettext）。

.EXAMPLE
    .\WindowOps.ps1 list
.EXAMPLE
    .\WindowOps.ps1 list -Pattern "记事本" -Json
.EXAMPLE
    .\WindowOps.ps1 info -Pattern "记事本"
.EXAMPLE
    .\WindowOps.ps1 type -Pattern "记事本" -Text "Hello，世界！"
.EXAMPLE
    .\WindowOps.ps1 keys -Pattern "记事本" -Keys "ctrl+a","ctrl+c"
.EXAMPLE
    .\WindowOps.ps1 click -Pattern "记事本" -X 200 -Y 150
.EXAMPLE
    .\WindowOps.ps1 gettext -Pattern "记事本" -ControlClass "Edit"
.EXAMPLE
    .\WindowOps.ps1 settext -Pattern "记事本" -Text "直接写入内容" -ControlClass "Edit"
.EXAMPLE
    .\WindowOps.ps1 close -Pattern "记事本"
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet('list', 'info', 'children', 'activate', 'type', 'keys', 'click',
                 'gettext', 'settext', 'close', 'foreground')]
    [string]$Action = 'list',

    [string]$Pattern,
    [string]$Class,
    [string]$ProcessName,
    [string]$Hwnd = '',
    [int]$ProcessId = 0,
    [switch]$Regex,
    [switch]$IncludeInvisible,
    [int]$Index = 0,

    [string]$Text,
    [string[]]$Keys,
    [int]$X = -1,
    [int]$Y = -1,
    [switch]$DoubleClick,
    [switch]$RightClick,

    [string]$ControlClass,
    [switch]$MessageMode,
    [switch]$NoActivate,
    [switch]$Json
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [Text.Encoding]::UTF8
    $OutputEncoding = [Text.Encoding]::UTF8
} catch { }

# ---------------------------------------------------------------------------
# Win32 API 封装（C# 5 语法，兼容 Windows PowerShell 5.1 的 Add-Type）
# ---------------------------------------------------------------------------
if (-not ('Win32Ext' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;

public class WinInfo
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
    public bool Enabled;
    public long Style;
    public long ExStyle;
    public IntPtr Owner;
}

public static class Win32Ext
{
    public const uint WM_SETTEXT      = 0x000C;
    public const uint WM_GETTEXT      = 0x000D;
    public const uint WM_GETTEXTLENGTH = 0x000E;
    public const uint WM_CHAR         = 0x0102;
    public const uint WM_CLOSE        = 0x0010;
    public const uint SCI_GETTEXT     = 2180;
    public const uint SCI_GETLENGTH   = 2006;

    public delegate bool EnumProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

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

    const uint INPUT_KEYBOARD          = 1;
    const uint KEYEVENTF_KEYUP         = 0x0002;
    const uint KEYEVENTF_UNICODE       = 0x0004;
    const uint MOUSEEVENTF_LEFTDOWN    = 0x0002;
    const uint MOUSEEVENTF_LEFTUP      = 0x0004;
    const uint MOUSEEVENTF_RIGHTDOWN   = 0x0008;
    const uint MOUSEEVENTF_RIGHTUP     = 0x0010;

    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent, EnumProc cb, IntPtr lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder sb, int max);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] static extern IntPtr GetWindowLongPtr(IntPtr h, int index);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr h, uint cmd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);
    [DllImport("user32.dll")] static extern void SwitchToThisWindow(IntPtr h, bool altTab);
    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wParam, string lParam);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wParam, StringBuilder lParam);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint msg, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags, int dx, int dy, uint data, IntPtr extraInfo);
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint count, INPUT[] inputs, int size);
    [DllImport("user32.dll")] static extern short VkKeyScan(char ch);

    static Dictionary<int, string> _procCache = new Dictionary<int, string>();

    static Win32Ext()
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

    static WinInfo Info(IntPtr h)
    {
        StringBuilder title = new StringBuilder(1024);
        GetWindowText(h, title, title.Capacity);
        StringBuilder cls = new StringBuilder(256);
        GetClassName(h, cls, cls.Capacity);

        uint pid;
        GetWindowThreadProcessId(h, out pid);

        RECT r;
        GetWindowRect(h, out r);

        WinInfo w = new WinInfo();
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
        w.Enabled = IsWindowEnabled(h);
        w.Style = GetWindowLongPtr(h, -16).ToInt64();   // GWL_STYLE
        w.ExStyle = GetWindowLongPtr(h, -20).ToInt64(); // GWL_EXSTYLE
        w.Owner = GetWindow(h, 4);                      // GW_OWNER
        return w;
    }

    public static List<WinInfo> GetWindows(bool visibleOnly)
    {
        List<WinInfo> list = new List<WinInfo>();
        EnumWindows(delegate(IntPtr h, IntPtr l)
        {
            if (!visibleOnly || IsWindowVisible(h)) list.Add(Info(h));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    public static List<WinInfo> GetChildren(IntPtr parent)
    {
        List<WinInfo> list = new List<WinInfo>();
        EnumChildWindows(parent, delegate(IntPtr h, IntPtr l)
        {
            list.Add(Info(h));
            return true;
        }, IntPtr.Zero);
        return list;
    }

    public static bool IsWindowHandle(IntPtr h)
    {
        return IsWindow(h);
    }

    public static WinInfo GetForegroundInfo()
    {
        return Info(GetForegroundWindow());
    }

    public static IntPtr GetForegroundHandle()
    {
        return GetForegroundWindow();
    }

    // 把窗口强制带到前台：AttachThreadInput 绕过前台锁定，失败再 SwitchToThisWindow
    public static bool ForceForeground(IntPtr h)
    {
        if (!IsWindow(h)) return false;
        if (IsIconic(h)) ShowWindow(h, 9);  // SW_RESTORE
        ShowWindow(h, 5);                   // SW_SHOW
        SetWindowPos(h, IntPtr.Zero, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0040); // NOSIZE|NOMOVE|SHOWWINDOW

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

    // 逐字符发送 Unicode 键盘事件，中文、emoji、任意符号都能输入
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
                if (name.Length == 1) return 0; // 单字符走 Unicode 分支
                return 0;
        }
    }

    // 发送组合键，例如 ctrl+shift+s / alt+f4 / enter
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
                // 单字符且无修饰键：直接走 Unicode，中文/符号都能打
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

    // 读取控件文本：WM_GETTEXT；Scintilla 控件回退到 SCI_GETTEXT
    public static string GetControlText(IntPtr h)
    {
        int len = (int)SendMessage(h, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero);
        if (len < 0) len = 0;
        StringBuilder sb = new StringBuilder(len + 2);
        SendMessage(h, WM_GETTEXT, new IntPtr(len + 1), sb);
        string text = sb.ToString();

        if (text.Length == 0)
        {
            StringBuilder cls = new StringBuilder(64);
            GetClassName(h, cls, cls.Capacity);
            if (cls.ToString().IndexOf("Scintilla", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                int slen = (int)SendMessage(h, SCI_GETLENGTH, IntPtr.Zero, IntPtr.Zero);
                if (slen > 0)
                {
                    StringBuilder buf = new StringBuilder(slen + 2);
                    SendMessage(h, SCI_GETTEXT, new IntPtr(slen + 1), buf);
                    text = buf.ToString();
                }
            }
        }
        return text;
    }

    public static void SetControlText(IntPtr h, string text)
    {
        SendMessage(h, WM_SETTEXT, IntPtr.Zero, text == null ? "" : text);
    }

    // 向控件投递 WM_CHAR（不需要前台焦点）
    public static void PostText(IntPtr h, string text)
    {
        if (text == null) return;
        foreach (char ch in text)
        {
            if (ch == '\r') continue;
            if (ch == '\n') { PostMessage(h, WM_CHAR, new IntPtr(0x0D), IntPtr.Zero); continue; }
            PostMessage(h, WM_CHAR, new IntPtr(ch), IntPtr.Zero);
        }
        Thread.Sleep(60);
    }

    public static void CloseWindow(IntPtr h)
    {
        PostMessage(h, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }
}
'@
}

# ---------------------------------------------------------------------------
# PowerShell 辅助函数
# ---------------------------------------------------------------------------
# -Hwnd 允许写十进制或 0x 十六进制
$HwndValue = 0L
if ($Hwnd) {
    if ($Hwnd -match '^0[xX]') { $HwndValue = [Convert]::ToInt64($Hwnd.Substring(2), 16) }
    else { $HwndValue = [long]$Hwnd }
}

function Get-TargetWindows {
    [CmdletBinding()]
    param(
        [string]$TitlePattern,
        [string]$ClassName,
        [string]$ProcName,
        [long]$WindowHandle = 0,
        [int]$PidValue = 0,
        [switch]$UseRegex,
        [switch]$IncludeHidden,
        [switch]$AllowUnfiltered
    )

    if (-not $AllowUnfiltered -and
        -not ($TitlePattern -or $ClassName -or $ProcName -or ($WindowHandle -ne 0) -or ($PidValue -ne 0))) {
        throw '请至少指定一个筛选条件：-Pattern / -Class / -ProcessName / -Hwnd / -ProcessId'
    }

    $result = @([Win32Ext]::GetWindows(-not $IncludeHidden.IsPresent))

    if ($WindowHandle -ne 0) { $result = @($result | Where-Object { $_.Handle.ToInt64() -eq $WindowHandle }) }
    if ($PidValue -ne 0)     { $result = @($result | Where-Object { $_.Pid -eq $PidValue }) }
    if ($ProcName)           { $result = @($result | Where-Object { $_.ProcessName -like $ProcName }) }
    if ($ClassName)          { $result = @($result | Where-Object { $_.ClassName -like $ClassName }) }
    if ($TitlePattern) {
        if ($UseRegex) {
            $result = @($result | Where-Object { $_.Title -match $TitlePattern })
        } else {
            $result = @($result | Where-Object { $_.Title.IndexOf($TitlePattern, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
        }
    }

    return @($result | Sort-Object Pid, Title)
}

function Get-OneTarget {
    param(
        [string]$TitlePattern,
        [string]$ClassName,
        [string]$ProcName,
        [long]$WindowHandle = 0,
        [int]$PidValue = 0,
        [switch]$UseRegex,
        [switch]$IncludeHidden,
        [int]$Pick = 0
    )

    $list = @(Get-TargetWindows -TitlePattern $TitlePattern -ClassName $ClassName -ProcName $ProcName `
                                -WindowHandle $WindowHandle -PidValue $PidValue `
                                -UseRegex:$UseRegex -IncludeHidden:$IncludeHidden)
    if ($list.Count -eq 0) {
        throw '没有找到匹配的窗口，请检查 -Pattern / -Class / -ProcessName / -Hwnd / -ProcessId'
    }
    if ($Pick -lt 0 -or $Pick -ge $list.Count) {
        throw "索引 $Pick 越界，共找到 $($list.Count) 个窗口"
    }
    return $list[$Pick]
}

function ConvertTo-WindowObject {
    param([Parameter(Mandatory)]$Window)

    [pscustomobject][ordered]@{
        Handle    = ('0x{0:X8}' -f $Window.Handle.ToInt64())
        PID       = $Window.Pid
        Process   = $Window.ProcessName
        Class     = $Window.ClassName
        Visible   = $Window.Visible
        Minimized = $Window.Minimized
        Enabled   = $Window.Enabled
        Rect      = ('{0},{1} {2}x{3}' -f $Window.Left, $Window.Top, $Window.Width, $Window.Height)
        Title     = $Window.Title
    }
}

function Get-TargetControl {
    param(
        [Parameter(Mandatory)]$Window,
        [string]$ClassNamePattern
    )

    $kids = @([Win32Ext]::GetChildren($Window.Handle))
    if ($kids.Count -eq 0) {
        throw "窗口 [$($Window.Title)] 中没有子控件"
    }

    if ($ClassNamePattern) {
        $matched = @($kids | Where-Object { $_.ClassName -like $ClassNamePattern })
        if ($matched.Count -eq 0) {
            throw "未找到类名匹配 '$ClassNamePattern' 的子控件"
        }
        return $matched[0]
    }

    $textish = @($kids | Where-Object {
        $_.ClassName -match '(?i)^(Edit|RichEdit|Scintilla|TextBox|Chrome_RenderWidgetHostHWND)'
    })
    if ($textish.Count -gt 0) { return $textish[0] }

    return ($kids | Sort-Object { $_.Width * $_.Height } -Descending)[0]
}

# ---------------------------------------------------------------------------
# 动作分发
# ---------------------------------------------------------------------------
function Invoke-WindowAction {
    $filter = @{
        TitlePattern  = $Pattern
        ClassName     = $Class
        ProcName      = $ProcessName
        WindowHandle  = $HwndValue
        PidValue      = $ProcessId
        UseRegex      = $Regex
        IncludeHidden = $IncludeInvisible
    }

    switch ($Action) {

        'list' {
            $list = @(Get-TargetWindows @filter -AllowUnfiltered)
            if ($list.Count -eq 0) { Write-Warning '没有匹配的窗口'; return }
            $objects = @($list | ForEach-Object { ConvertTo-WindowObject $_ })
            if ($Json) { $objects | ConvertTo-Json -Depth 5 } else { $objects }
        }

        'info' {
            $target = Get-OneTarget @filter -Pick $Index
            $kids = @([Win32Ext]::GetChildren($target.Handle) | ForEach-Object { ConvertTo-WindowObject $_ })
            if ($Json) {
                [pscustomobject]@{
                    Window   = (ConvertTo-WindowObject $target)
                    Style    = ('0x{0:X8}' -f $target.Style)
                    ExStyle  = ('0x{0:X8}' -f $target.ExStyle)
                    Owner    = ('0x{0:X8}' -f $target.Owner.ToInt64())
                    Children = $kids
                } | ConvertTo-Json -Depth 6
            } else {
                (ConvertTo-WindowObject $target) | Format-List | Out-String -Width 300
                ('Style = 0x{0:X8}   ExStyle = 0x{1:X8}   Owner = 0x{2:X8}' -f $target.Style, $target.ExStyle, $target.Owner.ToInt64())
                ''
                "子控件（$($kids.Count) 个）:"
                if ($kids.Count -gt 0) { $kids | Format-Table -AutoSize | Out-String -Width 300 }
            }
        }

        'children' {
            $target = Get-OneTarget @filter -Pick $Index
            $kids = @([Win32Ext]::GetChildren($target.Handle))
            if ($kids.Count -eq 0) { "窗口 [$($target.Title)] 没有子控件"; return }
            $objects = @($kids | ForEach-Object { ConvertTo-WindowObject $_ })
            if ($Json) { $objects | ConvertTo-Json -Depth 5 }
            else { $objects | Format-Table -AutoSize | Out-String -Width 300 }
        }

        'activate' {
            $target = Get-OneTarget @filter -Pick $Index
            $ok = [Win32Ext]::ForceForeground($target.Handle)
            "已激活窗口: [$($target.Title)]  句柄=0x$('{0:X}' -f $target.Handle.ToInt64())  成功=$ok"
        }

        'type' {
            if ([string]::IsNullOrEmpty($Text)) { throw '请用 -Text 指定要输入的文本' }
            $target = Get-OneTarget @filter -Pick $Index
            if (-not $NoActivate) { [void][Win32Ext]::ForceForeground($target.Handle) }
            if ($MessageMode) {
                $ctl = Get-TargetControl $target -ClassNamePattern $ControlClass
                [Win32Ext]::PostText($ctl.Handle, $Text)
                "已通过 WM_CHAR 向控件 [$($ctl.ClassName)] 发送 $($Text.Length) 个字符"
            } else {
                [Win32Ext]::TypeText($Text)
                "已通过 SendInput 向窗口 [$($target.Title)] 输入 $($Text.Length) 个字符"
            }
        }

        'keys' {
            if (-not $Keys) { throw '请用 -Keys 指定按键，例如 "ctrl+a"' }
            $target = Get-OneTarget @filter -Pick $Index
            if (-not $NoActivate) { [void][Win32Ext]::ForceForeground($target.Handle) }
            foreach ($k in $Keys) {
                [Win32Ext]::SendCombo($k)
                Start-Sleep -Milliseconds 100
            }
            "已向 [$($target.Title)] 发送按键: $($Keys -join ', ')"
        }

        'click' {
            $target = Get-OneTarget @filter -Pick $Index
            if (-not $NoActivate) { [void][Win32Ext]::ForceForeground($target.Handle) }
            if ($X -ge 0) { $px = $target.Left + $X } else { $px = $target.Left + [int]($target.Width / 2) }
            if ($Y -ge 0) { $py = $target.Top + $Y } else { $py = $target.Top + [int]($target.Height / 2) }
            [Win32Ext]::ClickAt([int]$px, [int]$py, $DoubleClick.IsPresent, $RightClick.IsPresent)
            "已在 [$($target.Title)] 内点击屏幕坐标 ($px, $py)"
        }

        'gettext' {
            $target = Get-OneTarget @filter -Pick $Index
            $ctl = Get-TargetControl $target -ClassNamePattern $ControlClass
            $content = [Win32Ext]::GetControlText($ctl.Handle)
            if ($Json) {
                [pscustomobject]@{
                    Window  = (ConvertTo-WindowObject $target)
                    Control = (ConvertTo-WindowObject $ctl)
                    Text    = $content
                } | ConvertTo-Json -Depth 6
            } else {
                "窗口: [$($target.Title)]"
                "控件: $($ctl.ClassName)  句柄=0x$('{0:X}' -f $ctl.Handle.ToInt64())"
                '--- 文本开始 ---'
                $content
                '--- 文本结束 ---'
            }
        }

        'settext' {
            if ($null -eq $Text) { throw '请用 -Text 指定要写入的文本' }
            $target = Get-OneTarget @filter -Pick $Index
            $ctl = Get-TargetControl $target -ClassNamePattern $ControlClass
            [Win32Ext]::SetControlText($ctl.Handle, $Text)
            "已通过 WM_SETTEXT 写入控件 [$($ctl.ClassName)]，共 $($Text.Length) 个字符"
        }

        'close' {
            $target = Get-OneTarget @filter -Pick $Index
            [Win32Ext]::CloseWindow($target.Handle)
            "已向 [$($target.Title)] 发送 WM_CLOSE"
        }

        'foreground' {
            $w = [Win32Ext]::GetForegroundInfo()
            if ($Json) { ConvertTo-WindowObject $w | ConvertTo-Json -Depth 5 }
            else { ConvertTo-WindowObject $w | Format-List | Out-String -Width 300 }
        }
    }
}

# 直接执行脚本时运行动作；被 demo 脚本 dot-source 时只加载函数
if ($MyInvocation.InvocationName -ne '.') {
    Invoke-WindowAction
}
