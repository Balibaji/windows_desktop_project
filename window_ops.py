#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 窗口枚举与自动化操作工具（Python 版，纯 ctypes，无第三方依赖）

功能与 PowerShell 版 WindowOps.ps1 对齐：
  * 枚举桌面上的顶层窗口：句柄、PID、进程名、窗口类名、标题、位置尺寸、
    可见 / 最小化 / 启用状态、窗口样式
  * 枚举窗口内的子控件（编辑框、按钮、静态文本……）
  * 激活窗口、模拟真实键盘输入（SendInput，支持中文 / Unicode）、组合键、
    鼠标点击、直接读写控件文本（WM_SETTEXT / WM_GETTEXT / WM_CHAR）、关闭窗口

示例：
  python window_ops.py list
  python window_ops.py list --pattern 记事本 --json
  python window_ops.py info --pattern 记事本
  python window_ops.py type --pattern 记事本 --text "Hello，世界！"
  python window_ops.py keys --pattern 记事本 --keys ctrl+a ctrl+c
  python window_ops.py click --pattern 记事本 --x 200 --y 150
  python window_ops.py gettext --pattern 记事本 --control-class Edit
  python window_ops.py settext --pattern 记事本 --text "直接写入" --control-class Edit
  python window_ops.py close --pattern 记事本

注意：真实键盘输入（SendInput）需要目标窗口在前台；
      --message-mode 则改用 PostMessage(WM_CHAR) 直接发给控件，不需要焦点。
"""

from __future__ import annotations

import argparse
import ctypes
import ctypes.wintypes as wt
import fnmatch
import json
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass
from typing import Iterable, Optional

# ---------------------------------------------------------------------------
# 控制台编码：让中文在管道 / 终端里都能正常输出
# ---------------------------------------------------------------------------
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace") #type: ignore
    except Exception:
        pass

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

try:
    user32.SetProcessDPIAware()
except Exception:
    pass

# ---------------------------------------------------------------------------
# Win32 常量
# ---------------------------------------------------------------------------
WM_SETTEXT = 0x000C
WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_CLOSE = 0x0010
WM_CHAR = 0x0102
SCI_GETTEXT = 2180
SCI_GETLENGTH = 2006

GW_OWNER = 4
GWL_STYLE = -16
GWL_EXSTYLE = -20

SW_SHOW = 5
SW_RESTORE = 9
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040

INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010

VK_SHIFT = 0x10
VK_CONTROL = 0x11
VK_MENU = 0x12
VK_LWIN = 0x5B
VK_RETURN = 0x0D
VK_TAB = 0x09

PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

# ---------------------------------------------------------------------------
# 函数原型
# ---------------------------------------------------------------------------
WNDENUMPROC = ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)

user32.EnumWindows.argtypes = (WNDENUMPROC, wt.LPARAM)
user32.EnumWindows.restype = wt.BOOL
user32.EnumChildWindows.argtypes = (wt.HWND, WNDENUMPROC, wt.LPARAM)
user32.EnumChildWindows.restype = wt.BOOL

user32.GetWindowTextW.argtypes = (wt.HWND, wt.LPWSTR, ctypes.c_int)
user32.GetWindowTextW.restype = ctypes.c_int
user32.GetClassNameW.argtypes = (wt.HWND, wt.LPWSTR, ctypes.c_int)
user32.GetClassNameW.restype = ctypes.c_int
user32.GetWindowThreadProcessId.argtypes = (wt.HWND, ctypes.POINTER(wt.DWORD))
user32.GetWindowThreadProcessId.restype = wt.DWORD

user32.IsWindow.argtypes = (wt.HWND,)
user32.IsWindow.restype = wt.BOOL
user32.IsWindowVisible.argtypes = (wt.HWND,)
user32.IsWindowVisible.restype = wt.BOOL
user32.IsWindowEnabled.argtypes = (wt.HWND,)
user32.IsWindowEnabled.restype = wt.BOOL
user32.IsIconic.argtypes = (wt.HWND,)
user32.IsIconic.restype = wt.BOOL

user32.GetWindowRect.argtypes = (wt.HWND, ctypes.POINTER(wt.RECT))
user32.GetWindowRect.restype = wt.BOOL
user32.GetWindowLongPtrW.argtypes = (wt.HWND, ctypes.c_int)
user32.GetWindowLongPtrW.restype = ctypes.c_ssize_t
user32.GetWindow.argtypes = (wt.HWND, wt.UINT)
user32.GetWindow.restype = wt.HWND

user32.GetForegroundWindow.restype = wt.HWND
user32.SetForegroundWindow.argtypes = (wt.HWND,)
user32.SetForegroundWindow.restype = wt.BOOL
user32.ShowWindow.argtypes = (wt.HWND, ctypes.c_int)
user32.ShowWindow.restype = wt.BOOL
user32.SetWindowPos.argtypes = (wt.HWND, wt.HWND, ctypes.c_int, ctypes.c_int,
                                ctypes.c_int, ctypes.c_int, wt.UINT)
user32.SetWindowPos.restype = wt.BOOL
user32.AttachThreadInput.argtypes = (wt.DWORD, wt.DWORD, wt.BOOL)
user32.AttachThreadInput.restype = wt.BOOL
user32.SwitchToThisWindow.argtypes = (wt.HWND, wt.BOOL)
user32.SwitchToThisWindow.restype = None

user32.SendMessageW.argtypes = (wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM)
user32.SendMessageW.restype = wt.LPARAM
user32.PostMessageW.argtypes = (wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM)
user32.PostMessageW.restype = wt.BOOL

# 专门用于传字符串 / 缓冲区的 SendMessageW 原型
SendMessageStr = ctypes.WINFUNCTYPE(wt.LPARAM, wt.HWND, wt.UINT, wt.WPARAM,
                                    ctypes.c_wchar_p)(("SendMessageW", user32))

user32.SetCursorPos.argtypes = (ctypes.c_int, ctypes.c_int)
user32.SetCursorPos.restype = wt.BOOL
user32.mouse_event.argtypes = (wt.DWORD, wt.LONG, wt.LONG, wt.DWORD, ctypes.c_size_t)
user32.mouse_event.restype = None
user32.VkKeyScanW.argtypes = (ctypes.c_wchar,)
user32.VkKeyScanW.restype = ctypes.c_short

kernel32.GetCurrentThreadId.restype = wt.DWORD
kernel32.OpenProcess.argtypes = (wt.DWORD, wt.BOOL, wt.DWORD)
kernel32.OpenProcess.restype = wt.HANDLE
kernel32.CloseHandle.argtypes = (wt.HANDLE,)
kernel32.QueryFullProcessImageNameW.argtypes = (wt.HANDLE, wt.DWORD, wt.LPWSTR,
                                                ctypes.POINTER(wt.DWORD))
kernel32.QueryFullProcessImageNameW.restype = wt.BOOL


# ---------------------------------------------------------------------------
# SendInput 结构体
# ---------------------------------------------------------------------------
class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wt.WORD),
        ("wScan", wt.WORD),
        ("dwFlags", wt.DWORD),
        ("time", wt.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wt.LONG),
        ("dy", wt.LONG),
        ("mouseData", wt.DWORD),
        ("dwFlags", wt.DWORD),
        ("time", wt.DWORD),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wt.DWORD),
        ("wParamL", wt.WORD),
        ("wParamH", wt.WORD),
    ]


class INPUTUNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _anonymous_ = ("u",)
    _fields_ = [("type", wt.DWORD), ("u", INPUTUNION)]


user32.SendInput.argtypes = (wt.UINT, ctypes.POINTER(INPUT), ctypes.c_int)
user32.SendInput.restype = wt.UINT

INPUT_SIZE = ctypes.sizeof(INPUT)


# ---------------------------------------------------------------------------
# 窗口信息
# ---------------------------------------------------------------------------
_proc_cache: dict[int, str] = {}


def _process_name(pid: int) -> str:
    if pid in _proc_cache:
        return _proc_cache[pid]
    name = ""
    handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if handle:
        try:
            buf = ctypes.create_unicode_buffer(1024)
            size = wt.DWORD(len(buf))
            if kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
                name = os.path.splitext(os.path.basename(buf.value))[0]
        finally:
            kernel32.CloseHandle(handle)
    _proc_cache[pid] = name
    return name


@dataclass
class WinInfo:
    hwnd: int
    title: str
    cls: str
    pid: int
    process: str
    left: int
    top: int
    width: int
    height: int
    visible: bool
    minimized: bool
    enabled: bool
    style: int
    exstyle: int
    owner: int

    def as_dict(self) -> dict:
        return {
            "Handle": f"0x{self.hwnd:08X}",
            "PID": self.pid,
            "Process": self.process,
            "Class": self.cls,
            "Visible": self.visible,
            "Minimized": self.minimized,
            "Enabled": self.enabled,
            "Rect": f"{self.left},{self.top} {self.width}x{self.height}",
            "Title": self.title,
        }


def _window_info(hwnd: int) -> WinInfo:
    title_buf = ctypes.create_unicode_buffer(1024)
    user32.GetWindowTextW(hwnd, title_buf, len(title_buf))
    class_buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_buf, len(class_buf))

    pid = wt.DWORD(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

    rect = wt.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))

    style = int(user32.GetWindowLongPtrW(hwnd, GWL_STYLE)) & 0xFFFFFFFF
    exstyle = int(user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)) & 0xFFFFFFFF
    owner = int(user32.GetWindow(hwnd, GW_OWNER) or 0)

    return WinInfo(
        hwnd=int(hwnd),
        title=title_buf.value,
        cls=class_buf.value,
        pid=int(pid.value),
        process=_process_name(int(pid.value)),
        left=rect.left,
        top=rect.top,
        width=rect.right - rect.left,
        height=rect.bottom - rect.top,
        visible=bool(user32.IsWindowVisible(hwnd)),
        minimized=bool(user32.IsIconic(hwnd)),
        enabled=bool(user32.IsWindowEnabled(hwnd)),
        style=style,
        exstyle=exstyle,
        owner=owner,
    )


def get_windows(visible_only: bool = True) -> list[WinInfo]:
    """枚举所有顶层窗口。"""
    result: list[WinInfo] = []

    @WNDENUMPROC
    def _callback(hwnd, _lparam):
        if not visible_only or user32.IsWindowVisible(hwnd):
            result.append(_window_info(hwnd))
        return True

    user32.EnumWindows(_callback, 0)
    return result


def get_children(parent: int) -> list[WinInfo]:
    """枚举某个窗口的全部子控件。"""
    result: list[WinInfo] = []

    @WNDENUMPROC
    def _callback(hwnd, _lparam):
        result.append(_window_info(hwnd))
        return True

    user32.EnumChildWindows(parent, _callback, 0)
    return result


def is_window(hwnd: int) -> bool:
    return bool(user32.IsWindow(hwnd))


def get_foreground() -> WinInfo:
    return _window_info(int(user32.GetForegroundWindow()))


def get_foreground_handle() -> int:
    return int(user32.GetForegroundWindow() or 0)


# ---------------------------------------------------------------------------
# 窗口操作
# ---------------------------------------------------------------------------
def force_foreground(hwnd: int) -> bool:
    """把窗口强制带到前台（AttachThreadInput 绕过前台锁定）。"""
    if not is_window(hwnd):
        return False
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    user32.ShowWindow(hwnd, SW_SHOW)
    user32.SetWindowPos(hwnd, None, 0, 0, 0, 0,
                        SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW)

    if get_foreground_handle() == hwnd:
        return True

    fg_thread = user32.GetWindowThreadProcessId(user32.GetForegroundWindow(), None)
    current = kernel32.GetCurrentThreadId()
    attached = False
    if fg_thread and fg_thread != current:
        attached = bool(user32.AttachThreadInput(current, fg_thread, True))

    ok = bool(user32.SetForegroundWindow(hwnd))

    if attached:
        user32.AttachThreadInput(current, fg_thread, False)

    time.sleep(0.05)
    if get_foreground_handle() != hwnd:
        user32.SwitchToThisWindow(hwnd, True)
        time.sleep(0.08)
    return get_foreground_handle() == hwnd


def _send_inputs(inputs: Iterable[INPUT]) -> None:
    items = list(inputs)
    if not items:
        return
    array = (INPUT * len(items))(*items)
    sent = user32.SendInput(len(items), array, INPUT_SIZE)
    if sent != len(items):
        raise ctypes.WinError(ctypes.get_last_error())


def _unicode_input(ch: str, keyup: bool) -> INPUT:
    item = INPUT()
    item.type = INPUT_KEYBOARD
    item.ki.wVk = 0
    item.ki.wScan = ord(ch)
    item.ki.dwFlags = KEYEVENTF_UNICODE | (KEYEVENTF_KEYUP if keyup else 0)
    return item


def _vk_input(vk: int, keyup: bool) -> INPUT:
    item = INPUT()
    item.type = INPUT_KEYBOARD
    item.ki.wVk = vk
    item.ki.wScan = 0
    item.ki.dwFlags = KEYEVENTF_KEYUP if keyup else 0
    return item


def type_text(text: str, interval: float = 0.008) -> None:
    """逐字符发送 Unicode 键盘事件，中文 / emoji / 符号都能输入。"""
    for ch in text:
        if ch == "\r":
            continue
        if ch == "\n":
            _send_inputs([_vk_input(VK_RETURN, False), _vk_input(VK_RETURN, True)])
        elif ch == "\t":
            _send_inputs([_vk_input(VK_TAB, False), _vk_input(VK_TAB, True)])
        else:
            _send_inputs([_unicode_input(ch, False), _unicode_input(ch, True)])
        time.sleep(interval)


_KEY_NAMES = {
    "backspace": 0x08, "bs": 0x08,
    "tab": 0x09,
    "enter": 0x0D, "return": 0x0D,
    "shift": VK_SHIFT,
    "ctrl": VK_CONTROL, "control": VK_CONTROL,
    "alt": VK_MENU,
    "pause": 0x13,
    "capslock": 0x14,
    "esc": 0x1B, "escape": 0x1B,
    "space": 0x20,
    "pageup": 0x21, "pgup": 0x21,
    "pagedown": 0x22, "pgdn": 0x22,
    "end": 0x23,
    "home": 0x24,
    "left": 0x25,
    "up": 0x26,
    "right": 0x27,
    "down": 0x28,
    "printscreen": 0x2C, "prtsc": 0x2C,
    "insert": 0x2D, "ins": 0x2D,
    "delete": 0x2E, "del": 0x2E,
    "win": VK_LWIN, "lwin": VK_LWIN,
    "apps": 0x5D,
    "multiply": 0x6A, "add": 0x6B, "subtract": 0x6D, "decimal": 0x6E, "divide": 0x6F,
}
for _i in range(1, 25):
    _KEY_NAMES[f"f{_i}"] = 0x70 + _i - 1


def _name_to_vk(name: str) -> int:
    lowered = name.lower()
    if lowered in _KEY_NAMES:
        return _KEY_NAMES[lowered]
    if len(name) == 1:
        scan = user32.VkKeyScanW(name)
        if scan != -1:
            return scan & 0xFF
    return 0


def press_key(key: str, ctrl: bool = False, alt: bool = False,
              shift: bool = False, win: bool = False) -> None:
    if not key:
        return

    # 单字符 + 无修饰键：直接走 Unicode，中文也能打
    if len(key) == 1 and not (ctrl or alt or win):
        type_text(key)
        return

    need_shift = False
    if len(key) == 1:
        scan = user32.VkKeyScanW(key)
        if scan == -1:
            return
        vk = scan & 0xFF
        need_shift = bool(scan & 0x100)
    else:
        vk = _name_to_vk(key)

    if not vk:
        return

    sequence: list[INPUT] = []
    if ctrl:
        sequence.append(_vk_input(VK_CONTROL, False))
    if alt:
        sequence.append(_vk_input(VK_MENU, False))
    if shift or need_shift:
        sequence.append(_vk_input(VK_SHIFT, False))
    if win:
        sequence.append(_vk_input(VK_LWIN, False))

    sequence.append(_vk_input(vk, False))
    sequence.append(_vk_input(vk, True))

    if win:
        sequence.append(_vk_input(VK_LWIN, True))
    if shift or need_shift:
        sequence.append(_vk_input(VK_SHIFT, True))
    if alt:
        sequence.append(_vk_input(VK_MENU, True))
    if ctrl:
        sequence.append(_vk_input(VK_CONTROL, True))

    _send_inputs(sequence)
    time.sleep(0.03)


def send_combo(combo: str) -> None:
    """发送组合键，例如 ctrl+shift+s / alt+f4 / enter。"""
    if not combo:
        return
    parts = combo.split("+")
    ctrl = alt = shift = win = False
    for modifier in parts[:-1]:
        m = modifier.strip().lower()
        if m in ("ctrl", "control"):
            ctrl = True
        elif m == "alt":
            alt = True
        elif m == "shift":
            shift = True
        elif m in ("win", "lwin"):
            win = True
    press_key(parts[-1].strip(), ctrl, alt, shift, win)


def click_at(x: int, y: int, double: bool = False, right: bool = False) -> None:
    user32.SetCursorPos(int(x), int(y))
    time.sleep(0.04)
    down = MOUSEEVENTF_RIGHTDOWN if right else MOUSEEVENTF_LEFTDOWN
    up = MOUSEEVENTF_RIGHTUP if right else MOUSEEVENTF_LEFTUP
    user32.mouse_event(down, 0, 0, 0, 0)
    user32.mouse_event(up, 0, 0, 0, 0)
    if double:
        time.sleep(0.07)
        user32.mouse_event(down, 0, 0, 0, 0)
        user32.mouse_event(up, 0, 0, 0, 0)
    time.sleep(0.05)


def get_control_text(hwnd: int) -> str:
    """读取控件文本：WM_GETTEXT；Scintilla 控件回退到 SCI_GETTEXT。"""
    length = int(user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0))
    if length < 0:
        length = 0
    buf = ctypes.create_unicode_buffer(length + 2)
    user32.SendMessageW(hwnd, WM_GETTEXT, length + 1, ctypes.addressof(buf))
    text = buf.value

    if not text:
        class_buf = ctypes.create_unicode_buffer(64)
        user32.GetClassNameW(hwnd, class_buf, len(class_buf))
        if "scintilla" in class_buf.value.lower():
            slen = int(user32.SendMessageW(hwnd, SCI_GETLENGTH, 0, 0))
            if slen > 0:
                sbuf = ctypes.create_unicode_buffer(slen + 2)
                user32.SendMessageW(hwnd, SCI_GETTEXT, slen + 1, ctypes.addressof(sbuf))
                text = sbuf.value
    return text


def set_control_text(hwnd: int, text: str) -> None:
    SendMessageStr(hwnd, WM_SETTEXT, 0, text or "")


def post_text(hwnd: int, text: str) -> None:
    """向控件投递 WM_CHAR（不需要前台焦点）。"""
    for ch in text:
        if ch == "\r":
            continue
        if ch == "\n":
            user32.PostMessageW(hwnd, WM_CHAR, 0x0D, 0)
        else:
            user32.PostMessageW(hwnd, WM_CHAR, ord(ch), 0)
    time.sleep(0.06)


def close_window(hwnd: int) -> None:
    user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)


# ---------------------------------------------------------------------------
# 筛选 / 选择
# ---------------------------------------------------------------------------
def select_windows(args: argparse.Namespace, allow_unfiltered: bool = False) -> list[WinInfo]:
    has_filter = any([args.pattern, args.class_name, args.process_name,
                      args.hwnd, args.pid])
    if not allow_unfiltered and not has_filter:
        raise SystemExit("请至少指定一个筛选条件：--pattern / --class / --process / --hwnd / --pid")

    windows = get_windows(not args.include_invisible)

    if args.hwnd:
        windows = [w for w in windows if w.hwnd == args.hwnd]
    if args.pid:
        windows = [w for w in windows if w.pid == args.pid]
    if args.process_name:
        windows = [w for w in windows
                   if fnmatch.fnmatch(w.process.lower(), args.process_name.lower())]
    if args.class_name:
        windows = [w for w in windows
                   if fnmatch.fnmatch(w.cls.lower(), args.class_name.lower())]
    if args.pattern:
        if args.regex:
            regex = re.compile(args.pattern, re.IGNORECASE)
            windows = [w for w in windows if regex.search(w.title)]
        else:
            needle = args.pattern.lower()
            windows = [w for w in windows if needle in w.title.lower()]

    windows.sort(key=lambda w: (w.pid, w.title))
    return windows


def pick_one(args: argparse.Namespace) -> WinInfo:
    windows = select_windows(args)
    if not windows:
        raise SystemExit("没有找到匹配的窗口，请检查筛选条件")
    if args.index < 0 or args.index >= len(windows):
        raise SystemExit(f"索引 {args.index} 越界，共找到 {len(windows)} 个窗口")
    return windows[args.index]


def pick_control(win: WinInfo, control_class: Optional[str] = None) -> WinInfo:
    children = get_children(win.hwnd)
    if not children:
        raise SystemExit(f"窗口 [{win.title}] 中没有子控件")

    if control_class:
        matched = [c for c in children
                   if fnmatch.fnmatch(c.cls.lower(), control_class.lower())]
        if not matched:
            raise SystemExit(f"未找到类名匹配 '{control_class}' 的子控件")
        return matched[0]

    textish = [c for c in children
               if re.match(r"^(edit|richedit|scintilla|textbox|chrome_render)", c.cls, re.I)]
    if textish:
        return textish[0]
    return max(children, key=lambda c: c.width * c.height)


# ---------------------------------------------------------------------------
# 输出
# ---------------------------------------------------------------------------
def _display_width(text: str) -> int:
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in text)


def _pad(text: str, width: int) -> str:
    return text + " " * max(0, width - _display_width(text))


def print_table(rows: list[list[str]], headers: list[str]) -> None:
    widths = [max(_display_width(str(row[i])) for row in [headers] + rows)
              for i in range(len(headers))]
    print("  ".join(_pad(h, w) for h, w in zip(headers, widths)))
    print("  ".join("-" * w for w in widths))
    for row in rows:
        print("  ".join(_pad(str(cell), w) for cell, w in zip(row, widths)))


def print_windows(windows: list[WinInfo]) -> None:
    rows = [[f"0x{w.hwnd:08X}", str(w.pid), w.process, w.cls,
             str(w.visible), f"{w.left},{w.top} {w.width}x{w.height}", w.title]
            for w in windows]
    print_table(rows, ["Handle", "PID", "Process", "Class", "Visible", "Rect", "Title"])


def print_detail(win: WinInfo) -> None:
    print(f"Handle    : 0x{win.hwnd:08X}")
    print(f"PID       : {win.pid}")
    print(f"Process   : {win.process}")
    print(f"Class     : {win.cls}")
    print(f"Visible   : {win.visible}")
    print(f"Minimized : {win.minimized}")
    print(f"Enabled   : {win.enabled}")
    print(f"Rect      : {win.left},{win.top} {win.width}x{win.height}")
    print(f"Style     : 0x{win.style:08X}")
    print(f"ExStyle   : 0x{win.exstyle:08X}")
    print(f"Owner     : 0x{win.owner:08X}")
    print(f"Title     : {win.title}")


# ---------------------------------------------------------------------------
# 命令行
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--pattern", help="按标题筛选（默认包含匹配，不区分大小写）")
    common.add_argument("--class", dest="class_name", help="按窗口类名筛选（支持通配符）")
    common.add_argument("--process", dest="process_name", help="按进程名筛选（支持通配符）")
    common.add_argument("--hwnd", type=lambda s: int(s, 0), help="按窗口句柄定位，如 0x1A2B 或 6699")
    common.add_argument("--pid", type=int, help="按进程 ID 定位")
    common.add_argument("--regex", action="store_true", help="标题按正则匹配")
    common.add_argument("--include-invisible", action="store_true", help="包含隐藏窗口")
    common.add_argument("--index", type=int, default=0, help="匹配多个时取第几个（从 0 开始）")
    common.add_argument("--json", action="store_true", help="以 JSON 输出")

    parser = argparse.ArgumentParser(
        description="Windows 窗口枚举与自动化操作工具（Python / ctypes）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = parser.add_subparsers(dest="action", required=True)

    sub.add_parser("list", parents=[common], help="列出窗口")
    sub.add_parser("info", parents=[common], help="窗口详细信息 + 子控件")
    sub.add_parser("children", parents=[common], help="只列出子控件")
    sub.add_parser("activate", parents=[common], help="激活窗口到前台")
    sub.add_parser("foreground", parents=[common], help="显示当前前台窗口")

    p_type = sub.add_parser("type", parents=[common], help="向窗口输入文本（真实键盘事件）")
    p_type.add_argument("--text", required=True, help="要输入的文本")
    p_type.add_argument("--message-mode", action="store_true",
                        help="改用 PostMessage(WM_CHAR)，不需要前台焦点")
    p_type.add_argument("--control-class", help="message-mode 下指定控件类名（支持通配符）")
    p_type.add_argument("--no-activate", action="store_true", help="输入前不激活窗口")

    p_keys = sub.add_parser("keys", parents=[common], help="发送组合键")
    p_keys.add_argument("--keys", nargs="+", required=True,
                        help="组合键，如 ctrl+a alt+f4 enter")
    p_keys.add_argument("--no-activate", action="store_true", help="发送前不激活窗口")

    p_click = sub.add_parser("click", parents=[common], help="在窗口内点击")
    p_click.add_argument("--x", type=int, help="相对窗口左上角的 X 坐标（默认居中）")
    p_click.add_argument("--y", type=int, help="相对窗口左上角的 Y 坐标（默认居中）")
    p_click.add_argument("--double", action="store_true", help="双击")
    p_click.add_argument("--right", action="store_true", help="右键")
    p_click.add_argument("--no-activate", action="store_true", help="点击前不激活窗口")

    p_get = sub.add_parser("gettext", parents=[common], help="读取控件文本")
    p_get.add_argument("--control-class", help="指定控件类名（支持通配符）")

    p_set = sub.add_parser("settext", parents=[common], help="直接写入控件文本（WM_SETTEXT）")
    p_set.add_argument("--text", required=True, help="要写入的文本")
    p_set.add_argument("--control-class", help="指定控件类名（支持通配符）")

    sub.add_parser("close", parents=[common], help="发送 WM_CLOSE 关闭窗口")

    return parser


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    action = args.action

    if action == "list":
        windows = select_windows(args, allow_unfiltered=True)
        if not windows:
            print("没有匹配的窗口")
            return 0
        if args.json:
            print(json.dumps([w.as_dict() for w in windows], ensure_ascii=False, indent=2))
        else:
            print_windows(windows)
        return 0

    if action == "foreground":
        win = get_foreground()
        if args.json:
            print(json.dumps(win.as_dict(), ensure_ascii=False, indent=2))
        else:
            print_detail(win)
        return 0

    if action == "info":
        win = pick_one(args)
        children = get_children(win.hwnd)
        if args.json:
            print(json.dumps({
                "Window": win.as_dict(),
                "Style": f"0x{win.style:08X}",
                "ExStyle": f"0x{win.exstyle:08X}",
                "Owner": f"0x{win.owner:08X}",
                "Children": [c.as_dict() for c in children],
            }, ensure_ascii=False, indent=2))
        else:
            print_detail(win)
            print()
            print(f"子控件（{len(children)} 个）:")
            if children:
                print_windows(children)
        return 0

    if action == "children":
        win = pick_one(args)
        children = get_children(win.hwnd)
        if not children:
            print(f"窗口 [{win.title}] 没有子控件")
            return 0
        if args.json:
            print(json.dumps([c.as_dict() for c in children], ensure_ascii=False, indent=2))
        else:
            print_windows(children)
        return 0

    if action == "activate":
        win = pick_one(args)
        ok = force_foreground(win.hwnd)
        print(f"已激活窗口: [{win.title}]  句柄=0x{win.hwnd:08X}  成功={ok}")
        return 0

    if action == "type":
        win = pick_one(args)
        if not args.no_activate:
            force_foreground(win.hwnd)
        if args.message_mode:
            control = pick_control(win, args.control_class)
            post_text(control.hwnd, args.text)
            print(f"已通过 WM_CHAR 向控件 [{control.cls}] 发送 {len(args.text)} 个字符")
        else:
            type_text(args.text)
            print(f"已通过 SendInput 向窗口 [{win.title}] 输入 {len(args.text)} 个字符")
        return 0

    if action == "keys":
        win = pick_one(args)
        if not args.no_activate:
            force_foreground(win.hwnd)
        for combo in args.keys:
            send_combo(combo)
            time.sleep(0.1)
        print(f"已向 [{win.title}] 发送按键: {', '.join(args.keys)}")
        return 0

    if action == "click":
        win = pick_one(args)
        if not args.no_activate:
            force_foreground(win.hwnd)
        x = win.left + (args.x if args.x is not None else win.width // 2)
        y = win.top + (args.y if args.y is not None else win.height // 2)
        click_at(x, y, args.double, args.right)
        print(f"已在 [{win.title}] 内点击屏幕坐标 ({x}, {y})")
        return 0

    if action == "gettext":
        win = pick_one(args)
        control = pick_control(win, args.control_class)
        content = get_control_text(control.hwnd)
        if args.json:
            print(json.dumps({
                "Window": win.as_dict(),
                "Control": control.as_dict(),
                "Text": content,
            }, ensure_ascii=False, indent=2))
        else:
            print(f"窗口: [{win.title}]")
            print(f"控件: {control.cls}  句柄=0x{control.hwnd:08X}")
            print("--- 文本开始 ---")
            print(content)
            print("--- 文本结束 ---")
        return 0

    if action == "settext":
        win = pick_one(args)
        control = pick_control(win, args.control_class)
        set_control_text(control.hwnd, args.text)
        print(f"已通过 WM_SETTEXT 写入控件 [{control.cls}]，共 {len(args.text)} 个字符")
        return 0

    if action == "close":
        win = pick_one(args)
        close_window(win.hwnd)
        print(f"已向 [{win.title}] 发送 WM_CLOSE")
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
