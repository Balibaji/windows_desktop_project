#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RDP / 位图窗口自动化工具（Python 版，纯标准库 + ctypes）

远程桌面的画面在本地只是一张位图，本地 UI Automation / WM_GETTEXT / WM_SETTEXT
都看不到远端控件，唯一能用的通道是“真实键鼠输入”（SendInput）+ 截图 / OCR 反馈。

本脚本在 window_ops.py 的基础上提供：
  find         查找 RDP 窗口（默认 mstsc / TscShell*，也支持任意窗口）
  status       窗口 + 客户区（位图区域）在屏幕上的位置、尺寸、DPI
  open-cmd     激活 RDP 窗口 → Win+R → 输入 cmd → 回车
  run-ps       在远端执行 PowerShell（短脚本用 EncodedCommand，长脚本用剪贴板粘贴）
  type         向远端的当前焦点窗口输入文本
  keys         向远端发送组合键
  click        在 RDP 位图区域内点击（默认客户区坐标，也可用窗口坐标）
  screenshot   截取 RDP 客户区为 PNG / BMP
  ocr          对截图做 OCR（调用 RdpOps.ps1 的 Windows.Media.Ocr）
  click-text   截图 → OCR → 找到文字 → 点击文字中心
  wait-text    轮询 OCR，直到文字出现或超时
  clip-set / clip-get / paste   本机剪贴板操作（RDP 默认启用剪贴板重定向）
  tunnel       发现 mstsc → 127.0.0.1:<port> 的本地端口隧道（AccessClient 场景）
  wait-port    等待本机端口进入 LISTEN（隧道就绪检查）
  health       窗口 + 隧道的综合健康检查（窗口存在 / 未挂起 / 未最小化 / 端口活着）

示例：
  python rdp_ops.py find
  python rdp_ops.py status
  python rdp_ops.py open-cmd
  python rdp_ops.py run-ps --script "Get-Date; whoami"
  python rdp_ops.py screenshot --out shot.png
  python rdp_ops.py click-text --text "确定"
  python rdp_ops.py wait-text --text "完成" --timeout 30

  python rdp_ops.py tunnel
  python rdp_ops.py tunnel --watch --json
  python rdp_ops.py tunnel --app AccessClient --timeout 30
  python rdp_ops.py wait-port --port 52341 --timeout 20
  python rdp_ops.py health --any --process mstsc --screen

本地测试（不连 RDP 也能验证截图 / OCR）：
  python rdp_ops.py find --any --process notepad
  python rdp_ops.py screenshot --any --process notepad --out notepad.png
  python rdp_ops.py ocr --image notepad.png

注意：
  * SendInput 需要目标窗口在前台，自动化期间不要动本机键鼠。
  * Win+R 在“窗口模式”下可能被本机截获（默认设置：仅全屏时发给远端）。
    脚本会检测本地“运行”对话框并报错，此时请把 mstsc 的
    “本地资源 → 键盘 → 应用 Windows 键组合”改为“在远程计算机上”，或使用全屏。
  * 截图只能截到屏幕上可见的内容，窗口不能被其他窗口遮挡。
"""

from __future__ import annotations

import argparse
import base64
import ctypes
import ctypes.wintypes as wt
import fnmatch
import json
import os
import re
import socket
import struct
import subprocess
import sys
import tempfile
import time
import zlib
from pathlib import Path
from typing import Optional

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import window_ops as wo  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # type: ignore
    except Exception:
        pass

user32 = wo.user32
kernel32 = wo.kernel32
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

# ---------------------------------------------------------------------------
# 常量 / Win32 原型
# ---------------------------------------------------------------------------
SRCCOPY = 0x00CC0020
DIB_RGB_COLORS = 0
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002

gdi32.CreateCompatibleDC.argtypes = (wt.HDC,)
gdi32.CreateCompatibleDC.restype = wt.HDC
gdi32.CreateCompatibleBitmap.argtypes = (wt.HDC, ctypes.c_int, ctypes.c_int)
gdi32.CreateCompatibleBitmap.restype = wt.HBITMAP
gdi32.SelectObject.argtypes = (wt.HDC, wt.HGDIOBJ)
gdi32.SelectObject.restype = wt.HGDIOBJ
gdi32.BitBlt.argtypes = (wt.HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int,
                         ctypes.c_int, wt.HDC, ctypes.c_int, ctypes.c_int, wt.DWORD)
gdi32.BitBlt.restype = wt.BOOL
gdi32.DeleteObject.argtypes = (wt.HGDIOBJ,)
gdi32.DeleteObject.restype = wt.BOOL
gdi32.DeleteDC.argtypes = (wt.HDC,)
gdi32.DeleteDC.restype = wt.BOOL
gdi32.GetDIBits.argtypes = (wt.HDC, wt.HBITMAP, wt.UINT, wt.UINT, ctypes.c_void_p,
                            ctypes.c_void_p, wt.UINT)
gdi32.GetDIBits.restype = ctypes.c_int

user32.GetDC.argtypes = (wt.HWND,)
user32.GetDC.restype = wt.HDC
user32.ReleaseDC.argtypes = (wt.HWND, wt.HDC)
user32.ReleaseDC.restype = ctypes.c_int
user32.GetClientRect.argtypes = (wt.HWND, ctypes.POINTER(wt.RECT))
user32.GetClientRect.restype = wt.BOOL
user32.ClientToScreen.argtypes = (wt.HWND, ctypes.POINTER(wt.POINT))
user32.ClientToScreen.restype = wt.BOOL

user32.OpenClipboard.argtypes = (wt.HWND,)
user32.OpenClipboard.restype = wt.BOOL
user32.EmptyClipboard.restype = wt.BOOL
user32.CloseClipboard.restype = wt.BOOL
user32.SetClipboardData.argtypes = (wt.UINT, wt.HANDLE)
user32.SetClipboardData.restype = wt.HANDLE
user32.GetClipboardData.argtypes = (wt.UINT,)
user32.GetClipboardData.restype = wt.HANDLE

kernel32.GlobalAlloc.argtypes = (wt.UINT, ctypes.c_size_t)
kernel32.GlobalAlloc.restype = wt.HGLOBAL
kernel32.GlobalLock.argtypes = (wt.HGLOBAL,)
kernel32.GlobalLock.restype = ctypes.c_void_p
kernel32.GlobalUnlock.argtypes = (wt.HGLOBAL,)
kernel32.GlobalUnlock.restype = wt.BOOL
kernel32.GlobalFree.argtypes = (wt.HGLOBAL,)
kernel32.GlobalFree.restype = wt.HGLOBAL


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wt.DWORD),
        ("biWidth", wt.LONG),
        ("biHeight", wt.LONG),
        ("biPlanes", wt.WORD),
        ("biBitCount", wt.WORD),
        ("biCompression", wt.DWORD),
        ("biSizeImage", wt.DWORD),
        ("biXPelsPerMeter", wt.LONG),
        ("biYPelsPerMeter", wt.LONG),
        ("biClrUsed", wt.DWORD),
        ("biClrImportant", wt.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wt.DWORD * 3)]


# ---------------------------------------------------------------------------
# 窗口定位
# ---------------------------------------------------------------------------
def is_rdp_window(win: wo.WinInfo) -> bool:
    proc = win.process.lower()
    cls = win.cls.lower()
    return proc in ("mstsc", "msrdc") or cls.startswith("tscshell") or cls.startswith("rail")


def _match_filters(win: wo.WinInfo, args: argparse.Namespace) -> bool:
    if args.hwnd and win.hwnd != args.hwnd:
        return False
    if args.pid and win.pid != args.pid:
        return False
    if args.process_name and not fnmatch.fnmatch(win.process.lower(),
                                                 args.process_name.lower()):
        return False
    if args.class_name and not fnmatch.fnmatch(win.cls.lower(),
                                               args.class_name.lower()):
        return False
    if args.pattern:
        if args.regex:
            if not re.search(args.pattern, win.title, re.IGNORECASE):
                return False
        elif args.pattern.lower() not in win.title.lower():
            return False
    return True


def select_windows(args: argparse.Namespace) -> list[wo.WinInfo]:
    windows = [w for w in wo.get_windows(not args.include_invisible)
               if _match_filters(w, args)]
    has_filter = any([args.pattern, args.class_name, args.process_name,
                      args.hwnd, args.pid])
    if not has_filter and not args.any:
        windows = [w for w in windows if is_rdp_window(w)]
    windows.sort(key=lambda w: (w.pid, w.title))
    return windows


def pick_window(args: argparse.Namespace) -> wo.WinInfo:
    windows = select_windows(args)
    if not windows:
        raise SystemExit(
            "没有找到目标窗口。\n"
            "  * 连上 RDP 后直接运行（默认找 mstsc / msrdc / TscShell*）；\n"
            "  * 或用 --any --pattern/--class/--process/--hwnd 指定任意窗口。"
        )
    if args.index < 0 or args.index >= len(windows):
        raise SystemExit(f"索引 {args.index} 越界，共找到 {len(windows)} 个窗口")
    return windows[args.index]


def get_client_rect(hwnd: int) -> tuple[int, int, int, int]:
    """返回客户区在屏幕上的 (left, top, width, height)。"""
    rect = wt.RECT()
    if not user32.GetClientRect(hwnd, ctypes.byref(rect)):
        raise ctypes.WinError(ctypes.get_last_error())
    origin = wt.POINT(0, 0)
    if not user32.ClientToScreen(hwnd, ctypes.byref(origin)):
        raise ctypes.WinError(ctypes.get_last_error())
    return (int(origin.x), int(origin.y),
            int(rect.right - rect.left), int(rect.bottom - rect.top))


def get_dpi_scale(hwnd: int) -> float:
    try:
        user32.GetDpiForWindow.argtypes = (wt.HWND,)
        user32.GetDpiForWindow.restype = wt.UINT
        return round(float(user32.GetDpiForWindow(hwnd)) / 96.0, 3)
    except AttributeError:
        return 1.0
    except OSError:
        return 1.0


def window_dict(win: wo.WinInfo) -> dict:
    data = win.as_dict()
    try:
        left, top, width, height = get_client_rect(win.hwnd)
        data.update({
            "ClientOrigin": f"{left},{top}",
            "ClientSize": f"{width}x{height}",
            "DpiScale": get_dpi_scale(win.hwnd),
            "Foreground": wo.get_foreground_handle() == win.hwnd,
            "IsRdp": is_rdp_window(win),
        })
    except Exception:
        pass
    return data


# ---------------------------------------------------------------------------
# 本地隧道感知（AccessClient / mstsc → 127.0.0.1:<port>）
# ---------------------------------------------------------------------------
# 堡垒机客户端（AccessClient 等）通常会在本机回环上开一个 TCP 代理端口，
# 再以 `mstsc /v:127.0.0.1:<port>` 拉起远程桌面。这里用 iphlpapi 的
# GetExtendedTcpTable 直接读取 TCP 表，避免依赖 netstat / Get-NetTCPConnection。
AF_INET = 2
AF_INET6 = 23
TCP_TABLE_OWNER_PID_ALL = 5
TCP_TABLE_OWNER_PID_LISTENER = 3
TCP_STATE_ESTABLISHED = 5

iphlpapi = ctypes.WinDLL("iphlpapi", use_last_error=True)
iphlpapi.GetExtendedTcpTable.argtypes = (
    ctypes.c_void_p, ctypes.POINTER(wt.DWORD), wt.BOOL, wt.ULONG,
    ctypes.c_int, wt.DWORD)
iphlpapi.GetExtendedTcpTable.restype = wt.DWORD

user32.IsHungAppWindow.argtypes = (wt.HWND,)
user32.IsHungAppWindow.restype = wt.BOOL

_LOOPBACK_HOSTS = {"127.0.0.1", "localhost", "::1", "0:0:0:0:0:0:0:1",
                   "::ffff:127.0.0.1"}
_V_RE = re.compile(r"/v:[\"']?(?P<host>\[[^\]]+\]|[^\"'\s:]+)(?::(?P<port>\d+))?",
                   re.IGNORECASE)
_RDP_ADDRESS_RE = re.compile(
    r"full address:s:(?P<host>\[[^\]]+\]|[^:\s]+)(?::(?P<port>\d+))?",
    re.IGNORECASE)


def _tcp_rows(table_class: int, family: int) -> list[dict]:
    """读取 TCP 表，返回 State / LocalAddress / LocalPort / Remote* / Pid。"""
    size = wt.DWORD(0)
    iphlpapi.GetExtendedTcpTable(None, ctypes.byref(size), False, family,
                                 table_class, 0)
    if not size.value:
        return []

    buffer = None
    for _ in range(3):
        buffer = ctypes.create_string_buffer(size.value)
        err = iphlpapi.GetExtendedTcpTable(buffer, ctypes.byref(size), False,
                                           family, table_class, 0)
        if err == 0:
            break
        if err != 122:  # ERROR_INSUFFICIENT_BUFFER
            raise OSError(err, f"GetExtendedTcpTable 失败 (family={family})")
    else:
        raise OSError("GetExtendedTcpTable 连续返回缓冲区不足")

    raw = buffer.raw
    count = struct.unpack_from("<I", raw, 0)[0]
    rows: list[dict] = []
    if family == AF_INET:
        offset, step, fmt = 4, 24, "<6I"
        for _ in range(count):
            state, laddr, lport, raddr, rport, pid = struct.unpack_from(fmt, raw, offset)
            offset += step
            rows.append({
                "State": state,
                "LocalAddress": socket.inet_ntoa(struct.pack("<I", laddr)),
                "LocalPort": socket.ntohs(lport & 0xFFFF),
                "RemoteAddress": socket.inet_ntoa(struct.pack("<I", raddr)),
                "RemotePort": socket.ntohs(rport & 0xFFFF),
                "Pid": pid,
            })
    else:
        offset, step, fmt = 4, 56, "<16sII16sIIII"
        for _ in range(count):
            (laddr, _ls, lport, raddr, _rs, rport,
             state, pid) = struct.unpack_from(fmt, raw, offset)
            offset += step
            rows.append({
                "State": state,
                "LocalAddress": socket.inet_ntop(socket.AF_INET6, laddr),
                "LocalPort": socket.ntohs(lport & 0xFFFF),
                "RemoteAddress": socket.inet_ntop(socket.AF_INET6, raddr),
                "RemotePort": socket.ntohs(rport & 0xFFFF),
                "Pid": pid,
            })
    return rows


def _all_tcp_rows(table_class: int) -> list[dict]:
    rows: list[dict] = []
    for family in (AF_INET, AF_INET6):
        try:
            rows.extend(_tcp_rows(table_class, family))
        except OSError:
            pass
    return rows


def tcp_connections() -> list[dict]:
    return _all_tcp_rows(TCP_TABLE_OWNER_PID_ALL)


def tcp_listeners() -> list[dict]:
    return _all_tcp_rows(TCP_TABLE_OWNER_PID_LISTENER)


def _ps_json(script: str, timeout: int = 25):
    """执行一小段 PowerShell 并解析 JSON（仅用于读取进程命令行）。"""
    command = ("[Console]::OutputEncoding=[Text.Encoding]::UTF8;"
               "$ProgressPreference='SilentlyContinue';" + script)
    proc = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive",
         "-ExecutionPolicy", "Bypass", "-Command", command],
        capture_output=True, text=True, encoding="utf-8-sig",
        errors="replace", timeout=timeout)
    output = (proc.stdout or "").strip()
    if proc.returncode != 0:
        raise RuntimeError((proc.stderr or output).strip() or "PowerShell 执行失败")
    if not output:
        return None
    try:
        return json.loads(output)
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"PowerShell 输出不是 JSON: {output[:300]}") from exc


def list_client_processes(names: Optional[list[str]] = None) -> list[dict]:
    """列出 RDP 客户端进程（默认 mstsc / msrdc）及其命令行。"""
    names = names or ["mstsc.exe", "msrdc.exe"]
    where = " OR ".join(f"Name='{n}'" for n in names)
    script = (
        f"$items=@(Get-CimInstance Win32_Process -Filter \"{where}\" | "
        "Select-Object ProcessId,ParentProcessId,Name,ExecutablePath,CommandLine);"
        "ConvertTo-Json -InputObject $items -Compress -Depth 3"
    )
    data = _ps_json(script)
    if not data:
        return []
    if isinstance(data, dict):
        data = [data]
    result = []
    for item in data:
        result.append({
            "ProcessId": int(item.get("ProcessId") or 0),
            "ParentProcessId": int(item.get("ParentProcessId") or 0),
            "Name": item.get("Name") or "",
            "ExecutablePath": item.get("ExecutablePath") or "",
            "CommandLine": item.get("CommandLine") or "",
        })
    return result


def parse_loopback_endpoint(text: str) -> Optional[tuple[str, int]]:
    """从 mstsc 命令行（/v:）或 .rdp 内容（full address:s:）解析回环端口。"""
    if not text:
        return None
    for regex in (_V_RE, _RDP_ADDRESS_RE):
        match = regex.search(text)
        if not match:
            continue
        host = (match.group("host") or "").strip().strip("[]").lower()
        if host not in _LOOPBACK_HOSTS:
            continue
        return host, int(match.group("port") or 3389)
    return None


def discover_tunnels(port: Optional[int] = None,
                     pid: Optional[int] = None) -> list[dict]:
    """发现 mstsc → 127.0.0.1 的本地端口隧道。

    端口优先取命令行里的 /v:127.0.0.1:<port>；如果 mstsc 是用 .rdp 文件
    启动的，则回退到“该进程到回环地址的 ESTABLISHED 连接”推断端口。
    """
    processes = list_client_processes()
    connections = tcp_connections()
    listeners = tcp_listeners()
    windows = wo.get_windows(False)

    listening_by_port: dict[int, list[dict]] = {}
    for row in listeners:
        listening_by_port.setdefault(row["LocalPort"], []).append(row)

    results: list[dict] = []
    for proc in processes:
        proc_pid = proc["ProcessId"]
        if pid and proc_pid != pid:
            continue

        command_line = proc.get("CommandLine") or ""
        declared = parse_loopback_endpoint(command_line)
        observed = sorted({
            row["RemotePort"] for row in connections
            if row["Pid"] == proc_pid
            and row["State"] == TCP_STATE_ESTABLISHED
            and row["RemoteAddress"] in _LOOPBACK_HOSTS
            and row["RemotePort"] > 0
        })

        ports: list[int] = []
        if declared:
            ports.append(declared[1])
        for value in observed:
            if value not in ports:
                ports.append(value)
        if port:
            ports = [value for value in ports if value == port]
        if not ports:
            continue

        window = next((w for w in windows if w.pid == proc_pid), None)
        parent_pid = proc.get("ParentProcessId") or 0
        parent_name = wo._process_name(parent_pid) if parent_pid else ""

        for value in ports:
            listener_rows = listening_by_port.get(value) or []
            listener_pids = sorted({row["Pid"] for row in listener_rows})
            listener_name = (wo._process_name(listener_pids[0])
                             if listener_pids else "")
            connected = any(
                row["Pid"] == proc_pid
                and row["State"] == TCP_STATE_ESTABLISHED
                and row["RemotePort"] == value
                and row["RemoteAddress"] in _LOOPBACK_HOSTS
                for row in connections)

            results.append({
                "Pid": proc_pid,
                "Process": proc.get("Name") or wo._process_name(proc_pid),
                "ParentPid": parent_pid,
                "ParentProcess": parent_name,
                "Port": value,
                "Endpoint": f"127.0.0.1:{value}",
                "Declared": bool(declared and declared[1] == value),
                "Listening": bool(listener_rows),
                "ListenerPid": listener_pids[0] if listener_pids else 0,
                "ListenerProcess": listener_name,
                "Connected": connected,
                "Ready": bool(listener_rows) and connected,
                "WindowHandle": f"0x{window.hwnd:08X}" if window else "",
                "WindowTitle": window.title if window else "",
                "CommandLine": command_line,
            })

    results.sort(key=lambda item: (item["Pid"], item["Port"]))
    return results


def filter_tunnels(tunnels: list[dict], app: Optional[str]) -> list[dict]:
    if not app:
        return tunnels
    pattern = app.lower()
    return [item for item in tunnels
            if fnmatch.fnmatch((item.get("ParentProcess") or "").lower(), pattern)
            or fnmatch.fnmatch((item.get("ListenerProcess") or "").lower(), pattern)]


def port_status(port: int) -> dict:
    """查询本机端口的监听状态（只认 127.0.0.1 / ::1 / 通配地址）。"""
    rows = [row for row in tcp_listeners() if row["LocalPort"] == port]
    rows = [row for row in rows
            if row["LocalAddress"] in ("127.0.0.1", "::1", "0.0.0.0", "::")]
    pids = sorted({row["Pid"] for row in rows})
    return {
        "Port": port,
        "Listening": bool(rows),
        "Addresses": sorted({row["LocalAddress"] for row in rows}),
        "Pid": pids[0] if pids else 0,
        "Process": wo._process_name(pids[0]) if pids else "",
    }


def wait_port(port: int, timeout: float = 30.0, interval: float = 0.5,
              connect: bool = False) -> dict:
    """轮询等待端口进入 LISTEN，返回最后一次探测结果。"""
    deadline = time.time() + timeout
    status = port_status(port)
    while True:
        if status["Listening"]:
            if connect:
                try:
                    with socket.create_connection(("127.0.0.1", port), timeout=2):
                        status["TcpConnect"] = True
                except OSError as exc:
                    status["TcpConnect"] = False
                    status["ConnectError"] = str(exc)
            return status
        if time.time() >= deadline:
            return status
        time.sleep(max(0.05, interval))
        status = port_status(port)


def window_health(win: wo.WinInfo, check_screen: bool = False,
                  screen_interval: float = 0.6,
                  port_override: Optional[int] = None) -> dict:
    """窗口 + 隧道的综合健康检查。"""
    data = window_dict(win)
    reasons: list[str] = []

    hung = bool(user32.IsHungAppWindow(win.hwnd))
    if hung:
        reasons.append("window_not_responding")
    if not win.visible:
        reasons.append("window_invisible")
    if win.minimized:
        reasons.append("window_minimized")

    left, top, width, height = get_client_rect(win.hwnd)
    if width <= 0 or height <= 0:
        reasons.append("client_area_empty")

    tunnels = discover_tunnels(port=port_override, pid=win.pid)
    tunnel = tunnels[0] if tunnels else None
    if not tunnel:
        reasons.append("tunnel_not_found")
    else:
        if not tunnel["Listening"]:
            reasons.append("tunnel_port_not_listening")
        if not tunnel["Connected"]:
            reasons.append("tunnel_not_connected")

    screen = None
    if check_screen:
        if win.minimized or width <= 0 or height <= 0:
            screen = {"Checked": False, "Reason": "window_unavailable"}
        else:
            first, _, _ = capture_client(win.hwnd, default_screenshot_path(),
                                         activate=False)
            first_bytes = first.read_bytes()
            time.sleep(max(0.1, screen_interval))
            second, _, _ = capture_client(win.hwnd, default_screenshot_path(),
                                          activate=False)
            second_bytes = second.read_bytes()
            screen = {
                "Checked": True,
                "Changed": first_bytes != second_bytes,
                "First": str(first),
                "Second": str(second),
            }
            if first_bytes == second_bytes:
                reasons.append("screen_frozen")

    return {
        "Healthy": not reasons,
        "Reasons": reasons,
        "Window": data,
        "Hung": hung,
        "ClientArea": {"Left": left, "Top": top,
                       "Width": width, "Height": height},
        "Tunnel": tunnel,
        "Screen": screen,
    }


# ---------------------------------------------------------------------------
# 截图：BitBlt 屏幕客户区 → PNG / BMP（纯标准库）
# ---------------------------------------------------------------------------
def _write_bmp(path: Path, width: int, height: int, bgr_bottom_up: bytes) -> None:
    file_size = 14 + 40 + len(bgr_bottom_up)
    file_header = struct.pack("<2sIHHI", b"BM", file_size, 0, 0, 54)
    info_header = struct.pack("<IiiHHIIiiII", 40, width, height, 1, 24, 0,
                              len(bgr_bottom_up), 2835, 2835, 0, 0)
    path.write_bytes(file_header + info_header + bgr_bottom_up)


def _write_png(path: Path, width: int, height: int, bgr_bottom_up: bytes) -> None:
    stride = (width * 3 + 3) & ~3
    rows = []
    for y in range(height - 1, -1, -1):
        row = bgr_bottom_up[y * stride:y * stride + width * 3]
        rgb = bytearray(width * 3)
        rgb[0::3] = row[2::3]   # R
        rgb[1::3] = row[1::3]   # G
        rgb[2::3] = row[0::3]   # B
        rows.append(b"\x00" + bytes(rgb))
    raw = b"".join(rows)

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
                     + chunk(b"IDAT", zlib.compress(raw, 6)) + chunk(b"IEND", b""))


def capture_client(hwnd: int, out_path: Path, activate: bool = True) -> tuple[Path, int, int]:
    """截取窗口客户区（RDP 的位图区域），保存为 PNG 或 BMP。"""
    if activate:
        wo.force_foreground(hwnd)
        time.sleep(0.25)

    left, top, width, height = get_client_rect(hwnd)
    if width <= 0 or height <= 0:
        raise RuntimeError(f"客户区尺寸无效: {width}x{height}")

    screen_dc = user32.GetDC(None)
    if not screen_dc:
        raise ctypes.WinError(ctypes.get_last_error())
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)
    bitmap = gdi32.CreateCompatibleBitmap(screen_dc, width, height)
    old_obj = gdi32.SelectObject(mem_dc, bitmap)
    try:
        if not gdi32.BitBlt(mem_dc, 0, 0, width, height, screen_dc,
                            left, top, SRCCOPY):
            raise ctypes.WinError(ctypes.get_last_error())

        # GetDIBits 要求位图未被选入 DC
        gdi32.SelectObject(mem_dc, old_obj)

        stride = (width * 3 + 3) & ~3
        buffer = ctypes.create_string_buffer(stride * height)
        info = BITMAPINFO()
        info.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        info.bmiHeader.biWidth = width
        info.bmiHeader.biHeight = height
        info.bmiHeader.biPlanes = 1
        info.bmiHeader.biBitCount = 24
        info.bmiHeader.biCompression = 0
        info.bmiHeader.biSizeImage = stride * height

        scanned = gdi32.GetDIBits(mem_dc, bitmap, 0, height,
                                  ctypes.cast(buffer, ctypes.c_void_p),
                                  ctypes.byref(info), DIB_RGB_COLORS)
        if scanned == 0:
            raise ctypes.WinError(ctypes.get_last_error())
    finally:
        try:
            gdi32.SelectObject(mem_dc, old_obj)
        except Exception:
            pass
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(mem_dc)
        user32.ReleaseDC(None, screen_dc)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if out_path.suffix.lower() == ".bmp":
        _write_bmp(out_path, width, height, buffer.raw)
    else:
        _write_png(out_path, width, height, buffer.raw)
    return out_path, width, height


def default_screenshot_path() -> Path:
    return Path(tempfile.gettempdir()) / f"rdp_shot_{time.strftime('%Y%m%d_%H%M%S')}.png"


# ---------------------------------------------------------------------------
# OCR（调用 PowerShell + Windows.Media.Ocr，保持零第三方依赖）
# ---------------------------------------------------------------------------
def ps_script_path(explicit: Optional[str] = None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve()
    return SCRIPT_DIR / "RdpOps.ps1"


def ocr_image(image: Path, language: Optional[str] = None,
              explicit_ps: Optional[str] = None, timeout: int = 90) -> dict:
    script = ps_script_path(explicit_ps)
    if not script.exists():
        raise RuntimeError(f"找不到 OCR 所需的 PowerShell 脚本: {script}")

    cmd = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", str(script), "ocr", "-Image", str(image), "-Json"]
    if language:
        cmd.extend(["-Language", language])

    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8-sig", errors="replace", timeout=timeout)
    output = (proc.stdout or "").strip()
    if proc.returncode != 0:
        detail = (proc.stderr or "").strip() or output
        raise RuntimeError(detail or f"OCR 失败（退出码 {proc.returncode}）")

    start = output.find("{")
    if start < 0:
        raise RuntimeError(f"OCR 输出不是 JSON: {output[:300]}")
    return json.loads(output[start:])


def find_matches(ocr: dict, needle: str, exact: bool = False) -> list[dict]:
    """优先匹配单词，其次匹配整行。返回的坐标是截图内像素坐标。"""
    needle_lower = needle.lower()
    words: list[dict] = []
    lines: list[dict] = []

    for line in ocr.get("Lines") or []:
        for word in line.get("Words") or []:
            text = str(word.get("Text", ""))
            hit = (text.lower() == needle_lower) if exact else (needle_lower in text.lower())
            if hit:
                words.append({
                    "Text": text,
                    "X": int(word.get("X", 0)),
                    "Y": int(word.get("Y", 0)),
                    "Width": int(word.get("Width", 0)),
                    "Height": int(word.get("Height", 0)),
                    "Source": "word",
                    "Line": line.get("Text", ""),
                })

        text = str(line.get("Text", ""))
        hit = (text.lower() == needle_lower) if exact else (needle_lower in text.lower())
        if hit:
            lines.append({
                "Text": text,
                "X": int(line.get("X", 0)),
                "Y": int(line.get("Y", 0)),
                "Width": int(line.get("Width", 0)),
                "Height": int(line.get("Height", 0)),
                "Source": "line",
                "Line": text,
            })

    return words or lines


# ---------------------------------------------------------------------------
# 剪贴板（ctypes，纯标准库）
# ---------------------------------------------------------------------------
def _open_clipboard(retries: int = 20) -> None:
    for _ in range(retries):
        if user32.OpenClipboard(None):
            return
        time.sleep(0.05)
    raise RuntimeError("无法打开剪贴板（可能被其他程序占用）")


def clip_set(text: str) -> None:
    _open_clipboard()
    handle = None
    try:
        user32.EmptyClipboard()
        data = (text + "\x00").encode("utf-16-le")
        handle = kernel32.GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not handle:
            raise ctypes.WinError(ctypes.get_last_error())
        pointer = kernel32.GlobalLock(handle)
        if not pointer:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            ctypes.memmove(pointer, data, len(data))
        finally:
            kernel32.GlobalUnlock(handle)
        if not user32.SetClipboardData(CF_UNICODETEXT, handle):
            raise ctypes.WinError(ctypes.get_last_error())
        handle = None  # 所有权已转交给系统
    finally:
        if handle:
            kernel32.GlobalFree(handle)
        user32.CloseClipboard()


def clip_get() -> str:
    _open_clipboard()
    try:
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        pointer = kernel32.GlobalLock(handle)
        if not pointer:
            return ""
        try:
            return ctypes.wstring_at(pointer)
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


# ---------------------------------------------------------------------------
# RDP 输入
# ---------------------------------------------------------------------------
def activate(hwnd: int) -> None:
    if not wo.force_foreground(hwnd):
        raise RuntimeError("无法把目标窗口带到前台")


def _open_run_dialog(hwnd: int, wait: float = 0.7) -> None:
    """在 RDP 里按 Win+R，并确认没有被本机截获。"""
    wo.send_combo("win+r")
    time.sleep(wait)

    foreground = wo.get_foreground()
    if foreground.hwnd != hwnd:
        # 本机“运行”对话框 / 开始菜单抢走了焦点：赶紧取消，避免误输入本机
        title = foreground.title.lower()
        if foreground.cls == "#32770" or "run" in title or "运行" in foreground.title:
            wo.send_combo("esc")
            raise RuntimeError(
                "Win+R 被本机截获，没有发到远端。\n"
                "请把 mstsc 的“本地资源 → 键盘 → 应用 Windows 键组合”"
                "改成“在远程计算机上”，或使用全屏连接。"
            )
        raise RuntimeError(
            f"Win+R 后前台窗口变成了 [{foreground.title}]，按键没有送到 RDP 会话"
        )


def open_cmd(hwnd: int, wait_ms: int = 900) -> None:
    activate(hwnd)
    _open_run_dialog(hwnd)
    wo.send_combo("ctrl+a")
    wo.type_text("cmd")
    wo.send_combo("enter")
    time.sleep(wait_ms / 1000.0)


def _paste_script(hwnd: int, script: str, wait_ms: int) -> None:
    activate(hwnd)
    _open_run_dialog(hwnd)
    wo.send_combo("ctrl+a")
    wo.type_text("powershell -NoProfile -NoExit")
    wo.send_combo("enter")
    time.sleep(1.5)

    clip_set(script)
    time.sleep(0.2)
    wo.send_combo("ctrl+v")
    time.sleep(0.4)
    wo.send_combo("enter")
    time.sleep(wait_ms / 1000.0)


def run_powershell(hwnd: int, script: str, encoded: bool = False,
                   wait_ms: int = 1200) -> str:
    """在远端执行 PowerShell 脚本，返回实际使用的方式。"""
    if not script.strip():
        raise ValueError("脚本内容为空")

    if encoded:
        payload = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        command = f"powershell -NoProfile -NoExit -EncodedCommand {payload}"
        # 本机“运行”对话框有长度限制，太长自动退回剪贴板方式
        if len(command) <= 200:
            activate(hwnd)
            _open_run_dialog(hwnd)
            wo.send_combo("ctrl+a")
            wo.type_text(command)
            wo.send_combo("enter")
            time.sleep(wait_ms / 1000.0)
            return "encoded"

    _paste_script(hwnd, script, wait_ms)
    return "clipboard"


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--pattern", help="按标题筛选（默认包含匹配）")
    common.add_argument("--class", dest="class_name", help="按窗口类名筛选（支持通配符）")
    common.add_argument("--process", dest="process_name", help="按进程名筛选（支持通配符）")
    common.add_argument("--hwnd", type=lambda s: int(s, 0), help="窗口句柄，如 0x1A2B")
    common.add_argument("--pid", type=int, help="进程 ID")
    common.add_argument("--regex", action="store_true", help="标题按正则匹配")
    common.add_argument("--include-invisible", action="store_true", help="包含隐藏窗口")
    common.add_argument("--index", type=int, default=0, help="匹配多个时取第几个（从 0 开始）")
    common.add_argument("--any", action="store_true",
                        help="不限定 mstsc，允许匹配任意窗口（本地测试用）")
    common.add_argument("--json", action="store_true", help="JSON 输出")

    parser = argparse.ArgumentParser(
        description="RDP / 位图窗口自动化工具（SendInput + 截图 + OCR）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    sub = parser.add_subparsers(dest="action", required=True)

    sub.add_parser("find", parents=[common], help="查找 RDP 窗口")
    sub.add_parser("status", parents=[common], help="窗口 + 客户区 + DPI 状态")

    p_open = sub.add_parser("open-cmd", parents=[common], help="在远端打开 cmd")
    p_open.add_argument("--wait", type=int, default=900, help="回车后等待毫秒数")

    p_ps = sub.add_parser("run-ps", parents=[common], help="在远端执行 PowerShell")
    source = p_ps.add_mutually_exclusive_group(required=True)
    source.add_argument("--script", help="直接给出 PowerShell 脚本内容")
    source.add_argument("--file", help="读取本机脚本文件内容，再粘贴到远端执行")
    source.add_argument("--remote-file", help="远端（或 \\\\tsclient 重定向）脚本路径，直接 -File 执行")
    p_ps.add_argument("--encoded", action="store_true",
                      help="短脚本用 -EncodedCommand（不依赖剪贴板）")
    p_ps.add_argument("--wait", type=int, default=1200, help="执行后等待毫秒数")

    p_type = sub.add_parser("type", parents=[common], help="向远端输入文本")
    p_type.add_argument("--text", required=True)
    p_type.add_argument("--no-activate", action="store_true")

    p_keys = sub.add_parser("keys", parents=[common], help="向远端发送组合键")
    p_keys.add_argument("--keys", nargs="+", required=True, help="如 ctrl+v enter win+r")
    p_keys.add_argument("--no-activate", action="store_true")

    p_click = sub.add_parser("click", parents=[common], help="点击位图区域")
    p_click.add_argument("--x", type=int, help="相对客户区左上角的 X（默认中心）")
    p_click.add_argument("--y", type=int, help="相对客户区左上角的 Y（默认中心）")
    p_click.add_argument("--window-relative", action="store_true",
                         help="坐标相对窗口外框，而不是客户区")
    p_click.add_argument("--double", action="store_true")
    p_click.add_argument("--right", action="store_true")
    p_click.add_argument("--no-activate", action="store_true")

    p_shot = sub.add_parser("screenshot", parents=[common], help="截取客户区")
    p_shot.add_argument("--out", help="输出文件（.png 或 .bmp），默认临时目录")
    p_shot.add_argument("--no-activate", action="store_true")

    p_ocr = sub.add_parser("ocr", parents=[common], help="OCR 识别截图")
    p_ocr.add_argument("--image", help="已有图片；不传则先截取目标窗口")
    p_ocr.add_argument("--language", help="OCR 语言，如 zh-CN / en-US（默认用户配置）")
    p_ocr.add_argument("--ps-script", help="RdpOps.ps1 路径")

    p_ct = sub.add_parser("click-text", parents=[common], help="OCR 找到文字并点击")
    p_ct.add_argument("--text", required=True, help="要查找的文字")
    p_ct.add_argument("--exact", action="store_true", help="整词/整行精确匹配")
    p_ct.add_argument("--match-index", type=int, default=0, help="第几个匹配（从 0 开始）")
    p_ct.add_argument("--double", action="store_true")
    p_ct.add_argument("--right", action="store_true")
    p_ct.add_argument("--dry-run", action="store_true", help="只输出坐标，不点击")
    p_ct.add_argument("--language", help="OCR 语言")
    p_ct.add_argument("--ps-script", help="RdpOps.ps1 路径")

    p_wait = sub.add_parser("wait-text", parents=[common], help="轮询 OCR 等待文字出现")
    p_wait.add_argument("--text", required=True)
    p_wait.add_argument("--exact", action="store_true")
    p_wait.add_argument("--timeout", type=float, default=15.0, help="超时秒数")
    p_wait.add_argument("--interval", type=float, default=0.8, help="轮询间隔秒数")
    p_wait.add_argument("--language", help="OCR 语言")
    p_wait.add_argument("--ps-script", help="RdpOps.ps1 路径")

    p_cs = sub.add_parser("clip-set", parents=[common], help="设置本机剪贴板")
    p_cs.add_argument("--text", required=True)
    sub.add_parser("clip-get", parents=[common], help="读取本机剪贴板")

    p_paste = sub.add_parser("paste", parents=[common], help="设置剪贴板并 Ctrl+V")
    p_paste.add_argument("--text", required=True)

    p_tunnel = sub.add_parser(
        "tunnel", help="发现 mstsc 的本地端口隧道（AccessClient 场景）")
    p_tunnel.add_argument("--port", type=int, help="只看指定本地端口")
    p_tunnel.add_argument("--pid", type=int, help="只看指定 mstsc PID")
    p_tunnel.add_argument("--app", help="按父进程/监听进程名过滤，如 AccessClient")
    p_tunnel.add_argument("--timeout", type=float, default=0.0,
                          help="等待至少一个隧道出现（秒，0 表示不等待）")
    p_tunnel.add_argument("--interval", type=float, default=1.0,
                          help="轮询间隔秒数")
    p_tunnel.add_argument("--watch", action="store_true",
                          help="持续监控，状态变化时输出（Ctrl+C 退出）")
    p_tunnel.add_argument("--launch", action="store_true",
                          help="对已声明但尚未连接的端口拉起 mstsc")
    p_tunnel.add_argument("--json", action="store_true", help="JSON 输出")

    p_wait = sub.add_parser("wait-port", help="等待本机端口进入 LISTEN")
    p_wait.add_argument("--port", type=int, required=True, help="本地端口")
    p_wait.add_argument("--timeout", type=float, default=30.0, help="超时秒数")
    p_wait.add_argument("--interval", type=float, default=0.5, help="轮询间隔秒数")
    p_wait.add_argument("--connect", action="store_true",
                        help="额外做一次 TCP 连接探测（会向 RDP 端口发起握手）")
    p_wait.add_argument("--json", action="store_true", help="JSON 输出")

    p_health = sub.add_parser(
        "health", parents=[common], help="窗口 + 隧道的综合健康检查")
    p_health.add_argument("--port", type=int,
                          help="手工指定隧道端口（默认从 mstsc 命令行/连接推断）")
    p_health.add_argument("--screen", action="store_true",
                          help="连续截图比较，检测画面是否卡死")
    p_health.add_argument("--screen-interval", type=float, default=0.6,
                          help="两次截图的间隔秒数")

    return parser


def _print_match(match: dict, screen_x: int, screen_y: int) -> None:
    print(f"匹配文字 : {match['Text']}")
    print(f"匹配来源 : {match['Source']}  (整行: {match.get('Line', '')})")
    print(f"图片坐标 : {match['X']},{match['Y']} {match['Width']}x{match['Height']}")
    print(f"屏幕坐标 : {screen_x},{screen_y}")


def _tunnel_line(item: dict) -> str:
    state = (f"LISTEN:{item['ListenerProcess'] or '?'}({item['ListenerPid']})"
             if item["Listening"] else "NO-LISTEN")
    state += " / " + ("CONNECTED" if item["Connected"] else "NOT-CONNECTED")
    parent = f"{item['ParentProcess'] or '?'}({item['ParentPid']})"
    return (f"PID {item['Pid']:<6} {item['Process']:<12} -> {item['Endpoint']:<21} "
            f"[{state}]  parent={parent}  {item['WindowTitle']}")


def _print_health(payload: dict) -> None:
    win = payload.get("Window") or {}
    print(f"健康     : {'OK' if payload['Healthy'] else 'FAIL'}")
    if payload["Reasons"]:
        print(f"问题     : {', '.join(payload['Reasons'])}")
    if win:
        print(f"窗口     : {win.get('Handle')} [{win.get('Title')}] PID {win.get('PID')}")
        print(f"客户区   : {win.get('ClientOrigin')} {win.get('ClientSize')} "
              f"DPI {win.get('DpiScale')}")
        print(f"状态     : visible={win.get('Visible')} "
              f"minimized={win.get('Minimized')} hung={payload.get('Hung')} "
              f"foreground={win.get('Foreground')}")
    tunnel = payload.get("Tunnel")
    if tunnel:
        print(f"隧道     : {tunnel['Endpoint']}  "
              f"listener={tunnel['ListenerProcess']}({tunnel['ListenerPid']})  "
              f"connected={tunnel['Connected']}")
    else:
        print("隧道     : 未发现")
    screen = payload.get("Screen")
    if screen and screen.get("Checked"):
        print(f"画面     : {'有变化' if screen['Changed'] else '无变化（可能卡死）'}")


def main(argv: Optional[list[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    action = args.action

    try:
        if action == "find":
            windows = select_windows(args)
            if not windows:
                print("没有匹配的窗口")
            elif args.json:
                print(json.dumps([window_dict(w) for w in windows],
                                 ensure_ascii=False, indent=2))
            else:
                wo.print_windows(windows)
            return 0

        if action == "status":
            win = pick_window(args)
            data = window_dict(win)
            if args.json:
                print(json.dumps(data, ensure_ascii=False, indent=2))
            else:
                wo.print_detail(win)
                print(f"Client    : {data.get('ClientOrigin')} {data.get('ClientSize')}")
                print(f"DPI Scale : {data.get('DpiScale')}")
                print(f"Foreground: {data.get('Foreground')}")
                print(f"Is RDP    : {data.get('IsRdp')}")
            return 0

        if action == "open-cmd":
            win = pick_window(args)
            open_cmd(win.hwnd, args.wait)
            print(f"已向 [{win.title}] 发送 Win+R → cmd")
            return 0

        if action == "run-ps":
            win = pick_window(args)
            if args.script is not None:
                script = args.script
            elif args.file:
                script = Path(args.file).read_text(encoding="utf-8-sig", errors="replace")
            else:
                remote = args.remote_file
                if not remote.lower().startswith("powershell"):
                    remote = (f'powershell -NoProfile -ExecutionPolicy Bypass '
                              f'-File "{remote}"')
                activate(win.hwnd)
                _open_run_dialog(win.hwnd)
                wo.send_combo("ctrl+a")
                wo.type_text(remote)
                wo.send_combo("enter")
                time.sleep(args.wait / 1000.0)
                print(f"已在远端执行文件: {args.remote_file}")
                return 0

            mode = run_powershell(win.hwnd, script, args.encoded, args.wait)
            print(f"已在远端启动 PowerShell（{mode} 方式），脚本 {len(script)} 字符")
            if mode == "clipboard":
                print("提示：剪贴板方式要求 mstsc 已启用“剪贴板”重定向。")
            return 0

        if action == "type":
            win = pick_window(args)
            if not args.no_activate:
                activate(win.hwnd)
            wo.type_text(args.text)
            print(f"已向 [{win.title}] 输入 {len(args.text)} 个字符")
            return 0

        if action == "keys":
            win = pick_window(args)
            if not args.no_activate:
                activate(win.hwnd)
            for combo in args.keys:
                wo.send_combo(combo)
                time.sleep(0.1)
            print(f"已向 [{win.title}] 发送按键: {', '.join(args.keys)}")
            return 0

        if action == "click":
            win = pick_window(args)
            if not args.no_activate:
                activate(win.hwnd)
            if args.window_relative:
                origin_x, origin_y = win.left, win.top
                width, height = win.width, win.height
            else:
                origin_x, origin_y, width, height = get_client_rect(win.hwnd)
            x = origin_x + (args.x if args.x is not None else width // 2)
            y = origin_y + (args.y if args.y is not None else height // 2)
            wo.click_at(x, y, args.double, args.right)
            print(f"已点击屏幕坐标 ({x}, {y})")
            return 0

        if action == "screenshot":
            win = pick_window(args)
            out = Path(args.out) if args.out else default_screenshot_path()
            path, width, height = capture_client(win.hwnd, out, not args.no_activate)
            if args.json:
                print(json.dumps({"Path": str(path), "Width": width,
                                  "Height": height}, ensure_ascii=False, indent=2))
            else:
                print(f"已保存截图: {path}  ({width}x{height})")
            return 0

        if action == "ocr":
            if args.image:
                image = Path(args.image)
            else:
                win = pick_window(args)
                image, _, _ = capture_client(win.hwnd, default_screenshot_path(),
                                             not args.no_activate)
            result = ocr_image(image, args.language, args.ps_script)
            if args.json:
                print(json.dumps(result, ensure_ascii=False, indent=2))
            else:
                print(result.get("Text", ""))
            return 0

        if action in ("click-text", "wait-text"):
            win = pick_window(args)
            activate(win.hwnd)

            deadline = time.time() + (args.timeout if action == "wait-text" else 0)
            match = None
            image = None
            while True:
                image, _, _ = capture_client(win.hwnd, default_screenshot_path(),
                                             activate=False)
                result = ocr_image(image, args.language, args.ps_script)
                matches = find_matches(result, args.text, args.exact)
                if matches:
                    index = args.match_index if action == "click-text" else 0
                    if index < len(matches):
                        match = matches[index]
                        break
                if action != "wait-text" or time.time() >= deadline:
                    break
                time.sleep(max(0.05, args.interval))

            if not match:
                print(f"未找到文字: {args.text}", file=sys.stderr)
                if action == "click-text":
                    print(f"最后一张截图: {image}", file=sys.stderr)
                return 1

            left, top, _, _ = get_client_rect(win.hwnd)
            screen_x = left + match["X"] + match["Width"] // 2
            screen_y = top + match["Y"] + match["Height"] // 2

            if action == "wait-text":
                if args.json:
                    print(json.dumps({"Found": True, "Match": match,
                                      "ScreenX": screen_x, "ScreenY": screen_y},
                                     ensure_ascii=False, indent=2))
                else:
                    _print_match(match, screen_x, screen_y)
                return 0

            if args.json:
                print(json.dumps({"Found": True, "Match": match,
                                  "ScreenX": screen_x, "ScreenY": screen_y,
                                  "Clicked": not args.dry_run},
                                 ensure_ascii=False, indent=2))
            else:
                _print_match(match, screen_x, screen_y)
            if not args.dry_run:
                wo.click_at(screen_x, screen_y, args.double, args.right)
                print("已点击")
            return 0

        if action == "clip-set":
            clip_set(args.text)
            print(f"已写入剪贴板，共 {len(args.text)} 个字符")
            return 0

        if action == "clip-get":
            content = clip_get()
            if args.json:
                print(json.dumps({"Text": content}, ensure_ascii=False, indent=2))
            else:
                print(content)
            return 0

        if action == "paste":
            win = pick_window(args)
            clip_set(args.text)
            time.sleep(0.15)
            activate(win.hwnd)
            wo.send_combo("ctrl+v")
            print(f"已粘贴 {len(args.text)} 个字符到 [{win.title}]")
            return 0

        if action == "tunnel":
            deadline = time.time() + args.timeout if args.timeout > 0 else None
            while True:
                tunnels = filter_tunnels(
                    discover_tunnels(args.port, args.pid), args.app)
                if tunnels or deadline is None or time.time() >= deadline:
                    break
                time.sleep(max(0.05, args.interval))

            if args.launch:
                for item in tunnels:
                    if not item["Connected"]:
                        subprocess.Popen(["mstsc", f"/v:127.0.0.1:{item['Port']}"])
                        print(f"已拉起 mstsc /v:127.0.0.1:{item['Port']}")

            if not args.watch:
                if args.json:
                    print(json.dumps(tunnels, ensure_ascii=False, indent=2))
                elif not tunnels:
                    print("没有发现本地隧道：mstsc 未连接 127.0.0.1，"
                          "或客户端进程不可见。")
                else:
                    for item in tunnels:
                        print(_tunnel_line(item))
                return 0 if tunnels else 1

            print("监控本地隧道中，Ctrl+C 退出 ...")
            last: dict[tuple[int, int], tuple] = {}
            while True:
                tunnels = filter_tunnels(
                    discover_tunnels(args.port, args.pid), args.app)
                current = {
                    (item["Pid"], item["Port"]):
                        (item["Listening"], item["Connected"], item["Ready"])
                    for item in tunnels
                }
                stamp = time.strftime("%H:%M:%S")
                for item in tunnels:
                    key = (item["Pid"], item["Port"])
                    if last.get(key) != current[key]:
                        print(f"[{stamp}] {_tunnel_line(item)}")
                for key in sorted(set(last) - set(current)):
                    print(f"[{stamp}] 隧道消失: PID {key[0]} 端口 {key[1]}")
                last = current
                time.sleep(max(0.2, args.interval))

        if action == "wait-port":
            status = wait_port(args.port, args.timeout, args.interval,
                               args.connect)
            if args.json:
                print(json.dumps(status, ensure_ascii=False, indent=2))
            elif status["Listening"]:
                extra = ""
                if args.connect:
                    extra = (" (TCP 连接成功)" if status.get("TcpConnect")
                             else " (TCP 连接失败)")
                print(f"端口 {args.port} 已就绪: "
                      f"{', '.join(status['Addresses'])}  "
                      f"进程 {status['Process']}({status['Pid']}){extra}")
            else:
                print(f"等待端口 {args.port} 超时（{args.timeout:g}s）",
                      file=sys.stderr)
            return 0 if status["Listening"] else 1

        if action == "health":
            windows = select_windows(args)
            if not windows:
                payload = {"Healthy": False, "Reasons": ["window_not_found"],
                           "Window": None, "Tunnel": None}
                if args.json:
                    print(json.dumps(payload, ensure_ascii=False, indent=2))
                else:
                    print("窗口未找到", file=sys.stderr)
                return 1
            if args.index < 0 or args.index >= len(windows):
                raise SystemExit(
                    f"索引 {args.index} 越界，共找到 {len(windows)} 个窗口")

            payload = window_health(windows[args.index], args.screen,
                                    args.screen_interval, args.port)
            if args.json:
                print(json.dumps(payload, ensure_ascii=False, indent=2))
            else:
                _print_health(payload)
            return 0 if payload["Healthy"] else 1

    except KeyboardInterrupt:
        return 130
    except Exception as exc:  # noqa: BLE001
        print(f"错误: {exc}", file=sys.stderr)
        return 1

    return 1


if __name__ == "__main__":
    sys.exit(main())
