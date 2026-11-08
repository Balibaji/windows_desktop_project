#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
window_ops.py 的端到端演示（Python 版）

依次演示：
  1. 枚举当前桌面上打开的窗口（句柄 / PID / 进程 / 类名 / 位置尺寸 / 标题）
  2. 启动一个控制台窗口作为演示目标，并按标题找到它
  3. 查看窗口详细信息（不止标题）
  4. 激活窗口 + SendInput 真实键盘输入（支持中文）+ 组合键
  5. 启动“字符映射表”，枚举子控件，用 WM_SETTEXT / WM_GETTEXT 读写控件文本
  6. 鼠标点击 + 真实键盘输入到指定控件
  7. 关闭演示窗口

用法：
    python demo.py
    python demo.py --keep-open     # 演示结束后保留窗口
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import time

import window_ops as wo

CONSOLE_TITLE = "WinOps-PyDemo"


def section(text: str) -> None:
    print()
    print("=" * 72)
    print(f"  {text}")
    print("=" * 72)


def wait_window(pattern: str | None = None, process: str | None = None,
                timeout: float = 15.0) -> wo.WinInfo | None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        for win in wo.get_windows():
            if pattern and pattern.lower() not in win.title.lower():
                continue
            if process and win.process.lower() != process.lower():
                continue
            return win
        time.sleep(0.2)
    return None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="window_ops.py 端到端演示")
    parser.add_argument("--keep-open", action="store_true", help="演示结束后保留窗口")
    args = parser.parse_args(argv)

    console_proc = None
    charmap_proc = None
    console_win = None
    charmap_win = None

    try:
        # ------------------------------------------------------------------
        section("1. 枚举当前打开的窗口（包含但不限于标题）")
        windows = wo.get_windows()
        print(f"共找到 {len(windows)} 个可见顶层窗口，下面列出前 8 个：")
        wo.print_windows(windows[:8])

        # ------------------------------------------------------------------
        section("2. 启动演示目标：一个控制台窗口")
        result_file = os.path.join(os.environ.get("TEMP", "."), "winops_py_demo.txt")
        result_file2 = os.path.join(os.environ.get("TEMP", "."), "winops_py_keys.txt")
        for path in (result_file, result_file2):
            try:
                os.remove(path)
            except OSError:
                pass

        console_proc = subprocess.Popen(
            ["cmd.exe", "/k", "title", CONSOLE_TITLE],
            creationflags=subprocess.CREATE_NEW_CONSOLE,
        )
        console_win = wait_window(pattern=CONSOLE_TITLE)
        if not console_win:
            raise RuntimeError("演示控制台窗口没有出现")

        print(f"按标题找到窗口: [{console_win.title.strip()}]")
        print(f"  句柄 = 0x{console_win.hwnd:08X}")
        print(f"  PID  = {console_win.pid}")
        print(f"  进程 = {console_win.process}")
        print(f"  类名 = {console_win.cls}")

        # ------------------------------------------------------------------
        section("3. 窗口详细信息")
        wo.print_detail(console_win)

        # ------------------------------------------------------------------
        section("4. 激活窗口 + SendInput 真实键盘输入（中文可输入）")
        activated = wo.force_foreground(console_win.hwnd)
        print(f"激活成功 = {activated}")
        time.sleep(0.3)

        command = f'echo Python-SendInput-OK > "{result_file}"'
        print(f"输入命令: {command}")
        wo.type_text(command)
        wo.send_combo("enter")
        time.sleep(1.2)

        if os.path.exists(result_file):
            with open(result_file, encoding="utf-8", errors="replace") as fh:
                print(f"控制台执行结果: {fh.read().strip()}")
        else:
            print("未检测到结果文件（键盘输入可能没有送达）")

        # ------------------------------------------------------------------
        section("5. 组合键：Ctrl+L 清屏，再用 Enter 执行命令")
        wo.send_combo("ctrl+l")
        time.sleep(0.2)
        command2 = f'echo Keys-OK > "{result_file2}"'
        wo.type_text(command2)
        wo.send_combo("enter")
        time.sleep(1.2)

        if os.path.exists(result_file2):
            with open(result_file2, encoding="utf-8", errors="replace") as fh:
                print(f"组合键执行结果: {fh.read().strip()}")
        else:
            print("未检测到结果文件")

        # ------------------------------------------------------------------
        section("6. 读写 GUI 控件文本（字符映射表）")
        charmap_path = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                                    "System32", "charmap.exe")
        charmap_proc = subprocess.Popen([charmap_path])
        charmap_win = wait_window(process="charmap")
        if not charmap_win:
            raise RuntimeError("字符映射表窗口没有出现")

        print(f"窗口: [{charmap_win.title}]   类名 = {charmap_win.cls}")
        children = wo.get_children(charmap_win.hwnd)
        print(f"共 {len(children)} 个子控件，前 6 个：")
        wo.print_windows(children[:6])

        control = wo.pick_control(charmap_win, "RICHEDIT*")
        print(f"选中的文本框控件: {control.cls}   句柄 = 0x{control.hwnd:08X}")

        payload = "Hello 世界，Python 通过 WM_SETTEXT 写入"
        wo.set_control_text(control.hwnd, payload)
        time.sleep(0.3)
        print(f"写入后 WM_GETTEXT 读回: {wo.get_control_text(control.hwnd)}")

        # ------------------------------------------------------------------
        section("7. 鼠标点击 + 真实键盘输入到控件")
        click_x = control.left + control.width // 2
        click_y = control.top + control.height // 2
        wo.force_foreground(charmap_win.hwnd)
        wo.click_at(click_x, click_y)
        time.sleep(0.2)
        wo.send_combo("ctrl+End")  # 把光标移到文本末尾
        time.sleep(0.15)

        wo.type_text(" + 键盘追加")
        time.sleep(0.4)
        print(f"点击并输入后读回: {wo.get_control_text(control.hwnd)}")

        # ------------------------------------------------------------------
        section("8. 演示结束")
        print("演示完成，下面关闭演示窗口。")

    except Exception as exc:  # noqa: BLE001
        print(f"\n演示出错: {exc}", file=sys.stderr)
        return 1
    finally:
        if args.keep_open:
            print("\n已保留演示窗口（--keep-open）。")
        else:
            if charmap_win:
                wo.close_window(charmap_win.hwnd)
            if console_win:
                wo.force_foreground(console_win.hwnd)
                time.sleep(0.2)
                wo.type_text("exit")
                wo.send_combo("enter")
                time.sleep(0.6)
            for proc in (charmap_proc, console_proc):
                if proc and proc.poll() is None:
                    try:
                        proc.kill()
                    except OSError:
                        pass
            print("演示窗口已关闭。")

    return 0


if __name__ == "__main__":
    sys.exit(main())
