#!/usr/bin/env python
"""Prewarm a WeChat group before running local summaries.

This script focuses the desktop WeChat window, opens the target group by name,
waits for recent cloud-backed content to sync locally, and optionally scrolls
the chat history to encourage older messages/media to land on disk.
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path

import pyperclip
import win32api
import win32con
import win32gui
import win32process
from PIL import ImageGrab
from pywinauto import Application, Desktop


WINDOW_TITLE = "微信"
WINDOW_CLASS = "Qt51514QWindowIcon"
OCR_SCRIPT = """
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null=[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$null=[Windows.Storage.FileAccessMode,Windows.Storage,ContentType=WindowsRuntime]
$null=[Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$null=[Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime]
$null=[Windows.Media.Ocr.OcrEngine,Windows.Foundation.FoundationContract,ContentType=WindowsRuntime]
$null=[Windows.Media.Ocr.OcrResult,Windows.Foundation.FoundationContract,ContentType=WindowsRuntime]
function Await($op,[Type]$t){$m=[System.WindowsRuntimeSystemExtensions].GetMethods()|Where-Object{$_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.IsGenericMethod}|Select-Object -First 1;$task=$m.MakeGenericMethod($t).Invoke($null,@($op));$task.Wait();$task.Result}
$imgPath=$env:WECHAT_OCR_IMAGE
$file=Await ([Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]::GetFileFromPathAsync($imgPath)) ([Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime])
$stream=Await ($file.OpenAsync([Windows.Storage.FileAccessMode,Windows.Storage,ContentType=WindowsRuntime]::Read)) ([Windows.Storage.Streams.IRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime])
$decoder=Await ([Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics.Imaging,ContentType=WindowsRuntime])
$bmp=Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap,Windows.Graphics.Imaging,ContentType=WindowsRuntime])
$engine=[Windows.Media.Ocr.OcrEngine,Windows.Foundation.FoundationContract,ContentType=WindowsRuntime]::TryCreateFromUserProfileLanguages()
$result=Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult,Windows.Foundation.FoundationContract,ContentType=WindowsRuntime])
Write-Output $result.Text
""".strip()


@dataclass(frozen=True)
class WindowRect:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top

    def point(self, x_ratio: float, y_ratio: float) -> tuple[int, int]:
        return (
            self.left + int(self.width * x_ratio),
            self.top + int(self.height * y_ratio),
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Open and prewarm a WeChat group before local analysis."
    )
    parser.add_argument("--group", required=True, help="Target WeChat group name.")
    parser.add_argument(
        "--wait-seconds",
        type=float,
        default=10.0,
        help="Seconds to wait after opening the chat. Default: 10.",
    )
    parser.add_argument(
        "--history-scrolls",
        type=int,
        default=6,
        help="How many wheel-up scrolls to perform in the message pane. Default: 6.",
    )
    parser.add_argument(
        "--history-scroll-pause",
        type=float,
        default=0.4,
        help="Pause in seconds between history scrolls. Default: 0.4.",
    )
    parser.add_argument(
        "--settle-after-scroll",
        type=float,
        default=4.0,
        help="Seconds to wait after scrolling. Default: 4.",
    )
    parser.add_argument(
        "--layout",
        choices=("split-right", "split-left", "restore"),
        default="split-right",
        help=(
            "How to place the WeChat window before operating. "
            "Default: split-right."
        ),
    )
    parser.add_argument(
        "--maximize",
        action="store_true",
        help="Maximize WeChat before operating. Overrides --layout.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print planned actions without touching the GUI.",
    )
    parser.add_argument(
        "--image-clicks",
        type=int,
        default=4,
        help="How many visible image slots to click for prewarming. Default: 4.",
    )
    parser.add_argument(
        "--image-click-pause",
        type=float,
        default=0.8,
        help="Pause in seconds after opening an image before closing it. Default: 0.8.",
    )
    return parser


def find_wechat_handle() -> int:
    windows = []
    for win in Desktop(backend="uia").windows():
        try:
            if win.window_text() == WINDOW_TITLE and win.class_name() == WINDOW_CLASS:
                rect = win.rectangle()
                area = max(0, rect.right - rect.left) * max(0, rect.bottom - rect.top)
                windows.append((area, win.handle))
        except Exception:
            continue
    if not windows:
        raise RuntimeError("未找到可见的微信主窗口，请先打开桌面微信。")
    windows.sort(reverse=True)
    return windows[0][1]


def get_monitor_work_rect(handle: int) -> WindowRect:
    monitor = win32api.MonitorFromWindow(handle, win32con.MONITOR_DEFAULTTONEAREST)
    info = win32api.GetMonitorInfo(monitor)
    left, top, right, bottom = info["Work"]
    return WindowRect(left, top, right, bottom)


def tile_wechat(handle: int, layout: str) -> None:
    work = get_monitor_work_rect(handle)
    width = work.width
    height = work.height
    half_width = max(900, width // 2)
    half_width = min(half_width, width)

    if layout == "split-right":
        left = work.right - half_width
    elif layout == "split-left":
        left = work.left
    else:
        win32gui.ShowWindow(handle, win32con.SW_RESTORE)
        return

    top = work.top
    win32gui.ShowWindow(handle, win32con.SW_RESTORE)
    time.sleep(0.2)
    win32gui.SetWindowPos(
        handle,
        win32con.HWND_TOP,
        left,
        top,
        half_width,
        height,
        win32con.SWP_SHOWWINDOW,
    )


def focus_wechat(handle: int, maximize: bool, layout: str) -> WindowRect:
    app = Application(backend="uia").connect(handle=handle)
    window = app.window(handle=handle)
    if maximize:
        win32gui.ShowWindow(handle, win32con.SW_MAXIMIZE)
        time.sleep(0.5)
    else:
        tile_wechat(handle, layout)
        time.sleep(0.4)
    window.set_focus()
    time.sleep(0.8)
    left, top, right, bottom = win32gui.GetWindowRect(handle)
    return WindowRect(left, top, right, bottom)


def click(point: tuple[int, int]) -> None:
    win32api.SetCursorPos(point)
    time.sleep(0.1)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.25)


def key_down(vk: int) -> None:
    win32api.keybd_event(vk, 0, 0, 0)


def key_up(vk: int) -> None:
    win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)


def hotkey(*keys: int) -> None:
    for key in keys:
        key_down(key)
    time.sleep(0.05)
    for key in reversed(keys):
        key_up(key)
    time.sleep(0.25)


def press(vk: int) -> None:
    key_down(vk)
    time.sleep(0.05)
    key_up(vk)
    time.sleep(0.25)


def scroll_up_at(point: tuple[int, int], times: int, pause: float) -> None:
    win32api.SetCursorPos(point)
    time.sleep(0.1)
    for _ in range(times):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, 120 * 4, 0)
        time.sleep(pause)


def build_key_image_points(rect: WindowRect, count: int) -> list[tuple[int, int]]:
    if count <= 0:
        return []
    x_ratio = 0.68
    start_y = 0.30
    step = 0.12
    points = []
    for index in range(count):
        y_ratio = min(0.78, start_y + step * index)
        points.append(
            (
                rect.left + round(rect.width * x_ratio),
                rect.top + round(rect.height * y_ratio),
            )
        )
    return points


def build_title_region(rect: WindowRect) -> tuple[int, int, int, int]:
    return (
        rect.left + round(rect.width * 0.32),
        rect.top + round(rect.height * 0.03),
        rect.left + round(rect.width * 0.88),
        rect.top + round(rect.height * 0.11),
    )


def normalize_visible_text(text: str) -> str:
    return "".join(ch for ch in text if ch.isalnum() or "\u4e00" <= ch <= "\u9fff")


def group_name_matches_ocr_text(group_name: str, ocr_text: str) -> bool:
    expected = normalize_visible_text(group_name)
    actual = normalize_visible_text(ocr_text)
    chinese = "".join(ch for ch in expected if "\u4e00" <= ch <= "\u9fff")
    digits = "".join(ch for ch in expected if ch.isdigit())
    return bool(actual) and chinese in actual and (not digits or digits in actual)


def run_ocr_on_image(image_path: Path) -> str:
    result = subprocess.run(
        ["powershell", "-NoProfile", "-Command", OCR_SCRIPT],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env={**os.environ, "WECHAT_OCR_IMAGE": str(image_path)},
        check=True,
    )
    return result.stdout.strip()


def read_current_chat_title_text(rect: WindowRect) -> str:
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
        image_path = Path(temp_file.name)
    try:
        ImageGrab.grab(bbox=build_title_region(rect)).save(image_path)
        return run_ocr_on_image(image_path)
    finally:
        image_path.unlink(missing_ok=True)


def assert_target_group_opened(rect: WindowRect, group_name: str) -> None:
    ocr_text = read_current_chat_title_text(rect)
    if group_name_matches_ocr_text(group_name, ocr_text):
        return
    press(win32con.VK_RETURN)
    time.sleep(0.6)
    ocr_text = read_current_chat_title_text(rect)
    if not group_name_matches_ocr_text(group_name, ocr_text):
        raise RuntimeError(f"未确认已切到目标群: {group_name}; OCR={ocr_text!r}")


def prewarm_key_images(rect: WindowRect, image_clicks: int, pause: float) -> None:
    for point in build_key_image_points(rect, image_clicks):
        click(point)
        time.sleep(pause)
        press(win32con.VK_ESCAPE)
        time.sleep(0.2)


def open_group(rect: WindowRect, group_name: str) -> None:
    search_box = rect.point(0.115, 0.08)

    click(search_box)
    hotkey(win32con.VK_CONTROL, 0x41)
    press(win32con.VK_BACK)
    pyperclip.copy(group_name)
    hotkey(win32con.VK_CONTROL, 0x56)
    time.sleep(0.9)
    press(win32con.VK_RETURN)
    time.sleep(0.6)
    press(win32con.VK_RETURN)
    time.sleep(0.6)
    assert_target_group_opened(rect, group_name)


def prewarm_history(rect: WindowRect, scrolls: int, pause: float) -> None:
    message_pane = rect.point(0.70, 0.55)
    scroll_up_at(message_pane, scrolls, pause)


def main() -> int:
    args = build_parser().parse_args()
    handle = find_wechat_handle()

    if args.dry_run:
        print(f"[dry-run] 微信句柄: {handle}")
        print(f"[dry-run] 目标群: {args.group}")
        print(f"[dry-run] 布局: {args.layout}")
        print(f"[dry-run] 等待: {args.wait_seconds}s")
        print(f"[dry-run] 历史滚动: {args.history_scrolls} 次")
        print(f"[dry-run] 滚动后等待: {args.settle_after_scroll}s")
        print(f"[dry-run] 最大化: {args.maximize}")
        return 0

    rect = focus_wechat(handle, maximize=args.maximize, layout=args.layout)
    print(f"已聚焦微信窗口: {handle} rect={rect}")

    open_group(rect, args.group)
    print(f"已切到群聊: {args.group}")

    time.sleep(args.wait_seconds)
    print(f"已等待 {args.wait_seconds} 秒让最近消息同步")

    if args.history_scrolls > 0:
        prewarm_history(rect, args.history_scrolls, args.history_scroll_pause)
        print(f"已向上滚动历史 {args.history_scrolls} 次")
        time.sleep(args.settle_after_scroll)
        print(f"滚动后已额外等待 {args.settle_after_scroll} 秒")

    if args.image_clicks > 0:
        prewarm_key_images(rect, args.image_clicks, args.image_click_pause)
        print(f"已预热可见关键图片 {args.image_clicks} 张")

    print("群同步预热完成")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("已取消")
        raise SystemExit(130)
    except Exception as exc:
        print(f"预热失败: {exc}", file=sys.stderr)
        raise SystemExit(1)
