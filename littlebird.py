#!/usr/bin/env python
"""Littlebird: file watcher -> GUI notifier for Codex CLI windows.

Dependencies:
  - watchdog
  - pyautogui

Usage:
  python littlebird.py

Notes:
  - Update WINDOW_TITLES to match your terminal window titles.
  - This script sends a short message and presses Enter in the target window.
"""

from __future__ import annotations

import os
import queue
import threading
import time
from pathlib import Path
import argparse

pyautogui = None
Observer = None
FileSystemEventHandler = None


# ---- Configuration ----
BASE_DIR = Path(__file__).resolve().parent
WATCH_PATHS = [
    BASE_DIR / ".claude" / "handoff",
    BASE_DIR / ".claude" / "workspace",
]

# ---- Notification routing + window targeting ----
# Each agent can use different matching strategy and optional tab hotkey.
# Matching fields are ANDed when provided.
#
# Examples:
# - manager: class + pid + hotkey Ctrl+Alt+1
# - uiux: title + hotkey Ctrl+Alt+2
# - backend: title only (no hotkey)
AGENT_WINDOWS = {
    "manager": {
        "hwnd": 2034144,  # Exact window handle (int). Overrides other matching if set.
        "class_name": "CASCADIA_HOSTING_WINDOW_CLASS",
        "pid": 23068,  # Set a PID number if you want to lock to one window.
        "title": "manager",  # Optional title substring.
        "hotkey": ("ctrl", "alt", "1"),
    },
    "backend": {
        "hwnd": 1120952,
        "class_name": "CASCADIA_HOSTING_WINDOW_CLASS",
        "pid": 2900,
        "title": "Backend",
        "hotkey": ("ctrl", "alt", "1"),
    },
    "uiux": {
        "hwnd": 112728508,
        "class_name": "CASCADIA_HOSTING_WINDOW_CLASS",
        "pid": 23068,
        "title": "UIUX",
        "hotkey": ("ctrl", "alt", "1"),
    },
}

DEBOUNCE_SECONDS = 1.0
DRY_RUN = False

# ---- Windows focus helpers (ctypes, no external deps) ----
import ctypes
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)

EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
GetWindowText = user32.GetWindowTextW
GetWindowTextLength = user32.GetWindowTextLengthW
GetClassName = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
IsWindowVisible = user32.IsWindowVisible
SetForegroundWindow = user32.SetForegroundWindow
ShowWindow = user32.ShowWindow

SW_RESTORE = 9


def _iter_windows():
    hwnds = []

    def callback(hwnd, _):
        if IsWindowVisible(hwnd):
            length = GetWindowTextLength(hwnd)
            if length > 0:
                buf = ctypes.create_unicode_buffer(length + 1)
                GetWindowText(hwnd, buf, length + 1)
                title = buf.value
            else:
                title = ""
            class_buf = ctypes.create_unicode_buffer(256)
            GetClassName(hwnd, class_buf, 256)
            class_name = class_buf.value
            pid = wintypes.DWORD()
            GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            hwnds.append((hwnd, title, class_name, int(pid.value)))
        return True

    EnumWindows(EnumWindowsProc(callback), 0)
    return hwnds


def find_window(config: dict) -> int | None:
    hwnd_exact = config.get("hwnd")
    if isinstance(hwnd_exact, int) and hwnd_exact > 0:
        return hwnd_exact
    title_sub = (config.get("title") or "").lower()
    class_name = (config.get("class_name") or "").lower()
    pid = config.get("pid")

    for hwnd, title, cls, win_pid in _iter_windows():
        if pid is not None and win_pid != pid:
            continue
        if class_name and class_name not in cls.lower():
            continue
        if title_sub and title_sub not in title.lower():
            continue
        return hwnd
    return None


def focus_window(hwnd: int) -> bool:
    if not hwnd:
        return False
    ShowWindow(hwnd, SW_RESTORE)
    return bool(SetForegroundWindow(hwnd))


# ---- Notification routing ----

def route_event(path: Path) -> tuple[str, str] | None:
    path_str = str(path)
    lower = path_str.lower()

    if lower.endswith("output.md") and "handoff" in lower:
        # Extract task id if possible: ...\handoff\{task}\output.md
        parts = path.parts
        task_id = ""
        try:
            handoff_idx = [p.lower() for p in parts].index("handoff")
            task_id = parts[handoff_idx + 1]
        except Exception:
            task_id = ""
        msg = f"[Manager] output updated: {task_id or path.name}"
        return "manager", msg

    if lower.endswith("notifications.md") and "workspace" in lower:
        if "backend" in lower:
            return "backend", "[Backend] new notification"
        if "ui-ux" in lower or "uiux" in lower:
            return "uiux", "[UI-UX] new notification"

    return None


# ---- Debounce queue ----

event_queue: "queue.Queue[Path]" = queue.Queue()
last_sent: dict[tuple[str, str], float] = {}
lock = threading.Lock()


def worker():
    while True:
        path = event_queue.get()
        if path is None:
            break

        routed = route_event(path)
        if not routed:
            continue

        target, message = routed
        key = (target, str(path))
        now = time.time()

        with lock:
            last_time = last_sent.get(key, 0.0)
            if now - last_time < DEBOUNCE_SECONDS:
                continue
            last_sent[key] = now

        if DRY_RUN:
            print(f"[DRY_RUN] {target}: {message}")
            continue

        config = AGENT_WINDOWS.get(target, {})
        hwnd = find_window(config)
        if not hwnd:
            print(f"[WARN] Window not found for {target}: {config}")
            continue

        if not focus_window(hwnd):
            print(f"[WARN] Failed to focus window for {target}: {config}")
            continue

        time.sleep(0.05)
        hotkey = config.get("hotkey")
        if hotkey:
            pyautogui.hotkey(*hotkey)
            time.sleep(0.05)
        pyautogui.typewrite(message)
        pyautogui.press("enter")


def load_deps():
    global pyautogui, Observer, FileSystemEventHandler
    try:
        import pyautogui as _pyautogui
    except ImportError as exc:
        raise SystemExit("pyautogui is required: pip install pyautogui") from exc
    try:
        from watchdog.events import FileSystemEventHandler as _FileSystemEventHandler
        from watchdog.observers import Observer as _Observer
    except ImportError as exc:
        raise SystemExit("watchdog is required: pip install watchdog") from exc

    pyautogui = _pyautogui
    Observer = _Observer
    FileSystemEventHandler = _FileSystemEventHandler


def list_windows():
    print("HWND=... | PID=... | CLASS=... | TITLE=...")
    for hwnd, title, cls, win_pid in _iter_windows():
        title_clean = title.replace("\t", " ").replace("\n", " ").strip()
        cls_clean = cls.replace("\t", " ").replace("\n", " ").strip()
        print(f"HWND={hwnd} | PID={win_pid} | CLASS={cls_clean} | TITLE={title_clean}")


def main():
    global DRY_RUN, DEBOUNCE_SECONDS
    parser = argparse.ArgumentParser(description="Littlebird file watcher")
    parser.add_argument("--list-windows", action="store_true", help="List visible windows")
    parser.add_argument("--dry-run", action="store_true", help="Print actions without typing")
    parser.add_argument("--debounce", type=float, default=DEBOUNCE_SECONDS, help="Debounce seconds")
    args = parser.parse_args()

    if args.list_windows:
        list_windows()
        return

    DRY_RUN = args.dry_run
    DEBOUNCE_SECONDS = args.debounce

    load_deps()

    class WatchHandler(FileSystemEventHandler):
        def on_modified(self, event):
            if not event.is_directory:
                event_queue.put(Path(event.src_path))

        def on_created(self, event):
            if not event.is_directory:
                event_queue.put(Path(event.src_path))

    for path in WATCH_PATHS:
        if not path.exists():
            print(f"[WARN] watch path missing: {path}")

    handler = WatchHandler()
    observer = Observer()

    for path in WATCH_PATHS:
        if path.exists():
            observer.schedule(handler, str(path), recursive=True)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    observer.start()
    print("Littlebird running. Press Ctrl+C to stop.")

    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        observer.stop()
    finally:
        observer.join()
        event_queue.put(None)


if __name__ == "__main__":
    main()
