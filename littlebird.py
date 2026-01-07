#!/usr/bin/env python
"""Littlebird: file watcher -> GUI notifier for Codex CLI windows.
Optimized for stability and race-condition prevention.
"""

from __future__ import annotations

import os
import queue
import threading
import time
import random
import logging
from pathlib import Path
from datetime import datetime
import argparse

pyautogui = None
Observer = None
FileSystemEventHandler = None


# ---- Logging Configuration ----
LOG_FILE = Path(__file__).resolve().parent / "littlebird.log"

def setup_logging():
    """設置日誌，同時輸出到終端和檔案"""
    # 清空舊的 handlers
    root_logger = logging.getLogger()
    root_logger.handlers = []
    root_logger.setLevel(logging.DEBUG)

    # 格式
    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # 終端輸出
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # 檔案輸出
    file_handler = logging.FileHandler(LOG_FILE, encoding='utf-8', mode='a')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    return root_logger

logger = logging.getLogger(__name__)

def log(msg: str, level: str = "INFO"):
    """統一的日誌函數"""
    if level == "ERROR":
        logger.error(msg)
    elif level == "WARN":
        logger.warning(msg)
    elif level == "DEBUG":
        logger.debug(msg)
    else:
        logger.info(msg)

# ---- Configuration ----
BASE_DIR = Path(__file__).resolve().parent
WATCH_PATHS = [
    BASE_DIR / ".agent-workspace",
]

# ---- Notification routing + window targeting ----
# 建議：移除 hwnd 硬編碼，依賴 title 和 class_name 即可
AGENT_WINDOWS = {
    "manager": {
        "class_name": "CASCADIA_HOSTING_WINDOW_CLASS",
        "title": "manager",  # 確保您的 Windows Terminal Tab 或標題包含此字串
        "hotkey": ("ctrl", "alt", "1"),
        "click": {"monitor": 2, "x": 395, "y": 253},
    },
    "backend": {
        "class_name": "CASCADIA_HOSTING_WINDOW_CLASS",
        "title": "Backend",
        "hotkey": ("ctrl", "alt", "1"),
        "click": {"monitor": 2, "x": 1436, "y": 260},
    },
    "uiux": {
        "class_name": "CASCADIA_HOSTING_WINDOW_CLASS",
        "title": "UIUX",
        "hotkey": ("ctrl", "alt", "1"),
        "click": {"monitor": 2, "x": 1445, "y": 766},
    },
}

DEBOUNCE_SECONDS = 1.0
DRY_RUN = False
RESTORE_FOCUS = True
RESTORE_CLIPBOARD = False # 建議 False，頻繁讀寫剪貼簿容易造成衝突

# 強化剪貼簿重試參數
CLIPBOARD_RETRY_COUNT = 10
CLIPBOARD_RETRY_DELAY_MAX = 0.2

STATUS_WAIT_SECONDS = 60.0
STATUS_POLL_INTERVAL_SECONDS = 0.25
IDLE_STATUS = "idle"
THINKING_STATUS = "thinking"

# ---- Windows focus helpers (ctypes) ----
import ctypes
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
GetWindowText = user32.GetWindowTextW
GetWindowTextLength = user32.GetWindowTextLengthW
GetClassName = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
IsWindowVisible = user32.IsWindowVisible
IsIconic = user32.IsIconic
SetForegroundWindow = user32.SetForegroundWindow
ShowWindow = user32.ShowWindow
GetForegroundWindow = user32.GetForegroundWindow
AttachThreadInput = user32.AttachThreadInput
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
GetCurrentThreadId = kernel32.GetCurrentThreadId
GetWindowRect = user32.GetWindowRect
MonitorFromWindow = user32.MonitorFromWindow
GetMonitorInfoW = user32.GetMonitorInfoW
EnumDisplayMonitors = user32.EnumDisplayMonitors
# Clipboard APIs - 必須設定正確的 argtypes/restype (64-bit 安全)
OpenClipboard = user32.OpenClipboard
OpenClipboard.argtypes = [wintypes.HWND]
OpenClipboard.restype = wintypes.BOOL

CloseClipboard = user32.CloseClipboard
CloseClipboard.restype = wintypes.BOOL

EmptyClipboard = user32.EmptyClipboard
EmptyClipboard.restype = wintypes.BOOL

SetClipboardData = user32.SetClipboardData
SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
SetClipboardData.restype = wintypes.HANDLE

GetClipboardData = user32.GetClipboardData
GetClipboardData.argtypes = [wintypes.UINT]
GetClipboardData.restype = wintypes.HANDLE

IsClipboardFormatAvailable = user32.IsClipboardFormatAvailable
IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
IsClipboardFormatAvailable.restype = wintypes.BOOL

# Global Memory APIs - 關鍵：HGLOBAL 是 64-bit handle
GlobalAlloc = kernel32.GlobalAlloc
GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
GlobalAlloc.restype = wintypes.HGLOBAL

GlobalLock = kernel32.GlobalLock
GlobalLock.argtypes = [wintypes.HGLOBAL]
GlobalLock.restype = wintypes.LPVOID

GlobalUnlock = kernel32.GlobalUnlock
GlobalUnlock.argtypes = [wintypes.HGLOBAL]
GlobalUnlock.restype = wintypes.BOOL

GlobalSize = kernel32.GlobalSize
GlobalSize.argtypes = [wintypes.HGLOBAL]
GlobalSize.restype = ctypes.c_size_t

GetLastError = kernel32.GetLastError
GetLastError.restype = wintypes.DWORD

# BlockInput API - 鎖定用戶輸入（需要管理員權限）
BlockInput = user32.BlockInput
BlockInput.argtypes = [wintypes.BOOL]
BlockInput.restype = wintypes.BOOL

# Shell32 for admin check
shell32 = ctypes.WinDLL("shell32", use_last_error=True)
IsUserAnAdmin = shell32.IsUserAnAdmin
IsUserAnAdmin.restype = wintypes.BOOL

SW_RESTORE = 9
CF_UNICODETEXT = 13
GMEM_MOVEABLE = 0x0002
MONITOR_DEFAULTTONEAREST = 2

# ---- 輸入法切換 ----
GetKeyboardLayout = user32.GetKeyboardLayout
LoadKeyboardLayoutW = user32.LoadKeyboardLayoutW
ActivateKeyboardLayout = user32.ActivateKeyboardLayout
SendMessageW = user32.SendMessageW

WM_INPUTLANGCHANGEREQUEST = 0x0050
INPUTLANGCHANGE_SYSCHARSET = 0x0001
HKL_NEXT = 1
HKL_PREV = 0
KLF_ACTIVATE = 0x00000001

# 英文輸入法代碼
EN_US_LAYOUT = "00000409"  # 美式英文
ENGLISH_LANG_ID = 0x0409   # 英文語言識別碼

# 輸入法切換重試參數
IME_SWITCH_RETRY_COUNT = 5
IME_SWITCH_RETRY_DELAY = 0.15

# 緊急解鎖熱鍵
ABORT_HOTKEY = '<ctrl>+<alt>+q'


def is_admin() -> bool:
    """檢查是否以管理員權限運行"""
    try:
        return IsUserAnAdmin() != 0
    except Exception:
        return False


def _get_current_keyboard_layout(hwnd: int = None) -> int:
    """取得鍵盤佈局的語言識別碼 (LANGID)

    Args:
        hwnd: 目標視窗句柄。如果提供，會取得該視窗所屬執行緒的鍵盤佈局；
              如果為 None，會使用當前前景視窗。
    """
    try:
        # 如果沒有指定 hwnd，使用當前前景視窗
        if hwnd is None:
            hwnd = GetForegroundWindow()

        # GetWindowThreadProcessId 的返回值是 thread ID
        # 第二個參數是 output parameter 用來接收 process ID（我們不需要）
        target_thread = GetWindowThreadProcessId(hwnd, None)
        if not target_thread:
            log(f"[WARN] Failed to get thread ID for hwnd={hwnd}", "WARN")
            return 0

        # GetKeyboardLayout 傳入 thread ID，返回該執行緒的鍵盤佈局
        # 注意：需要傳入視窗的 thread ID，而不是 0（littlebird 自己的 thread）
        hkl = GetKeyboardLayout(target_thread)
        if hkl:
            # 取低 16 位作為 LANGID
            lang_id = hkl & 0xFFFF
            return lang_id
    except Exception as e:
        log(f"[WARN] Failed to get keyboard layout: {e}", "WARN")
    return 0


def _is_english_input(hwnd: int = None) -> bool:
    """檢查目標視窗的輸入法是否為英文"""
    lang_id = _get_current_keyboard_layout(hwnd)
    # 0x0409 = 英文 (美國), 0x0809 = 英文 (英國), etc.
    # 檢查主要語言是否為英文 (語言識別碼的低 10 位 = 0x09 表示英文)
    is_english = (lang_id & 0x00FF) == 0x09
    log(f"[IME] Layout LANGID: 0x{lang_id:04X}, is_english={is_english}")
    return is_english


def _switch_to_english_input(hwnd: int = None) -> bool:
    """切換到英文輸入法，並驗證切換成功

    Args:
        hwnd: 目標視窗句柄，用於驗證該視窗的輸入法狀態
    """
    for attempt in range(IME_SWITCH_RETRY_COUNT):
        # 先檢查當前是否已經是英文（使用目標視窗的輸入法狀態）
        if _is_english_input(hwnd):
            if attempt > 0:
                log(f"[IME] Verified: English input active (attempt {attempt + 1})")
            else:
                log("[IME] Already in English input mode")
            return True

        log(f"[IME] Switching to English input (attempt {attempt + 1}/{IME_SWITCH_RETRY_COUNT})...")

        try:
            # 方法 1：載入並啟用英文輸入法（系統全域）
            hkl = LoadKeyboardLayoutW(EN_US_LAYOUT, KLF_ACTIVATE)
            if hkl:
                ActivateKeyboardLayout(hkl, 0)
                time.sleep(IME_SWITCH_RETRY_DELAY)

                # 驗證切換結果
                if _is_english_input(hwnd):
                    log(f"[IME] Successfully switched to English (method 1)")
                    return True

            # 方法 2：發送訊息給目標視窗，請求切換輸入法
            if hwnd:
                hkl_for_msg = LoadKeyboardLayoutW(EN_US_LAYOUT, 0)
                if hkl_for_msg:
                    SendMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, hkl_for_msg)
                    time.sleep(IME_SWITCH_RETRY_DELAY)

                    # 驗證切換結果
                    if _is_english_input(hwnd):
                        log(f"[IME] Successfully switched to English (method 2)")
                        return True

            # 方法 3：嘗試用 PostMessageW 非同步方式
            PostMessageW = user32.PostMessageW
            if hwnd:
                hkl_for_post = LoadKeyboardLayoutW(EN_US_LAYOUT, 0)
                if hkl_for_post:
                    PostMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, hkl_for_post)
                    time.sleep(IME_SWITCH_RETRY_DELAY * 2)  # 給多一點時間

                    if _is_english_input(hwnd):
                        log(f"[IME] Successfully switched to English (method 3)")
                        return True

        except Exception as e:
            log(f"[WARN] IME switch attempt {attempt + 1} failed: {e}", "WARN")

        time.sleep(IME_SWITCH_RETRY_DELAY)

    # 所有嘗試都失敗
    log(f"[ERROR] Failed to switch to English input after {IME_SWITCH_RETRY_COUNT} attempts", "ERROR")
    return False


# ---- Input Blocker ----

class InputBlocker:
    """管理輸入鎖定狀態，防止用戶操作干擾自動化流程"""

    def __init__(self):
        self._locked = False
        self._lock = threading.Lock()
        self._abort_requested = False
        self._admin_checked = False
        self._is_admin = False

    def _check_admin(self) -> bool:
        """檢查管理員權限（只檢查一次）"""
        if not self._admin_checked:
            self._is_admin = is_admin()
            self._admin_checked = True
            if not self._is_admin:
                log("[WARN] 未以管理員權限運行，輸入鎖定功能將被停用", "WARN")
        return self._is_admin

    def lock(self) -> bool:
        """鎖定用戶輸入"""
        with self._lock:
            if self._locked:
                return True  # 已經鎖定

            if not self._check_admin():
                return False

            try:
                result = BlockInput(True)
                if result:
                    self._locked = True
                    self._abort_requested = False
                    log("[LOCK] 用戶輸入已鎖定")
                    return True
                else:
                    err = GetLastError()
                    log(f"[ERROR] BlockInput(True) 失敗 (LastError={err})", "ERROR")
                    return False
            except Exception as e:
                log(f"[ERROR] BlockInput 異常: {e}", "ERROR")
                return False

    def unlock(self) -> bool:
        """解鎖用戶輸入"""
        with self._lock:
            if not self._locked:
                return True  # 本來就沒鎖定

            try:
                BlockInput(False)
                self._locked = False
                log("[UNLOCK] 用戶輸入已解鎖")
                return True
            except Exception as e:
                log(f"[ERROR] BlockInput(False) 異常: {e}", "ERROR")
                return False

    def request_abort(self):
        """請求中止當前操作並解鎖"""
        log("[ABORT] 收到中止請求")
        self._abort_requested = True
        self.unlock()

    def reset_abort(self):
        """重置中止狀態"""
        self._abort_requested = False

    @property
    def is_locked(self) -> bool:
        return self._locked

    @property
    def abort_requested(self) -> bool:
        return self._abort_requested


# ---- Tray Icon ----

class TrayIcon:
    """系統托盤圖示管理，顯示輸入鎖定狀態"""

    GREEN = (0, 200, 0)
    RED = (200, 0, 0)

    def __init__(self, input_blocker: InputBlocker):
        self.input_blocker = input_blocker
        self.icon = None
        self._thread = None
        self._pystray = None
        self._pil_available = False

    def _load_deps(self):
        """載入依賴庫"""
        try:
            import pystray
            from PIL import Image, ImageDraw
            self._pystray = pystray
            self._pil_available = True
            return True
        except ImportError as e:
            log(f"[WARN] 無法載入托盤圖示依賴: {e}", "WARN")
            log("[INFO] 請安裝: pip install pystray Pillow")
            return False

    def _create_icon_image(self, color):
        """創建圓形燈號圖示"""
        from PIL import Image, ImageDraw
        size = 64
        image = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        # 繪製圓形燈號
        margin = 4
        draw.ellipse([margin, margin, size - margin, size - margin], fill=color)
        return image

    def _create_menu(self):
        """創建右鍵選單"""
        return self._pystray.Menu(
            self._pystray.MenuItem('Littlebird 輸入鎖定', lambda: None, enabled=False),
            self._pystray.Menu.SEPARATOR,
            self._pystray.MenuItem('強制解鎖 (Ctrl+Alt+Q)', self._on_unlock),
            self._pystray.MenuItem('退出', self._on_exit)
        )

    def update_status(self, locked: bool):
        """更新圖示顏色"""
        if self.icon and self._pil_available:
            color = self.RED if locked else self.GREEN
            self.icon.icon = self._create_icon_image(color)
            self.icon.title = "Littlebird: 已鎖定 (Ctrl+Alt+Q 解鎖)" if locked else "Littlebird: 運行中"

    def start(self):
        """啟動托盤圖示（非阻塞）"""
        if not self._load_deps():
            return False

        try:
            self.icon = self._pystray.Icon(
                "littlebird",
                icon=self._create_icon_image(self.GREEN),
                title="Littlebird: 運行中",
                menu=self._create_menu()
            )
            self._thread = threading.Thread(target=self.icon.run, daemon=True)
            self._thread.start()
            log("[OK] 托盤圖示已啟動")
            return True
        except Exception as e:
            log(f"[ERROR] 托盤圖示啟動失敗: {e}", "ERROR")
            return False

    def stop(self):
        """停止托盤圖示"""
        if self.icon:
            try:
                self.icon.stop()
            except Exception:
                pass

    def _on_unlock(self):
        """選單：強制解鎖"""
        self.input_blocker.request_abort()
        self.update_status(False)

    def _on_exit(self):
        """選單：退出"""
        self.input_blocker.unlock()
        self.stop()
        os._exit(0)


# ---- Hotkey Listener ----

class HotkeyListener:
    """全域熱鍵監聽，用於緊急解鎖"""

    def __init__(self, input_blocker: InputBlocker, tray_icon: TrayIcon):
        self.input_blocker = input_blocker
        self.tray_icon = tray_icon
        self._listener = None

    def start(self):
        """啟動熱鍵監聽"""
        try:
            from pynput.keyboard import GlobalHotKeys

            def on_abort():
                log("[HOTKEY] Ctrl+Alt+Q 被按下 - 中止操作")
                self.input_blocker.request_abort()
                if self.tray_icon:
                    self.tray_icon.update_status(False)

            self._listener = GlobalHotKeys({ABORT_HOTKEY: on_abort})
            self._listener.start()
            log(f"[OK] 熱鍵監聽已啟動 ({ABORT_HOTKEY})")
            return True
        except ImportError as e:
            log(f"[WARN] 無法載入熱鍵監聽依賴: {e}", "WARN")
            log("[INFO] 請安裝: pip install pynput")
            return False
        except Exception as e:
            log(f"[ERROR] 熱鍵監聽啟動失敗: {e}", "ERROR")
            return False

    def stop(self):
        """停止熱鍵監聽"""
        if self._listener:
            try:
                self._listener.stop()
            except Exception:
                pass


# ---- 全域實例 ----
input_blocker: InputBlocker | None = None
tray_icon: TrayIcon | None = None
hotkey_listener: HotkeyListener | None = None


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", wintypes.LONG),
        ("top", wintypes.LONG),
        ("right", wintypes.LONG),
        ("bottom", wintypes.LONG),
    ]


class MONITORINFOEXW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", RECT),
        ("rcWork", RECT),
        ("dwFlags", wintypes.DWORD),
        ("szDevice", wintypes.WCHAR * 32),
    ]


MonitorEnumProc = ctypes.WINFUNCTYPE(
    ctypes.c_int,
    wintypes.HMONITOR,
    wintypes.HDC,
    ctypes.POINTER(RECT),
    wintypes.LPARAM,
)


def _list_monitors() -> list[dict]:
    monitors: list[dict] = []

    def callback(hmonitor, _hdc, _lprc, _data):
        info = MONITORINFOEXW()
        info.cbSize = ctypes.sizeof(info)
        if GetMonitorInfoW(hmonitor, ctypes.byref(info)):
            monitors.append(
                {
                    "handle": hmonitor,
                    "left": info.rcMonitor.left,
                    "top": info.rcMonitor.top,
                    "right": info.rcMonitor.right,
                    "bottom": info.rcMonitor.bottom,
                    "device": info.szDevice,
                }
            )
        return True

    EnumDisplayMonitors(0, 0, MonitorEnumProc(callback), 0)
    return monitors


def _monitor_for_window(hwnd: int) -> tuple[int | None, dict | None]:
    hmonitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST)
    monitors = _list_monitors()
    for idx, mon in enumerate(monitors, start=1):
        if mon["handle"] == hmonitor:
            return idx, mon
    return None, None


def _release_modifiers():
    """釋放所有修飾鍵，避免組合鍵殘留導致 Enter 變成 Shift+Enter 等問題"""
    if not pyautogui:
        return
    pyautogui.keyUp('ctrl')
    pyautogui.keyUp('alt')
    pyautogui.keyUp('shift')
    time.sleep(0.05)


def _click_focus(hwnd: int, click_cfg: dict) -> bool:
    if not pyautogui:
        return False
    monitor_index = click_cfg.get("monitor")
    x = click_cfg.get("x")
    y = click_cfg.get("y")
    if monitor_index is None or x is None or y is None:
        log("[WARN] Click config missing monitor/x/y.", "WARN")
        return False

    monitors = _list_monitors()
    if monitor_index < 1 or monitor_index > len(monitors):
        log(f"[WARN] Click monitor index out of range: {monitor_index}", "WARN")
        return False

    mon = monitors[monitor_index - 1]
    abs_x = mon["left"] + int(x)
    abs_y = mon["top"] + int(y)
    pyautogui.click(abs_x, abs_y)
    _release_modifiers()  # 防呆：點擊後釋放修飾鍵
    time.sleep(0.1)
    return GetForegroundWindow() == hwnd


def _get_clipboard_text() -> str | None:
    # 讀取如果失敗就直接放棄，不要卡太久
    if not OpenClipboard(None):
        return None
    try:
        if not IsClipboardFormatAvailable(CF_UNICODETEXT):
            return ""
        handle = GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        ptr = GlobalLock(handle)
        if not ptr:
            return ""
        try:
            size_bytes = GlobalSize(handle)
            if not size_bytes:
                return ""
            max_chars = max(0, int(size_bytes // 2) - 1)
            if max_chars <= 0:
                return ""
            try:
                return ctypes.wstring_at(ptr, max_chars)
            except OSError:
                return ""
        finally:
            GlobalUnlock(handle)
    finally:
        CloseClipboard()


def _set_clipboard_text(text: str) -> bool:
    """Robust clipboard set with exponential backoff."""
    for i in range(CLIPBOARD_RETRY_COUNT):
        if OpenClipboard(None):
            break
        # 隨機延遲避免多個進程死鎖
        time.sleep(random.uniform(0.01, CLIPBOARD_RETRY_DELAY_MAX))
    else:
        err = GetLastError()
        log(f"[ERROR] Failed to open clipboard after {CLIPBOARD_RETRY_COUNT} attempts (LastError={err})", "ERROR")
        return False

    try:
        if not EmptyClipboard():
            err = GetLastError()
            log(f"[ERROR] EmptyClipboard failed (LastError={err})", "ERROR")
            return False

        data = (text + "\0").encode("utf-16-le")
        h_mem = GlobalAlloc(GMEM_MOVEABLE, len(data))
        if not h_mem:
            err = GetLastError()
            log(f"[ERROR] GlobalAlloc failed (LastError={err})", "ERROR")
            return False

        ptr = GlobalLock(h_mem)
        if not ptr:
            err = GetLastError()
            log(f"[ERROR] GlobalLock failed (LastError={err})", "ERROR")
            return False

        try:
            ctypes.memmove(ptr, data, len(data))
        finally:
            GlobalUnlock(h_mem)

        if not SetClipboardData(CF_UNICODETEXT, h_mem):
            err = GetLastError()
            log(f"[ERROR] SetClipboardData failed (LastError={err})", "ERROR")
            return False

        return True
    except Exception as e:
        log(f"[ERROR] Clipboard Exception: {e}", "ERROR")
        return False
    finally:
        CloseClipboard()


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
    title_sub = (config.get("title") or "").lower()
    class_name = (config.get("class_name") or "").lower()
    pid = config.get("pid")
    hwnd_exact = config.get("hwnd")

    # 1. 如果有 HWND 且有效，優先使用 (但不建議依賴)
    if isinstance(hwnd_exact, int) and hwnd_exact > 0:
        if IsWindowVisible(hwnd_exact):
            return hwnd_exact

    # 2. 搜尋匹配
    candidates = []
    for hwnd, title, cls, win_pid in _iter_windows():
        if class_name and class_name not in cls.lower():
            continue
        if title_sub and title_sub not in title.lower():
            continue
        if pid is not None and win_pid != pid:
            continue
        candidates.append(hwnd)

    if not candidates:
        return None
    
    # 如果找到多個，回傳第一個 (假設設定檔夠精確)
    return candidates[0]


def force_focus_window(hwnd: int) -> bool:
    """Aggressively attempt to focus the window."""
    if not hwnd:
        return False
    
    current_fg = GetForegroundWindow()
    if current_fg == hwnd:
        return True

    # 如果最小化，先還原
    if IsIconic(hwnd):
        ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.1)

    # 嘗試直接切換
    SetForegroundWindow(hwnd)
    
    # 檢查是否成功
    for _ in range(5):
        if GetForegroundWindow() == hwnd:
            return True
        time.sleep(0.05)
        # 如果還沒切過去，再次嘗試
        SetForegroundWindow(hwnd)
    
    # 如果還是失敗，嘗試 AttachThreadInput hack (進階手段，通常不需要但備用)
    # 這裡保持簡單，失敗就回傳 False
    return GetForegroundWindow() == hwnd


def _status_path_for(target: str) -> Path | None:
    if target == "backend":
        return BASE_DIR / ".agent-workspace" / "backend" / "current.md"
    if target == "uiux":
        return BASE_DIR / ".agent-workspace" / "ui-ux" / "current.md"
    if target == "manager":
        return BASE_DIR / ".agent-workspace" / "manager" / "current.md"
    return None


def _read_status(path: Path) -> str:
    if not path.exists():
        return ""
    try:
        # 增加 retry 避免檔案鎖定
        text = path.read_text(encoding="utf-8", errors="ignore")
        lines = text.splitlines()
        for idx, line in enumerate(lines):
            if line.strip().lower() == "## 狀態":
                for next_line in lines[idx + 1 :]:
                    if next_line.strip():
                        return next_line.strip().lower()
                break
    except Exception:
        pass
    return ""


def _write_status(path: Path, status: str) -> None:
    try:
        if not path.exists():
            path.parent.mkdir(parents=True, exist_ok=True)
            content = (
                f"# {path.parent.name.capitalize()} Agent - 當前工作\n\n"
                "## 狀態\n"
                f"{status}\n\n"
                "## 當前任務\n"
                "（無）\n\n"
                "## 進行中的變更\n"
                "（無）\n\n"
                "## 待處理項目\n"
                "（無）\n"
            )
            path.write_text(content, encoding="utf-8")
            return

        lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
        for idx, line in enumerate(lines):
            if line.strip().lower() == "## 狀態":
                for j in range(idx + 1, len(lines)):
                    if lines[j].strip():
                        lines[j] = status
                        break
                else:
                    lines.insert(idx + 1, status)
                break
        else:
            lines.extend(["", "## 狀態", status])
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    except Exception as e:
        log(f"[WARN] Failed to write status: {e}", "WARN")


def _wait_for_idle(path: Path) -> bool:
    start = time.time()
    while True:
        status = _read_status(path)
        if status == "" or status == IDLE_STATUS:
            return True
        if time.time() - start >= STATUS_WAIT_SECONDS:
            return False
        time.sleep(STATUS_POLL_INTERVAL_SECONDS)


def route_event(path: Path) -> tuple[str, str] | None:
    path_str = str(path)
    lower = path_str.lower()

    # 檢查是否在 .agent-workspace 目錄下
    if ".agent-workspace" not in lower:
        return None

    if lower.endswith("output.md") and "handoff" in lower:
        parts = path.parts
        task_id = ""
        try:
            handoff_idx = [p.lower() for p in parts].index("handoff")
            # 簡單防禦 index out of range
            if handoff_idx + 1 < len(parts):
                task_id = parts[handoff_idx + 1]
        except Exception:
            task_id = ""
        msg = f"[Manager] output updated: {task_id or path.name}"
        return "manager", msg

    if lower.endswith("notifications.md"):
        if "backend" in lower:
            return "backend", "[Backend] new notification"
        if "ui-ux" in lower or "uiux" in lower:
            return "uiux", "[UI-UX] new notification"

    return None


# ---- Agent completion monitoring ----
COMPLETION_POLL_INTERVAL = 3.0  # 每 3 秒檢查一次
COMPLETION_TIMEOUT = 300.0  # 最多等待 5 分鐘

def _wait_for_agent_completion(agent: str) -> bool:
    """等待 Agent 狀態變回 idle，返回是否成功"""
    status_path = _status_path_for(agent)
    if not status_path:
        return False

    start = time.time()
    while time.time() - start < COMPLETION_TIMEOUT:
        status = _read_status(status_path)
        if status == IDLE_STATUS or status == "":
            return True
        log(f"[WAIT] {agent} still {status}, checking again in {COMPLETION_POLL_INTERVAL}s...")
        time.sleep(COMPLETION_POLL_INTERVAL)

    log(f"[WARN] Timeout waiting for {agent} to complete", "WARN")
    return False


def _notify_manager_completion(agent: str):
    """通知 Manager 某個 Agent 完成了任務"""
    message = f"[{agent.upper()}] Task completed, please review"
    log(f"[NOTIFY] Sending completion notice to manager: {message}")

    # 直接調用處理函數通知 Manager
    with gui_action_lock:
        _process_single_event("manager", message, wait_for_completion=False)


# ---- Debounce queue ----

event_queue: "queue.Queue[Path]" = queue.Queue()
last_sent: dict[tuple[str, str], float] = {}

# 【重要】這個 Lock 必須保護「整個」GUI 操作流程，避免 A 視窗切到一半被 B 視窗搶走
gui_action_lock = threading.Lock()


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

        # 第一層 Debounce 檢查 (快速過濾)
        # 注意：這裡不鎖 gui_action_lock，因為我們只想鎖住真正的 GUI 操作
        # 但 last_sent 的讀寫需要 thread safety，雖然 Python dict 操作是 atomic 的，
        # 為了保險起見，可以加一個小鎖，或者乾脆讓 gui_action_lock 涵蓋大一點。
        # 為了效能，這裡暫時不鎖，允許極少數的 race condition 進入下一層。
        
        last_time = last_sent.get(key, 0.0)
        if now - last_time < DEBOUNCE_SECONDS:
            continue
        last_sent[key] = now

        if DRY_RUN:
            log(f"[DRY_RUN] {target}: {message}")
            continue

        # 【關鍵修正】開始 GUI 操作前，取得全域鎖。
        # 這確保了在操作 Manager 視窗的整個過程中，UIUX 的事件必須排隊等待。
        with gui_action_lock:
            _process_single_event(target, message)


def _process_single_event(target: str, message: str, wait_for_completion: bool = True):
    """
    發送訊息給目標 Agent。
    wait_for_completion: 如果為 True 且目標不是 Manager，會等待 Agent 完成並通知 Manager
    """
    global input_blocker, tray_icon

    config = AGENT_WINDOWS.get(target, {})
    hwnd = find_window(config)

    if not hwnd:
        log(f"[WARN] Window not found for {target}: {config}", "WARN")
        return

    status_path = _status_path_for(target)
    if status_path and not _wait_for_idle(status_path):
        log(f"[WARN] Target {target} not idle; skipping message.", "WARN")
        return

    prev_hwnd = GetForegroundWindow() if RESTORE_FOCUS else None

    # === 輸入鎖定：操作前鎖定用戶輸入 ===
    locked = False
    if input_blocker:
        input_blocker.reset_abort()  # 重置中止狀態
        locked = input_blocker.lock()
        if locked and tray_icon:
            tray_icon.update_status(True)

    try:
        # === 檢查是否被中止 ===
        if input_blocker and input_blocker.abort_requested:
            log(f"[ABORT] Operation aborted before start")
            return

        # 嘗試取得焦點
        if not force_focus_window(hwnd):
            click_cfg = config.get("click")
            if click_cfg and _click_focus(hwnd, click_cfg):
                pass
            else:
                log(f"[WARN] Failed to focus window for {target}. Aborting to avoid mis-typing.", "WARN")
                return

        # 焦點切換後給一點點緩衝，等待 Windows 動畫或 Input Queue 就緒
        time.sleep(0.2)

        # === 檢查是否被中止 ===
        if input_blocker and input_blocker.abort_requested:
            log(f"[ABORT] Operation aborted after focus")
            return

        hotkey = config.get("hotkey")
        if hotkey:
            pyautogui.hotkey(*hotkey)
            _release_modifiers()  # 防呆：hotkey 後釋放修飾鍵
            time.sleep(0.1)

        # 切換到英文輸入法，避免中文輸入法干擾
        # 如果切換失敗，中止操作以避免中文輸入造成卡死
        if not _switch_to_english_input(hwnd):
            log(f"[ERROR] Cannot switch to English input for {target}. Aborting to prevent IME issues.", "ERROR")
            return
        time.sleep(0.1)

        # === 檢查是否被中止 ===
        if input_blocker and input_blocker.abort_requested:
            log(f"[ABORT] Operation aborted before clipboard")
            return

        # 保存原始剪貼簿 (如果需要)
        original_clip = _get_clipboard_text() if RESTORE_CLIPBOARD else None

        # 寫入剪貼簿
        if _set_clipboard_text(message):
            # 貼上
            pyautogui.hotkey("ctrl", "v")
            _release_modifiers()  # 防呆：貼上後釋放修飾鍵
            time.sleep(0.2)  # 增加延遲，等待貼上完成

            # 恢復剪貼簿
            if original_clip is not None:
                # 不用 retry 太多次，如果不重要就跳過
                time.sleep(0.1)
                _set_clipboard_text(original_clip)
        else:
            # 萬不得已才用 typewrite，且加上 interval
            log("[WARN] Clipboard absolutely unavailable; falling back to slow typewrite.", "WARN")
            # 切換到英文輸入法通常很難控制，這裡只能祈禱
            pyautogui.typewrite(message, interval=0.01)
            _release_modifiers()  # 防呆：typewrite 後釋放修飾鍵

        # === 檢查是否被中止 ===
        if input_blocker and input_blocker.abort_requested:
            log(f"[ABORT] Operation aborted before Enter")
            return

        # 按 Enter 前再次確認焦點在目標視窗
        if GetForegroundWindow() != hwnd:
            log(f"[WARN] Focus lost before Enter, re-focusing {target}...", "WARN")
            if not force_focus_window(hwnd):
                log(f"[ERROR] Failed to re-focus {target}, Enter may go to wrong window!", "ERROR")
            time.sleep(0.1)

        # 發送訊息並驗證
        max_retries = 3
        send_success = False

        for attempt in range(max_retries):
            # === 檢查是否被中止 ===
            if input_blocker and input_blocker.abort_requested:
                log(f"[ABORT] Operation aborted during retry loop")
                return

            # 嚴格確認焦點在目標視窗
            current_fg = GetForegroundWindow()
            if current_fg != hwnd:
                log(f"[RETRY {attempt+1}] Focus wrong (current={current_fg}, target={hwnd}), re-focusing {target}...")
                force_focus_window(hwnd)
                time.sleep(0.3)

                # 再次確認
                current_fg = GetForegroundWindow()
                if current_fg != hwnd:
                    log(f"[ERROR] Still not focused after retry, current={current_fg}", "ERROR")
                    continue  # 跳過這次嘗試，進入下一次重試

            # 焦點確認正確，確保修飾鍵已釋放，然後按 Enter
            log(f"[SEND] Focus confirmed (hwnd={hwnd}), pressing Enter for {target} (attempt {attempt+1})")

            _release_modifiers()  # 防呆：Enter 前釋放修飾鍵
            pyautogui.press("enter")

            # 對非 Manager：發送後設定 thinking，然後驗證
            if target != "manager" and status_path:
                time.sleep(0.3)  # 等待一下
                _write_status(status_path, THINKING_STATUS)

                # 驗證狀態確實被設定
                time.sleep(0.1)
                actual_status = _read_status(status_path)
                if actual_status == THINKING_STATUS:
                    log(f"[OK] Message sent to {target} (verified: status={actual_status})")
                    send_success = True
                    break
                else:
                    log(f"[WARN] Status verification failed: expected '{THINKING_STATUS}', got '{actual_status}'", "WARN")
            else:
                # Manager 不需要驗證
                if status_path and target == "manager":
                    _write_status(status_path, IDLE_STATUS)
                log(f"[OK] Message sent to {target}")
                send_success = True
                break

        if not send_success:
            log(f"[ERROR] Failed to send message to {target} after {max_retries} attempts", "ERROR")
            return

        # 復原焦點
        if RESTORE_FOCUS and prev_hwnd and prev_hwnd != hwnd:
            time.sleep(0.1)
            try:
                SetForegroundWindow(prev_hwnd)
            except Exception:
                pass

        # 對非 Manager 的 Agent，等待完成後通知 Manager
        if wait_for_completion and target != "manager":
            log(f"[MONITOR] Starting completion monitor for {target}...")

            def monitor_and_notify():
                if _wait_for_agent_completion(target):
                    log(f"[COMPLETE] {target} finished, notifying manager...")
                    _notify_manager_completion(target)
                else:
                    log(f"[WARN] {target} did not complete in time", "WARN")

            monitor_thread = threading.Thread(target=monitor_and_notify, daemon=True)
            monitor_thread.start()

    finally:
        # === 輸入解鎖：無論如何都要解鎖 ===
        if input_blocker and locked:
            input_blocker.unlock()
            if tray_icon:
                tray_icon.update_status(False)


def load_deps():
    global pyautogui, Observer, FileSystemEventHandler
    try:
        import pyautogui as _pyautogui
        # 調整 pyautogui 的保護設定
        _pyautogui.FAILSAFE = True
        _pyautogui.PAUSE = 0.05 # 每個動作後的微小暫停
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
    log("MON | HWND         | PID   | TITLE")
    log("-" * 70)
    for hwnd, title, cls, win_pid in _iter_windows():
        title_clean = title.replace("\t", " ").replace("\n", " ").strip()
        if "CASCADIA" in cls or "Terminal" in title or "Code" in title:
            mon_idx, _ = _monitor_for_window(hwnd)
            mon_str = f"{mon_idx}" if mon_idx is not None else "-"
            log(f"{mon_str:<3} | {hwnd:<12} | {win_pid:<5} | {title_clean[:40]}")


def main():
    global DRY_RUN, DEBOUNCE_SECONDS
    global input_blocker, tray_icon, hotkey_listener

    # 初始化日誌
    setup_logging()
    log(f"=== Littlebird started at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
    log(f"Log file: {LOG_FILE}")

    parser = argparse.ArgumentParser(description="Littlebird file watcher")
    parser.add_argument("--list-windows", action="store_true", help="List visible windows")
    parser.add_argument("--dry-run", action="store_true", help="Print actions without typing")
    parser.add_argument("--debounce", type=float, default=DEBOUNCE_SECONDS, help="Debounce seconds")
    parser.add_argument("--no-lock", action="store_true", help="Disable input locking (run without admin)")
    args = parser.parse_args()

    if args.list_windows:
        list_windows()
        return

    DRY_RUN = args.dry_run
    DEBOUNCE_SECONDS = args.debounce

    load_deps()

    # === 初始化輸入鎖定系統 ===
    if not args.no_lock:
        # 檢查管理員權限
        if is_admin():
            log("[OK] 以管理員權限運行，輸入鎖定功能已啟用")
        else:
            log("[WARN] 未以管理員權限運行", "WARN")
            log("[INFO] 輸入鎖定功能需要管理員權限")
            log("[INFO] 請以管理員身份重新運行，或使用 --no-lock 參數跳過此功能")

        # 初始化 InputBlocker
        input_blocker = InputBlocker()

        # 初始化 TrayIcon
        tray_icon = TrayIcon(input_blocker)
        tray_icon.start()

        # 初始化 HotkeyListener
        hotkey_listener = HotkeyListener(input_blocker, tray_icon)
        hotkey_listener.start()

        log(f"[INFO] 緊急解鎖熱鍵: Ctrl+Alt+Q")
    else:
        log("[INFO] 輸入鎖定功能已停用 (--no-lock)")
        input_blocker = None
        tray_icon = None
        hotkey_listener = None

    class WatchHandler(FileSystemEventHandler):
        def on_modified(self, event):
            if not event.is_directory:
                event_queue.put(Path(event.src_path))

        def on_created(self, event):
            if not event.is_directory:
                event_queue.put(Path(event.src_path))

        def on_moved(self, event):
            # Claude Code 使用原子寫入：寫入 .tmp -> 刪除原檔 -> 重命名 .tmp
            # 所以需要監聽 moved 事件，dest_path 才是最終檔案路徑
            if not event.is_directory:
                event_queue.put(Path(event.dest_path))

    for path in WATCH_PATHS:
        if not path.exists():
            log(f"[WARN] watch path missing: {path}", "WARN")

    handler = WatchHandler()
    observer = Observer()

    for path in WATCH_PATHS:
        if path.exists():
            observer.schedule(handler, str(path), recursive=True)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

    observer.start()
    log(f"Littlebird running. Watching {len(WATCH_PATHS)} paths.")
    log("Press Ctrl+C to stop.")

    try:
        while True:
            time.sleep(0.5)
    except KeyboardInterrupt:
        log("[INFO] 收到中斷信號，正在停止...")
        observer.stop()
    finally:
        observer.join()
        event_queue.put(None)

        # === 清理輸入鎖定系統 ===
        if input_blocker:
            input_blocker.unlock()  # 確保解鎖
        if hotkey_listener:
            hotkey_listener.stop()
        if tray_icon:
            tray_icon.stop()

        log("[OK] Littlebird 已停止")


if __name__ == "__main__":
    main()
