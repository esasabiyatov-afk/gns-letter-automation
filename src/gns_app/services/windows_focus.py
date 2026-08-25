from __future__ import annotations

import ctypes
import os
import sys
import threading
import time
from ctypes import wintypes
from typing import Iterable


SW_RESTORE = 9
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2


def focus_window_handle(hwnd: int) -> bool:
    """Best-effort foreground activation compatible with Windows 8.1."""
    if sys.platform != "win32" or not hwnd:
        return False
    try:
        user32 = ctypes.windll.user32
        user32.ShowWindow(wintypes.HWND(hwnd), SW_RESTORE)
        flags = SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW
        user32.SetWindowPos(
            wintypes.HWND(hwnd),
            wintypes.HWND(HWND_TOPMOST),
            0,
            0,
            0,
            0,
            flags,
        )
        user32.SetWindowPos(
            wintypes.HWND(hwnd),
            wintypes.HWND(HWND_NOTOPMOST),
            0,
            0,
            0,
            0,
            flags,
        )
        return bool(user32.SetForegroundWindow(wintypes.HWND(hwnd)))
    except (AttributeError, OSError, ValueError):
        return False


def _window_text(user32, hwnd: int) -> str:
    length = int(user32.GetWindowTextLengthW(wintypes.HWND(hwnd)) or 0)
    buffer = ctypes.create_unicode_buffer(max(1, length + 1))
    user32.GetWindowTextW(wintypes.HWND(hwnd), buffer, len(buffer))
    return buffer.value


def _window_class(user32, hwnd: int) -> str:
    buffer = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(wintypes.HWND(hwnd), buffer, len(buffer))
    return buffer.value


def find_and_focus_window(
    *,
    process_id: int | None = None,
    title_parts: Iterable[str] = (),
    class_parts: Iterable[str] = (),
) -> bool:
    if sys.platform != "win32":
        return False
    user32 = ctypes.windll.user32
    normalized_titles = tuple(
        value.casefold() for value in title_parts if value
    )
    normalized_classes = tuple(
        value.casefold() for value in class_parts if value
    )
    matches: list[int] = []
    callback_type = ctypes.WINFUNCTYPE(
        wintypes.BOOL,
        wintypes.HWND,
        wintypes.LPARAM,
    )

    @callback_type
    def callback(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True
        if process_id is not None:
            owner = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if int(owner.value) != int(process_id):
                return True
        title = _window_text(user32, hwnd).casefold()
        class_name = _window_class(user32, hwnd).casefold()
        if normalized_titles and not any(part in title for part in normalized_titles):
            return True
        if normalized_classes and not any(
            part in class_name for part in normalized_classes
        ):
            return True
        matches.append(int(hwnd))
        return False

    user32.EnumWindows(callback, 0)
    return bool(matches and focus_window_handle(matches[0]))


def start_foreground_watcher(
    *,
    process_id: int | None = None,
    title_parts: Iterable[str] = (),
    class_parts: Iterable[str] = (),
    timeout_seconds: float = 15.0,
) -> threading.Thread | None:
    if sys.platform != "win32":
        return None
    titles = tuple(title_parts)
    classes = tuple(class_parts)

    def watch() -> None:
        deadline = time.monotonic() + timeout_seconds
        while time.monotonic() < deadline:
            if find_and_focus_window(
                process_id=process_id,
                title_parts=titles,
                class_parts=classes,
            ):
                return
            time.sleep(0.15)

    thread = threading.Thread(
        target=watch,
        name="gns-window-focus",
        daemon=True,
    )
    thread.start()
    return thread


def focus_next_dialog_for_current_process(
    timeout_seconds: float = 20.0,
) -> threading.Thread | None:
    return start_foreground_watcher(
        process_id=os.getpid(),
        class_parts=("#32770",),
        timeout_seconds=timeout_seconds,
    )
