from __future__ import annotations

import ctypes
import os
import sys
import threading
import time
from dataclasses import dataclass
from ctypes import wintypes
from typing import Iterable


SW_RESTORE = 9
SWP_NOSIZE = 0x0001
SWP_NOMOVE = 0x0002
SWP_SHOWWINDOW = 0x0040
HWND_TOPMOST = -1
HWND_NOTOPMOST = -2
BM_CLICK = 0x00F5
IDYES = 6
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000

OUTLOOK_CERTIFICATE_TITLE_MARKERS = (
    "internet security warning",
    "security warning",
    "предупреждение безопасности интернета",
    "предупреждение системы безопасности в интернете",
    "предупреждение системы безопасности",
)
OUTLOOK_CERTIFICATE_TEXT_MARKERS = (
    "certificate",
    "сертификат",
)
OUTLOOK_CERTIFICATE_PROBLEM_MARKERS = (
    "expired",
    "not yet valid",
    "cannot be verified",
    "could not be verified",
    "истек",
    "просроч",
    "еще не действителен",
    "ещё не действителен",
    "не может быть проверен",
    "не удалось проверить",
    "отсутствует отношение доверия",
    "корневом сертификате",
)
OUTLOOK_CERTIFICATE_YES_BUTTONS = frozenset(
    {
        "да",
        "yes",
        "продолжить",
        "continue",
    }
)


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


def _looks_like_outlook_certificate_warning(
    title: str,
    child_texts: Iterable[str],
) -> bool:
    normalized_title = title.casefold()
    normalized_text = "\n".join(child_texts).casefold()
    return (
        any(marker in normalized_title for marker in OUTLOOK_CERTIFICATE_TITLE_MARKERS)
        and any(marker in normalized_text for marker in OUTLOOK_CERTIFICATE_TEXT_MARKERS)
        and any(marker in normalized_text for marker in OUTLOOK_CERTIFICATE_PROBLEM_MARKERS)
    )


def _is_outlook_certificate_yes_button(text: str, control_id: int) -> bool:
    normalized_text = text.replace("&", "").strip().casefold()
    return (
        control_id == IDYES
        or normalized_text in OUTLOOK_CERTIFICATE_YES_BUTTONS
    )


def _window_process_image(user32, kernel32, hwnd: int) -> str:
    process_id = wintypes.DWORD()
    user32.GetWindowThreadProcessId(
        wintypes.HWND(hwnd), ctypes.byref(process_id)
    )
    if not process_id.value:
        return ""
    process = kernel32.OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION,
        False,
        process_id.value,
    )
    if not process:
        return ""
    try:
        buffer = ctypes.create_unicode_buffer(32768)
        size = wintypes.DWORD(len(buffer))
        if not kernel32.QueryFullProcessImageNameW(
            process, 0, buffer, ctypes.byref(size)
        ):
            return ""
        return buffer.value
    finally:
        kernel32.CloseHandle(process)


def confirm_outlook_certificate_dialog_once() -> bool:
    """Confirm only a recognized invalid-certificate dialog owned by Outlook."""
    if sys.platform != "win32":
        return False
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
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
        if _window_class(user32, hwnd) != "#32770":
            return True
        image = _window_process_image(user32, kernel32, hwnd).casefold()
        if not image.endswith("\\outlook.exe"):
            return True

        child_texts: list[str] = []
        yes_buttons: list[int] = []
        child_callback_type = ctypes.WINFUNCTYPE(
            wintypes.BOOL,
            wintypes.HWND,
            wintypes.LPARAM,
        )

        @child_callback_type
        def child_callback(child, _child_lparam):
            text = _window_text(user32, child)
            if text:
                child_texts.append(text)
            if (
                _window_class(user32, child).casefold() == "button"
                and _is_outlook_certificate_yes_button(
                    text,
                    int(user32.GetDlgCtrlID(wintypes.HWND(child))),
                )
            ):
                yes_buttons.append(int(child))
            return True

        user32.EnumChildWindows(
            wintypes.HWND(hwnd), child_callback, 0
        )
        if (
            yes_buttons
            and _looks_like_outlook_certificate_warning(
                _window_text(user32, hwnd), child_texts
            )
        ):
            matches.append(yes_buttons[0])
            return False
        return True

    user32.EnumWindows(callback, 0)
    if not matches:
        return False
    return bool(
        user32.PostMessageW(
            wintypes.HWND(matches[0]), BM_CLICK, 0, 0
        )
    )


@dataclass(slots=True)
class OutlookCertificateDialogWatcher:
    thread: threading.Thread
    stop_event: threading.Event
    confirmed_event: threading.Event

    def finish(self, grace_seconds: float = 5.0) -> bool:
        self.confirmed_event.wait(max(0.0, grace_seconds))
        self.stop_event.set()
        self.thread.join(timeout=1.0)
        return self.confirmed_event.is_set()


def start_outlook_certificate_dialog_watcher(
    *,
    timeout_seconds: float = 20.0,
) -> OutlookCertificateDialogWatcher | None:
    """Watch briefly and confirm only Outlook's known certificate warning."""
    if sys.platform != "win32":
        return None
    stop_event = threading.Event()
    confirmed_event = threading.Event()

    def watch() -> None:
        deadline = time.monotonic() + timeout_seconds
        while not stop_event.is_set() and time.monotonic() < deadline:
            if confirm_outlook_certificate_dialog_once():
                confirmed_event.set()
                return
            stop_event.wait(0.15)

    thread = threading.Thread(
        target=watch,
        name="gns-outlook-certificate-confirmation",
        daemon=True,
    )
    watcher = OutlookCertificateDialogWatcher(
        thread=thread,
        stop_event=stop_event,
        confirmed_event=confirmed_event,
    )
    thread.start()
    return watcher


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
