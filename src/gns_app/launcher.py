from __future__ import annotations

import ctypes
import os
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from collections.abc import Callable
from types import TracebackType

from gns_app.diagnostics import record_event, record_exception


class ApplicationAlreadyRunningError(RuntimeError):
    pass


class WindowsSingleInstance:
    """Не позволяет запустить два сервера из двух окон START.bat."""

    MUTEX_NAME = "Local\\GNSLetterAutomationServer"
    ERROR_ALREADY_EXISTS = 183

    def __init__(self) -> None:
        self._handle: int | None = None

    def __enter__(self) -> "WindowsSingleInstance":
        if os.name != "nt":
            return self
        kernel32 = ctypes.windll.kernel32
        create_mutex = kernel32.CreateMutexW
        create_mutex.argtypes = [ctypes.c_void_p, ctypes.c_bool, ctypes.c_wchar_p]
        create_mutex.restype = ctypes.c_void_p
        close_handle = kernel32.CloseHandle
        close_handle.argtypes = [ctypes.c_void_p]
        close_handle.restype = ctypes.c_bool
        handle = create_mutex(None, False, self.MUTEX_NAME)
        if not handle:
            raise OSError("Не удалось создать блокировку приложения.")
        if kernel32.GetLastError() == self.ERROR_ALREADY_EXISTS:
            close_handle(handle)
            raise ApplicationAlreadyRunningError(
                "Приложение уже запущено. Используйте открытое окно браузера."
            )
        self._handle = handle
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_value: BaseException | None,
        traceback: TracebackType | None,
    ) -> None:
        if self._handle is not None:
            ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(self._handle))
            self._handle = None


def supervise(
    run_child: Callable[[], int],
    *,
    wait: Callable[[float], None] = time.sleep,
    restart_delay_seconds: float = 3,
) -> int:
    """Перезапускает сервер после аварийного, но не штатного завершения."""

    while True:
        exit_code = run_child()
        if exit_code == 0:
            return 0
        record_event(
            "launcher",
            "server_exit",
            "crashed",
            details={
                "exit_code": exit_code,
                "restart_delay_seconds": restart_delay_seconds,
            },
        )
        print(
            "Сервер неожиданно завершился. "
            f"Повторный запуск через {restart_delay_seconds:g} сек.",
            flush=True,
        )
        wait(restart_delay_seconds)


def _run_worker(module_name: str, arguments: list[str]) -> int:
    sys.argv = [module_name, *arguments]
    if module_name == "gns_app.services.pdf_service":
        from gns_app.services.pdf_service import main as worker_main
    elif module_name == "gns_app.services.outlook_service":
        from gns_app.services.outlook_service import main as worker_main
    elif module_name == "gns_app.services.scanner_service":
        from gns_app.services.scanner_service import main as worker_main
    else:
        record_event(
            "launcher",
            "worker_dispatch",
            "rejected",
            details={"module": module_name},
        )
        return 64
    return int(worker_main())


def _open_application_when_ready(process: subprocess.Popen[object]) -> None:
    url = "http://127.0.0.1:8765"
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline and process.poll() is None:
        try:
            with urllib.request.urlopen(f"{url}/health", timeout=1) as response:
                if response.status == 200:
                    webbrowser.open(url)
                    record_event("launcher", "browser_open", "success")
                    return
        except (OSError, TimeoutError):
            time.sleep(0.2)
    record_event("launcher", "browser_open", "server_not_ready")


def _server_command() -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--server"]
    return [sys.executable, "-m", "gns_app.main"]


def _run_server_process(command: list[str], *, open_browser: bool) -> int:
    try:
        process = subprocess.Popen(command)
    except OSError as exc:
        record_exception("launcher", "server_start", exc)
        return 1
    if open_browser:
        threading.Thread(
            target=_open_application_when_ready,
            args=(process,),
            name="gns-browser-opener",
            daemon=True,
        ).start()
    return int(process.wait())


def main(arguments: list[str] | None = None) -> int:
    arguments = list(sys.argv[1:] if arguments is None else arguments)
    if arguments[:1] == ["--server"]:
        from gns_app.main import run

        run()
        return 0
    if arguments[:1] == ["--worker-module"]:
        if len(arguments) < 2:
            return 64
        return _run_worker(arguments[1], arguments[2:])

    command = _server_command()
    browser_pending = "--background" not in arguments
    record_event(
        "launcher",
        "application_start",
        "started",
        details={"frozen": bool(getattr(sys, "frozen", False))},
    )

    def run_child() -> int:
        nonlocal browser_pending
        result = _run_server_process(command, open_browser=browser_pending)
        browser_pending = False
        return result

    try:
        with WindowsSingleInstance():
            return supervise(run_child)
    except ApplicationAlreadyRunningError as exc:
        webbrowser.open("http://127.0.0.1:8765")
        record_event("launcher", "duplicate_start", "already_running")
        print(str(exc), flush=True)
        return 2
    except KeyboardInterrupt:
        record_event("launcher", "application_stop", "user_requested")
        return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except BaseException as exc:
        if isinstance(exc, SystemExit):
            raise
        record_exception("launcher", "unhandled_exception", exc)
        raise
