from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path

from gns_app.diagnostics import record_event, record_exception
from gns_app.runtime_commands import module_command
from gns_app.services.windows_focus import (
    focus_next_dialog_for_current_process,
)


WIA_SCANNER_DEVICE = 1
# Не навязываем профиль поверх выбора в штатном окне WIA: цвет, DPI, лоток и
# другие аппаратные параметры остаются за настройками драйвера.
WIA_UNSPECIFIED_INTENT = 0
# У WIA нет настройки физической скорости сканирования. Это единственный bias
# в сторону меньшего (и обычно быстрее передаваемого) результата; фактическую
# скорость всё равно определяет драйвер.
WIA_MINIMIZE_SIZE = 65536
WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"


class ScannerError(RuntimeError):
    """Безопасная для показа сотруднику ошибка сканера."""


class ScannerCancelled(ScannerError):
    pass


class ScannerService:
    def acquire_a4(self, destination: Path, timeout_seconds: int = 900) -> Path:
        started = time.monotonic()
        record_event("scanner", "wia_acquire", "started")
        if os.name != "nt":
            record_event("scanner", "wia_acquire", "unsupported_platform")
            raise ScannerError("Сканирование через WIA доступно только в Windows")
        destination.parent.mkdir(parents=True, exist_ok=True)
        command = module_command(
            "gns_app.services.scanner_service",
            "--wia-worker",
            str(destination),
        )
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout_seconds,
                creationflags=creationflags,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            record_exception(
                "scanner",
                "wia_acquire",
                exc,
                details={"duration_ms": int((time.monotonic() - started) * 1000)},
            )
            raise ScannerError("Окно сканирования не завершено вовремя") from exc
        message = _worker_message(result.stdout)
        if result.returncode == 2:
            record_event(
                "scanner",
                "wia_acquire",
                "cancelled",
                details={"duration_ms": int((time.monotonic() - started) * 1000)},
            )
            raise ScannerCancelled(message or "Сканирование отменено")
        if result.returncode != 0 or not destination.is_file():
            record_event(
                "scanner",
                "wia_acquire",
                "error",
                details={
                    "duration_ms": int((time.monotonic() - started) * 1000),
                    "worker_exit_code": result.returncode,
                    "output_created": destination.is_file(),
                },
            )
            raise ScannerError(message or "Не удалось получить изображение со сканера")
        record_event(
            "scanner",
            "wia_acquire",
            "success",
            details={"duration_ms": int((time.monotonic() - started) * 1000)},
        )
        return destination


def _worker_message(output: str) -> str:
    try:
        payload = json.loads(output.strip() or "{}")
    except json.JSONDecodeError:
        return ""
    return str(payload.get("message") or "")


def _wia_worker(destination: Path) -> int:
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print(json.dumps({"message": "Компонент WIA/pywin32 не установлен"}))
        return 1

    pythoncom.CoInitialize()
    try:
        dialog = win32com.client.Dispatch("WIA.CommonDialog")
        # WIA запускается из фонового HTTP-запроса. Наблюдатель поднимает
        # штатный диалог сканера поверх браузера, не нажимая кнопки за
        # сотрудника.
        focus_next_dialog_for_current_process()
        image = dialog.ShowAcquireImage(
            WIA_SCANNER_DEVICE,
            WIA_UNSPECIFIED_INTENT,
            WIA_MINIMIZE_SIZE,
            WIA_FORMAT_PNG,
            False,
            True,
            False,
        )
        if image is None:
            print(json.dumps({"message": "Сканирование отменено"}))
            return 2
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.unlink(missing_ok=True)
        image.SaveFile(str(destination))
        print(json.dumps({"message": "ok"}))
        return 0
    except Exception as exc:
        message = str(exc)
        if "cancel" in message.casefold() or "отмен" in message.casefold():
            print(json.dumps({"message": "Сканирование отменено"}))
            return 2
        print(json.dumps({"message": "WIA не смог получить лист со сканера"}))
        return 1
    finally:
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--wia-worker", type=Path)
    args = parser.parse_args()
    if args.wia_worker:
        return _wia_worker(args.wia_worker)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
