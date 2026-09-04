from __future__ import annotations

import argparse
import json
import os
import subprocess
import time
from pathlib import Path

from gns_app.diagnostics import record_event, record_exception
from gns_app.runtime_commands import module_command
from gns_app.services.windows_focus import (
    focus_next_dialog_for_current_process,
)


WIA_SCANNER_DEVICE = 1
# У WIA нет настройки физической скорости сканирования. Это единственный bias
# в сторону меньшего (и обычно быстрее передаваемого) результата; фактическую
# скорость всё равно определяет драйвер.
WIA_MINIMIZE_SIZE = 65536
WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
WIA_IPS_CUR_INTENT = 6146
WIA_IPS_XRES = 6147
WIA_IPS_YRES = 6148

DEFAULT_SCANNER_DPI = 150
SCANNER_DPI_CHOICES = (150, 200, 300)
DEFAULT_SCANNER_COLOR_MODE = "grayscale"
SCANNER_COLOR_INTENTS = {
    "color": 1,
    "grayscale": 2,
    "black_white": 4,
}


class ScannerError(RuntimeError):
    """Безопасная для показа сотруднику ошибка сканера."""


class ScannerCancelled(ScannerError):
    pass


class ScannerService:
    def acquire_a4(
        self,
        destination: Path,
        timeout_seconds: int = 900,
        *,
        dpi: int = DEFAULT_SCANNER_DPI,
        color_mode: str = DEFAULT_SCANNER_COLOR_MODE,
    ) -> Path:
        dpi, color_mode = _validate_scanner_profile(dpi, color_mode)
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
            "--dpi",
            str(dpi),
            "--color-mode",
            color_mode,
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


def _validate_scanner_profile(dpi: int, color_mode: str) -> tuple[int, str]:
    try:
        normalized_dpi = int(dpi)
    except (TypeError, ValueError) as exc:
        raise ScannerError("Указано некорректное разрешение сканирования") from exc
    if normalized_dpi not in SCANNER_DPI_CHOICES:
        raise ScannerError(
            "Разрешение сканирования должно быть 150, 200 или 300 DPI"
        )
    normalized_mode = str(color_mode).strip().casefold()
    if normalized_mode not in SCANNER_COLOR_INTENTS:
        raise ScannerError("Указан неизвестный режим изображения сканера")
    return normalized_dpi, normalized_mode


def _first_wia_item(device):
    try:
        items = device.Items
    except Exception as exc:
        raise ScannerError(
            "WIA-драйвер не предоставил область сканирования"
        ) from exc

    item_method = getattr(items, "Item", None)
    if callable(item_method):
        try:
            return item_method(1)
        except Exception:
            pass
    try:
        return items[1]
    except Exception as exc:
        raise ScannerError(
            "WIA-драйвер не предоставил область сканирования"
        ) from exc


def _wia_property(item, property_id: int):
    try:
        properties = item.Properties
    except Exception:
        return None

    item_method = getattr(properties, "Item", None)
    if callable(item_method):
        for key in (str(property_id), property_id):
            try:
                return item_method(key)
            except Exception:
                pass
    for key in (str(property_id), property_id):
        try:
            return properties[key]
        except Exception:
            pass
    try:
        for candidate in properties:
            candidate_id = getattr(candidate, "PropertyID", None)
            if candidate_id is None:
                candidate_id = getattr(candidate, "PropertyId", None)
            if int(candidate_id) == property_id:
                return candidate
    except Exception:
        pass
    return None


def _set_wia_property(item, property_id: int, value: int, label: str) -> None:
    prop = _wia_property(item, property_id)
    if prop is None:
        raise ScannerError(
            f"WIA-драйвер не поддерживает настройку «{label}». "
            "Выберите другое значение или загрузите готовый файл."
        )
    try:
        prop.Value = value
    except Exception as exc:
        raise ScannerError(
            f"WIA-драйвер не поддерживает выбранное значение «{label}». "
            "Выберите другое значение или загрузите готовый файл."
        ) from exc


def _apply_wia_profile(item, dpi: int, color_mode: str) -> None:
    # Сначала задаём тип изображения: некоторые драйверы при этом обновляют
    # доступный диапазон разрешений. DPI устанавливается после режима.
    intent = SCANNER_COLOR_INTENTS[color_mode] | WIA_MINIMIZE_SIZE
    _set_wia_property(
        item,
        WIA_IPS_CUR_INTENT,
        intent,
        "режим изображения",
    )
    _set_wia_property(item, WIA_IPS_XRES, dpi, "разрешение по горизонтали")
    _set_wia_property(item, WIA_IPS_YRES, dpi, "разрешение по вертикали")


def _wia_worker(
    destination: Path,
    dpi: int = DEFAULT_SCANNER_DPI,
    color_mode: str = DEFAULT_SCANNER_COLOR_MODE,
) -> int:
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print(json.dumps({"message": "Компонент WIA/pywin32 не установлен"}))
        return 1

    pythoncom.CoInitialize()
    try:
        dpi, color_mode = _validate_scanner_profile(dpi, color_mode)
        dialog = win32com.client.Dispatch("WIA.CommonDialog")
        # WIA запускается из фонового HTTP-запроса. Наблюдатель поднимает
        # штатный диалог сканера поверх браузера, не нажимая кнопки за
        # сотрудника.
        focus_next_dialog_for_current_process()
        device = dialog.ShowSelectDevice(
            WIA_SCANNER_DEVICE,
            False,
            False,
        )
        if device is None:
            print(json.dumps({"message": "Сканирование отменено"}))
            return 2
        item = _first_wia_item(device)
        _apply_wia_profile(item, dpi, color_mode)
        image = dialog.ShowTransfer(item, WIA_FORMAT_PNG, False)
        if image is None:
            print(json.dumps({"message": "Сканирование отменено"}))
            return 2
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.unlink(missing_ok=True)
        image.SaveFile(str(destination))
        print(json.dumps({"message": "ok"}))
        return 0
    except ScannerError as exc:
        print(json.dumps({"message": str(exc)}, ensure_ascii=False))
        return 1
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
    parser.add_argument("--dpi", type=int, default=DEFAULT_SCANNER_DPI)
    parser.add_argument(
        "--color-mode",
        choices=tuple(SCANNER_COLOR_INTENTS),
        default=DEFAULT_SCANNER_COLOR_MODE,
    )
    args = parser.parse_args()
    if args.wia_worker:
        return _wia_worker(args.wia_worker, args.dpi, args.color_mode)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
