from __future__ import annotations

import os
from pathlib import Path
from typing import Any


WD_WINDOW_STATE_NORMAL = 0
WD_WINDOW_STATE_MINIMIZE = 2


class WordDesktopError(RuntimeError):
    """Ошибка открытия готового ответа в установленном Microsoft Word."""


def _same_document_path(value: Any, expected: Path) -> bool:
    try:
        return Path(str(value)).resolve() == expected
    except (OSError, RuntimeError, ValueError):
        return False


def open_word_document(path: Path) -> None:
    """Open the document through Word COM and bring it to the foreground."""
    source = path.resolve()
    if not source.is_file():
        raise WordDesktopError("Файл готового ответа не найден")
    if os.name != "nt":
        raise WordDesktopError("Открытие Microsoft Word доступно только в Windows")

    try:
        import pythoncom  # type: ignore[import-not-found]
        from win32com import client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise WordDesktopError(
            "Не установлен локальный компонент связи с Microsoft Word"
        ) from exc

    pythoncom.CoInitialize()
    try:
        application = None
        get_active = getattr(client, "GetActiveObject", None)
        if callable(get_active):
            try:
                application = get_active("Word.Application")
            except Exception:
                application = None
        if application is None:
            application = client.Dispatch("Word.Application")

        application.Visible = True
        documents = application.Documents
        document = None
        try:
            count = int(documents.Count)
        except Exception:
            count = 0
        for index in range(1, count + 1):
            candidate = documents.Item(index)
            if _same_document_path(getattr(candidate, "FullName", ""), source):
                document = candidate
                break
        if document is None:
            document = documents.Open(str(source))

        document.Activate()
        application.Activate()
        try:
            if int(application.WindowState) == WD_WINDOW_STATE_MINIMIZE:
                application.WindowState = WD_WINDOW_STATE_NORMAL
        except Exception:
            pass
        application.Activate()
    except WordDesktopError:
        raise
    except Exception as exc:
        raise WordDesktopError(
            "Microsoft Word не смог открыть готовый ответ"
        ) from exc
    finally:
        pythoncom.CoUninitialize()
