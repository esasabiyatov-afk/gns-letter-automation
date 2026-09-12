from __future__ import annotations

import os
from pathlib import Path


class WordDesktopError(RuntimeError):
    """Ошибка открытия готового ответа в установленном Microsoft Word."""


def open_word_document(path: Path) -> None:
    """Open through the Windows file association, as Explorer does."""
    source = path.resolve()
    if not source.is_file():
        raise WordDesktopError("Файл готового ответа не найден")
    if os.name != "nt":
        raise WordDesktopError("Открытие Microsoft Word доступно только в Windows")

    try:
        os.startfile(str(source))
    except OSError as exc:
        raise WordDesktopError(
            "Windows не смогла открыть готовый ответ штатной программой"
        ) from exc


def print_word_document(path: Path) -> None:
    """Send one document to Word's registered print command."""
    source = path.resolve()
    if not source.is_file():
        raise WordDesktopError("Файл готового ответа не найден")
    if os.name != "nt":
        raise WordDesktopError("Печать Word доступна только в Windows")
    try:
        os.startfile(str(source), "print")
    except OSError as exc:
        raise WordDesktopError(
            "Windows не смогла передать готовый ответ на печать"
        ) from exc
