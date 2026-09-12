from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

from gns_app.runtime_commands import module_command


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


def exact_word_page_count(path: Path, timeout_seconds: int = 45) -> int | None:
    """Ask installed Word for the final layout page count in an isolated worker."""
    source = path.resolve()
    if not source.is_file() or os.name != "nt":
        return None
    command = module_command(
        "gns_app.services.word_desktop", "--page-count-worker", str(source)
    )
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout_seconds,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            check=False,
        )
        payload = json.loads(result.stdout.strip() or "{}")
        count = int(payload.get("page_count") or 0)
    except (OSError, subprocess.TimeoutExpired, ValueError, json.JSONDecodeError):
        return None
    return count if result.returncode == 0 and count > 0 else None


def _word_page_count_worker(path: Path) -> int:
    word = None
    document = None
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            str(path.resolve()), ReadOnly=True, AddToRecentFiles=False
        )
        document.Repaginate()
        page_count = int(document.ComputeStatistics(2))  # wdStatisticPages
        print(json.dumps({"page_count": page_count}), flush=True)
        return 0 if page_count > 0 else 1
    except Exception:
        print(json.dumps({"page_count": 0}), flush=True)
        return 1
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit(False)
            except Exception:
                pass
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:
            pass


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--page-count-worker", type=Path)
    args = parser.parse_args(arguments)
    if args.page_count_worker:
        return _word_page_count_worker(args.page_count_worker)
    return 64


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
