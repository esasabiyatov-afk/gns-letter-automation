from __future__ import annotations

import hashlib
import os
import re
from pathlib import Path
from typing import BinaryIO


SAFE_NAME_RE = re.compile(r"[^0-9A-Za-zА-Яа-яЁёҢңӨөҮү._ -]+")


class StorageError(ValueError):
    pass


def sanitize_filename(filename: str) -> str:
    source = Path(filename or "document.pdf").name
    cleaned = SAFE_NAME_RE.sub("_", source).strip(" .")
    return cleaned[:160] or "document.pdf"


def save_pdf_stream(
    stream: BinaryIO,
    destination: Path,
    max_bytes: int,
) -> tuple[str, int]:
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.with_suffix(destination.suffix + ".part")
    digest = hashlib.sha256()
    total = 0

    try:
        with temporary.open("wb") as output:
            while chunk := stream.read(1024 * 1024):
                total += len(chunk)
                if total > max_bytes:
                    raise StorageError(
                        f"PDF превышает ограничение {max_bytes // (1024 * 1024)} МБ"
                    )
                digest.update(chunk)
                output.write(chunk)

        with temporary.open("rb") as saved:
            if saved.read(5) != b"%PDF-":
                raise StorageError("Файл не имеет сигнатуру PDF")

        os.replace(temporary, destination)
        return digest.hexdigest(), total
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def ensure_within(path: Path, root: Path) -> Path:
    resolved_path = path.resolve()
    resolved_root = root.resolve()
    if resolved_path != resolved_root and resolved_root not in resolved_path.parents:
        raise StorageError("Попытка доступа вне рабочей папки")
    return resolved_path

