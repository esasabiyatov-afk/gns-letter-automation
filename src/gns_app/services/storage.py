from __future__ import annotations

import hashlib
import os
import re
import warnings
from pathlib import Path
from typing import BinaryIO

from PIL import Image, UnidentifiedImageError


SAFE_NAME_RE = re.compile(r"[^0-9A-Za-zА-Яа-яЁёҢңӨөҮү._ -]+")
SUPPORTED_INPUT_EXTENSIONS = frozenset(
    {".pdf", ".bmp", ".png", ".jpg", ".jpeg"}
)
IMAGE_INPUT_EXTENSIONS = SUPPORTED_INPUT_EXTENSIONS - {".pdf"}
IMAGE_FORMATS_BY_EXTENSION = {
    ".bmp": frozenset({"BMP"}),
    ".png": frozenset({"PNG"}),
    ".jpg": frozenset({"JPEG"}),
    ".jpeg": frozenset({"JPEG"}),
}
MAX_INPUT_IMAGE_PIXELS = 60_000_000


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


def validate_input_document(path: Path, suffix: str) -> None:
    normalized_suffix = suffix.casefold()
    if normalized_suffix == ".pdf":
        with path.open("rb") as saved:
            if saved.read(5) != b"%PDF-":
                raise StorageError("Файл не имеет сигнатуру PDF")
        return
    if normalized_suffix not in IMAGE_INPUT_EXTENSIONS:
        raise StorageError("Неподдерживаемый формат входящего документа")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(path) as image:
                detected_format = str(image.format or "").upper()
                if (
                    detected_format
                    not in IMAGE_FORMATS_BY_EXTENSION[normalized_suffix]
                ):
                    raise StorageError(
                        "Расширение изображения не соответствует содержимому файла"
                    )
                width, height = image.size
                if (
                    width < 1
                    or height < 1
                    or width * height > MAX_INPUT_IMAGE_PIXELS
                ):
                    raise StorageError("Недопустимый размер входящего изображения")
                if int(getattr(image, "n_frames", 1) or 1) != 1:
                    raise StorageError(
                        "Поддерживаются только однокадровые изображения"
                    )
                image.verify()
    except StorageError:
        raise
    except (Image.DecompressionBombError, Image.DecompressionBombWarning):
        raise StorageError("Входящее изображение слишком большое") from None
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise StorageError("Не удалось прочитать входящее изображение") from exc


def save_input_stream(
    stream: BinaryIO,
    destination: Path,
    max_bytes: int,
    suffix: str,
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
                        f"Документ превышает ограничение "
                        f"{max_bytes // (1024 * 1024)} МБ"
                    )
                digest.update(chunk)
                output.write(chunk)
        validate_input_document(temporary, suffix)
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
