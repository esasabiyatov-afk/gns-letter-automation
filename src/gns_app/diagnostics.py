from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import platform
import re
import sys
import threading
import traceback
import zipfile
from datetime import datetime, timezone
from importlib import metadata
from pathlib import Path
from typing import Any
from uuid import uuid4


APP_VERSION = "0.1.1-office-test"
_MAX_LOG_BYTES = 2 * 1024 * 1024
_MAX_ARCHIVES = 3
_WRITE_LOCK = threading.Lock()
_EMAIL_RE = re.compile(r"(?<![\w.+-])[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}(?![\w.-])")
_INN_RE = re.compile(r"(?<!\d)\d{14}(?!\d)")
_WINDOWS_USER_RE = re.compile(r"(?i)([A-Z]:\\Users\\)[^\\\s]+")
_SECRET_TEXT_RE = re.compile(
    r"(?i)\b(password|passwd|парол[ья]?|token|токен|authorization|cookie|csrf)\b"
    r"\s*[:=]\s*[^\s,;]+"
)
_SENSITIVE_KEY_PARTS = (
    "password",
    "passwd",
    "login",
    "username",
    "token",
    "cookie",
    "authorization",
    "csrf",
    "body",
    "document",
    "attachment_path",
    "file_path",
    "taxpayer",
    "inn",
    "pin",
    "fio",
    "name",
    "email",
    "subject",
)
_SAFE_PACKAGE_NAMES = (
    "fastapi",
    "uvicorn",
    "pypdf",
    "pypdfium2",
    "pillow",
    "opencv-python-headless",
    "pywin32",
    "tesserocr",
    "zxing-cpp",
)
_PACKAGE_IMPORT_NAMES = {
    "fastapi": "fastapi",
    "uvicorn": "uvicorn",
    "pypdf": "pypdf",
    "pypdfium2": "pypdfium2",
    "pillow": "PIL",
    "opencv-python-headless": "cv2",
    "pywin32": "pythoncom",
    "tesserocr": "tesserocr",
    "zxing-cpp": "zxingcpp",
}


def _default_runtime_dir() -> Path:
    configured = os.environ.get("GNS_RUNTIME_DIR", "").strip()
    if configured:
        return Path(configured).resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "runtime"
    return Path(__file__).resolve().parents[2] / "runtime"


def _safe_text(value: object, *, limit: int = 240) -> str:
    text = str(value).replace("\r", " ").replace("\n", " ")
    text = _EMAIL_RE.sub("[email removed]", text)
    text = _INN_RE.sub("[id removed]", text)
    text = _WINDOWS_USER_RE.sub(r"\1[user]", text)
    text = _SECRET_TEXT_RE.sub(r"\1=[removed]", text)
    return text[:limit]


def _safe_details(details: dict[str, Any] | None) -> dict[str, Any]:
    safe: dict[str, Any] = {}
    for key, value in (details or {}).items():
        normalized_key = str(key).casefold()
        if any(part in normalized_key for part in _SENSITIVE_KEY_PARTS):
            safe[str(key)] = "[removed]"
            continue
        if value is None or isinstance(value, (bool, int, float)):
            safe[str(key)] = value
        elif isinstance(value, str):
            safe[str(key)] = _safe_text(value)
        elif isinstance(value, (list, tuple)):
            safe[str(key)] = [
                _safe_text(item, limit=80) for item in value[:20]
            ]
        else:
            safe[str(key)] = _safe_text(type(value).__name__)
    return safe


def _log_path(runtime_dir: Path | None = None) -> Path:
    root = (runtime_dir or _default_runtime_dir()).resolve()
    path = root / "diagnostics" / "gns-diagnostics.jsonl"
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _rotate(path: Path) -> None:
    if not path.exists() or path.stat().st_size < _MAX_LOG_BYTES:
        return
    oldest = path.with_name(f"{path.name}.{_MAX_ARCHIVES}")
    oldest.unlink(missing_ok=True)
    for index in range(_MAX_ARCHIVES - 1, 0, -1):
        source = path.with_name(f"{path.name}.{index}")
        if source.exists():
            source.replace(path.with_name(f"{path.name}.{index + 1}"))
    path.replace(path.with_name(f"{path.name}.1"))


def record_event(
    component: str,
    event: str,
    status: str,
    *,
    details: dict[str, Any] | None = None,
    runtime_dir: Path | None = None,
) -> None:
    """Append one privacy-safe diagnostic event.

    Callers must pass counts and technical states, never document contents or
    credentials. A second redaction layer protects against accidental values.
    Logging failures are deliberately non-fatal for the main workflow.
    """

    try:
        path = _log_path(runtime_dir)
        payload = {
            "timestamp": datetime.now(timezone.utc).astimezone().isoformat(
                timespec="seconds"
            ),
            "version": APP_VERSION,
            "component": _safe_text(component, limit=60),
            "event": _safe_text(event, limit=80),
            "status": _safe_text(status, limit=40),
            "details": _safe_details(details),
        }
        line = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
        with _WRITE_LOCK:
            _rotate(path)
            with path.open("a", encoding="utf-8", newline="\n") as stream:
                stream.write(line + "\n")
                stream.flush()
    except (OSError, TypeError, ValueError):
        return


def record_exception(
    component: str,
    event: str,
    error: BaseException,
    *,
    details: dict[str, Any] | None = None,
    runtime_dir: Path | None = None,
) -> None:
    frames = traceback.extract_tb(error.__traceback__)[-12:]
    frame_labels = [
        f"{Path(frame.filename).name}:{frame.lineno}:{frame.name}"
        for frame in frames
    ]
    fingerprint_source = "|".join(
        [type(error).__name__, *frame_labels]
    ).encode("utf-8", errors="replace")
    safe = dict(details or {})
    safe.update(
        {
            "error_type": type(error).__name__,
            "error_fingerprint": hashlib.sha256(fingerprint_source).hexdigest()[:16],
            "frames": frame_labels,
        }
    )
    record_event(
        component,
        event,
        "error",
        details=safe,
        runtime_dir=runtime_dir,
    )


def _package_versions() -> dict[str, str]:
    result: dict[str, str] = {}
    for package_name in _SAFE_PACKAGE_NAMES:
        try:
            result[package_name] = metadata.version(package_name)
        except metadata.PackageNotFoundError:
            module_name = _PACKAGE_IMPORT_NAMES[package_name]
            try:
                bundled = importlib.util.find_spec(module_name) is not None
            except (ImportError, AttributeError, ValueError):
                bundled = False
            result[package_name] = "bundled" if bundled else "not-installed"
    return result


def system_report(extra: dict[str, Any] | None = None) -> dict[str, Any]:
    return {
        "generated_at": datetime.now(timezone.utc).astimezone().isoformat(
            timespec="seconds"
        ),
        "application_version": APP_VERSION,
        "frozen_executable": bool(getattr(sys, "frozen", False)),
        "operating_system": platform.system(),
        "operating_system_release": platform.release(),
        "operating_system_version": _safe_text(platform.version(), limit=160),
        "architecture": platform.machine(),
        "python_version": platform.python_version(),
        "packages": _package_versions(),
        "checks": _safe_details(extra),
    }


def create_diagnostic_bundle(
    runtime_dir: Path,
    *,
    extra: dict[str, Any] | None = None,
) -> Path:
    """Create an archive containing only safe diagnostics and system metadata."""

    root = runtime_dir.resolve()
    diagnostics_dir = root / "diagnostics"
    exports_dir = diagnostics_dir / "exports"
    exports_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    destination = exports_dir / f"GNS-diagnostics-{timestamp}-{uuid4().hex[:6]}.zip"
    report = json.dumps(
        system_report(extra),
        ensure_ascii=False,
        indent=2,
    )
    readme = (
        "Диагностический архив ГНС.\n"
        "Он не содержит базу SQLite, PDF/Word, тексты писем, ИНН, логины, "
        "пароли, cookies или токены.\n"
    )
    with zipfile.ZipFile(
        destination,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as archive:
        archive.writestr("system-report.json", report)
        archive.writestr("README.txt", readme)
        if diagnostics_dir.exists():
            for candidate in sorted(diagnostics_dir.glob("gns-diagnostics.jsonl*")):
                if candidate.is_file():
                    archive.write(candidate, f"logs/{candidate.name}")
    record_event(
        "diagnostics",
        "bundle_created",
        "success",
        runtime_dir=root,
    )
    return destination
