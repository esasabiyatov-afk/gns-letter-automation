from __future__ import annotations

import os
import zipfile
from pathlib import Path
from urllib.parse import urlparse

import httpx


class OfficialDocumentError(ValueError):
    pass


class OfficialDocumentClient:
    def __init__(
        self,
        allowed_hosts: frozenset[str],
        allowed_paths: frozenset[str],
        max_bytes: int = 30 * 1024 * 1024,
    ):
        self.allowed_hosts = allowed_hosts
        self.allowed_paths = allowed_paths
        self.max_bytes = max_bytes

    def download(self, url: str, destination_dir: Path) -> Path:
        self._validate_url(url)
        destination_dir.mkdir(parents=True, exist_ok=True)
        temporary = destination_dir / "official.download"

        try:
            with httpx.Client(
                follow_redirects=True,
                timeout=httpx.Timeout(25.0, connect=10.0),
                headers={"User-Agent": "GNS-Letter-Automation/0.1"},
            ) as client:
                with client.stream("GET", url) as response:
                    response.raise_for_status()
                    self._validate_url(str(response.url))
                    total = 0
                    with temporary.open("wb") as output:
                        for chunk in response.iter_bytes(1024 * 256):
                            total += len(chunk)
                            if total > self.max_bytes:
                                raise OfficialDocumentError(
                                    "Официальный файл превышает допустимый размер"
                                )
                            output.write(chunk)

            extension = self._detect_extension(temporary)
            final_path = destination_dir / f"official{extension}"
            os.replace(temporary, final_path)
            return final_path
        except Exception:
            temporary.unlink(missing_ok=True)
            raise

    def _validate_url(self, url: str) -> None:
        parsed = urlparse(url)
        host = (parsed.hostname or "").lower()
        if (
            parsed.scheme != "https"
            or host not in self.allowed_hosts
            or parsed.path not in self.allowed_paths
        ):
            raise OfficialDocumentError(
                "Адрес официального документа не входит в разрешённый список"
            )

    @staticmethod
    def _detect_extension(path: Path) -> str:
        with path.open("rb") as stream:
            signature = stream.read(8)
        if signature.startswith(b"%PDF-"):
            return ".pdf"
        if signature.startswith(b"PK"):
            try:
                with zipfile.ZipFile(path) as archive:
                    names = set(archive.namelist())
                if "[Content_Types].xml" in names and "word/document.xml" in names:
                    return ".docx"
            except zipfile.BadZipFile as exc:
                raise OfficialDocumentError(
                    "Официальный DOCX повреждён"
                ) from exc
        raise OfficialDocumentError(
            "Официальный сервер вернул неподдерживаемый тип файла"
        )

