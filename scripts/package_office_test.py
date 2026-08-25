from __future__ import annotations

import argparse
import hashlib
import json
import platform
from datetime import datetime
from pathlib import Path

import py7zr


PART_BYTES = 20 * 1024 * 1024


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def write_manifest(release_dir: Path) -> Path:
    manifest_path = release_dir / "release-manifest.json"
    files = []
    for path in sorted(release_dir.rglob("*")):
        if path.is_file() and path != manifest_path:
            files.append(
                {
                    "path": path.relative_to(release_dir).as_posix(),
                    "bytes": path.stat().st_size,
                    "sha256": sha256_file(path),
                }
            )
    payload = {
        "build": "GNS-Test-Win8.1",
        "application_version": "0.1.1-office-test",
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "builder_os": platform.platform(),
        "files": files,
    }
    manifest_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return manifest_path


def split_archive(archive_path: Path) -> list[Path]:
    parts: list[Path] = []
    with archive_path.open("rb") as source:
        index = 1
        while chunk := source.read(PART_BYTES):
            part = archive_path.with_name(f"{archive_path.name}.{index:03d}")
            part.write_bytes(chunk)
            parts.append(part)
            index += 1
    return parts


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("release_dir", type=Path)
    parser.add_argument("output_dir", type=Path)
    arguments = parser.parse_args()
    release_dir = arguments.release_dir.resolve()
    output_dir = arguments.output_dir.resolve()
    if not (release_dir / "GNS-Test-Win8.1.exe").is_file():
        raise SystemExit("Test EXE was not found")
    output_dir.mkdir(parents=True, exist_ok=True)
    write_manifest(release_dir)

    archive_path = output_dir / "GNS-Test-Win8.1-office-0.1.1.7z"
    archive_path.unlink(missing_ok=True)
    for old_part in output_dir.glob(f"{archive_path.name}.*"):
        old_part.unlink()
    filters = [{"id": py7zr.FILTER_LZMA2, "preset": 9}]
    with py7zr.SevenZipFile(archive_path, "w", filters=filters) as archive:
        archive.writeall(release_dir, arcname=release_dir.name)
    with py7zr.SevenZipFile(archive_path, "r") as archive:
        if archive.test() is not None:
            raise SystemExit("7z verification failed")

    parts = split_archive(archive_path) if archive_path.stat().st_size > PART_BYTES else []
    checksums = [f"{sha256_file(archive_path)}  {archive_path.name}"]
    checksums.extend(f"{sha256_file(part)}  {part.name}" for part in parts)
    (output_dir / "SHA256SUMS.txt").write_text(
        "\n".join(checksums) + "\n",
        encoding="utf-8",
    )
    print(
        json.dumps(
            {
                "archive": str(archive_path),
                "bytes": archive_path.stat().st_size,
                "parts": [str(part) for part in parts],
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
