from __future__ import annotations

import argparse
import hashlib
from pathlib import Path


MAIL_PART_BYTES = 14 * 1024 * 1024


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("archive", type=Path)
    parser.add_argument("output_dir", type=Path)
    arguments = parser.parse_args()
    archive = arguments.archive.resolve()
    output_dir = arguments.output_dir.resolve()
    if not archive.is_file():
        raise SystemExit("Archive was not found")
    output_dir.mkdir(parents=True, exist_ok=True)

    parts: list[Path] = []
    with archive.open("rb") as source:
        index = 1
        while chunk := source.read(MAIL_PART_BYTES):
            part = output_dir / f"{archive.name}.{index:03d}.txt"
            if part.exists():
                raise SystemExit(f"Output already exists: {part}")
            part.write_bytes(chunk)
            parts.append(part)
            index += 1

    joined_digest = hashlib.sha256()
    for part in parts:
        joined_digest.update(part.read_bytes())
    source_digest = sha256_file(archive)
    if joined_digest.hexdigest() != source_digest:
        raise SystemExit("Split verification failed")

    checksums = [f"{sha256_file(part)}  {part.name}" for part in parts]
    (output_dir / "EMAIL_PARTS_SHA256.txt").write_text(
        "\n".join(checksums) + "\n",
        encoding="utf-8",
    )
    (output_dir / "HOW_TO_OPEN.txt").write_text(
        "1. Save every attached .txt file into one folder.\n"
        "2. Remove only the final .txt extension from every file.\n"
        "   Example: .7z.001.txt becomes .7z.001\n"
        "3. Open the .001 file with 7-Zip or WinRAR and extract the folder.\n"
        "4. Do not run the EXE directly from the archive.\n",
        encoding="utf-8",
    )
    print(f"parts={len(parts)} sha256={source_digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
