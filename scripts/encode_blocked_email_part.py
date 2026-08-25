from __future__ import annotations

import argparse
import base64
import hashlib
import shutil
from pathlib import Path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("blocked_part", type=Path)
    parser.add_argument("output_dir", type=Path)
    arguments = parser.parse_args()
    blocked_part = arguments.blocked_part.resolve()
    output_dir = arguments.output_dir.resolve()
    if not blocked_part.is_file():
        raise SystemExit("Blocked archive part was not found")
    output_dir.mkdir(parents=True, exist_ok=True)

    encoded = base64.encodebytes(blocked_part.read_bytes())
    midpoint = len(encoded) // 2
    split_at = encoded.find(b"\n", midpoint)
    if split_at < 0:
        raise SystemExit("Could not split Base64 safely")
    split_at += 1
    part_a = output_dir / "GNS-blocked-001-A.txt"
    part_b = output_dir / "GNS-blocked-001-B.txt"
    part_a.write_bytes(encoded[:split_at])
    part_b.write_bytes(encoded[split_at:])

    restored = base64.decodebytes(part_a.read_bytes() + part_b.read_bytes())
    restored_digest = hashlib.sha256(restored).hexdigest()
    source_digest = sha256_file(blocked_part)
    if restored_digest != source_digest:
        raise SystemExit("Base64 verification failed")

    restored_name = blocked_part.name
    copied_parts: list[str] = []
    stem = restored_name[:-3] if restored_name.endswith("001") else ""
    if stem:
        for index in range(2, 100):
            sibling = blocked_part.with_name(f"{stem}{index:03d}")
            if not sibling.is_file():
                break
            destination = output_dir / f"{sibling.name}.txt"
            shutil.copy2(sibling, destination)
            copied_parts.append(destination.name)
    (output_dir / "RESTORE_001.txt").write_text(
        "Сохраните GNS-blocked-001-A.txt и GNS-blocked-001-B.txt "
        "и все файлы частей 002-005 в одну папку.\n"
        "У файлов 002-005 удалите только последнее расширение .txt.\n\n"
        "Откройте Командную строку в этой папке и выполните:\n\n"
        "copy /b GNS-blocked-001-A.txt+GNS-blocked-001-B.txt "
        "GNS-blocked-001-base64.txt\n"
        "certutil -decode GNS-blocked-001-base64.txt "
        f"{restored_name}\n\n"
        f"Затем откройте {restored_name} через 7-Zip или WinRAR.\n"
        f"Ожидаемый SHA-256 восстановленной 001: {source_digest}\n",
        encoding="utf-8",
    )
    print(
        f"a={part_a.stat().st_size} b={part_b.stat().st_size} "
        f"copied={len(copied_parts)} sha256={source_digest}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
