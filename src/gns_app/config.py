from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date
from pathlib import Path


@dataclass(frozen=True, slots=True)
class Settings:
    project_root: Path
    runtime_dir: Path
    database_path: Path
    uploads_dir: Path
    previews_dir: Path
    official_dir: Path
    responses_dir: Path
    source_templates_dir: Path
    period_threshold: date
    max_upload_bytes: int
    allowed_qr_hosts: frozenset[str]
    allowed_qr_paths: frozenset[str]
    auto_download_official: bool
    employee_name: str

    @classmethod
    def load(cls) -> "Settings":
        project_root = Path(__file__).resolve().parents[2]
        runtime_dir = Path(
            os.environ.get("GNS_RUNTIME_DIR", project_root / "runtime")
        ).resolve()
        threshold_text = os.environ.get("GNS_PERIOD_THRESHOLD", "2019-01-01")

        return cls(
            project_root=project_root,
            runtime_dir=runtime_dir,
            database_path=runtime_dir / "gns.sqlite3",
            uploads_dir=runtime_dir / "uploads",
            previews_dir=runtime_dir / "previews",
            official_dir=runtime_dir / "official",
            responses_dir=runtime_dir / "responses",
            source_templates_dir=project_root / "УГНС",
            period_threshold=date.fromisoformat(threshold_text),
            max_upload_bytes=int(
                os.environ.get("GNS_MAX_UPLOAD_BYTES", str(150 * 1024 * 1024))
            ),
            allowed_qr_hosts=frozenset({"qr.salyk.kg"}),
            allowed_qr_paths=frozenset({"/getsti010decission"}),
            auto_download_official=(
                os.environ.get("GNS_AUTO_DOWNLOAD_OFFICIAL", "false").lower()
                == "true"
            ),
            employee_name=os.environ.get("GNS_EMPLOYEE_NAME", "").strip(),
        )

    def ensure_directories(self) -> None:
        for path in (
            self.runtime_dir,
            self.uploads_dir,
            self.previews_dir,
            self.official_dir,
            self.responses_dir,
        ):
            path.mkdir(parents=True, exist_ok=True)


settings = Settings.load()
