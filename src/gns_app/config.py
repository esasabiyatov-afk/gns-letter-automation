from __future__ import annotations

import os
import sys
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
    inbox_dir: Path
    source_templates_dir: Path
    period_threshold: date
    max_upload_bytes: int
    allowed_qr_hosts: frozenset[str]
    allowed_qr_paths: frozenset[str]
    auto_download_official: bool
    employee_name: str
    auto_registry_check: bool = True
    ocr_fast_data_dir: Path | None = None
    ocr_best_data_dir: Path | None = None
    processing_workers: int = 2
    abs_mode: str = "fake"
    tolubay_base_url: str = "https://ob.tolubay.kg"
    tolubay_ca_file: Path | None = None
    tolubay_timeout_seconds: float = 30.0

    @classmethod
    def load(cls) -> "Settings":
        project_root = (
            Path(sys.executable).resolve().parent
            if getattr(sys, "frozen", False)
            else Path(__file__).resolve().parents[2]
        )
        bundle_root = Path(getattr(sys, "_MEIPASS", project_root)).resolve()
        runtime_dir = Path(
            os.environ.get("GNS_RUNTIME_DIR", project_root / "runtime")
        ).resolve()
        local_app_data = Path(os.environ.get("LOCALAPPDATA", runtime_dir)).resolve()
        packaged_model_root = bundle_root / "models"
        default_model_root = (
            packaged_model_root
            if getattr(sys, "frozen", False) and packaged_model_root.is_dir()
            else local_app_data / "GNSLetterAutomation" / "models"
        )
        ocr_model_root = Path(
            os.environ.get(
                "GNS_OCR_MODEL_DIR",
                default_model_root,
            )
        ).resolve()
        threshold_text = os.environ.get("GNS_PERIOD_THRESHOLD", "2019-01-01")
        abs_mode = os.environ.get("GNS_ABS_MODE", "fake").strip().casefold()
        if abs_mode not in {"fake", "tolubay"}:
            raise ValueError("GNS_ABS_MODE должен быть fake или tolubay")
        tolubay_ca_text = os.environ.get("GNS_TOLUBAY_CA_FILE", "").strip()

        return cls(
            project_root=project_root,
            runtime_dir=runtime_dir,
            database_path=runtime_dir / "gns.sqlite3",
            uploads_dir=runtime_dir / "uploads",
            previews_dir=runtime_dir / "previews",
            official_dir=runtime_dir / "official",
            responses_dir=runtime_dir / "responses",
            inbox_dir=Path(
                os.environ.get(
                    "GNS_INBOX_DIR",
                    project_root / "Входящие",
                )
            ).resolve(),
            source_templates_dir=(
                bundle_root / "УГНС"
                if getattr(sys, "frozen", False)
                else project_root / "УГНС"
            ),
            period_threshold=date.fromisoformat(threshold_text),
            max_upload_bytes=int(
                os.environ.get("GNS_MAX_UPLOAD_BYTES", str(150 * 1024 * 1024))
            ),
            allowed_qr_hosts=frozenset({"qr.salyk.kg"}),
            allowed_qr_paths=frozenset({"/getsti010decission"}),
            auto_download_official=(
                os.environ.get("GNS_AUTO_DOWNLOAD_OFFICIAL", "true").lower()
                == "true"
            ),
            employee_name=os.environ.get("GNS_EMPLOYEE_NAME", "").strip(),
            auto_registry_check=(
                os.environ.get("GNS_AUTO_REGISTRY_CHECK", "true").lower()
                == "true"
            ),
            ocr_fast_data_dir=ocr_model_root / "tessdata_fast",
            ocr_best_data_dir=ocr_model_root / "tessdata_best",
            processing_workers=max(
                1,
                min(2, int(os.environ.get("GNS_PROCESSING_WORKERS", "2"))),
            ),
            abs_mode=abs_mode,
            tolubay_base_url=os.environ.get(
                "GNS_TOLUBAY_BASE_URL",
                "https://ob.tolubay.kg",
            ).strip(),
            tolubay_ca_file=(
                Path(tolubay_ca_text).resolve()
                if tolubay_ca_text
                else None
            ),
            tolubay_timeout_seconds=max(
                5.0,
                min(
                    120.0,
                    float(os.environ.get("GNS_TOLUBAY_TIMEOUT_SECONDS", "30")),
                ),
            ),
        )

    def ensure_directories(self) -> None:
        for path in (
            self.runtime_dir,
            self.uploads_dir,
            self.previews_dir,
            self.official_dir,
            self.responses_dir,
            self.inbox_dir,
        ):
            path.mkdir(parents=True, exist_ok=True)


settings = Settings.load()
