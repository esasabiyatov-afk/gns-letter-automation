from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

from gns_app.config import Settings
from gns_app.database import Database
from gns_app.services.workflow import WorkflowService


@pytest.fixture
def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


@pytest.fixture
def test_settings(tmp_path: Path, project_root: Path) -> Settings:
    runtime = tmp_path / "runtime"
    settings = Settings(
        project_root=project_root,
        runtime_dir=runtime,
        database_path=runtime / "test.sqlite3",
        uploads_dir=runtime / "uploads",
        previews_dir=runtime / "previews",
        official_dir=runtime / "official",
        responses_dir=runtime / "responses",
        source_templates_dir=project_root / "УГНС",
        period_threshold=date(2019, 1, 1),
        max_upload_bytes=150 * 1024 * 1024,
        allowed_qr_hosts=frozenset({"qr.salyk.kg"}),
        allowed_qr_paths=frozenset({"/getsti010decission"}),
        auto_download_official=False,
        employee_name="Гапарова Э.",
    )
    settings.ensure_directories()
    return settings


@pytest.fixture
def workflow(test_settings: Settings) -> WorkflowService:
    database = Database(test_settings.database_path)
    database.initialize()
    return WorkflowService(database, test_settings)

