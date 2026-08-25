from __future__ import annotations

import json
import zipfile

from gns_app.diagnostics import create_diagnostic_bundle, record_event


def test_diagnostic_log_redacts_sensitive_values(tmp_path):
    runtime = tmp_path / "runtime"
    record_event(
        "abs",
        "check",
        "error",
        details={
            "username": "employee",
            "password": "secret",
            "unexpected": "user@example.com 12345678901234 password=hidden",
            "record_count": 2,
        },
        runtime_dir=runtime,
    )

    payload = json.loads(
        (runtime / "diagnostics" / "gns-diagnostics.jsonl")
        .read_text(encoding="utf-8")
        .strip()
    )
    serialized = json.dumps(payload, ensure_ascii=False)
    assert "employee" not in serialized
    assert "secret" not in serialized
    assert "user@example.com" not in serialized
    assert "12345678901234" not in serialized
    assert "hidden" not in serialized
    assert payload["details"]["record_count"] == 2


def test_diagnostic_bundle_excludes_working_data(tmp_path):
    runtime = tmp_path / "runtime"
    (runtime / "uploads").mkdir(parents=True)
    (runtime / "uploads" / "real.pdf").write_bytes(b"sensitive")
    (runtime / "gns.sqlite3").write_bytes(b"database")
    record_event("application", "startup", "ready", runtime_dir=runtime)

    bundle = create_diagnostic_bundle(runtime, extra={"abs_mode": "tolubay"})

    with zipfile.ZipFile(bundle) as archive:
        names = set(archive.namelist())
        assert "system-report.json" in names
        assert "README.txt" in names
        assert "logs/gns-diagnostics.jsonl" in names
        assert not any(name.endswith(".pdf") for name in names)
        assert not any(name.endswith(".sqlite3") for name in names)
