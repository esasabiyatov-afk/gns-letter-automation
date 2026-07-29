import pytest

from gns_app.database import Database


def test_audit_rejects_credentials(tmp_path):
    db = Database(tmp_path / "audit.sqlite3")
    db.initialize()
    with pytest.raises(ValueError):
        db.audit(
            "case",
            "case-1",
            "bad_event",
            {"password": "secret"},
        )


def test_audit_accepts_safe_abs_result(tmp_path):
    db = Database(tmp_path / "audit.sqlite3")
    db.initialize()
    db.audit(
        "case",
        "case-1",
        "abs_checked",
        {"status": "not_found", "is_fake": True},
    )
    events = db.fetch_all("SELECT * FROM audit_events")
    assert len(events) == 1
    assert "secret" not in events[0]["payload_json"]


def test_initialize_corrects_legacy_qr_source(tmp_path):
    db = Database(tmp_path / "migration.sqlite3")
    db.initialize()
    db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('upload-1', 'sample.pdf', 'sample.pdf', 'hash', 1,
                  'ready', '2026-01-01T00:00:00+00:00')
        """
    )
    db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind, fields_confirmed,
            created_at, updated_at
        ) VALUES ('case-1', 'upload-1', 'ready_for_abs', 'qr_official', 1,
                  '2026-01-01T00:00:00+00:00',
                  '2026-01-01T00:00:00+00:00')
        """
    )

    db.initialize()

    assert db.fetch_one(
        "SELECT source_kind FROM cases WHERE id = 'case-1'"
    )["source_kind"] == "manual"
