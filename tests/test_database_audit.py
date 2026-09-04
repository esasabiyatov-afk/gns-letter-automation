import json
import sqlite3

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


def test_initialize_adds_abs_account_summary_columns_to_legacy_database(tmp_path):
    path = tmp_path / "legacy-abs.sqlite3"
    with sqlite3.connect(path) as connection:
        connection.execute(
            """
            CREATE TABLE taxpayers (
                id TEXT PRIMARY KEY,
                case_id TEXT NOT NULL,
                display_order INTEGER NOT NULL,
                name TEXT NOT NULL,
                inn TEXT NOT NULL,
                name_source TEXT NOT NULL,
                inn_source TEXT NOT NULL,
                manually_confirmed INTEGER NOT NULL DEFAULT 0,
                abs_result TEXT,
                abs_account_result TEXT,
                odb_result TEXT,
                registry_status TEXT,
                registry_name TEXT,
                registry_director TEXT,
                registry_checked_at TEXT,
                registry_provider TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )

    Database(path).initialize()

    with sqlite3.connect(path) as connection:
        columns = {
            row[1] for row in connection.execute("PRAGMA table_info(taxpayers)")
        }
    assert "abs_active_account_count" in columns
    assert "abs_closed_account_count" in columns


def test_initialize_adds_original_sender_to_legacy_outlook_messages(tmp_path):
    path = tmp_path / "legacy-outlook.sqlite3"
    with sqlite3.connect(path) as connection:
        connection.execute(
            """
            CREATE TABLE outlook_messages (
                source_key TEXT PRIMARY KEY,
                sender_smtp TEXT NOT NULL,
                received_at TEXT NOT NULL,
                status TEXT NOT NULL,
                attachment_count INTEGER NOT NULL DEFAULT 0,
                pdf_attachment_count INTEGER NOT NULL DEFAULT 0,
                error_message TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            INSERT INTO outlook_messages(
                source_key, sender_smtp, received_at, status,
                created_at, updated_at
            ) VALUES ('legacy-message', 'reception@bank.kg',
                      '2026-08-28T09:00:00+00:00', 'completed',
                      '2026-08-28T09:00:00+00:00',
                      '2026-08-28T09:00:00+00:00')
            """
        )

    db = Database(path)
    db.initialize()

    message = db.fetch_one(
        "SELECT sender_smtp, original_sender_smtp "
        "FROM outlook_messages WHERE source_key = 'legacy-message'"
    )
    assert message == {
        "sender_smtp": "reception@bank.kg",
        "original_sender_smtp": None,
    }


def test_initialize_marks_legacy_upload_source_as_unknown(tmp_path):
    path = tmp_path / "legacy-upload.sqlite3"
    with sqlite3.connect(path) as connection:
        connection.execute(
            """
            CREATE TABLE uploads (
                id TEXT PRIMARY KEY,
                original_filename TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                sha256 TEXT NOT NULL,
                page_count INTEGER NOT NULL,
                status TEXT NOT NULL,
                issue_message TEXT,
                created_at TEXT NOT NULL,
                completed_at TEXT
            )
            """
        )
        connection.execute(
            """
            INSERT INTO uploads(
                id, original_filename, stored_path, sha256,
                page_count, status, created_at
            ) VALUES ('legacy-upload', 'old.pdf', 'old.pdf', 'hash', 1,
                      'ready', '2026-08-21T09:00:00+00:00')
            """
        )

    db = Database(path)
    db.initialize()

    assert db.fetch_one(
        "SELECT intake_source FROM uploads WHERE id = 'legacy-upload'"
    )["intake_source"] == "legacy"


def test_initialize_cleans_edge_quote_and_audits_change(tmp_path):
    db = Database(tmp_path / "cleanup.sqlite3")
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
            id, upload_id, status, source_kind, district_place,
            fields_confirmed, created_at, updated_at
        ) VALUES ('case-1', 'upload-1', 'ready_for_abs', 'qr_official',
                  'по городу Балыкчы Ысык-Кульской области"', 1,
                  '2026-01-01T00:00:00+00:00',
                  '2026-01-01T00:00:00+00:00')
        """
    )

    db.initialize()

    case = db.fetch_one("SELECT * FROM cases WHERE id = 'case-1'")
    assert case["district_place"] == (
        "по городу Балыкчы Ысык-Кульской области"
    )
    event = db.fetch_one(
        """
        SELECT * FROM audit_events
        WHERE entity_id = 'case-1'
          AND event_type = 'district_place_edge_noise_removed'
        """
    )
    assert event is not None


def test_initialize_backfills_letters_for_legacy_split_response(tmp_path):
    path = tmp_path / "legacy-response.sqlite3"
    with sqlite3.connect(path) as connection:
        connection.executescript(
            """
            CREATE TABLE response_groups (
                id TEXT PRIMARY KEY,
                business_date TEXT NOT NULL,
                group_key TEXT NOT NULL,
                abs_bucket TEXT NOT NULL,
                status TEXT NOT NULL,
                district_place TEXT NOT NULL,
                recipient_position TEXT NOT NULL,
                recipient_full_name TEXT NOT NULL,
                recipient_display_name TEXT NOT NULL,
                employee_name TEXT NOT NULL,
                taxpayer_count INTEGER NOT NULL,
                response_path TEXT,
                response_page_overflow INTEGER,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE audit_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entity_type TEXT NOT NULL,
                entity_id TEXT NOT NULL,
                event_type TEXT NOT NULL,
                actor TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            """
        )
        connection.execute(
            """
            INSERT INTO response_groups(
                id, business_date, group_key, abs_bucket, status,
                district_place, recipient_position, recipient_full_name,
                recipient_display_name, employee_name, taxpayer_count,
                response_path, created_at, updated_at
            ) VALUES ('legacy-group', '2026-08-22', 'key', 'not_found',
                      'created', 'по району', 'Начальник', 'Тестов Тест',
                      'Тестову Т.', 'Сотрудник', 5, 'response.docx',
                      '2026-08-22T01:00:00+00:00',
                      '2026-08-22T01:00:00+00:00')
            """
        )
        connection.execute(
            """
            INSERT INTO audit_events(
                entity_type, entity_id, event_type, actor,
                payload_json, created_at
            ) VALUES ('response_group', 'legacy-group',
                      'grouped_response_created', 'system', ?,
                      '2026-08-22T01:00:00+00:00')
            """,
            (json.dumps({"taxpayers_per_page": 2}),),
        )

    database = Database(path)
    database.initialize()

    group = database.fetch_one(
        "SELECT * FROM response_groups WHERE id = 'legacy-group'"
    )
    letters = database.fetch_all(
        "SELECT * FROM response_letters "
        "WHERE response_group_id = 'legacy-group' ORDER BY letter_order"
    )
    assert group["taxpayers_per_letter"] == 2
    assert group["opened_for_print_at"] is None
    assert [letter["taxpayer_count"] for letter in letters] == [2, 2, 1]
