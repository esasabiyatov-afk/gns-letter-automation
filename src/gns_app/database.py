from __future__ import annotations

import json
import sqlite3
from contextlib import contextmanager
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Iterator, Sequence
from uuid import uuid4

from gns_app.text_cleanup import clean_location, clean_taxpayer_name


def utc_now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


SCHEMA = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS uploads (
    id TEXT PRIMARY KEY,
    original_filename TEXT NOT NULL,
    stored_path TEXT NOT NULL,
    sha256 TEXT NOT NULL,
    page_count INTEGER NOT NULL,
    status TEXT NOT NULL,
    issue_message TEXT,
    created_at TEXT NOT NULL,
    completed_at TEXT
);

CREATE INDEX IF NOT EXISTS idx_uploads_created_at
ON uploads(created_at DESC);

CREATE TABLE IF NOT EXISTS cases (
    id TEXT PRIMARY KEY,
    upload_id TEXT NOT NULL REFERENCES uploads(id) ON DELETE CASCADE,
    status TEXT NOT NULL,
    source_kind TEXT,
    qr_payload_hash TEXT,
    official_document_path TEXT,
    district_place TEXT,
    recipient_position TEXT,
    recipient_full_name TEXT,
    recipient_display_name TEXT,
    period_start TEXT,
    period_end TEXT,
    period_route TEXT,
    employee_name TEXT,
    fields_confirmed INTEGER NOT NULL DEFAULT 0,
    abs_status TEXT NOT NULL DEFAULT 'not_checked',
    response_status TEXT,
    official_parse_version INTEGER NOT NULL DEFAULT 0,
    response_path TEXT,
    response_page_overflow INTEGER,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_cases_upload ON cases(upload_id);
CREATE INDEX IF NOT EXISTS idx_cases_qr_hash ON cases(upload_id, qr_payload_hash);
CREATE INDEX IF NOT EXISTS idx_cases_status ON cases(status);

CREATE TABLE IF NOT EXISTS pages (
    id TEXT PRIMARY KEY,
    upload_id TEXT NOT NULL REFERENCES uploads(id) ON DELETE CASCADE,
    case_id TEXT REFERENCES cases(id) ON DELETE SET NULL,
    page_number INTEGER NOT NULL,
    preview_path TEXT,
    enhanced_preview_path TEXT,
    page_type TEXT NOT NULL DEFAULT 'unknown',
    type_confidence REAL NOT NULL DEFAULT 0,
    quality_score REAL,
    qr_status TEXT NOT NULL DEFAULT 'not_started',
    qr_payload_hash TEXT,
    qr_safe_url TEXT,
    qr_method TEXT,
    ocr_status TEXT NOT NULL DEFAULT 'not_started',
    ocr_confidence REAL NOT NULL DEFAULT 0,
    extracted_text TEXT,
    type_evidence_text TEXT,
    type_evidence_method TEXT,
    status TEXT NOT NULL,
    issue_code TEXT,
    issue_message TEXT,
    manual_confirmed INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(upload_id, page_number)
);

CREATE INDEX IF NOT EXISTS idx_pages_upload ON pages(upload_id, page_number);
CREATE INDEX IF NOT EXISTS idx_pages_status ON pages(status);
CREATE INDEX IF NOT EXISTS idx_pages_case ON pages(case_id);

CREATE TABLE IF NOT EXISTS taxpayers (
    id TEXT PRIMARY KEY,
    case_id TEXT NOT NULL REFERENCES cases(id) ON DELETE CASCADE,
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
);

CREATE INDEX IF NOT EXISTS idx_taxpayers_case
ON taxpayers(case_id, display_order);

CREATE TABLE IF NOT EXISTS audit_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entity_type TEXT NOT NULL,
    entity_id TEXT NOT NULL,
    event_type TEXT NOT NULL,
    actor TEXT NOT NULL,
    payload_json TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_audit_entity
ON audit_events(entity_type, entity_id, created_at);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS outlook_messages (
    source_key TEXT PRIMARY KEY,
    sender_smtp TEXT NOT NULL,
    received_at TEXT NOT NULL,
    status TEXT NOT NULL,
    attachment_count INTEGER NOT NULL DEFAULT 0,
    pdf_attachment_count INTEGER NOT NULL DEFAULT 0,
    error_message TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_outlook_messages_received
ON outlook_messages(received_at DESC);

CREATE INDEX IF NOT EXISTS idx_outlook_messages_status
ON outlook_messages(status, updated_at);

CREATE TABLE IF NOT EXISTS outlook_attachments (
    id TEXT PRIMARY KEY,
    message_key TEXT NOT NULL
        REFERENCES outlook_messages(source_key) ON DELETE CASCADE,
    attachment_index INTEGER NOT NULL,
    original_filename TEXT NOT NULL,
    stored_path TEXT,
    sha256 TEXT,
    size_bytes INTEGER,
    status TEXT NOT NULL,
    upload_id TEXT REFERENCES uploads(id) ON DELETE SET NULL,
    error_message TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(message_key, attachment_index)
);

CREATE INDEX IF NOT EXISTS idx_outlook_attachments_message
ON outlook_attachments(message_key, attachment_index);

CREATE TABLE IF NOT EXISTS employee_profiles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    name_key TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_employee_profiles_name
ON employee_profiles(name);

CREATE TABLE IF NOT EXISTS gns_offices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    office_key TEXT NOT NULL UNIQUE,
    office_name TEXT NOT NULL,
    district_place TEXT NOT NULL,
    postal_address TEXT,
    email_address TEXT,
    aliases_json TEXT NOT NULL DEFAULT '[]',
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_gns_offices_district
ON gns_offices(district_place);

CREATE TABLE IF NOT EXISTS response_groups (
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
    taxpayers_per_letter INTEGER NOT NULL DEFAULT 1,
    response_path TEXT,
    response_page_overflow INTEGER,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_response_groups_day
ON response_groups(business_date, created_at DESC);

CREATE TABLE IF NOT EXISTS response_group_cases (
    response_group_id TEXT NOT NULL
        REFERENCES response_groups(id) ON DELETE CASCADE,
    case_id TEXT NOT NULL REFERENCES cases(id) ON DELETE RESTRICT,
    PRIMARY KEY(response_group_id, case_id)
);

CREATE TABLE IF NOT EXISTS response_group_taxpayers (
    response_group_id TEXT NOT NULL
        REFERENCES response_groups(id) ON DELETE CASCADE,
    display_order INTEGER NOT NULL,
    source_case_id TEXT NOT NULL REFERENCES cases(id) ON DELETE RESTRICT,
    name TEXT NOT NULL,
    inn TEXT NOT NULL,
    PRIMARY KEY(response_group_id, display_order)
);

CREATE TABLE IF NOT EXISTS response_letters (
    id TEXT PRIMARY KEY,
    response_group_id TEXT NOT NULL
        REFERENCES response_groups(id) ON DELETE CASCADE,
    letter_order INTEGER NOT NULL,
    taxpayer_start_order INTEGER NOT NULL,
    taxpayer_count INTEGER NOT NULL,
    outgoing_number TEXT,
    assigned_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(response_group_id, letter_order)
);

CREATE INDEX IF NOT EXISTS idx_response_letters_group
ON response_letters(response_group_id, letter_order);

CREATE UNIQUE INDEX IF NOT EXISTS idx_response_letters_outgoing_number
ON response_letters(outgoing_number)
WHERE outgoing_number IS NOT NULL AND outgoing_number != '';

CREATE TABLE IF NOT EXISTS signed_response_scans (
    id TEXT PRIMARY KEY,
    response_letter_id TEXT NOT NULL
        REFERENCES response_letters(id) ON DELETE CASCADE,
    status TEXT NOT NULL,
    source TEXT NOT NULL,
    original_filename TEXT NOT NULL,
    original_path TEXT NOT NULL,
    pdf_path TEXT NOT NULL,
    sha256 TEXT NOT NULL,
    size_bytes INTEGER NOT NULL,
    page_count INTEGER NOT NULL DEFAULT 1,
    correct_letter_confirmed INTEGER NOT NULL DEFAULT 0,
    signature_confirmed INTEGER NOT NULL DEFAULT 0,
    bank_seal_confirmed INTEGER NOT NULL DEFAULT 0,
    confirmed_by TEXT,
    confirmed_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_signed_response_scans_letter
ON signed_response_scans(response_letter_id, created_at DESC);

CREATE UNIQUE INDEX IF NOT EXISTS idx_signed_response_scans_active
ON signed_response_scans(response_letter_id)
WHERE status IN ('needs_confirmation', 'confirmed');

CREATE TABLE IF NOT EXISTS outlook_outgoing_messages (
    id TEXT PRIMARY KEY,
    response_letter_id TEXT NOT NULL
        REFERENCES response_letters(id) ON DELETE CASCADE,
    signed_scan_id TEXT NOT NULL
        REFERENCES signed_response_scans(id) ON DELETE RESTRICT,
    status TEXT NOT NULL,
    recipient_email TEXT NOT NULL,
    subject TEXT NOT NULL,
    attachment_name TEXT NOT NULL,
    attachment_sha256 TEXT NOT NULL,
    outlook_entry_id TEXT,
    error_message TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    UNIQUE(signed_scan_id)
);

CREATE INDEX IF NOT EXISTS idx_outlook_outgoing_letter
ON outlook_outgoing_messages(response_letter_id, created_at DESC);
"""


class Database:
    def __init__(self, path: Path):
        self.path = path

    @contextmanager
    def connect(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.path, timeout=30)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute("PRAGMA journal_mode = WAL")
        try:
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def initialize(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connect() as connection:
            connection.executescript(SCHEMA)
            # Миграция прототипа: ранняя версия помечала сам факт наличия QR
            # как официальный источник, даже если документ по ссылке не скачан.
            connection.execute(
                """
                UPDATE cases
                SET source_kind = 'qr_link'
                WHERE source_kind = 'qr_official'
                  AND official_document_path IS NULL
                  AND fields_confirmed = 0
                """
            )
            connection.execute(
                """
                UPDATE cases
                SET source_kind = 'manual'
                WHERE official_document_path IS NULL
                  AND fields_confirmed = 1
                """
            )
            self._ensure_taxpayer_columns(connection)
            self._ensure_page_columns(connection)
            self._ensure_case_columns(connection)
            self._ensure_response_group_columns(connection)
            self._ensure_gns_office_columns(connection)
            self._clean_legacy_district_places(connection)
            self._normalize_legacy_taxpayer_names(connection)

    @staticmethod
    def _ensure_page_columns(connection: sqlite3.Connection) -> None:
        existing = {
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(pages)"
            ).fetchall()
        }
        additions = {
            "type_evidence_text": "TEXT",
            "type_evidence_method": "TEXT",
        }
        for name, column_type in additions.items():
            if name not in existing:
                connection.execute(
                    f"ALTER TABLE pages ADD COLUMN {name} {column_type}"
                )

    @staticmethod
    def _ensure_case_columns(connection: sqlite3.Connection) -> None:
        existing = {
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(cases)"
            ).fetchall()
        }
        if "official_parse_version" not in existing:
            connection.execute(
                "ALTER TABLE cases ADD COLUMN "
                "official_parse_version INTEGER NOT NULL DEFAULT 0"
            )
        if "period_route" not in existing:
            connection.execute(
                "ALTER TABLE cases ADD COLUMN period_route TEXT"
            )
        if "response_page_overflow" not in existing:
            connection.execute(
                "ALTER TABLE cases ADD COLUMN "
                "response_page_overflow INTEGER"
            )

    @staticmethod
    def _ensure_taxpayer_columns(connection: sqlite3.Connection) -> None:
        existing = {
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(taxpayers)"
            ).fetchall()
        }
        additions = {
            "odb_result": "TEXT",
            "abs_account_result": "TEXT",
            "registry_status": "TEXT",
            "registry_name": "TEXT",
            "registry_director": "TEXT",
            "registry_checked_at": "TEXT",
            "registry_provider": "TEXT",
        }
        for name, column_type in additions.items():
            if name not in existing:
                connection.execute(
                    f"ALTER TABLE taxpayers ADD COLUMN {name} {column_type}"
                )

    @staticmethod
    def _ensure_gns_office_columns(connection: sqlite3.Connection) -> None:
        existing = {
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(gns_offices)"
            ).fetchall()
        }
        if "email_address" not in existing:
            connection.execute(
                "ALTER TABLE gns_offices ADD COLUMN email_address TEXT"
            )

    @staticmethod
    def _ensure_response_group_columns(connection: sqlite3.Connection) -> None:
        existing = {
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(response_groups)"
            ).fetchall()
        }
        if "response_page_overflow" not in existing:
            connection.execute(
                "ALTER TABLE response_groups ADD COLUMN "
                "response_page_overflow INTEGER"
            )
        if "taxpayers_per_letter" not in existing:
            connection.execute(
                "ALTER TABLE response_groups ADD COLUMN "
                "taxpayers_per_letter INTEGER"
            )
        groups = connection.execute(
            """
            SELECT id, taxpayer_count, taxpayers_per_letter, created_at
            FROM response_groups
            ORDER BY created_at, id
            """
        ).fetchall()
        for group in groups:
            per_letter = int(group["taxpayers_per_letter"] or 0)
            if per_letter <= 0:
                per_letter = int(group["taxpayer_count"] or 1)
                event = connection.execute(
                    """
                    SELECT payload_json
                    FROM audit_events
                    WHERE entity_type = 'response_group'
                      AND entity_id = ?
                      AND event_type = 'grouped_response_created'
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (group["id"],),
                ).fetchone()
                if event:
                    try:
                        payload = json.loads(event["payload_json"])
                        candidate = int(payload.get("taxpayers_per_page") or 0)
                        if candidate > 0:
                            per_letter = candidate
                    except (TypeError, ValueError, json.JSONDecodeError):
                        pass
                connection.execute(
                    "UPDATE response_groups SET taxpayers_per_letter = ? "
                    "WHERE id = ?",
                    (per_letter, group["id"]),
                )
            existing_letters = connection.execute(
                "SELECT COUNT(*) AS count FROM response_letters "
                "WHERE response_group_id = ?",
                (group["id"],),
            ).fetchone()["count"]
            if existing_letters:
                continue
            taxpayer_count = max(1, int(group["taxpayer_count"] or 1))
            letter_order = 1
            start_order = 1
            while start_order <= taxpayer_count:
                count = min(per_letter, taxpayer_count - start_order + 1)
                connection.execute(
                    """
                    INSERT INTO response_letters(
                        id, response_group_id, letter_order,
                        taxpayer_start_order, taxpayer_count,
                        created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        uuid4().hex,
                        group["id"],
                        letter_order,
                        start_order,
                        count,
                        group["created_at"],
                        group["created_at"],
                    ),
                )
                letter_order += 1
                start_order += count

    @staticmethod
    def _clean_legacy_district_places(
        connection: sqlite3.Connection,
    ) -> None:
        rows = connection.execute(
            """
            SELECT id, district_place
            FROM cases
            WHERE district_place IS NOT NULL
            """
        ).fetchall()
        for row in rows:
            original = row["district_place"]
            cleaned = clean_location(original)
            if cleaned == original:
                continue
            now = utc_now()
            connection.execute(
                """
                UPDATE cases
                SET district_place = ?, updated_at = ?
                WHERE id = ?
                """,
                (cleaned, now, row["id"]),
            )
            connection.execute(
                """
                INSERT INTO audit_events(
                    entity_type, entity_id, event_type,
                    actor, payload_json, created_at
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    "case",
                    row["id"],
                    "district_place_edge_noise_removed",
                    "system",
                    json.dumps(
                        {"before": original, "after": cleaned},
                        ensure_ascii=False,
                    ),
                    now,
                ),
            )

    @staticmethod
    def _normalize_legacy_taxpayer_names(
        connection: sqlite3.Connection,
    ) -> None:
        for column in ("name", "registry_name"):
            rows = connection.execute(
                f"SELECT id, {column} FROM taxpayers "
                f"WHERE {column} IS NOT NULL AND {column} != ''"
            ).fetchall()
            for row in rows:
                cleaned = clean_taxpayer_name(row[column])
                if cleaned != row[column]:
                    connection.execute(
                        f"UPDATE taxpayers SET {column} = ?, updated_at = ? "
                        "WHERE id = ?",
                        (cleaned, utc_now(), row["id"]),
                    )

    def execute(self, sql: str, parameters: Sequence[Any] = ()) -> None:
        with self.connect() as connection:
            connection.execute(sql, parameters)

    def executemany(
        self, sql: str, parameters: Sequence[Sequence[Any]]
    ) -> None:
        with self.connect() as connection:
            connection.executemany(sql, parameters)

    def fetch_one(
        self, sql: str, parameters: Sequence[Any] = ()
    ) -> dict[str, Any] | None:
        with self.connect() as connection:
            row = connection.execute(sql, parameters).fetchone()
        return dict(row) if row else None

    def fetch_all(
        self, sql: str, parameters: Sequence[Any] = ()
    ) -> list[dict[str, Any]]:
        with self.connect() as connection:
            rows = connection.execute(sql, parameters).fetchall()
        return [dict(row) for row in rows]

    def audit(
        self,
        entity_type: str,
        entity_id: str,
        event_type: str,
        payload: dict[str, Any] | None = None,
        actor: str = "system",
    ) -> None:
        safe_payload = payload or {}
        forbidden = {"password", "login", "username", "credential", "token"}
        if any(key.lower() in forbidden for key in safe_payload):
            raise ValueError("Учётные данные запрещено записывать в журнал")

        self.execute(
            """
            INSERT INTO audit_events(
                entity_type, entity_id, event_type, actor, payload_json, created_at
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                entity_type,
                entity_id,
                event_type,
                actor,
                json.dumps(safe_payload, ensure_ascii=False),
                utc_now(),
            ),
        )
