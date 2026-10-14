from __future__ import annotations

import json
from datetime import UTC, date, datetime, timedelta, timezone

from gns_app.database import Database
from gns_app.time_display import (
    local_date,
    local_datetime,
    local_day_end_utc,
    local_day_start_utc,
    outlook_time_to_utc,
)


def _local_zone():
    return datetime.now(UTC).astimezone().tzinfo or timezone.utc


def test_utc_timestamp_is_shown_in_computer_local_time():
    source = "2026-09-13T01:15:00+00:00"
    expected = datetime.fromisoformat(source).astimezone().strftime(
        "%d.%m.%Y %H:%M"
    )

    assert local_datetime(source) == expected


def test_outlook_local_timestamp_is_not_shifted_again():
    source = datetime(2026, 9, 13, 7, 15, tzinfo=_local_zone()).isoformat()

    assert local_datetime(source) == "13.09.2026 07:15"


def test_naive_outlook_timestamp_remains_local():
    assert local_datetime("2026-09-13T07:15:00") == "13.09.2026 07:15"


def test_seconds_and_unparseable_values():
    assert local_datetime("2026-09-13T07:15:42", seconds=True) == (
        "13.09.2026 07:15:42"
    )
    assert local_datetime("") == ""
    assert local_datetime(None) == ""
    assert local_datetime("не дата") == "не дата"


def test_calendar_date_uses_single_format():
    assert local_date("2026-01-05") == "05.01.2026"
    assert local_date(date(2026, 1, 5)) == "05.01.2026"
    assert local_date(None) == ""


def test_outlook_com_time_labelled_utc_is_treated_as_local_wall_clock():
    # pywin32 returns Outlook's local wall-clock time with a UTC label.
    com_value = datetime(2026, 9, 13, 9, 30, tzinfo=UTC)

    stored = outlook_time_to_utc(com_value)

    assert local_datetime(stored) == "13.09.2026 09:30"
    expected_utc = datetime(2026, 9, 13, 9, 30).astimezone().astimezone(UTC)
    assert datetime.fromisoformat(stored) == expected_utc


def test_local_day_bounds_cover_exactly_one_local_day():
    start = datetime.fromisoformat(local_day_start_utc("2026-09-13"))
    end = datetime.fromisoformat(local_day_end_utc("2026-09-13"))

    assert start.astimezone().strftime("%d.%m.%Y %H:%M") == "13.09.2026 00:00"
    assert end - start == timedelta(days=1)


def _seed_legacy_outlook_rows(path):
    database = Database(path)
    database.initialize()
    with database.connect() as connection:
        # Simulate a database created before the fix.
        connection.execute(
            "DELETE FROM settings WHERE key = ?",
            (Database.OUTLOOK_TIME_FIX_SETTING,),
        )
        connection.execute(
            "INSERT INTO outlook_messages(source_key, sender_smtp, received_at, "
            "status, attachment_count, pdf_attachment_count, created_at, "
            "updated_at) VALUES ('m1', 'a@b.kg', '2026-09-13T09:30:00+00:00', "
            "'completed', 1, 1, '2026-09-13T03:31:00+00:00', "
            "'2026-09-13T03:31:00+00:00')"
        )
    return database


def test_legacy_outlook_timestamps_are_fixed_once(tmp_path):
    database = _seed_legacy_outlook_rows(tmp_path / "gns.sqlite3")
    legacy_sent = "2026-09-13T10:00:00+00:00"
    manual_sent = "2026-09-13T04:00:00+00:00"
    with database.connect() as connection:
        connection.execute("PRAGMA foreign_keys = OFF")
        required = [
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(outlook_outgoing_messages)"
            )
            if row["notnull"] and row["dflt_value"] is None and not row["pk"]
        ]
        for row_id, sent_at in (
            ("out-outlook", legacy_sent),
            ("out-manual", manual_sent),
        ):
            values = {name: row_id for name in required}
            values.update({"id": row_id, "status": "sent", "sent_at": sent_at})
            connection.execute(
                f"INSERT INTO outlook_outgoing_messages({', '.join(values)}) "
                f"VALUES ({', '.join('?' for _ in values)})",
                tuple(values.values()),
            )
        connection.execute(
            "INSERT INTO audit_events(entity_type, entity_id, event_type, actor, "
            "payload_json, created_at) VALUES ('outlook_outgoing_message', "
            "'out-outlook', 'outlook_sent_confirmed', 'system', ?, ?)",
            (json.dumps({"sent_at": legacy_sent}), "2026-09-13T04:01:00+00:00"),
        )

    database.initialize()
    database.initialize()

    received = database.fetch_one(
        "SELECT received_at FROM outlook_messages WHERE source_key = 'm1'"
    )["received_at"]
    sent = {
        row["id"]: row["sent_at"]
        for row in database.fetch_all(
            "SELECT id, sent_at FROM outlook_outgoing_messages"
        )
    }
    assert local_datetime(received) == "13.09.2026 09:30"
    assert local_datetime(sent["out-outlook"]) == "13.09.2026 10:00"
    # Real UTC values written by the app itself stay untouched.
    assert sent["out-manual"] == manual_sent
    events = database.fetch_all(
        "SELECT payload_json FROM audit_events "
        "WHERE event_type = 'outlook_local_timestamps_fixed'"
    )
    assert len(events) == 1
    assert json.loads(events[0]["payload_json"]) == {
        "received_at": 1,
        "sent_at": 1,
        "resent_at": 0,
    }


def test_new_database_is_marked_without_rewriting(tmp_path):
    database = Database(tmp_path / "gns.sqlite3")
    database.initialize()

    assert database.fetch_one(
        "SELECT 1 AS ok FROM settings WHERE key = ?",
        (Database.OUTLOOK_TIME_FIX_SETTING,),
    )
    assert not database.fetch_all(
        "SELECT 1 FROM audit_events "
        "WHERE event_type = 'outlook_local_timestamps_fixed'"
    )
