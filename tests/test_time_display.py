from __future__ import annotations

from datetime import UTC, datetime, timezone

from gns_app.time_display import local_datetime


def test_utc_timestamp_is_shown_in_computer_local_time():
    source = "2026-09-13T01:15:00+00:00"
    expected = datetime.fromisoformat(source).astimezone().strftime(
        "%Y-%m-%d %H:%M"
    )

    assert local_datetime(source) == expected


def test_outlook_local_timestamp_is_not_shifted_again():
    local_zone = datetime.now(UTC).astimezone().tzinfo or timezone.utc
    source = datetime(2026, 9, 13, 7, 15, tzinfo=local_zone).isoformat()

    assert local_datetime(source) == "2026-09-13 07:15"


def test_naive_outlook_timestamp_remains_local():
    assert local_datetime("2026-09-13T07:15:00") == "2026-09-13 07:15"
