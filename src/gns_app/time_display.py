from __future__ import annotations

from datetime import UTC, date, datetime, time, timedelta

# Единый формат времени в интерфейсе.
DATE_FORMAT = "%d.%m.%Y"
DATETIME_FORMAT = "%d.%m.%Y %H:%M"
DATETIME_SECONDS_FORMAT = "%d.%m.%Y %H:%M:%S"


def _parse(value: str | datetime) -> datetime | None:
    if isinstance(value, datetime):
        return value
    try:
        return datetime.fromisoformat(str(value).strip().replace("Z", "+00:00"))
    except ValueError:
        return None


def local_datetime(value: str | datetime | None, seconds: bool = False) -> str:
    """Format an ISO timestamp in the computer's local time zone."""
    if value is None or value == "":
        return ""
    parsed = _parse(value)
    if parsed is None:
        return str(value)
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone()
    return parsed.strftime(DATETIME_SECONDS_FORMAT if seconds else DATETIME_FORMAT)


def local_date(value: str | date | datetime | None) -> str:
    """Format a calendar date (or the local date of a timestamp)."""
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        parsed: date = value.astimezone() if value.tzinfo else value
    elif isinstance(value, date):
        parsed = value
    else:
        text = str(value).strip()
        try:
            parsed = date.fromisoformat(text)
        except ValueError:
            timestamp = _parse(text)
            if timestamp is None:
                return text
            parsed = timestamp.astimezone() if timestamp.tzinfo else timestamp
    return parsed.strftime(DATE_FORMAT)


def outlook_time_to_utc(value: datetime) -> str:
    """Convert an Outlook COM timestamp to a stored UTC ISO string.

    pywin32 returns Outlook ReceivedTime/SentOn as the local wall-clock time
    but labels it UTC, so the label is dropped and the computer zone applied.
    """
    local = value.replace(tzinfo=None).astimezone()
    return local.astimezone(UTC).isoformat(timespec="seconds")


def local_day_start_utc(day: str | date) -> str:
    """UTC ISO string of local midnight that starts the given calendar day."""
    parsed = day if isinstance(day, date) else date.fromisoformat(str(day))
    start = datetime.combine(parsed, time()).astimezone()
    return start.astimezone(UTC).isoformat(timespec="seconds")


def local_day_end_utc(day: str | date) -> str:
    """UTC ISO string of local midnight that follows the given calendar day."""
    parsed = day if isinstance(day, date) else date.fromisoformat(str(day))
    return local_day_start_utc(parsed + timedelta(days=1))
