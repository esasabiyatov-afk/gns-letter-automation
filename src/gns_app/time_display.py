from __future__ import annotations

from datetime import datetime


def local_datetime(value: str | datetime | None, seconds: bool = False) -> str:
    """Format an ISO timestamp in the computer's local time zone."""
    if value is None:
        return ""
    source = value if isinstance(value, datetime) else str(value).strip()
    try:
        parsed = (
            source
            if isinstance(source, datetime)
            else datetime.fromisoformat(source.replace("Z", "+00:00"))
        )
    except ValueError:
        text = str(source).replace("T", " ")
        return text[:19 if seconds else 16]
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone()
    return parsed.strftime("%Y-%m-%d %H:%M:%S" if seconds else "%Y-%m-%d %H:%M")
