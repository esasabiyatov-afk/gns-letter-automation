from __future__ import annotations

import re


# Эти знаки могут прилипнуть к реквизиту, когда регулярное выражение берёт
# текст из кавычек вокруг полного наименования учреждения. Внутренние кавычки
# не удаляются: они законны, например, в наименовании ОсОО "Чардж".
LOCATION_EDGE_NOISE = "\"'«»„“”‹›,;:|"


def clean_location(value: str | None) -> str:
    if not value:
        return ""
    normalized = re.sub(r"\s+", " ", value).strip()
    return normalized.strip(LOCATION_EDGE_NOISE).strip()
