from __future__ import annotations

import re
import unicodedata


# Эти знаки могут прилипнуть к реквизиту, когда регулярное выражение берёт
# текст из кавычек вокруг полного наименования учреждения. Внутренние кавычки
# не удаляются: они законны, например, в наименовании ОсОО "Чардж".
LOCATION_EDGE_NOISE = "\"'«»„“”‹›,;:|"


def clean_location(value: str | None) -> str:
    if not value:
        return ""
    normalized = re.sub(r"\s+", " ", value).strip()
    return normalized.strip(LOCATION_EDGE_NOISE).strip()


def clean_taxpayer_name(value: str | None) -> str:
    """Нормализует отображение наименования без изменения его смысла."""
    if not value:
        return ""
    normalized = re.sub(r"\s+", " ", value).strip()
    normalized = re.sub(
        r"\bобщество\s+с\s+ограниченной\s+ответственностью\b",
        "ОсОО",
        normalized,
        flags=re.IGNORECASE,
    )
    return normalized


def normalize_search_text(value: object) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    text = text.replace("ё", "е")
    return " ".join(re.findall(r"[0-9a-zа-яңөү]+", text))


def _bounded_edit_distance(left: str, right: str, maximum: int) -> int:
    if abs(len(left) - len(right)) > maximum:
        return maximum + 1
    previous = list(range(len(right) + 1))
    for left_index, left_char in enumerate(left, 1):
        current = [left_index]
        for right_index, right_char in enumerate(right, 1):
            current.append(
                min(
                    current[-1] + 1,
                    previous[right_index] + 1,
                    previous[right_index - 1] + (left_char != right_char),
                )
            )
        if min(current) > maximum:
            return maximum + 1
        previous = current
    return previous[-1]


def fuzzy_search_match(value: object, query: object) -> int:
    """Prefix and typo-tolerant search for user-facing history only."""

    haystack = normalize_search_text(value)
    needle = normalize_search_text(query)
    if not needle:
        return 1
    if needle in haystack:
        return 1
    words = haystack.split()
    for token in needle.split():
        if any(word.startswith(token) or token in word for word in words):
            continue
        if len(token) < 4:
            return 0
        maximum = 1 if len(token) < 7 else 2
        if not any(
            _bounded_edit_distance(token, word, maximum) <= maximum
            for word in words
            if abs(len(token) - len(word)) <= maximum
        ):
            return 0
    return 1
