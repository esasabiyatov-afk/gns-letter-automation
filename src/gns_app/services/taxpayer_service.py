from __future__ import annotations

import re
from enum import StrEnum


class TaxpayerKind(StrEnum):
    LEGAL_ENTITY = "legal_entity"
    INDIVIDUAL = "individual"
    UNKNOWN = "unknown"


LEGAL_FORM_RE = re.compile(
    r"(?:^|\s)(?:осоо|ооо|оао|ао|зао|жчк|аак|жак|коо|"
    r"общество\s+с\s+ограниченной\s+ответственностью|"
    r"открытое\s+акционерное\s+общество|"
    r"закрытое\s+акционерное\s+общество|"
    r"государственное\s+предприятие|муниципальное\s+предприятие|"
    r"учреждение|фонд|ассоциация|кооператив|филиал|"
    r"представительство|дипломатическое\s+представительство)"
    r"(?:\s|$|[«\"'])",
    re.IGNORECASE,
)
ENTREPRENEUR_RE = re.compile(
    r"^\s*(?:ип\b|индивидуальн(?:ый|ая)\s+предпринимател(?:ь|я))",
    re.IGNORECASE,
)


def classify_taxpayer(name: str, inn: str) -> TaxpayerKind:
    """Classify conservatively; contradictory markers stay unknown."""
    digits = re.sub(r"\D", "", inn)
    if len(digits) != 14:
        return TaxpayerKind.UNKNOWN
    has_legal_form = bool(LEGAL_FORM_RE.search(name.strip()))
    has_entrepreneur_form = bool(ENTREPRENEUR_RE.search(name.strip()))
    prefix = digits[0]
    if prefix in {"0", "3", "5"}:
        return (
            TaxpayerKind.UNKNOWN
            if has_entrepreneur_form
            else TaxpayerKind.LEGAL_ENTITY
        )
    if prefix == "4":
        # В открытых описаниях этот префикс встречается и у иностранных
        # физлиц, и у иностранных филиалов. Без формы организации не гадаем.
        if has_entrepreneur_form:
            return TaxpayerKind.UNKNOWN
        return (
            TaxpayerKind.LEGAL_ENTITY
            if has_legal_form
            else TaxpayerKind.UNKNOWN
        )
    if has_legal_form:
        return TaxpayerKind.UNKNOWN
    if prefix in {"1", "2"}:
        return TaxpayerKind.INDIVIDUAL
    return TaxpayerKind.UNKNOWN


def response_taxpayer_name(name: str, inn: str) -> str:
    cleaned = " ".join(name.split())
    if classify_taxpayer(cleaned, inn) != TaxpayerKind.INDIVIDUAL:
        return cleaned
    entrepreneur_prefix = ENTREPRENEUR_RE.match(cleaned)
    if entrepreneur_prefix:
        suffix = cleaned[entrepreneur_prefix.end() :].lstrip()
        return f"ИП {suffix}".rstrip()
    return f"ИП {cleaned}"
