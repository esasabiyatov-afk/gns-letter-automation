from __future__ import annotations

import re

from gns_app.domain import ClassificationResult, PageType


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.casefold())


class PageClassifier:
    DECISION_MARKERS = {
        "решение": 3,
        "audit sti-010": 4,
        "sti-010": 3,
        "раздел i": 2,
        "информация о проверяемом налогоплательщике": 3,
        "номер принятого решения": 2,
        "102": 1,
        "103": 1,
        "104": 1,
        "900": 1,
    }
    LETTER_MARKERS = {
        "запрашивает информацию": 4,
        "о налогоплательщике": 3,
        "наименование:": 2,
        "инн:": 1,
        "период:": 3,
        "с уважением": 1,
        "зам. начальника управления": 2,
    }
    DECISION_IDENTITY_MARKERS = {
        "audit sti-010",
        "sti-010",
    }
    DECISION_STRUCTURE_MARKERS = {
        "решение",
        "раздел i",
        "информация о проверяемом налогоплательщике",
        "номер принятого решения",
    }

    def classify(
        self, text: str, quality_score: float
    ) -> ClassificationResult:
        normalized = _normalize(text)
        decision_score, decision_reasons = self._score(
            normalized, self.DECISION_MARKERS
        )
        letter_score, letter_reasons = self._score(
            normalized, self.LETTER_MARKERS
        )

        best_score = max(decision_score, letter_score)
        difference = abs(decision_score - letter_score)
        if best_score < 4 or difference < 2 or quality_score < 0.12:
            reasons = ["Недостаточно согласованных признаков типа страницы"]
            reasons.extend(decision_reasons[:2])
            reasons.extend(letter_reasons[:2])
            return ClassificationResult(PageType.UNKNOWN, 0.0, reasons)

        page_type = (
            PageType.DECISION
            if decision_score > letter_score
            else PageType.LETTER
        )
        raw_confidence = min(0.92, 0.45 + difference / 14 + best_score / 40)
        confidence = min(raw_confidence, 0.78)
        reasons = (
            decision_reasons if page_type == PageType.DECISION else letter_reasons
        )
        automatic_terminal = bool(
            page_type == PageType.DECISION
            and quality_score >= 0.35
            and decision_score >= 8
            and difference >= 4
            and len(decision_reasons) >= 3
            and any(
                marker in normalized
                for marker in self.DECISION_IDENTITY_MARKERS
            )
            and any(
                marker in normalized
                for marker in self.DECISION_STRUCTURE_MARKERS
            )
        )
        return ClassificationResult(
            page_type=page_type,
            confidence=round(confidence, 3),
            reasons=reasons,
            automatic_terminal=automatic_terminal,
        )

    @staticmethod
    def _score(
        normalized: str, markers: dict[str, int]
    ) -> tuple[int, list[str]]:
        score = 0
        reasons: list[str] = []
        for marker, weight in markers.items():
            if marker in normalized:
                score += weight
                reasons.append(f"Найден признак: {marker}")
        return score, reasons
