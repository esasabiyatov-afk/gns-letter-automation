from __future__ import annotations

import re

from gns_app.domain import ClassificationResult, PageType, VisualPageEvidence


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.casefold())


def _contains_marker(normalized: str, marker: str) -> bool:
    return bool(
        re.search(
            rf"(?<!\w){re.escape(marker)}(?!\w)",
            normalized,
        )
    )


class PageClassifier:
    DECISION_MARKERS = {
        "решение": 3,
        "audit sti-010": 4,
        "sti-010": 3,
        "раздел i": 2,
        "информация о проверяемом налогоплательщике": 3,
        "номер принятого решения": 2,
        "о предоставлении информации об операциях": 3,
        "проводимых на счетах организаций": 2,
        "основание запроса": 2,
        "оформлено органом налоговой службы": 2,
        "принято решение о предоставлении": 2,
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
        "о предоставлении информации об операциях",
        "проводимых на счетах организаций",
        "основание запроса",
        "оформлено органом налоговой службы",
        "принято решение о предоставлении",
    }
    DECISION_EXCLUSIVE_STRUCTURE_MARKERS = {
        "раздел i",
        "информация о проверяемом налогоплательщике",
        "номер принятого решения",
        "основание запроса",
        "оформлено органом налоговой службы",
        "принято решение о предоставлении",
    }
    DECISION_FORM_CODE_MARKERS = {
        "102",
        "103",
        "104",
        "900",
    }

    def classify(
        self,
        text: str,
        quality_score: float,
        visual_evidence: VisualPageEvidence | None = None,
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
        has_identity = any(
            _contains_marker(normalized, marker)
            for marker in self.DECISION_IDENTITY_MARKERS
        )
        exclusive_structure_count = sum(
            _contains_marker(normalized, marker)
            for marker in self.DECISION_EXCLUSIVE_STRUCTURE_MARKERS
        )
        form_code_count = sum(
            _contains_marker(normalized, marker)
            for marker in self.DECISION_FORM_CODE_MARKERS
        )
        has_decision_title = _contains_marker(normalized, "решение")
        structure_count = sum(
            _contains_marker(normalized, marker)
            for marker in self.DECISION_STRUCTURE_MARKERS
        )
        visual_form_support = bool(
            visual_evidence
            and visual_evidence.decision_layout
            and visual_evidence.confidence >= 0.82
            and decision_score >= 5
            and difference >= 4
            and len(decision_reasons) >= 2
            and quality_score >= 0.25
            and (
                (has_decision_title and has_identity)
                or (has_decision_title and structure_count >= 2)
                or (has_decision_title and form_code_count >= 3)
                or (
                    has_identity
                    and (
                        exclusive_structure_count >= 1
                        or form_code_count >= 1
                    )
                )
                or (
                    exclusive_structure_count >= 2
                    and form_code_count >= 1
                )
            )
        )
        if (
            best_score < 4
            or difference < 2
            or quality_score < 0.12
        ) and not visual_form_support:
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
        has_strong_form_structure = bool(
            has_decision_title
            and exclusive_structure_count >= 1
            and form_code_count >= 3
            and decision_score >= 8
            and difference >= 5
        )
        automatic_terminal = bool(
            page_type == PageType.DECISION
            and quality_score >= 0.35
            and difference >= 4
            and (
                visual_form_support
                or (
                    decision_score >= 8
                    and len(decision_reasons) >= 3
                    and (
                        (
                            has_identity
                            and exclusive_structure_count >= 1
                            and form_code_count >= 2
                        )
                        or has_strong_form_structure
                    )
                )
            )
        )
        if visual_form_support:
            confidence = max(confidence, 0.9)
            reasons.extend(visual_evidence.reasons)
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
            if _contains_marker(normalized, marker):
                score += weight
                reasons.append(f"Найден признак: {marker}")
        return score, reasons
