from gns_app.domain import PageType, VisualPageEvidence
from gns_app.services.classifier import PageClassifier


def test_classifies_decision_by_structural_markers():
    text = """
    AUDIT STI-010
    РЕШЕНИЕ
    РАЗДЕЛ I. ИНФОРМАЦИЯ О ПРОВЕРЯЕМОМ НАЛОГОПЛАТЕЛЬЩИКЕ
    102 ИНН 103 ФИО 104 налоговый орган 900 номер принятого решения
    """
    result = PageClassifier().classify(text, quality_score=0.8)
    assert result.page_type == PageType.DECISION
    assert result.automatic_terminal


def test_strong_form_structure_survives_missing_sti_title_in_ocr():
    text = """
    РЕШЕНИЕ
    РАЗДЕЛ I. ИНФОРМАЦИЯ О ПРОВЕРЯЕМОМ НАЛОГОПЛАТЕЛЬЩИКЕ
    102 ИНН 103 ФИО 104 налоговый орган
    РАЗДЕЛ III. ОФОРМЛЕНО ОРГАНОМ НАЛОГОВОЙ СЛУЖБЫ
    900 Номер принятого решения
    """

    result = PageClassifier().classify(text, quality_score=1.0)

    assert result.page_type == PageType.DECISION
    assert result.automatic_terminal


def test_codes_and_decision_title_without_form_structure_stay_manual():
    text = "РЕШЕНИЕ 102 103 104 900"

    result = PageClassifier().classify(text, quality_score=1.0)

    assert result.page_type == PageType.DECISION
    assert not result.automatic_terminal


def test_numbers_inside_inn_are_not_treated_as_form_codes():
    text = "РЕШЕНИЕ ИНН 1234102103104900"

    result = PageClassifier().classify(text, quality_score=1.0)

    assert not result.automatic_terminal
    assert not any("признак: 102" in reason for reason in result.reasons)


def test_classifies_letter_by_letter_markers():
    text = """
    Управление Государственной налоговой службы запрашивает информацию.
    о налогоплательщике:
    Наименование: ОсОО Тест
    ИНН: 12345678901234
    Период: с 01.01.2024 по 01.01.2025
    Зам. начальника управления
    """
    result = PageClassifier().classify(text, quality_score=0.8)
    assert result.page_type == PageType.LETTER


def test_does_not_guess_blurred_or_conflicting_page():
    text = "РЕШЕНИЕ Наименование: ИНН: Период:"
    result = PageClassifier().classify(text, quality_score=0.05)
    assert result.page_type == PageType.UNKNOWN
    assert result.confidence == 0
    assert not result.automatic_terminal


def test_decision_title_without_structure_is_not_automatic_terminal():
    text = "РЕШЕНИЕ STI-010"

    result = PageClassifier().classify(text, quality_score=0.8)

    assert result.page_type == PageType.DECISION
    assert not result.automatic_terminal


def test_audit_title_without_form_structure_is_not_automatic_terminal():
    text = "AUDIT STI-010 РЕШЕНИЕ"

    result = PageClassifier().classify(text, quality_score=0.8)

    assert result.page_type == PageType.DECISION
    assert not result.automatic_terminal


def test_letter_reference_to_sti_decision_stays_letter():
    text = """
    На основании Решения STI-010 № 008-2026-010-5845
    запрашивает информацию об операциях на счетах.
    О налогоплательщике:
    Наименование: ОсОО Тест
    ИНН: 12345678901234
    Период: с 01.01.2024 по 01.01.2025
    """

    result = PageClassifier().classify(text, quality_score=0.8)

    assert result.page_type == PageType.LETTER
    assert not result.automatic_terminal


def test_damaged_letter_reference_cannot_auto_complete_as_decision():
    text = """
    STI-010 РЕШЕНИЕ
    О ПРЕДОСТАВЛЕНИИ ИНФОРМАЦИИ ОБ ОПЕРАЦИЯХ
    """

    result = PageClassifier().classify(text, quality_score=0.8)

    assert not result.automatic_terminal


def test_low_quality_decision_is_not_automatic_terminal():
    text = """
    AUDIT STI-010 РЕШЕНИЕ РАЗДЕЛ I
    ИНФОРМАЦИЯ О ПРОВЕРЯЕМОМ НАЛОГОПЛАТЕЛЬЩИКЕ 102 103 104 900
    """

    result = PageClassifier().classify(text, quality_score=0.2)

    assert result.page_type == PageType.DECISION
    assert not result.automatic_terminal


def test_visual_form_structure_rescues_damaged_decision_ocr():
    text = """
    РЕШЕНИЕ
    О ПРЕДОСТАВЛЕНИИ ИНФОРМАЦИИ ОБ ОПЕРАЦИЯХ
    ПРИНЯТО РЕШЕНИЕ О ПРЕДОСТАВЛЕНИИ
    Период:
    """
    visual = VisualPageEvidence(
        decision_layout=True,
        confidence=0.95,
        horizontal_line_groups=24,
        vertical_line_groups=16,
    )

    result = PageClassifier().classify(text, 0.8, visual)

    assert result.page_type == PageType.DECISION
    assert result.automatic_terminal


def test_visual_grid_alone_cannot_turn_sti_reference_into_decision():
    visual = VisualPageEvidence(
        decision_layout=True,
        confidence=0.95,
        horizontal_line_groups=24,
        vertical_line_groups=16,
    )
    text = "На основании решения STI-010 запрашивает информацию"

    result = PageClassifier().classify(text, 0.8, visual)

    assert not result.automatic_terminal


def test_scanner_streaks_cannot_turn_letter_into_decision():
    visual = VisualPageEvidence(
        decision_layout=True,
        confidence=0.95,
        horizontal_line_groups=11,
        vertical_line_groups=9,
    )
    text = """
    На основании Решения STI-010 запрашивает информацию.
    О налогоплательщике:
    Наименование: ОсОО Тест
    ИНН: 12345678901234
    Период: с 01.01.2024 по 01.01.2025
    """

    result = PageClassifier().classify(text, 0.8, visual)

    assert result.page_type == PageType.LETTER
    assert not result.automatic_terminal
