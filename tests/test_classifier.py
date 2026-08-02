from gns_app.domain import PageType
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


def test_low_quality_decision_is_not_automatic_terminal():
    text = """
    AUDIT STI-010 РЕШЕНИЕ РАЗДЕЛ I
    ИНФОРМАЦИЯ О ПРОВЕРЯЕМОМ НАЛОГОПЛАТЕЛЬЩИКЕ 102 103 104 900
    """

    result = PageClassifier().classify(text, quality_score=0.2)

    assert result.page_type == PageType.DECISION
    assert not result.automatic_terminal
