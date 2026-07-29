from __future__ import annotations

import zipfile

from docx import Document

from gns_app.services.word_service import WordTemplateService


def _all_word_text(path) -> str:
    document = Document(str(path))
    parts = [paragraph.text for paragraph in document.paragraphs]
    with zipfile.ZipFile(path) as archive:
        parts.extend(
            archive.read(name).decode("utf-8", errors="ignore")
            for name in archive.namelist()
            if name.startswith("word/") and name.endswith(".xml")
        )
    return "\n".join(parts)


def test_renders_single_response_without_placeholders(
    tmp_path, project_root
):
    service = WordTemplateService(project_root / "УГНС")
    output = tmp_path / "single.docx"
    service.render(
        output,
        {
            "district_place": "по Ленинскому району города Бишкек",
            "recipient_position": "Зам. начальника управления",
            "recipient_display_name": "Телтаеву Р. З.",
            "employee_name": "Гапарова Э.",
        },
        [{"name": 'ОсОО "Тест"', "inn": "12345678901234"}],
    )
    text = _all_word_text(output)
    assert output.exists()
    assert "12345678901234" in text
    assert 'ОсОО "Тест"' in text
    assert "[Дата.Сегодня]" not in text
    assert "[ФИО.Исп]" not in text


def test_renders_multiple_taxpayers_on_separate_lines(
    tmp_path, project_root
):
    service = WordTemplateService(project_root / "УГНС")
    output = tmp_path / "multi.docx"
    service.render(
        output,
        {
            "district_place": "по Ленинскому району города Бишкек",
            "recipient_position": "Зам. начальника управления",
            "recipient_display_name": "Телтаеву Р. З.",
            "employee_name": "Гапарова Э.",
        },
        [
            {"name": 'ОсОО "Первый"', "inn": "12345678901234"},
            {"name": 'ОсОО "Второй"', "inn": "23456789012345"},
        ],
    )
    text = _all_word_text(output)
    assert output.exists()
    assert "12345678901234" in text
    assert "23456789012345" in text
    assert "[Перечисления.Субьект]" not in text
    assert "[ИНН.Субьект]" not in text

