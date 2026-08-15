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
        [{"name": 'ОсОО "Тест"', "inn": "02312201410117"}],
    )
    text = _all_word_text(output)
    assert output.exists()
    assert "02312201410117" in text
    assert 'ОсОО "Тест"' in text
    assert "Заместителю начальника управления" in text
    assert "Зам. начальника управления" not in text
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
            {"name": 'ОсОО "Первый"', "inn": "02312201410117"},
            {"name": 'ОсОО "Второй"', "inn": "02801202210185"},
        ],
    )
    text = _all_word_text(output)
    assert output.exists()
    assert "02312201410117" in text
    assert "02801202210185" in text
    assert "Заместителю начальника управления" in text
    assert "[Перечисления.Субьект]" not in text
    assert "[ИНН.Субьект]" not in text


def test_response_adds_ip_prefix_for_person(tmp_path, project_root):
    service = WordTemplateService(project_root / "УГНС")
    output = tmp_path / "person.docx"

    service.render(
        output,
        {
            "district_place": "по Ленинскому району города Бишкек",
            "recipient_position": "Начальник управления",
            "recipient_display_name": "Телтаеву Р. З.",
            "employee_name": "Гапарова Э.",
        },
        [{"name": "Ишен кызы Саида", "inn": "10207200101109"}],
    )

    text = _all_word_text(output)
    assert "ИП Ишен кызы Саида" in text
    assert "ИП ИП" not in text


def test_split_response_repeats_full_letter_on_new_page(tmp_path, project_root):
    service = WordTemplateService(project_root / "УГНС")
    output = tmp_path / "split.docx"
    case = {
        "district_place": "по Ленинскому району города Бишкек",
        "recipient_position": "Зам. начальника управления",
        "recipient_display_name": "Телтаеву Р. З.",
        "employee_name": "Гапарова Э.",
    }
    taxpayers = [
        {"name": f'ОсОО "Тест {index}"', "inn": f"{index:014d}"}
        for index in range(1, 5)
    ]

    service.render_pages(output, case, taxpayers, taxpayers_per_page=2)

    text = _all_word_text(output)
    assert text.count("по Ленинскому району города Бишкек") >= 2
    assert text.count("Настоящим ЗАО АКБ") >= 2
    assert all(item["inn"] in text for item in taxpayers)
