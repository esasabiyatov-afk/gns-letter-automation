from gns_app.services.extractor import FieldExtractor


def test_official_district_does_not_span_two_document_pages():
    text = """
    Управление Государственной налоговой службы по Кара-Суйскому району
    Ошской области
    РЕШЕНИЕ
    Длинный текст решения без вводной фразы.

    Управление Государственной налоговой службы по Кара-Суйскому району
    Ошской области в соответствии со статьёй 146 запрашивает информацию.
    Наименование: Общество с ограниченной ответственностью "Хе Син"
    ИНН: 02312201410117
    Период: с 01.01.2022 по 30.06.2026
    Зам. начальника управления Торобек уулу Сталбек
    """

    fields = FieldExtractor().extract_official_letter(text)

    assert fields.district_place == (
        "по Кара-Суйскому району Ошской области"
    )
    assert fields.recipient_full_name == "Торобек уулу Сталбек"
    assert fields.taxpayers[0].inn == "02312201410117"
    assert fields.confidence == 0.96


def test_official_district_removes_only_stray_edge_quote():
    text = """
    Учреждение "Управление Государственной налоговой службы по городу
    Балыкчы Ысык-Кульской области" в соответствии со статьёй 146
    запрашивает информацию.
    Наименование: Общество с ограниченной ответственностью "Чардж"
    ИНН: 01608201810051
    Период: с 01.01.2020 по 30.06.2026
    Зам. начальника управления Омурбеков Нурлан Муратович
    """

    fields = FieldExtractor().extract_official_letter(text)

    assert fields.district_place == (
        "по г.Балыкчы Ысык-Кульской области"
    )
    assert fields.taxpayers[0].name == (
        'Общество с ограниченной ответственностью "Чардж"'
    )


def test_official_recipient_can_have_two_name_parts():
    text = """
    Учреждение "Управление государственной налоговой службы по Ноокатскому
    району Ошской области" в соответствии со статьёй 146 запрашивает информацию.
    Наименование: Косимжонов Зухриддин Абдилрузалиевич
    ИНН: 20111200100492
    Период: с 01.10.2024 по 21.08.2025
    Зам. начальника управления Жоробеков Тынчтыкбек
    """

    fields = FieldExtractor().extract_official_letter(text)

    assert fields.recipient_position == "Зам. начальника управления"
    assert fields.recipient_full_name == "Жоробеков Тынчтыкбек"
    assert fields.confidence == 0.96


def test_official_recipient_accepts_any_explicit_position_from_qr_text():
    text = """
    Учреждение "Управление государственной налоговой службы по Ат-Башинскому
    району Нарынской области" в соответствии со статьёй 146 запрашивает информацию.
    Наименование: ОсОО "Тест"
    ИНН: 02312201410117
    Период: с 01.01.2022 по 30.06.2026
    Начальник отдела Телтаев Рахатбек Замирбекович
    """

    fields = FieldExtractor().extract_official_letter(text)

    assert fields.recipient_position == "Начальник отдела"
    assert fields.recipient_full_name == "Телтаев Рахатбек Замирбекович"
    assert fields.confidence == 0.96


def test_scan_recipient_can_be_prefilled_from_full_page_ocr():
    text = """
    === Текстовый слой PDF ===
    Наименование: Общество с ограниченной ответственностью "Жер Компаниясы"
    ИНН: 01412201510188
    Период: с 14.12.2015 по 09.09.2025

    === OCR изображения всей страницы ===
    Зам. начальника управления
    Омошев Максат Тологонович
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert fields.recipient_position == "Зам. начальника управления"
    assert fields.recipient_full_name == "Омошев Максат Тологонович"


def test_scan_prefers_name_inn_pair_followed_by_period():
    text = """
    === OCR изображения всей страницы ===
    Наименование: Ткачев Владимир Александрович
    ИНН: 21303198400213
    Период: с 29.04.2023 по 18.09.2025

    === Дополнительная OCR-область 1 ===
    Наименование: Ткачев Владимир Александрович
    ИНН: 42209199214499

    === Дополнительная OCR-область 2 ===
    Наименование: Ткачев Владимир Александрович
    ИНН: 21303198400213
    Период: с 29.04.2023 по 18.09.2025
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert [(item.name, item.inn) for item in fields.taxpayers] == [
        ("Ткачев Владимир Александрович", "21303198400213")
    ]


def test_scan_does_not_use_bank_stamp_label_as_recipient_name():
    text = """
    Зам. начальника управления
    Количество листов
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert fields.recipient_position == "Зам. начальника управления"
    assert fields.recipient_full_name is None


def test_scan_uses_visible_14_digits_when_inn_label_is_damaged():
    text = """
    === OCR изображения всей страницы ===
    Наименование: Общественный фонд "Бакай-Ата"
    ПИ: 02801202010240
    Период: с 28.01.2020 по 08.09.2025
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert [(item.name, item.inn) for item in fields.taxpayers] == [
        ('Общественный фонд "Бакай-Ата"', "02801202010240")
    ]


def test_scan_associates_period_before_name_with_visible_inn():
    text = """
    === OCR изображения всей страницы ===
    Период: с 28.01.2020 по 08.09.2025
    Наименование: Общественный фонд "Бакай-Ата"
    ПИ: 02801202010240
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert [(item.name, item.inn) for item in fields.taxpayers] == [
        ('Общественный фонд "Бакай-Ата"', "02801202010240")
    ]


def test_scan_uses_same_visible_inn_repeated_by_independent_ocr_passes():
    text = """
    === Текстовый слой PDF ===
    ИНН: 02801202010240
    Период: с 28.01.2020 по 08.09.2025

    === OCR изображения всей страницы ===
    ПИ: 02801202010240
    сриод: с 28.01.2020 по 08.09.2025
    Наименование: Общественный фонд "Бакай-Ата"
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert [(item.name, item.inn) for item in fields.taxpayers] == [
        ('Общественный фонд "Бакай-Ата"', "02801202010240")
    ]


def test_scan_uses_damaged_inn_label_with_spaces_between_digits():
    text = """
    Наименование: Общественный фонд "Бакай-Ата"
    АГИ: 0280120201 0240
    Период: с 28.01.2020 по 08.09.2025
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert fields.taxpayers[0].inn == "02801202010240"


def test_scan_prefills_position_even_when_stamp_hides_recipient_surname():
    text = """
    Зам. начальника управления
    мканов Улан Маратович
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert fields.recipient_position == "Зам. начальника управления"
    assert fields.recipient_full_name is None


def test_scan_recovers_recipient_after_short_stamp_noise_lines():
    text = """
    === Дополнительная OCR-область 1 ===
    Зам. начальника управления.
    КО.
    ии
    Омурбеков Нурлан Муратович
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert fields.recipient_full_name == "Омурбеков Нурлан Муратович"


def test_official_name_wrapped_onto_second_line_is_not_truncated():
    # Реальный случай из пакета банковских запросов УГНС по Кара-Суйскому
    # району: длинное наименование ОсОО визуально переносится на вторую
    # строку без отдельной метки. Раньше регулярка обрывала имя на первом
    # переводе строки и теряла хвост названия ("Строй Компани" вместо
    # "Строй Компани Каракол").
    text = """
    Управление Государственной налоговой службы по Кара-Суйскому району
    Ошской области в соответствии со статьёй 146 запрашивает информацию.
    Наименование: Общество с ограниченной ответственностью "Строй Компани
    Каракол"
    ИНН: 01110202410416
    Период: с 01.01.2022 по 30.06.2026
    Зам. начальника управления Торобек уулу Сталбек
    """

    fields = FieldExtractor().extract_official_letter(text)

    assert fields.taxpayers[0].name == (
        'Общество с ограниченной ответственностью "Строй Компани Каракол"'
    )
    assert fields.taxpayers[0].inn == "01110202410416"
    assert fields.confidence == 0.96


def test_official_name_wrapped_with_blank_line_before_inn_label():
    # Тот же перенос имени, но с пустой строкой перед меткой ИНН (как в
    # официальном docx-шаблоне ГНС, где между блоками оставлен отступ).
    text = """
    Управление Государственной налоговой службы по Кара-Суйскому району
    Ошской области в соответствии со статьёй 146 запрашивает информацию.
    Наименование: Общество с ограниченной ответственностью "Строй Компани
    Каракол"

    ИНН: 01110202410416
    Период: с 01.01.2022 по 30.06.2026
    Зам. начальника управления Торобек уулу Сталбек
    """

    fields = FieldExtractor().extract_official_letter(text)

    assert fields.taxpayers[0].name == (
        'Общество с ограниченной ответственностью "Строй Компани Каракол"'
    )


def test_scan_does_not_use_unlabelled_number_without_period_boundary():
    text = """
    Наименование: Общественный фонд "Бакай-Ата"
    Служебный номер: 02801202010240
    """

    fields = FieldExtractor().extract_scan_letter(text)

    assert fields.taxpayers == []
