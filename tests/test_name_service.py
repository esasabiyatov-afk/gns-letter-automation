from gns_app.services.name_service import NameService


def test_declines_clear_russian_male_name():
    service = NameService()
    assert (
        service.recipient_display("Телтаев Рахатбек Замирбекович")
        == "Телтаеву Р. З."
    )


def test_declines_clear_russian_female_name():
    service = NameService()
    assert (
        service.recipient_display("Иванова Елена Сергеевна")
        == "Ивановой Е. С."
    )


def test_keeps_uulu_and_kyzy_constructions():
    service = NameService()
    assert service.recipient_display("Амантур уулу Жанат") == "Амантур уулу Ж."
    assert service.recipient_display("Айжан кызы Нурия") == "Айжан кызы Н."


def test_preserves_kyrgyz_unicode_letters():
    service = NameService()
    result = service.recipient_display("Өмүрбек уулу Үсөн")
    assert result == "Өмүрбек уулу Ү."
    assert "Ө" in result
    assert "Ү" in result


def test_declines_position_head_but_keeps_dependent_genitive():
    service = NameService()
    assert service.position_display(
        "Начальник управления по работе с клиентами"
    ) == "Начальнику управления по работе с клиентами"
    assert service.position_display(
        "Заместитель начальника управления"
    ) == "Заместителю начальника управления"


def test_declines_agreeing_words_before_position_head():
    service = NameService()
    assert service.position_display(
        "Главный государственный налоговый инспектор управления"
    ) == (
        "Главному государственному налоговому инспектору управления"
    )
    assert service.position_display(
        "Старший специалист отдела"
    ) == "Старшему специалисту отдела"
    assert service.position_display(
        "Заведующий отделом"
    ) == "Заведующему отделом"


def test_expands_common_position_abbreviations_safely():
    service = NameService()
    assert service.position_display(
        "Зам. начальника управления"
    ) == "Заместителю начальника управления"
    assert service.position_display(
        "И. о. начальника управления"
    ) == "Исполняющему обязанности начальника управления"
    assert service.position_display(
        "Врио начальника управления"
    ) == "Временно исполняющему обязанности начальника управления"


def test_keeps_already_dative_or_unrecognized_position():
    service = NameService()
    assert service.position_display(
        "Начальнику управления"
    ) == "Начальнику управления"
    assert service.position_display("Башкы адис Өмүрбек") == (
        "Башкы адис Өмүрбек"
    )


def test_declines_common_bank_and_government_positions():
    service = NameService()
    examples = {
        "Первый заместитель председателя правления": (
            "Первому заместителю председателя правления"
        ),
        "Управляющий филиалом": "Управляющему филиалом",
        "Генеральный директор ОсОО «Тест»": (
            "Генеральному директору ОсОО «Тест»"
        ),
        "Председатель правления": "Председателю правления",
        "Руководитель аппарата": "Руководителю аппарата",
        "Советник директора": "Советнику директора",
        "Ведущий специалист сектора": "Ведущему специалисту сектора",
    }
    for source, expected in examples.items():
        assert service.position_display(source) == expected
