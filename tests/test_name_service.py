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

