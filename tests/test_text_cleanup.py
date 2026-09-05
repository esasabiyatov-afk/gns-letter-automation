import pytest

from gns_app.text_cleanup import clean_location


@pytest.mark.parametrize(
    ("source", "expected"),
    [
        ("город Ош", "г.Ош Ошской области"),
        ("по городу Ош", "по г.Ош Ошской области"),
        ("из города Ош", "из г.Ош Ошской области"),
        ("в городе Бишкек", "в г.Бишкек"),
        ("городом Манас", "г.Манас Джалал-Абадской области"),
        ("по г. Ош", "по г.Ош Ошской области"),
        (
            "по г. Каракол Иссык-Кульской области",
            "по г.Каракол Иссык-Кульской области",
        ),
    ],
)
def test_city_words_are_shortened(source, expected):
    assert clean_location(source) == expected
