from pathlib import Path
from types import SimpleNamespace

from PIL import Image

from gns_app.domain import OcrResult, OcrStatus
from gns_app.services.ocr_service import OcrService


GARBLED_EMBEDDED_TEXT = """
lt'- KbIPI-bI3PECIIYEJI}TKACbIHbIH
MIIHII C TPJIEP KAE}IHETI4HE
KAPAIIITYY MAMJTEKETTI'TK CAJIbIK KbI3MATI,I
YTIPABJIEH I,IE fOCYAAPCTBEHHOII HAJTOfOBOfr C,rvXSlt
HauueHoBaHr4e : VIII Ilmqe- KbI3bI Cavha
tr4HH: 1020120010.1109
fleprao4: c 28.04.2023 no 22.07.2026
"""


def test_colored_artifact_mask_removes_blue_but_keeps_black_print():
    image = Image.new("RGB", (3, 1))
    image.putdata([(20, 20, 20), (40, 80, 180), (245, 245, 245)])

    cleaned = OcrService._suppress_colored_artifacts(image)

    assert [cleaned.getpixel((x, 0)) for x in range(3)] == [
        (20, 20, 20),
        (255, 255, 255),
        (245, 245, 245),
    ]


def test_high_resolution_region_preparation_is_binary():
    image = Image.new("RGB", (45, 45), (235, 235, 235))
    for x in range(10, 35):
        image.putpixel((x, 22), (20, 20, 20))

    prepared = OcrService._prepare_high_resolution_regions(image)

    assert prepared.mode == "L"
    values = {
        prepared.getpixel((x, y))
        for y in range(prepared.height)
        for x in range(prepared.width)
    }
    assert values.issubset({0, 255})


def test_recipient_signature_allows_sparse_ocr_blank_line():
    text = (
        "Зам. начальника управления\n\n"
        "Омошев Максат Тологонович"
    )

    assert OcrService._recipient_field_signature(text) == (
        "recipient:зам. начальника управления|"
        "омошев максат тологонович",
    )


def test_latin_garbled_embedded_layer_is_rejected():
    assert not OcrService._embedded_text_is_reliable(
        GARBLED_EMBEDDED_TEXT
    )


def test_valid_russian_and_kyrgyz_embedded_layer_is_accepted():
    text = (
        "Кыргыз Республикасынын Мамлекеттик салык кызматы. "
        "Управление Государственной налоговой службы запрашивает "
        "информацию о налогоплательщике. Наименование жана мезгил."
    )

    assert OcrService._embedded_text_is_reliable(text)


def test_partial_embedded_layer_is_supplemented_by_full_page_ocr(monkeypatch):
    embedded = (
        "Управление государственной налоговой службы запрашивает сведения. "
        "Наименование: ОсОО Тест ИНН: 01412201510188 "
        "Период: с 14.12.2015 по 09.09.2025"
    )
    image_text = (
        embedded
        + "\nЗам. начальника управления Омошев Максат Тологонович"
    )
    service = OcrService(
        SimpleNamespace(
            extract_embedded_text=lambda _path, _page: embedded
        ),
        fast_data_dir=Path("fast"),
        best_data_dir=Path("best"),
    )
    monkeypatch.setattr(service, "_model_available", lambda _path: True)
    monkeypatch.setattr(
        service,
        "_recognize_image",
        lambda _image, _data, model, **_kwargs: OcrResult(
            status=OcrStatus.COMPLETED,
            text=image_text,
            confidence=0.8 if model == "fast" else 0.82,
            language=f"rus+kir {model}",
        ),
    )

    result = service.recognize(
        Path("source.pdf"), 1, Path("page.jpg")
    )

    assert "=== Текстовый слой PDF ===" in result.text
    assert "=== OCR изображения всей страницы ===" in result.text
    assert "Омошев Максат Тологонович" in result.text
    assert result.recipient_fields_agree
    assert result.critical_fields_agree


def test_recipient_signature_accepts_name_on_next_ocr_line():
    signature = OcrService._recipient_field_signature(
        "Зам. начальника управления\n"
        "Омошев Максат Тологонович"
    )

    assert signature == (
        "recipient:зам. начальника управления|омошев максат тологонович",
    )


def test_rejected_embedded_layer_falls_back_to_local_rus_kir(monkeypatch):
    pdf_service = SimpleNamespace(
        extract_embedded_text=lambda _path, _page: GARBLED_EMBEDDED_TEXT
    )
    service = OcrService(
        pdf_service,
        fast_data_dir=Path("fast"),
        best_data_dir=Path("best"),
    )
    local_result = OcrResult(
        status=OcrStatus.COMPLETED,
        text=(
            "Мамлекеттик салык кызматы. "
            "Наименование: ИП Ишен кызы Саида. "
            "ИНН: 10207200101109. Период: с 28.04.2023 по 22.07.2026. "
        )
        * 3,
        confidence=0.81,
        language="rus+kir (Tesseract fast)",
    )
    monkeypatch.setattr(
        service,
        "_model_available",
        lambda path: path == Path("fast"),
    )
    monkeypatch.setattr(
        service,
        "_recognize_image",
        lambda _image, _data, _model, **_kwargs: local_result,
    )

    result = service.recognize(
        Path("source.pdf"),
        1,
        Path("page.jpg"),
    )

    assert result.status == OcrStatus.COMPLETED
    assert result.language == "rus+kir (Tesseract fast)"
    assert "Ишен кызы Саида" in result.text
    assert "Повреждённый текстовый слой PDF отклонён" in result.issue
    assert not result.critical_fields_agree


def test_two_local_models_must_agree_on_all_critical_fields(monkeypatch):
    pdf_service = SimpleNamespace(
        extract_embedded_text=lambda _path, _page: ""
    )
    service = OcrService(
        pdf_service,
        fast_data_dir=Path("fast"),
        best_data_dir=Path("best"),
    )
    fast = OcrResult(
        status=OcrStatus.COMPLETED,
        text=(
            "Наименование: ОсОО Тест ИНН: 12345678901234 "
            "Период: с 01.01.2020 по 01.01.2026"
        ),
        confidence=0.80,
        language="rus+kir fast",
    )
    best = OcrResult(
        status=OcrStatus.COMPLETED,
        text=(
            "Наименование: ОсОО Тест ИНН: 12345678901235 "
            "Период: с 01.01.2020 по 01.01.2026"
        ),
        confidence=0.82,
        language="rus+kir best",
    )
    monkeypatch.setattr(service, "_model_available", lambda _path: True)
    monkeypatch.setattr(
        service,
        "_recognize_image",
        lambda _image, _data, model, **_kwargs: (
            fast if model == "fast" else best
        ),
    )

    result = service.recognize(
        Path("source.pdf"), 1, Path("page.jpg")
    )

    assert not result.critical_fields_agree
    assert not result.taxpayer_fields_agree
    assert result.period_fields_agree
    assert "не подтвердили одинаково" in result.issue


def test_two_local_models_can_mark_exact_critical_consensus(monkeypatch):
    pdf_service = SimpleNamespace(
        extract_embedded_text=lambda _path, _page: ""
    )
    service = OcrService(
        pdf_service,
        fast_data_dir=Path("fast"),
        best_data_dir=Path("best"),
    )
    text = (
        "Наименование: ОсОО Тест ИНН: 12345678901234 "
        "Период: с 01.01.2020 по 01.01.2026"
    )
    monkeypatch.setattr(service, "_model_available", lambda _path: True)
    monkeypatch.setattr(
        service,
        "_recognize_image",
        lambda _image, _data, model, **_kwargs: OcrResult(
            status=OcrStatus.COMPLETED,
            text=text,
            confidence=0.80 if model == "fast" else 0.82,
            language=f"rus+kir {model}",
        ),
    )

    result = service.recognize(
        Path("source.pdf"), 1, Path("page.jpg")
    )

    assert result.critical_fields_agree
    assert "всё равно нужно сверить" in result.issue
