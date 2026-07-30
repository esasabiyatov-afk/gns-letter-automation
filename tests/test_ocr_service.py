from pathlib import Path
from types import SimpleNamespace

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
        lambda _image, _data, _model: local_result,
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
