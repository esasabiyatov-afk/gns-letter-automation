from __future__ import annotations

from pathlib import Path

import pytest

from gns_app.config import Settings
from gns_app.domain import OcrStatus, PageType, QrStatus
from gns_app.services.classifier import PageClassifier
from gns_app.services.extractor import FieldExtractor
from gns_app.services.ocr_service import OcrService
from gns_app.services.pdf_service import PdfService
from gns_app.services.qr_service import QrService


def test_problematic_bank_pdf_qr_and_rus_kir_ocr(
    project_root: Path,
    tmp_path: Path,
):
    source = project_root / "Запрос (Банк).pdf"
    if not source.is_file():
        pytest.skip("Локальный регрессионный PDF отсутствует")

    settings = Settings.load()
    fast_data = settings.ocr_fast_data_dir
    if not fast_data or not (fast_data / "kir.traineddata").is_file():
        pytest.skip("Локальные модели rus+kir не установлены")

    pdf = PdfService()
    qr = QrService(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
    )
    rendered = {}
    for page_number in (1, 2, 52):
        rendered[page_number] = pdf.render_page(
            source,
            page_number,
            tmp_path / f"page-{page_number}.jpg",
            tmp_path / f"page-{page_number}-enhanced.jpg",
        )

    decision_qr = qr.decode(rendered[1].preview_path)
    letter_qr = qr.decode(rendered[2].preview_path)
    later_letter_qr = qr.decode(rendered[52].preview_path)
    assert decision_qr.status == QrStatus.FOUND
    assert letter_qr.status == QrStatus.FOUND
    assert later_letter_qr.status == QrStatus.FOUND
    assert decision_qr.payload_hash == letter_qr.payload_hash

    ocr = OcrService(
        pdf,
        settings.ocr_fast_data_dir,
        settings.ocr_best_data_dir,
    )
    result = ocr.recognize(source, 2, rendered[2].preview_path)
    assert result.status == OcrStatus.COMPLETED
    assert "02312201410117" in result.text.replace(" ", "")
    assert "Хе Син" in result.text
    assert "ө" in result.text or "ү" in result.text

    classification = PageClassifier().classify(
        result.text,
        rendered[2].quality_score,
    )
    fields = FieldExtractor().extract_scan_letter(result.text)
    assert classification.page_type == PageType.LETTER
    assert fields.taxpayers[0].inn == "02312201410117"
    assert fields.period_start == "2022-01-01"
    assert fields.period_end == "2026-06-30"
