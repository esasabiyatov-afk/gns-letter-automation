from __future__ import annotations

from pathlib import Path

from gns_app.domain import (
    CaseStatus,
    ClassificationResult,
    OcrResult,
    OcrStatus,
    PageType,
    QrDecodeResult,
    QrStatus,
)


def test_sample_pdf_end_to_end_without_guessing(
    workflow,
    sample_pdf: Path,
    monkeypatch,
):
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)

    monkeypatch.setattr(
        workflow.qr,
        "decode",
        lambda _path: QrDecodeResult(
            status=QrStatus.FOUND,
            payload="https://qr.salyk.kg/getsti010decission?id=synthetic",
            payload_hash="synthetic-workflow-case",
            safe_url="https://qr.salyk.kg/getsti010decission",
            method="synthetic-test",
        ),
    )
    monkeypatch.setattr(
        workflow.ocr,
        "recognize",
        lambda *_args, **_kwargs: OcrResult(
            status=OcrStatus.COMPLETED,
            text="Тестовое письмо ГНС",
            confidence=0.99,
            language="synthetic",
        ),
    )
    monkeypatch.setattr(
        workflow.classifier,
        "classify",
        lambda *_args, **_kwargs: ClassificationResult(
            page_type=PageType.LETTER,
            confidence=0.99,
        ),
    )

    workflow.process_upload(upload_id)
    upload = workflow.get_upload(upload_id)
    pages = workflow.get_upload_pages(upload_id)

    assert upload["page_count"] == 2
    assert len(pages) == 2
    assert {page["qr_status"] for page in pages} == {"found"}
    assert all(
        page["status"]
        not in {"registered", "preview_ready", "processing"}
        for page in pages
    )

    letter_page = next(
        page for page in pages if page["page_type"] == "letter"
    )
    assert workflow.get_case(letter_page["case_id"])["source_kind"] == "qr_link"
    case_id = workflow.confirm_page(
        letter_page["id"],
        page_type="letter",
        district_place="по Ленинскому району города Бишкек",
        recipient_position="Зам. начальника управления",
        recipient_full_name="Телтаев Рахатбек Замирбекович",
        recipient_display_name="",
        period_start="2019-11-14",
        period_end="2025-09-10",
        employee_name="Гапарова Э.",
        critical_fields_verified=True,
        taxpayers=[
            {
                "name": (
                    'Филиал Общества с ограниченной ответственностью '
                    '"ВИТЕЛ 11" в Кыргызской Республике'
                ),
                "inn": "01411201910186",
            }
        ],
    )
    case = workflow.get_case(case_id)
    assert case["recipient_display_name"] == "Телтаеву Р. З."
    assert case["source_kind"] == "manual"
    assert case["status"] == CaseStatus.READY_FOR_RESPONSE
    assert workflow.get_taxpayers(case_id)[0]["abs_result"] == "not_found"

    response = workflow.generate_response(case_id)
    assert response.exists()
    assert workflow.get_case(case_id)["status"] == CaseStatus.RESPONSE_CREATED
