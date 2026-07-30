from dataclasses import replace
from pathlib import Path

from gns_app.database import Database
from gns_app.domain import (
    CaseStatus,
    ExtractedFields,
    ExtractedTaxpayer,
    ValueSource,
)
from gns_app.services.workflow import WorkflowService


def test_complete_official_qr_needs_no_manual_confirmation(
    test_settings,
    project_root: Path,
    monkeypatch,
):
    settings = replace(test_settings, employee_name="")
    database = Database(settings.database_path)
    database.initialize()
    workflow = WorkflowService(database, settings)

    sample = project_root / "УГНС" / "пример письма.pdf"
    with sample.open("rb") as stream:
        upload_id = workflow.create_upload(sample.name, stream)
    case_id, _ = workflow._ensure_qr_case(upload_id, "verified-qr-hash")

    fields = ExtractedFields(
        district_place="по Ысык-Атинскому району Чуйской области",
        recipient_position="Зам. начальника управления",
        recipient_full_name="Мевазов Юсуп Харсанович",
        period_start="2023-04-28",
        period_end="2026-07-22",
        taxpayers=[
            ExtractedTaxpayer(
                name="ИП Ишен кызы Саида",
                inn="10207200101109",
                confidence=0.99,
            )
        ],
        confidence=0.96,
    )
    monkeypatch.setattr(workflow, "_official_text", lambda _path: "official")
    monkeypatch.setattr(
        workflow.extractor,
        "extract_official_letter",
        lambda _text: fields,
    )

    complete = workflow._apply_official_document(case_id, sample)
    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(
            period_start="2020-01-01",
            period_end="2020-12-31",
            taxpayers=[
                ExtractedTaxpayer(
                    name="Ошибочный OCR",
                    inn="00000000000000",
                    confidence=0.5,
                )
            ],
        ),
    )

    case = workflow.get_case(case_id)
    taxpayers = workflow.get_taxpayers(case_id)
    assert complete
    assert case["source_kind"] == ValueSource.QR_OFFICIAL
    assert case["fields_confirmed"] == 1
    assert case["status"] == CaseStatus.READY_FOR_ABS
    assert case["employee_name"] is None
    assert case["period_start"] == "2023-04-28"
    assert case["period_end"] == "2026-07-22"
    assert taxpayers[0]["name"] == "ИП Ишен кызы Саида"
    assert taxpayers[0]["inn"] == "10207200101109"
    assert taxpayers[0]["manually_confirmed"] == 1
