from dataclasses import replace
from pathlib import Path

import pytest

from gns_app.database import Database
from gns_app.domain import (
    CaseStatus,
    ExtractedFields,
    ExtractedTaxpayer,
    PageStatus,
    QrStatus,
    ValueSource,
)
from gns_app.services.workflow import WorkflowService
from gns_app.services.workflow import WorkflowValidationError


def test_complete_official_qr_needs_no_manual_confirmation(
    test_settings,
    sample_pdf: Path,
    monkeypatch,
):
    settings = replace(test_settings, employee_name="")
    database = Database(settings.database_path)
    database.initialize()
    workflow = WorkflowService(database, settings)

    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
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

    complete = workflow._apply_official_document(case_id, sample_pdf)
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
    assert case["status"] == CaseStatus.READY_FOR_RESPONSE
    assert case["abs_status"] == "not_found"
    assert case["employee_name"] is None
    assert case["period_start"] == "2023-04-28"
    assert case["period_end"] == "2026-07-22"
    assert taxpayers[0]["name"] == "ИП Ишен кызы Саида"
    assert taxpayers[0]["inn"] == "10207200101109"
    assert taxpayers[0]["abs_result"] == "not_found"
    assert taxpayers[0]["manually_confirmed"] == 1


def test_stale_review_page_is_completed_when_official_qr_is_complete(
    workflow,
    sample_pdf: Path,
):
    workflow.initialize_employee_profiles()
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
    page = workflow.get_upload_pages(upload_id)[0]
    case_id, _ = workflow._ensure_qr_case(upload_id, "complete-official-qr")

    workflow.db.execute(
        """
        UPDATE cases
        SET source_kind = ?, fields_confirmed = 1,
            official_document_path = ?, status = ?
        WHERE id = ?
        """,
        (
            ValueSource.QR_OFFICIAL,
            str(sample_pdf),
            CaseStatus.READY_FOR_ABS,
            case_id,
        ),
    )
    workflow.db.execute(
        """
        UPDATE pages
        SET case_id = ?, qr_status = ?, status = ?,
            issue_code = ?, issue_message = ?
        WHERE id = ?
        """,
        (
            case_id,
            QrStatus.FOUND,
            PageStatus.NEEDS_REVIEW,
            "official_pending",
            "Устаревшее требование подтверждения",
            page["id"],
        ),
    )

    repaired = workflow.reconcile_official_qr_pages(upload_id)
    refreshed = workflow.get_page(page["id"])

    assert repaired == 1
    assert refreshed["status"] == PageStatus.COMPLETED
    assert refreshed["issue_code"] is None
    assert refreshed["issue_message"] is None

    with pytest.raises(
        WorkflowValidationError,
        match="уже обработана",
    ):
        workflow.confirm_page(page["id"], page_type="letter")


def test_incomplete_official_document_is_reparsed_after_parser_update(
    workflow, monkeypatch
):
    workflow.initialize_employee_profiles()
    upload_id = "official-reparse-upload"
    now = "2026-08-02T00:00:00+00:00"
    official_path = workflow.settings.official_dir / "case-reparse" / "official.pdf"
    official_path.parent.mkdir(parents=True, exist_ok=True)
    official_path.write_bytes(b"%PDF-test")
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, 'packet.pdf', 'packet.pdf', ?, 1, 'needs_review', ?)
        """,
        (upload_id, "hash-reparse", now),
    )
    case_id, _ = workflow._ensure_qr_case(upload_id, "reparse-hash")
    page_id = "official-reparse-page"
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, case_id, page_number, page_type,
            qr_status, status, created_at, updated_at
        ) VALUES (?, ?, ?, 1, 'letter', 'found', 'needs_review', ?, ?)
        """,
        (page_id, upload_id, case_id, now, now),
    )
    workflow.db.execute(
        """
        UPDATE cases
        SET source_kind = ?, official_document_path = ?,
            fields_confirmed = 0, official_parse_version = 0
        WHERE id = ?
        """,
        (ValueSource.QR_OFFICIAL, str(official_path), case_id),
    )
    fields = ExtractedFields(
        district_place="по Ноокатскому району Ошской области",
        recipient_position="Зам. начальника управления",
        recipient_full_name="Жоробеков Тынчтыкбек",
        period_start="2024-10-01",
        period_end="2025-08-21",
        taxpayers=[
            ExtractedTaxpayer(
                name="Косимжонов Зухриддин Абдилрузалиевич",
                inn="20111200100492",
                confidence=0.96,
            )
        ],
        confidence=0.96,
    )
    monkeypatch.setattr(workflow, "_official_text", lambda path: "official")
    monkeypatch.setattr(
        workflow.extractor,
        "extract_official_letter",
        lambda text: fields,
    )

    completed = workflow.reprocess_incomplete_official_documents()

    case = workflow.get_case(case_id)
    page = workflow.get_page(page_id)
    assert completed == 1
    assert case["fields_confirmed"] == 1
    assert case["official_parse_version"] == workflow.OFFICIAL_PARSER_VERSION
    assert page["status"] == PageStatus.COMPLETED
