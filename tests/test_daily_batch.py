from __future__ import annotations

from uuid import uuid4

from docx import Document

from gns_app.database import utc_now
from gns_app.domain import (
    CaseStatus,
    ExtractedFields,
    ExtractedTaxpayer,
    PageStatus,
    PageType,
    QrStatus,
)


def _insert_ready_case(workflow, inn: str, name: str) -> str:
    upload_id = uuid4().hex
    case_id = uuid4().hex
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, ?, ?, ?, 1, 'ready', ?)
        """,
        (upload_id, f"{case_id}.pdf", f"{case_id}.pdf", case_id, now),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind, district_place,
            recipient_position, recipient_full_name,
            recipient_display_name, period_start, period_end,
            employee_name, fields_confirmed, abs_status,
            created_at, updated_at
        ) VALUES (?, ?, 'ready_for_abs', 'qr_official',
                  'по Ленинскому району города Бишкек',
                  'Зам. начальника управления',
                  'Телтаев Рахатбек Замирбекович',
                  'Телтаеву Р. З.', '2020-01-01', '2026-01-01',
                  'Гапарова Э.', 1, 'not_checked', ?, ?)
        """,
        (case_id, upload_id, now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn,
            name_source, inn_source, manually_confirmed,
            created_at, updated_at
        ) VALUES (?, ?, 1, ?, ?, 'qr_official', 'qr_official', 1, ?, ?)
        """,
        (uuid4().hex, case_id, name, inn, now, now),
    )
    return case_id


def test_today_batch_checks_and_groups_by_recipient(workflow):
    first = _insert_ready_case(
        workflow, "12345678901234", 'ОсОО "Первый"'
    )
    second = _insert_ready_case(
        workflow, "23456789012345", 'ОсОО "Второй"'
    )

    summary = workflow.check_abs_today("batch-user", "one-time-secret")
    overview = workflow.today_overview()

    assert summary["case_count"] == 2
    assert len(overview["not_found_groups"]) == 1
    group = overview["not_found_groups"][0]
    assert group["case_count"] == 2
    assert group["taxpayer_count"] == 2
    assert group["can_generate"]

    group_id, output = workflow.generate_daily_response(group["group_key"])
    document_text = "\n".join(
        paragraph.text for paragraph in Document(output).paragraphs
    )

    assert group_id
    assert 'ОсОО "Первый"' in document_text
    assert 'ОсОО "Второй"' in document_text
    assert "12345678901234" in document_text
    assert "23456789012345" in document_text
    taxpayer_lines = [
        line.strip()
        for line in document_text.splitlines()
        if "ИНН:" in line
    ]
    assert taxpayer_lines[0].startswith("1. ")
    assert taxpayer_lines[1].startswith("2. ")
    assert workflow.get_case(first)["status"] == CaseStatus.RESPONSE_CREATED
    assert workflow.get_case(second)["status"] == CaseStatus.RESPONSE_CREATED
    assert not workflow.today_overview()["not_found_groups"]
    audit_text = "\n".join(
        event["payload_json"] for event in workflow.get_audit()
    )
    assert "batch-user" not in audit_text
    assert "one-time-secret" not in audit_text


def test_found_taxpayer_is_separate_and_never_uses_absence_template(workflow):
    _insert_ready_case(
        workflow, "11111111111111", 'ОсОО "Найденный"'
    )

    workflow.check_abs_today("batch-user", "one-time-secret")
    overview = workflow.today_overview()

    assert not overview["not_found_groups"]
    assert len(overview["found_groups"]) == 1
    assert not overview["found_groups"][0]["can_generate"]
    assert "отдельный утверждённый шаблон" in " ".join(
        overview["found_groups"][0]["issues"]
    )


def test_ocr_only_decision_stays_in_manual_review(workflow):
    status, issue_code, _ = workflow._page_outcome(
        PageType.DECISION,
        0.99,
        QrStatus.NOT_FOUND,
        False,
        None,
    )

    assert status == PageStatus.NEEDS_REVIEW
    assert issue_code == "manual_review_required"


def test_ocr_disagreement_does_not_prefill_critical_fields(workflow):
    upload_id = uuid4().hex
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, 'scan.pdf', 'scan.pdf', 'hash', 1, 'ready', ?)
        """,
        (upload_id, now),
    )
    case_id = workflow._create_scan_case(upload_id, "page-for-audit")
    fields = ExtractedFields(
        period_start="2020-01-01",
        period_end="2026-01-01",
        taxpayers=[
            ExtractedTaxpayer(
                name='ОсОО "Сомнительный"',
                inn="12345678901234",
                confidence=0.5,
            )
        ],
    )

    workflow._prefill_scan_case(
        case_id,
        fields,
        critical_fields_agree=False,
    )

    case = workflow.get_case(case_id)
    assert case["period_start"] is None
    assert case["period_end"] is None
    assert workflow.get_taxpayers(case_id) == []
