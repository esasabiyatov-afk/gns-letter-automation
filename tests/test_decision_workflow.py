from gns_app.database import utc_now
from gns_app.domain import PageStatus


def _insert_upload(workflow, upload_id: str):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, 'packet.pdf', 'packet.pdf', ?, 2, 'processing', ?)
        """,
        (upload_id, f"hash-{upload_id}", utc_now()),
    )


def _insert_page(
    workflow,
    page_id: str,
    upload_id: str,
    page_number: int,
    page_type: str,
    status: str,
    text: str = "",
    quality_score: float = 1.0,
):
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, page_number, page_type, type_confidence,
            quality_score, qr_status, ocr_status, extracted_text,
            status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, 0.78, ?, 'not_found', 'completed', ?, ?, ?, ?)
        """,
        (
            page_id,
            upload_id,
            page_number,
            page_type,
            quality_score,
            text,
            status,
            now,
            now,
        ),
    )


def test_confident_decision_without_any_letter_requires_review(workflow):
    _insert_upload(workflow, "decision-only")
    _insert_page(
        workflow,
        "decision-page",
        "decision-only",
        1,
        "decision",
        "completed",
    )

    changed = workflow._require_letter_for_scan_only_decisions(
        "decision-only"
    )
    page = workflow.get_page("decision-page")

    assert changed == 1
    assert page["status"] == PageStatus.NEEDS_REVIEW
    assert page["issue_code"] == "confirmed_letter_missing"


def test_confident_decision_does_not_block_packet_with_letter(workflow):
    _insert_upload(workflow, "letter-and-decision")
    _insert_page(
        workflow,
        "letter-page",
        "letter-and-decision",
        1,
        "letter",
        "needs_review",
    )
    _insert_page(
        workflow,
        "paired-decision",
        "letter-and-decision",
        2,
        "decision",
        "completed",
    )

    changed = workflow._require_letter_for_scan_only_decisions(
        "letter-and-decision"
    )

    assert changed == 0
    assert workflow.get_page("paired-decision")["status"] == (
        PageStatus.COMPLETED
    )


def test_legacy_confident_decision_is_removed_from_review(workflow):
    decision_text = """
    РЕШЕНИЕ РАЗДЕЛ I. ИНФОРМАЦИЯ О ПРОВЕРЯЕМОМ НАЛОГОПЛАТЕЛЬЩИКЕ
    102 ИНН 103 ФИО 104 налоговый орган 900 Номер принятого решения
    """
    _insert_upload(workflow, "legacy-packet")
    _insert_page(
        workflow,
        "legacy-letter",
        "legacy-packet",
        1,
        "letter",
        "completed",
    )
    _insert_page(
        workflow,
        "legacy-decision",
        "legacy-packet",
        2,
        "decision",
        "needs_review",
        text=decision_text,
    )

    changed = workflow.reconcile_confident_scan_decisions("legacy-packet")
    page = workflow.get_page("legacy-decision")

    assert changed == 1
    assert page["status"] == PageStatus.COMPLETED
    assert page["issue_code"] is None


def test_legacy_uncertain_decision_gets_explanatory_reason(workflow):
    _insert_upload(workflow, "uncertain-packet")
    _insert_page(
        workflow,
        "uncertain-letter",
        "uncertain-packet",
        1,
        "letter",
        "completed",
    )
    _insert_page(
        workflow,
        "uncertain-decision",
        "uncertain-packet",
        2,
        "decision",
        "needs_review",
        text="РЕШЕНИЕ 102 103 104 900",
    )

    changed = workflow.reconcile_confident_scan_decisions(
        "uncertain-packet"
    )
    page = workflow.get_page("uncertain-decision")

    assert changed == 1
    assert page["status"] == PageStatus.NEEDS_REVIEW
    assert page["issue_code"] == "decision_type_not_confident"
