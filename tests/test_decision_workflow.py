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
):
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, page_number, page_type, type_confidence,
            qr_status, ocr_status, status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, 0.78, 'not_found', 'completed', ?, ?, ?)
        """,
        (page_id, upload_id, page_number, page_type, status, now, now),
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
