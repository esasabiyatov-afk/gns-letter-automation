from gns_app.database import utc_now
from gns_app.domain import PageStatus, VisualPageEvidence


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


def test_decision_can_be_removed_directly_from_review_queue(workflow):
    _insert_upload(workflow, "queue-decision")
    _insert_page(
        workflow,
        "queue-decision-page",
        "queue-decision",
        1,
        "unknown",
        "needs_review",
    )
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO cases(id, upload_id, status, source_kind,
                          fields_confirmed, created_at, updated_at)
        VALUES ('queue-draft-case', 'queue-decision', 'needs_review',
                'ocr_scan', 0, ?, ?)
        """,
        (now, now),
    )
    workflow.db.execute(
        "UPDATE pages SET case_id = 'queue-draft-case' "
        "WHERE id = 'queue-decision-page'"
    )

    removed = workflow.mark_page_type_from_queue(
        "queue-decision-page", "decision"
    )

    assert removed
    page = workflow.get_page("queue-decision-page")
    assert page["page_type"] == "decision"
    assert page["status"] == PageStatus.MANUALLY_CONFIRMED
    assert page["case_id"] is None
    assert workflow.get_case("queue-draft-case") is None


def test_manually_selected_non_letter_can_return_to_review(workflow):
    _insert_upload(workflow, "undo-page-type")
    _insert_page(
        workflow,
        "undo-page-type-page",
        "undo-page-type",
        1,
        "unknown",
        "needs_review",
    )
    workflow.mark_page_type_from_queue(
        "undo-page-type-page", "attachment"
    )

    reopened = workflow.reopen_page_type_review("undo-page-type-page")

    assert reopened["page_type"] == "unknown"
    assert reopened["status"] == PageStatus.NEEDS_REVIEW
    assert not reopened["manual_confirmed"]
    assert reopened["issue_code"] == "page_type_reopened"
    assert workflow.get_upload("undo-page-type")["status"] == "needs_review"
    assert "undo-page-type-page" in {
        page["id"] for page in workflow.list_review_pages()
    }


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


def test_confirmed_official_qr_counts_as_letter_without_scan_ocr(workflow):
    upload_id = "official-letter-and-decision"
    _insert_upload(workflow, upload_id)
    _insert_page(
        workflow,
        "official-unknown-page",
        upload_id,
        1,
        "unknown",
        "completed",
    )
    _insert_page(
        workflow,
        "official-paired-decision",
        upload_id,
        2,
        "decision",
        "completed",
    )
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind, official_document_path,
            fields_confirmed, created_at, updated_at
        ) VALUES (?, ?, 'ready_for_abs', 'qr_official', ?, 1, ?, ?)
        """,
        ("official-case", upload_id, "official.pdf", now, now),
    )
    workflow.db.execute(
        """
        UPDATE pages SET case_id = ?, qr_status = 'found',
            ocr_status = 'skipped_official'
        WHERE id = ?
        """,
        ("official-case", "official-unknown-page"),
    )

    changed = workflow._require_letter_for_scan_only_decisions(upload_id)

    assert changed == 0
    assert workflow.get_page("official-paired-decision")["status"] == (
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


def test_completed_scan_decision_is_reopened_if_rule_is_not_safe(workflow):
    _insert_upload(workflow, "unsafe-completed-packet")
    _insert_page(
        workflow,
        "safe-letter",
        "unsafe-completed-packet",
        1,
        "letter",
        "completed",
    )
    _insert_page(
        workflow,
        "unsafe-completed-decision",
        "unsafe-completed-packet",
        2,
        "decision",
        "completed",
        text=(
            "STI-010 РЕШЕНИЕ "
            "О ПРЕДОСТАВЛЕНИИ ИНФОРМАЦИИ ОБ ОПЕРАЦИЯХ"
        ),
    )

    changed = workflow.reconcile_confident_scan_decisions(
        "unsafe-completed-packet"
    )
    page = workflow.get_page("unsafe-completed-decision")

    assert changed == 1
    assert page["status"] == PageStatus.NEEDS_REVIEW
    assert page["issue_code"] == "decision_type_not_confident"


def test_unknown_visual_decision_is_removed_from_review(
    workflow, monkeypatch
):
    decision_text = """
    РЕШЕНИЕ
    О ПРЕДОСТАВЛЕНИИ ИНФОРМАЦИИ ОБ ОПЕРАЦИЯХ
    ПРИНЯТО РЕШЕНИЕ О ПРЕДОСТАВЛЕНИИ
    Период:
    """
    _insert_upload(workflow, "visual-packet")
    _insert_page(
        workflow,
        "visual-letter",
        "visual-packet",
        1,
        "letter",
        "completed",
    )
    _insert_page(
        workflow,
        "visual-decision",
        "visual-packet",
        2,
        "unknown",
        "needs_review",
        text=decision_text,
    )
    monkeypatch.setattr(
        workflow,
        "_visual_evidence_for_page",
        lambda page: VisualPageEvidence(
            decision_layout=True,
            confidence=0.95,
            horizontal_line_groups=24,
            vertical_line_groups=16,
        ),
    )

    changed = workflow.reconcile_confident_scan_decisions("visual-packet")
    page = workflow.get_page("visual-decision")

    assert changed == 1
    assert page["page_type"] == "decision"
    assert page["status"] == PageStatus.COMPLETED
    assert page["issue_code"] is None


def test_header_ocr_evidence_is_persisted_for_rotated_decision(
    workflow, monkeypatch, tmp_path
):
    _insert_upload(workflow, "rotated-packet")
    _insert_page(
        workflow,
        "rotated-letter",
        "rotated-packet",
        1,
        "letter",
        "completed",
    )
    _insert_page(
        workflow,
        "rotated-decision",
        "rotated-packet",
        2,
        "unknown",
        "needs_review",
        text="",
    )
    preview = tmp_path / "rotated.jpg"
    preview.write_bytes(b"preview")
    monkeypatch.setattr(
        workflow,
        "_visual_evidence_for_page",
        lambda page: VisualPageEvidence(
            decision_layout=True,
            confidence=0.95,
            horizontal_line_groups=20,
            vertical_line_groups=8,
        ),
    )
    monkeypatch.setattr(
        workflow,
        "_preview_path_for_page",
        lambda page: preview,
    )
    monkeypatch.setattr(
        workflow.ocr,
        "recognize_type_markers",
        lambda path: (
            "РЕШЕНИЕ. Проводимых на счетах организаций"
        ),
    )

    changed = workflow.reconcile_confident_scan_decisions("rotated-packet")
    page = workflow.get_page("rotated-decision")

    assert changed == 1
    assert page["page_type"] == "decision"
    assert page["status"] == PageStatus.COMPLETED
    assert page["type_evidence_method"] == "local_header_ocr"
    assert "РЕШЕНИЕ" in page["type_evidence_text"]
