from __future__ import annotations

from io import BytesIO
from pathlib import Path
from uuid import uuid4

import pytest
from docx import Document
from PIL import Image
from pypdf import PdfReader, PdfWriter

from gns_app.database import utc_now
from gns_app.domain import (
    AbsCheckResult,
    AbsStatus,
    CaseStatus,
    ExtractedFields,
    ExtractedTaxpayer,
    PageStatus,
    PageType,
    QrStatus,
)
from gns_app.services.workflow import WorkflowValidationError
from gns_app.services.outlook_service import (
    OutlookDraftResult,
    OutlookIntegrationError,
    OutlookOutgoingService,
    OutlookService,
)


def _insert_ready_case(
    workflow,
    inn: str,
    name: str,
    *,
    source_kind: str = "qr_official",
    district_place: str = "по Ленинскому району города Бишкек",
    recipient_full_name: str = "Телтаев Рахатбек Замирбекович",
    recipient_display_name: str = "Телтаеву Р. З.",
) -> str:
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
        ) VALUES (?, ?, 'ready_for_abs', ?, ?,
                  'Зам. начальника управления', ?, ?,
                  '2020-01-01', '2026-01-01',
                  'Гапарова Э.', 1, 'not_checked', ?, ?)
        """,
        (
            case_id,
            upload_id,
            source_kind,
            district_place,
            recipient_full_name,
            recipient_display_name,
            now,
            now,
        ),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn,
            name_source, inn_source, manually_confirmed,
            created_at, updated_at
        ) VALUES (?, ?, 1, ?, ?, ?, ?, 1, ?, ?)
        """,
        (
            uuid4().hex,
            case_id,
            name,
            inn,
            source_kind,
            source_kind,
            now,
            now,
        ),
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


def test_same_inn_from_qr_and_manual_case_appears_once_in_response(workflow):
    inn = "12345678901234"
    name = 'ОсОО "Один налогоплательщик"'
    _insert_ready_case(workflow, inn, name, source_kind="qr_official")
    _insert_ready_case(workflow, inn, name, source_kind="manual")

    workflow.check_abs_today("batch-user", "one-time-secret")
    groups = workflow.today_overview()["not_found_groups"]

    assert len(groups) == 1
    assert groups[0]["case_count"] == 2
    assert groups[0]["taxpayer_count"] == 1
    assert groups[0]["can_generate"]

    _, output = workflow.generate_daily_response(groups[0]["group_key"])
    document_text = "\n".join(
        paragraph.text for paragraph in Document(output).paragraphs
    )
    assert document_text.count(inn) == 1
    assert document_text.count(name) == 1


def test_same_inn_with_different_names_blocks_response(workflow):
    inn = "12345678901234"
    _insert_ready_case(
        workflow,
        inn,
        'ОсОО "Название из QR"',
        source_kind="qr_official",
    )
    _insert_ready_case(
        workflow,
        inn,
        'ОсОО "Название после ручной проверки"',
        source_kind="manual",
    )

    workflow.check_abs_today("batch-user", "one-time-secret")
    groups = workflow.today_overview()["not_found_groups"]

    assert len(groups) == 1
    assert groups[0]["taxpayer_count"] == 1
    assert not groups[0]["can_generate"]
    assert any(inn in issue for issue in groups[0]["issues"])


def test_same_inn_for_another_recipient_stays_in_separate_response(workflow):
    inn = "12345678901234"
    name = 'ОсОО "Один налогоплательщик"'
    _insert_ready_case(workflow, inn, name, source_kind="qr_official")
    _insert_ready_case(
        workflow,
        inn,
        name,
        source_kind="manual",
        recipient_full_name="Асанова Айгуль Токтогуловна",
        recipient_display_name="Асановой А. Т.",
    )

    workflow.check_abs_today("batch-user", "one-time-secret")
    groups = workflow.today_overview()["not_found_groups"]

    assert len(groups) == 2
    assert all(group["taxpayer_count"] == 1 for group in groups)
    assert all(group["can_generate"] for group in groups)


def test_abs_session_reuses_credentials_only_in_memory(workflow):
    first = _insert_ready_case(
        workflow, "12345678901234", 'ОсОО "Первый"'
    )
    second = _insert_ready_case(
        workflow, "23456789012345", 'ОсОО "Второй"'
    )

    workflow.check_abs(first, "session-user", "session-secret")
    assert workflow.abs_session_active()
    workflow.check_abs(second)

    assert workflow.get_case(second)["status"] == CaseStatus.READY_FOR_RESPONSE
    audit_text = "\n".join(event["payload_json"] for event in workflow.get_audit())
    assert "session-user" not in audit_text
    assert "session-secret" not in audit_text


def test_real_abs_credentials_are_not_reused_between_checks(workflow):
    class RealReadOnlyGateway:
        is_fake = False
        supports_session = False

        @staticmethod
        def check(username, password, taxpayers):
            return AbsCheckResult(
                status=AbsStatus.NOT_FOUND,
                taxpayers=[
                    {
                        "inn": item["inn"],
                        "name": item["name"],
                        "result": AbsStatus.NOT_FOUND,
                    }
                    for item in taxpayers
                ],
                message="Счета не найдены.",
                is_fake=False,
            )

    workflow.abs = RealReadOnlyGateway()
    first = _insert_ready_case(workflow, "12345678901234", "Первый")
    second = _insert_ready_case(workflow, "23456789012345", "Второй")

    workflow.check_abs(first, "employee", "one-time-secret")

    assert not workflow.abs_session_active()
    audit = workflow.get_audit()
    event = next(item for item in audit if item["event_type"] == "abs_checked")
    assert "employee" not in event["payload_json"]
    assert "one-time-secret" not in event["payload_json"]
    with pytest.raises(WorkflowValidationError, match="логин и пароль"):
        workflow.check_abs(second)


def test_found_taxpayer_is_separate_and_never_uses_absence_template(workflow):
    _insert_ready_case(
        workflow, "11111111111111", 'ОсОО "Найденный"'
    )

    workflow.check_abs_today("batch-user", "one-time-secret")
    overview = workflow.today_overview()

    assert not overview["not_found_groups"]
    assert len(overview["found_groups"]) == 1
    assert not overview["found_groups"][0]["can_generate"]
    assert not overview["found_groups"][0]["issues"]


def test_ocr_only_decision_stays_in_manual_review(workflow):
    status, issue_code, _ = workflow._page_outcome(
        PageType.DECISION,
        0.99,
        QrStatus.NOT_FOUND,
        False,
        None,
    )

    assert status == PageStatus.NEEDS_REVIEW
    assert issue_code == "decision_type_not_confident"


def test_confident_ocr_decision_can_finish_without_manual_fields(workflow):
    status, issue_code, issue_message = workflow._page_outcome(
        PageType.DECISION,
        0.78,
        QrStatus.NOT_FOUND,
        False,
        None,
        True,
    )

    assert status == PageStatus.COMPLETED
    assert issue_code is None
    assert issue_message is None


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


def test_generate_all_ready_daily_responses_creates_one_file_per_recipient(
    workflow,
):
    first = _insert_ready_case(
        workflow,
        "12345678901234",
        'ОсОО "Первый"',
        district_place="по Ленинскому району города Бишкек",
        recipient_full_name="Телтаев Рахатбек Замирбекович",
        recipient_display_name="Телтаеву Р. З.",
    )
    second = _insert_ready_case(
        workflow,
        "23456789012345",
        'ОсОО "Второй"',
        district_place="по Свердловскому району города Бишкек",
        recipient_full_name="Асанова Айгуль Токтогуловна",
        recipient_display_name="Асановой А. Т.",
    )

    workflow.check_abs_today("batch-user", "one-time-secret")
    overview = workflow.today_overview()
    assert len(overview["not_found_groups"]) == 2

    summary = workflow.generate_all_ready_daily_responses()

    assert len(summary["created"]) == 2
    assert summary["errors"] == []
    assert workflow.get_case(first)["status"] == CaseStatus.RESPONSE_CREATED
    assert workflow.get_case(second)["status"] == CaseStatus.RESPONSE_CREATED
    assert not workflow.today_overview()["not_found_groups"]

    # Повторный вызов не должен падать и не находит новых готовых групп.
    empty_summary = workflow.generate_all_ready_daily_responses()
    assert empty_summary == {"created": [], "errors": []}


def test_split_response_registers_one_letter_per_word_section(workflow):
    for index in range(1, 5):
        _insert_ready_case(
            workflow,
            f"{index:014d}",
            f'ОсОО "Тест {index}"',
        )

    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    group_id, _ = workflow.generate_daily_response(
        group["group_key"], taxpayers_per_page=2
    )

    letters = workflow.db.fetch_all(
        "SELECT * FROM response_letters WHERE response_group_id = ? "
        "ORDER BY letter_order",
        (group_id,),
    )
    assert [letter["letter_order"] for letter in letters] == [1, 2]
    assert [letter["taxpayer_start_order"] for letter in letters] == [1, 3]
    assert [letter["taxpayer_count"] for letter in letters] == [2, 2]
    assert all(letter["outgoing_number"] is None for letter in letters)


def test_batch_outgoing_numbers_are_previewed_assigned_and_written_to_word(
    workflow,
):
    for index in range(1, 5):
        _insert_ready_case(
            workflow,
            f"{index:014d}",
            f'ОсОО "Тест {index}"',
        )

    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    group_id, output = workflow.generate_daily_response(
        group["group_key"], taxpayers_per_page=2
    )
    preview = workflow.preview_outgoing_numbers("9544")

    assert preview["count"] == 2
    assert preview["first_number"] == "9544"
    assert preview["last_number"] == "9545"
    assert [
        item["proposed_number"] for item in preview["letters"]
    ] == ["9544", "9545"]

    summary = workflow.assign_outgoing_numbers(
        "9544",
        [item["id"] for item in preview["letters"]],
        actor="Тестовый сотрудник",
    )

    assert summary["letter_count"] == 2
    letters = workflow.db.fetch_all(
        "SELECT * FROM response_letters WHERE response_group_id = ? "
        "ORDER BY letter_order",
        (group_id,),
    )
    assert [letter["outgoing_number"] for letter in letters] == [
        "9544",
        "9545",
    ]
    document_text = "\n".join(
        paragraph.text for paragraph in Document(output).paragraphs
    )
    assert "04-1/9544" in document_text
    assert "04-1/9545" in document_text
    events = workflow.db.fetch_all(
        "SELECT * FROM audit_events "
        "WHERE event_type = 'outgoing_number_assigned' ORDER BY id"
    )
    assert len(events) == 2
    assert all(event["actor"] == "Тестовый сотрудник" for event in events)


def test_outgoing_number_correction_is_unique_and_regenerates_word(workflow):
    _insert_ready_case(workflow, "12345678901234", 'ОсОО "Первый"')
    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    group_id, output = workflow.generate_daily_response(group["group_key"])
    letter = workflow.db.fetch_one(
        "SELECT * FROM response_letters WHERE response_group_id = ?",
        (group_id,),
    )

    workflow.set_outgoing_number(letter["id"], "9544")
    workflow.set_outgoing_number(letter["id"], "9600")

    refreshed = workflow.db.fetch_one(
        "SELECT * FROM response_letters WHERE id = ?", (letter["id"],)
    )
    assert refreshed["outgoing_number"] == "9600"
    document_text = "\n".join(
        paragraph.text for paragraph in Document(output).paragraphs
    )
    assert "04-1/9600" in document_text
    assert "04-1/9544" not in document_text
    changed = workflow.db.fetch_one(
        "SELECT * FROM audit_events "
        "WHERE entity_id = ? AND event_type = 'outgoing_number_changed'",
        (letter["id"],),
    )
    assert changed is not None


def test_outgoing_number_cannot_be_reused_for_another_letter(workflow):
    _insert_ready_case(
        workflow,
        "12345678901234",
        'ОсОО "Первый"',
        recipient_full_name="Телтаев Рахатбек Замирбекович",
        recipient_display_name="Телтаеву Р. З.",
    )
    _insert_ready_case(
        workflow,
        "23456789012345",
        'ОсОО "Второй"',
        recipient_full_name="Асанова Айгуль Токтогуловна",
        recipient_display_name="Асановой А. Т.",
    )
    workflow.check_abs_today("batch-user", "one-time-secret")
    workflow.generate_all_ready_daily_responses()
    letters = workflow.list_outgoing_letters()

    workflow.set_outgoing_number(letters[0]["id"], "9544")
    with pytest.raises(WorkflowValidationError, match="уже используется"):
        workflow.set_outgoing_number(letters[1]["id"], "9544")


def test_today_page_previews_and_confirms_outgoing_number_range(
    workflow, monkeypatch
):
    from starlette.testclient import TestClient

    from gns_app import main

    _insert_ready_case(workflow, "12345678901234", 'ОсОО "Первый"')
    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    workflow.generate_daily_response(group["group_key"])
    letter = workflow.list_outgoing_letters(only_unnumbered=True)[0]
    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)

    preview = client.get("/today?outgoing_start=9544")

    assert preview.status_code == 200
    assert "Будут заняты номера 9544–9544" in preview.text
    assert "Подтвердить и вставить номера в Word" in preview.text

    assigned = client.post(
        "/today/outgoing-numbers/assign",
        data={"first_number": "9544", "letter_ids": [letter["id"]]},
        follow_redirects=False,
    )

    assert assigned.status_code == 303
    refreshed = workflow.db.fetch_one(
        "SELECT outgoing_number FROM response_letters WHERE id = ?",
        (letter["id"],),
    )
    assert refreshed["outgoing_number"] == "9544"


def test_outgoing_number_preview_rejects_range_overflow(
    workflow, monkeypatch
):
    monkeypatch.setattr(
        workflow,
        "list_outgoing_letters",
        lambda *args, **kwargs: [{"id": "one"}, {"id": "two"}],
    )

    with pytest.raises(WorkflowValidationError, match="превышает"):
        workflow.preview_outgoing_numbers("999999999")


def test_today_page_opens_generated_response_in_word(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main

    _insert_ready_case(workflow, "12345678901234", 'ОсОО "Первый"')
    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    group_id, output = workflow.generate_daily_response(group["group_key"])
    opened: list[str] = []
    focus_requests: list[dict] = []
    monkeypatch.setattr(main, "workflow", workflow)
    monkeypatch.setattr(main.os, "startfile", opened.append, raising=False)
    monkeypatch.setattr(
        main,
        "start_foreground_watcher",
        lambda **kwargs: focus_requests.append(kwargs),
    )
    client = TestClient(main.app)

    response = client.post(
        f"/response-groups/{group_id}/open-for-print",
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert opened == [str(output)]
    assert focus_requests[0]["title_parts"] == (output.stem,)
    assert focus_requests[0]["class_parts"] == ("opusapp",)


def _ready_response_letter(workflow):
    _insert_ready_case(workflow, "12345678901234", 'ОсОО "Первый"')
    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    group_id, _ = workflow.generate_daily_response(group["group_key"])
    letter = workflow.db.fetch_one(
        "SELECT * FROM response_letters WHERE response_group_id = ?",
        (group_id,),
    )
    return group_id, letter


def _png_scan() -> BytesIO:
    stream = BytesIO()
    Image.new("RGB", (1240, 1754), "white").save(stream, "PNG")
    stream.seek(0)
    return stream


def test_signed_a4_scan_is_registered_previewed_and_confirmed(workflow):
    _, letter = _ready_response_letter(workflow)

    scan = workflow.register_signed_response_scan(
        letter["id"],
        "Подписанный ответ.png",
        _png_scan(),
        actor="Тестовый сотрудник",
    )

    assert scan["status"] == "needs_confirmation"
    assert scan["source"] == "upload"
    assert len(PdfReader(scan["pdf_path"]).pages) == 1
    overview_letter = workflow.today_overview()["generated_groups"][0][
        "letters"
    ][0]
    assert overview_letter["signed_scan"]["id"] == scan["id"]
    with pytest.raises(WorkflowValidationError, match="Подтвердите"):
        workflow.confirm_signed_response_scan(
            scan["id"],
            correct_letter=True,
            signature_present=True,
            bank_seal_present=False,
        )

    workflow.confirm_signed_response_scan(
        scan["id"],
        correct_letter=True,
        signature_present=True,
        bank_seal_present=True,
        actor="Тестовый сотрудник",
    )

    confirmed = workflow.get_signed_response_scan(scan["id"])
    assert confirmed["status"] == "confirmed"
    assert confirmed["confirmed_by"] == "Тестовый сотрудник"
    assert confirmed["correct_letter_confirmed"] == 1
    assert confirmed["signature_confirmed"] == 1
    assert confirmed["bank_seal_confirmed"] == 1


def test_new_signed_scan_supersedes_previous_without_deleting_it(workflow):
    _, letter = _ready_response_letter(workflow)
    first = workflow.register_signed_response_scan(
        letter["id"], "first.png", _png_scan()
    )
    second = workflow.register_signed_response_scan(
        letter["id"], "second.png", _png_scan()
    )

    assert workflow.get_signed_response_scan(first["id"])["status"] == "superseded"
    assert workflow.get_signed_response_scan(second["id"])["status"] == "needs_confirmation"
    assert Path(first["original_path"]).is_file()


def test_signed_scan_preserves_multi_page_pdf(workflow):
    _, letter = _ready_response_letter(workflow)
    writer = PdfWriter()
    writer.add_blank_page(width=595, height=842)
    writer.add_blank_page(width=595, height=842)
    stream = BytesIO()
    writer.write(stream)
    stream.seek(0)

    scan = workflow.register_signed_response_scan(
        letter["id"], "two-pages.pdf", stream
    )

    assert scan["page_count"] == 2
    assert len(PdfReader(scan["pdf_path"]).pages) == 2
    assert scan["status"] == "needs_confirmation"


def test_device_scan_is_registered_through_wia_gateway(workflow, monkeypatch):
    _, letter = _ready_response_letter(workflow)

    def fake_acquire(destination, timeout_seconds=900):
        Image.new("RGB", (1240, 1754), "white").save(destination, "PNG")
        return destination

    monkeypatch.setattr(workflow.scanner, "acquire_a4", fake_acquire)

    scan = workflow.acquire_signed_response_scan(letter["id"])

    assert scan["source"] == "wia"
    assert scan["status"] == "needs_confirmation"


def test_signed_scan_upload_and_confirmation_routes(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main

    group_id, letter = _ready_response_letter(workflow)
    png = _png_scan().getvalue()
    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)

    uploaded = client.post(
        f"/response-letters/{letter['id']}/scan-upload",
        data={"group_id": group_id},
        files={"scan_file": ("signed.png", png, "image/png")},
        follow_redirects=False,
    )

    assert uploaded.status_code == 303
    scan = workflow.db.fetch_one(
        "SELECT * FROM signed_response_scans WHERE response_letter_id = ?",
        (letter["id"],),
    )
    preview = client.get(f"/signed-response-scans/{scan['id']}")
    assert preview.status_code == 200
    assert preview.headers["content-type"] == "application/pdf"

    incomplete = client.post(
        f"/signed-response-scans/{scan['id']}/confirm",
        data={
            "group_id": group_id,
            "correct_letter": "true",
            "signature_present": "true",
        },
        follow_redirects=False,
    )
    assert incomplete.status_code == 303
    assert workflow.get_signed_response_scan(scan["id"])["status"] == "needs_confirmation"

    confirmed = client.post(
        f"/signed-response-scans/{scan['id']}/confirm",
        data={
            "group_id": group_id,
            "correct_letter": "true",
            "signature_present": "true",
            "bank_seal_present": "true",
        },
        follow_redirects=False,
    )
    assert confirmed.status_code == 303
    assert workflow.get_signed_response_scan(scan["id"])["status"] == "confirmed"


def test_confirmed_multi_page_scan_creates_one_idempotent_outlook_draft(
    workflow,
):
    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()
    _, letter = _ready_response_letter(workflow)
    workflow.set_outgoing_number(letter["id"], "9544")

    writer = PdfWriter()
    writer.add_blank_page(width=595, height=842)
    writer.add_blank_page(width=595, height=842)
    stream = BytesIO()
    writer.write(stream)
    stream.seek(0)
    scan = workflow.register_signed_response_scan(
        letter["id"], "Подписанный ответ.pdf", stream
    )
    workflow.confirm_signed_response_scan(
        scan["id"],
        correct_letter=True,
        signature_present=True,
        bank_seal_present=True,
    )

    class DraftGateway:
        def __init__(self):
            self.requests = []

        def create_draft(self, **request):
            assert request["attachment_path"].is_file()
            self.requests.append(request)
            return OutlookDraftResult(
                entry_id="outlook-entry-1",
                recipient_email=request["recipient_email"],
                subject=request["subject"],
                attachment_name=request["attachment_name"],
                existing=len(self.requests) > 1,
            )

    gateway = DraftGateway()
    outgoing = OutlookOutgoingService(
        workflow.db,
        workflow.settings,
        OutlookService(gateway),
    )
    outgoing.update_subject_template(
        "Ответ № {outgoing_number} — {office_name}"
    )

    created = outgoing.create_draft(workflow, letter["id"])
    reopened = outgoing.create_draft(workflow, letter["id"])

    assert created["status"] == "draft_created"
    assert created["recipient_email"] == "002lenin@sti.gov.kg"
    assert created["subject"] == (
        "Ответ № 9544 — УГНС по Ленинскому району"
    )
    assert created["attachment_name"] == "Ответ_исх_9544.pdf"
    assert not created["existing_outlook_draft"]
    assert reopened["existing_outlook_draft"]
    assert len(gateway.requests) == 2
    assert gateway.requests[0]["draft_key"] == f"gns-scan-{scan['id']}"
    assert workflow.db.fetch_one(
        "SELECT COUNT(*) AS count FROM outlook_outgoing_messages"
    )["count"] == 1


def test_outlook_draft_requires_confirmed_scan(workflow):
    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()
    _, letter = _ready_response_letter(workflow)
    workflow.set_outgoing_number(letter["id"], "9545")
    workflow.register_signed_response_scan(
        letter["id"], "scan.png", _png_scan()
    )
    outgoing = OutlookOutgoingService(
        workflow.db,
        workflow.settings,
        OutlookService(object()),
    )

    with pytest.raises(OutlookIntegrationError, match="подтвердите"):
        outgoing.create_draft(workflow, letter["id"])


def test_found_abs_questionnaire_waits_for_manual_account_result(workflow):
    case_id = _insert_ready_case(
        workflow, "12345678901234", 'ОсОО "Первый"'
    )
    workflow.db.execute(
        "UPDATE taxpayers SET abs_result = 'found' WHERE case_id = ?",
        (case_id,),
    )
    workflow.db.execute(
        "UPDATE cases SET status = 'needs_review', abs_status = 'found' "
        "WHERE id = ?",
        (case_id,),
    )

    result = workflow.confirm_abs_account_taxpayer(
        case_id, "12345678901234", "not_found"
    )

    assert result["next_status"] == "ready_for_response"
    taxpayer = workflow.get_taxpayers(case_id)[0]
    assert taxpayer["abs_result"] == "found"
    assert taxpayer["abs_account_result"] == "not_found"
    assert workflow.get_case(case_id)["status"] == "ready_for_response"


def test_real_abs_account_counts_are_persisted_without_auto_confirmation(
    workflow, monkeypatch
):
    from bs4 import BeautifulSoup
    from starlette.testclient import TestClient

    from gns_app import main

    class AccountSummaryGateway:
        is_fake = False
        supports_session = False

        @staticmethod
        def check(username, password, taxpayers):
            assert username == "employee"
            assert password == "one-time-secret"
            return AbsCheckResult(
                status=AbsStatus.FOUND,
                taxpayers=[
                    {
                        "inn": taxpayers[0]["inn"],
                        "name": taxpayers[0]["name"],
                        "result": AbsStatus.FOUND,
                        "active_account_count": 2,
                        "closed_account_count": 1,
                    }
                ],
                message="Анкета найдена.",
                is_fake=False,
            )

    workflow.abs = AccountSummaryGateway()
    case_id = _insert_ready_case(
        workflow, "12345678901234", 'ОсОО "Первый"'
    )

    workflow.check_abs(case_id, "employee", "one-time-secret")

    taxpayer = workflow.get_taxpayers(case_id)[0]
    assert taxpayer["abs_active_account_count"] == 2
    assert taxpayer["abs_closed_account_count"] == 1
    assert taxpayer["abs_account_result"] is None
    assert workflow.get_case(case_id)["status"] == "needs_review"
    overview_taxpayer = workflow.today_overview()["found_groups"][0][
        "taxpayers"
    ][0]
    assert overview_taxpayer["abs_active_account_count"] == 2
    assert overview_taxpayer["abs_closed_account_count"] == 1
    monkeypatch.setattr(main, "workflow", workflow)
    response = TestClient(main.app).get("/today")
    assert response.status_code == 200
    visible_text = " ".join(
        BeautifulSoup(response.text, "html.parser")
        .get_text(" ", strip=True)
        .split()
    )
    assert "АБС прочитала счета: активных — 2, закрытых — 1." in visible_text
    assert "Есть счёт" in response.text
    assert "Счёта нет" in response.text


def test_found_abs_questionnaire_with_account_stays_manual(workflow):
    case_id = _insert_ready_case(
        workflow, "12345678901234", 'ОсОО "Первый"'
    )
    workflow.db.execute(
        "UPDATE taxpayers SET abs_result = 'found' WHERE case_id = ?",
        (case_id,),
    )
    workflow.db.execute(
        "UPDATE cases SET status = 'needs_review' WHERE id = ?",
        (case_id,),
    )

    result = workflow.confirm_abs_account_taxpayer(
        case_id, "12345678901234", "found"
    )

    assert result["next_status"] == "needs_review"


def test_office_spelling_aliases_share_one_response_group(workflow):
    workflow.initialize_gns_offices()
    first = _insert_ready_case(
        workflow,
        "12345678901234",
        'ОсОО "Первый"',
        district_place="по г. Балыкчы Иссык-Кульской области",
    )
    second = _insert_ready_case(
        workflow,
        "23456789012345",
        'ОсОО "Второй"',
        district_place="по городу Балыкчы Ысык-Кульской области",
    )
    for case_id in (first, second):
        workflow.db.execute(
            "UPDATE cases SET status = 'ready_for_response' WHERE id = ?",
            (case_id,),
        )
        workflow.db.execute(
            "UPDATE taxpayers SET abs_result = 'not_found' WHERE case_id = ?",
            (case_id,),
        )

    groups = workflow.today_overview()["not_found_groups"]

    assert len(groups) == 1
    assert groups[0]["taxpayer_count"] == 2
    assert groups[0]["gns_office_key"]
