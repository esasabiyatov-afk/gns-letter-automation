from __future__ import annotations

from urllib.parse import unquote

from gns_app import main
from gns_app.database import utc_now


def _insert_upload(workflow, upload_id: str, page_count: int = 2):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, 'packet.pdf', 'packet.pdf', ?, ?, 'processing', ?)
        """,
        (upload_id, f"hash-{upload_id}", page_count, utc_now()),
    )


def _insert_review_page(workflow, page_id: str, upload_id: str, page_number: int):
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, page_number, page_type, type_confidence,
            quality_score, qr_status, ocr_status, extracted_text,
            status, created_at, updated_at
        ) VALUES (?, ?, ?, 'unknown', 0.5, 0.9, 'not_found', 'completed',
                  '', 'needs_review', ?, ?)
        """,
        (page_id, upload_id, page_number, now, now),
    )


def _confirm_as_other(page_id: str):
    """Call the confirm_review route directly, bypassing FastAPI's Form
    dependency injection (which only fires for real HTTP requests), so every
    Form(...) field needs an explicit value here."""
    return main.confirm_review(
        page_id,
        page_type="other",
        district_place="",
        recipient_position="",
        recipient_full_name="",
        recipient_display_name="",
        period_start="",
        period_end="",
        period_route="",
        employee_name="",
        critical_fields_verified=False,
        taxpayer_name=[],
        taxpayer_inn=[],
    )


def test_confirming_a_page_redirects_straight_to_the_next_one(
    workflow, monkeypatch
):
    monkeypatch.setattr(main, "workflow", workflow)
    _insert_upload(workflow, "batch-upload", page_count=2)
    _insert_review_page(workflow, "page-one", "batch-upload", 1)
    _insert_review_page(workflow, "page-two", "batch-upload", 2)

    response = _confirm_as_other("page-one")

    assert response.status_code == 303
    location = response.headers["location"]
    path, _, query = location.partition("?")
    assert path == "/review/page-two"
    assert "Осталось проверить: 1" in unquote(query)
    assert workflow.get_page("page-one")["status"] != "needs_review"


def test_confirming_the_last_page_goes_to_the_empty_queue(
    workflow, monkeypatch
):
    monkeypatch.setattr(main, "workflow", workflow)
    _insert_upload(workflow, "single-upload", page_count=1)
    _insert_review_page(workflow, "only-page", "single-upload", 1)

    response = _confirm_as_other("only-page")

    assert response.status_code == 303
    location = response.headers["location"]
    path, _, query = location.partition("?")
    assert path == "/"
    assert "tab=review" in query
    assert "Проблемных страниц больше нет" in unquote(query)


def test_create_all_route_reports_zero_when_nothing_is_ready(
    workflow, monkeypatch
):
    monkeypatch.setattr(main, "workflow", workflow)

    response = main.create_all_grouped_responses()

    assert response.status_code == 303
    location = response.headers["location"]
    path, _, query = location.partition("?")
    assert path == "/"
    assert "tab=responses" in query
    assert "response_view=prepare" in query
    assert "Нет готовых групп" in unquote(query)


def test_review_page_does_not_render_duplicate_registry_button(
    workflow, monkeypatch
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)

    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('registry-dup-upload', 'f.pdf', 'f.pdf', 'hdup',
                  1, 'processing', ?)
        """,
        (now,),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind,
            fields_confirmed, created_at, updated_at
        ) VALUES ('registry-dup-case', 'registry-dup-upload',
                  'needs_review', 'ocr_scan', 0, ?, ?)
        """,
        (now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn,
            name_source, inn_source, manually_confirmed,
            registry_status, registry_name, registry_provider,
            created_at, updated_at
        ) VALUES ('registry-dup-taxpayer', 'registry-dup-case', 1,
                  'ОсОО "Черновик"', '02312201410117',
                  'ocr_scan', 'ocr_scan', 0,
                  'mismatch', 'ОсОО "Официальное"', 'ОсОО.KG', ?, ?)
        """,
        (now, now),
    )
    _insert_review_page(
        workflow, "registry-dup-page", "registry-dup-upload", 1
    )
    workflow.db.execute(
        "UPDATE pages SET case_id = ? WHERE id = ?",
        ("registry-dup-case", "registry-dup-page"),
    )

    client = TestClient(main.app)
    response = client.get("/review/registry-dup-page")

    assert response.status_code == 200
    # До фикса кнопка встречалась дважды: один раз как статичный HTML из
    # review_page.html, второй раз — как результат живой JS-проверки.
    # Статичной версии больше нет, "живая" кнопка не рендерится на
    # сервере вообще (она появляется в браузере после AJAX-запроса).
    assert "Использовать название из реестра" not in response.text
    assert 'class="registry-live-result"' in response.text


def test_conflicting_ocr_inns_render_one_taxpayer_with_choices(
    workflow, monkeypatch
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)
    now = utc_now()
    _insert_upload(workflow, "conflict-upload", page_count=1)
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind,
            fields_confirmed, created_at, updated_at
        ) VALUES ('conflict-case', 'conflict-upload',
                  'needs_review', 'ocr_scan', 0, ?, ?)
        """,
        (now, now),
    )
    workflow.db.executemany(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn,
            name_source, inn_source, manually_confirmed,
            created_at, updated_at
        ) VALUES (?, 'conflict-case', ?, ?, ?,
                  'ocr_scan', 'ocr_scan', 0, ?, ?)
        """,
        [
            ("candidate-one", 1, "", "12904195800139", now, now),
            (
                "candidate-two",
                2,
                "Мамарасулова Минохжатхон Мамиржановна",
                "12904195890139",
                now,
                now,
            ),
        ],
    )
    _insert_review_page(workflow, "conflict-page", "conflict-upload", 1)
    workflow.db.execute(
        "UPDATE pages SET case_id = 'conflict-case' WHERE id = 'conflict-page'"
    )

    response = TestClient(main.app).get("/review/conflict-page")

    assert response.status_code == 200
    assert response.text.count('name="taxpayer_name"') == 1
    assert response.text.count('data-ocr-inn-candidate') == 2
    assert "12904195800139" in response.text
    assert "12904195890139" in response.text
    assert "OCR распознал разные варианты ИНН" in response.text


def _insert_orphaned_scan_case(workflow, prefix: str) -> tuple[str, str]:
    """A scan case left behind after its source sheet was marked as other."""
    now = utc_now()
    upload_id = f"{prefix}-upload"
    case_id = f"{prefix}-case"
    page_id = f"{prefix}-page"
    _insert_upload(workflow, upload_id, page_count=1)
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, page_number, page_type, type_confidence,
            quality_score, qr_status, ocr_status, extracted_text,
            status, manual_confirmed, created_at, updated_at
        ) VALUES (?, ?, 1, 'other', 1, 0.9, 'not_found', 'completed',
                  'Текст письма', 'manually_confirmed', 1, ?, ?)
        """,
        (page_id, upload_id, now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind, district_place,
            fields_confirmed, created_at, updated_at
        ) VALUES (?, ?, 'needs_review', 'ocr_scan',
                  'по Жайыльскому району Чуйской области', 0, ?, ?)
        """,
        (case_id, upload_id, now, now),
    )
    workflow.db.executemany(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn,
            name_source, inn_source, manually_confirmed,
            created_at, updated_at
        ) VALUES (?, ?, ?, 'Ткачев Владимир Александрович', ?,
                  'ocr_scan', 'ocr_scan', 0, ?, ?)
        """,
        [
            (f"{prefix}-taxpayer-one", case_id, 1, "21393198400213", now, now),
            (f"{prefix}-taxpayer-two", case_id, 2, "21303198400213", now, now),
        ],
    )
    workflow.db.audit(
        "case",
        case_id,
        "scan_case_created",
        {"source_page_id": page_id},
    )
    return case_id, page_id


def test_orphaned_scan_case_reopens_its_source_page_for_normal_review(
    workflow, monkeypatch
):
    from starlette.testclient import TestClient

    case_id, page_id = _insert_orphaned_scan_case(workflow, "reopen")

    page = workflow.reopen_incomplete_case_review(
        case_id, actor="Тестовый сотрудник"
    )

    assert page["id"] == page_id
    assert page["case_id"] == case_id
    assert page["page_type"] == "letter"
    assert page["status"] == "needs_review"
    assert page["manual_confirmed"] == 0
    assert "Заполните и подтвердите" in page["issue_message"]
    assert workflow.manual_review_overview()["case_tasks"] == []

    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)
    review = client.get(f"/review/{page_id}")

    assert review.status_code == 200
    assert "Заполните данные письма" in review.text
    assert 'value="letter"' in review.text
    assert review.text.count('name="taxpayer_name"') == 1
    assert review.text.count('data-ocr-inn-candidate') == 2
    assert "21393198400213" in review.text
    assert "21303198400213" in review.text

    events = workflow.db.fetch_all(
        "SELECT event_type FROM audit_events WHERE entity_id = ?",
        (case_id,),
    )
    assert "case_review_reopened" in {event["event_type"] for event in events}


def test_reopen_route_uses_the_normal_review_page(workflow, monkeypatch):
    from starlette.testclient import TestClient

    case_id, page_id = _insert_orphaned_scan_case(workflow, "reopen-route")
    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)

    queue = client.get("/?tab=review")
    assert queue.status_code == 200
    assert f'action="/cases/{case_id}/reopen-review"' in queue.text
    assert f'href="/cases/{case_id}"' not in queue.text

    response = client.post(
        f"/cases/{case_id}/reopen-review",
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == f"/review/{page_id}"


def test_restart_returns_one_case_to_full_manual_review(workflow):
    case_id, page_id = _insert_orphaned_scan_case(workflow, "restart")
    now = utc_now()
    workflow.db.execute(
        """
        UPDATE cases
        SET fields_confirmed = 1, status = 'ready_for_response',
            abs_status = 'not_found', response_status = 'grouped',
            response_path = 'old-response.docx', updated_at = ?
        WHERE id = ?
        """,
        (now, case_id),
    )
    workflow.db.execute(
        """
        UPDATE taxpayers
        SET manually_confirmed = 1, abs_result = 'not_found',
            odb_result = 'not_found', registry_status = 'found',
            registry_name = 'Старые данные', updated_at = ?
        WHERE case_id = ?
        """,
        (now, case_id),
    )

    page = workflow.restart_case_manual_review(
        case_id, actor="Тестовый сотрудник"
    )

    assert page["id"] == page_id
    assert page["status"] == "needs_review"
    assert page["manual_confirmed"] == 0
    restarted = workflow.get_case(case_id)
    assert restarted["status"] == "needs_review"
    assert restarted["fields_confirmed"] == 0
    assert restarted["abs_status"] == "not_checked"
    assert restarted["response_path"] is None
    taxpayer = workflow.get_taxpayers(case_id)[0]
    assert taxpayer["manually_confirmed"] == 0
    assert taxpayer["abs_result"] is None
    assert taxpayer["odb_result"] is None
    assert taxpayer["registry_status"] is None
    events = workflow.db.fetch_all(
        "SELECT event_type FROM audit_events WHERE entity_id = ?", (case_id,)
    )
    assert "case_manual_review_restarted" in {
        event["event_type"] for event in events
    }


def test_restart_route_opens_the_same_manual_form(workflow, monkeypatch):
    from starlette.testclient import TestClient

    case_id, page_id = _insert_orphaned_scan_case(workflow, "restart-route")
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).post(
        f"/cases/{case_id}/restart-manual-review",
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"] == f"/review/{page_id}"


def test_marking_reopened_scan_case_as_other_removes_the_orphan(workflow):
    case_id, page_id = _insert_orphaned_scan_case(workflow, "cleanup")
    workflow.reopen_incomplete_case_review(case_id)

    workflow.confirm_page(page_id, page_type="other")

    assert workflow.get_case(case_id) is None
    assert workflow.db.fetch_all(
        "SELECT id FROM taxpayers WHERE case_id = ?", (case_id,)
    ) == []
    assert workflow.manual_review_overview()["case_tasks"] == []
