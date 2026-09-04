from __future__ import annotations

import re
from pathlib import Path

import pytest

from gns_app import main
from gns_app.database import utc_now


def _insert_review_page(
    workflow,
    *,
    upload_id: str,
    page_id: str,
    page_number: int,
) -> None:
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, 'interface-test.pdf', 'interface-test.pdf', ?,
                  1, 'processing', ?)
        """,
        (upload_id, f"hash-{upload_id}", now),
    )
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, page_number, page_type, type_confidence,
            quality_score, qr_status, ocr_status, extracted_text,
            status, issue_message, created_at, updated_at
        ) VALUES (?, ?, ?, 'unknown', 0.4, 0.5, 'not_found',
                  'completed', '', 'needs_review',
                  'Нужно определить тип страницы.', ?, ?)
        """,
        (page_id, upload_id, page_number, now, now),
    )


def _insert_upload_workspace_pages(workflow, upload_id: str) -> None:
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, 'workspace.pdf', 'workspace.pdf', ?,
                  3, 'needs_review', ?)
        """,
        (upload_id, f"hash-{upload_id}", now),
    )
    for page_number, status, issue in (
        (1, "completed", None),
        (2, "needs_review", "Нужно проверить тип страницы."),
        (3, "technical_error", "Не удалось завершить обработку."),
    ):
        workflow.db.execute(
            """
            INSERT INTO pages(
                id, upload_id, page_number, preview_path, page_type,
                type_confidence, quality_score, qr_status, ocr_status,
                status, issue_message, created_at, updated_at
            ) VALUES (?, ?, ?, ?, 'unknown', 0.4, 0.75,
                      'not_found', 'completed', ?, ?, ?, ?)
            """,
            (
                f"upload-page-{page_number}",
                upload_id,
                page_number,
                f"page-{page_number}.jpg",
                status,
                issue,
                now,
                now,
            ),
        )


def test_work_incoming_is_compact_and_mail_first(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get("/")

    assert response.status_code == 200
    assert '<details id="employee-entry"' not in response.text
    assert 'data-testid="incoming-workspace"' in response.text
    assert 'data-testid="work-tab-incoming"' in response.text
    assert "Получить письма" in response.text
    assert "Добавить вручную" in response.text
    assert "Проверить папку" in response.text
    assert "Что делает приложение" not in response.text
    assert 'data-testid="manual-review-workspace"' not in response.text
    assert 'data-testid="responses-workspace"' not in response.text

    employee_response = TestClient(main.app).get("/?employee=1")
    assert '<details id="employee-entry"' in employee_response.text


def test_work_incoming_paginates_50_pdfs_and_keeps_sort(workflow, monkeypatch):
    from starlette.testclient import TestClient

    rows = []
    for index in range(51):
        rows.append(
            (
                f"incoming-{index:02d}",
                f"document-{index:02d}.pdf",
                f"document-{index:02d}.pdf",
                f"incoming-hash-{index:02d}",
                f"2026-09-04T08:{index:02d}:00+00:00",
            )
        )
    workflow.db.executemany(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES (?, ?, ?, ?, 1, 'completed', ?)
        """,
        rows,
    )
    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)

    newest = client.get("/?tab=incoming")
    assert newest.text.count('data-testid="incoming-row"') == 50
    assert "Всего PDF: 51" in newest.text
    assert "document-50.pdf" in newest.text
    assert "document-00.pdf" not in newest.text
    assert "1 из 2" in newest.text
    assert "incoming_page=2" in newest.text
    assert 'value="newest" selected' in newest.text

    newest_second = client.get(
        "/?tab=incoming&incoming_sort=newest&incoming_page=2"
    )
    assert newest_second.text.count('data-testid="incoming-row"') == 1
    assert "document-00.pdf" in newest_second.text
    assert "document-50.pdf" not in newest_second.text
    assert "2 из 2" in newest_second.text

    oldest = client.get("/?tab=incoming&incoming_sort=oldest")
    assert oldest.text.count('data-testid="incoming-row"') == 50
    assert "document-00.pdf" in oldest.text
    assert "document-50.pdf" not in oldest.text
    assert 'value="oldest" selected' in oldest.text
    assert "incoming_sort=oldest&amp;incoming_page=2" in oldest.text


def test_upload_page_combines_pdf_and_compact_page_list(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    upload_id = "upload-workspace"
    _insert_upload_workspace_pages(workflow, upload_id)
    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)
    script = (
        Path(main.__file__).parent / "static" / "app.js"
    ).read_text(encoding="utf-8")

    response = client.get(f"/uploads/{upload_id}")

    assert response.status_code == 200
    assert 'data-testid="upload-workspace"' in response.text
    assert 'data-testid="upload-page-pdf"' in response.text
    assert (
        'src="/pages/upload-page-2/pdf#page=1&zoom=page-fit"'
        in response.text
    )
    assert response.text.count('data-testid="upload-page-row"') == 3
    assert re.search(
        r'href="/uploads/upload-workspace\?page=2#page-viewer"'
        r'[^>]+aria-current="page"',
        response.text,
    )
    assert "Открыть превью" not in response.text
    assert "Открыть сохранённое превью" not in response.text
    assert 'target="_blank"' not in response.text
    assert "Проверить" in response.text
    assert "Повторить" in response.text
    assert 'name="return_page" value="2"' in response.text
    assert '.upload-page-row.is-selected' in script
    assert "scrollIntoView" in script

    selected = client.get(f"/uploads/{upload_id}?page=1")

    assert selected.status_code == 200
    assert (
        'src="/pages/upload-page-1/pdf#page=1&zoom=page-fit"'
        in selected.text
    )
    assert re.search(
        r'href="/uploads/upload-workspace\?page=1#page-viewer"'
        r'[^>]+aria-current="page"',
        selected.text,
    )
    assert 'http-equiv="refresh"' not in selected.text
    assert client.get(f"/uploads/{upload_id}?page=99").status_code == 404

    monkeypatch.setattr(workflow, "process_upload", lambda *_args: None)
    rejected = client.post(
        "/pages/upload-page-1/reprocess",
        follow_redirects=False,
    )
    assert rejected.status_code == 303
    assert rejected.headers["location"].startswith(
        f"/uploads/{upload_id}?page=1&error="
    )
    assert rejected.headers["location"].endswith("#page-viewer")

    reprocessed = client.post(
        "/pages/upload-page-3/reprocess",
        follow_redirects=False,
    )
    assert reprocessed.status_code == 303
    assert reprocessed.headers["location"].startswith(
        f"/uploads/{upload_id}?page=3&message="
    )
    assert reprocessed.headers["location"].endswith("#page-viewer")

    batch_reprocessed = client.post(
        f"/uploads/{upload_id}/reprocess",
        data={"return_page": "2"},
        follow_redirects=False,
    )
    assert batch_reprocessed.status_code == 303
    assert batch_reprocessed.headers["location"].startswith(
        f"/uploads/{upload_id}?page=2&message="
    )
    assert batch_reprocessed.headers["location"].endswith("#page-viewer")


def test_upload_embedded_page_pdf_is_served_inline(
    workflow,
    sample_pdf,
    monkeypatch,
):
    from starlette.testclient import TestClient

    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
    source_page = workflow.get_upload_pages(upload_id)[0]
    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)

    page_response = client.get(f"/pages/{source_page['id']}/pdf")

    assert page_response.status_code == 200
    assert page_response.headers["content-type"] == "application/pdf"
    assert page_response.headers["content-disposition"].startswith("inline;")
    assert page_response.content.startswith(b"%PDF-")


def test_unknown_review_type_has_no_letter_default(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)
    _insert_review_page(
        workflow,
        upload_id="unknown-type-upload",
        page_id="unknown-type-page",
        page_number=1,
    )

    response = TestClient(main.app).get("/review")

    assert response.status_code == 200
    assert re.search(
        r'<option value="" disabled\s+selected>Тип</option>',
        response.text,
    )
    assert ">Применить<" not in response.text
    assert ">Проверить<" in response.text


def test_quick_page_type_choice_offers_immediate_undo(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)
    _insert_review_page(
        workflow,
        upload_id="quick-undo-upload",
        page_id="quick-undo-page",
        page_number=1,
    )
    client = TestClient(main.app)

    selected = client.post(
        "/review/quick-undo-page/type",
        data={"page_type": "other"},
        follow_redirects=False,
    )

    assert selected.status_code == 303
    assert "undo_page_id=quick-undo-page" in selected.headers["location"]
    notice = client.get(selected.headers["location"])
    assert 'action="/pages/quick-undo-page/reopen-type"' in notice.text
    assert "Отменить выбор" in notice.text

    reopened = client.post(
        "/pages/quick-undo-page/reopen-type",
        follow_redirects=False,
    )
    assert reopened.status_code == 303
    assert reopened.headers["location"] == "/review/quick-undo-page"
    assert workflow.get_page("quick-undo-page")["status"] == "needs_review"


def test_review_page_shows_queue_position_and_next_page(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)
    _insert_review_page(
        workflow,
        upload_id="first-navigation-upload",
        page_id="first-navigation-page",
        page_number=1,
    )
    _insert_review_page(
        workflow,
        upload_id="second-navigation-upload",
        page_id="second-navigation-page",
        page_number=2,
    )

    response = TestClient(main.app).get("/review/first-navigation-page")

    assert response.status_code == 200
    assert "1 из 2" in response.text
    assert 'href="/review/second-navigation-page"' in response.text
    assert 'class="review-submit-bar"' in response.text
    assert "PDF · Alt+1" in response.text
    assert "zoom=page-fit" in response.text


def test_work_sections_are_server_rendered_one_at_a_time(workflow, monkeypatch):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)

    client = TestClient(main.app)
    review = client.get("/?tab=review")
    responses = client.get("/?tab=responses&response_view=prepare")

    assert review.status_code == 200
    assert 'data-testid="manual-review-workspace"' in review.text
    assert 'data-testid="incoming-workspace"' not in review.text
    assert 'data-testid="responses-workspace"' not in review.text
    assert responses.status_code == 200
    assert 'data-testid="responses-workspace"' in responses.text
    assert 'data-testid="incoming-workspace"' not in responses.text
    assert 'data-testid="manual-review-workspace"' not in responses.text
    assert 'data-workflow-panel=' not in responses.text
    assert "Параметры Word" not in responses.text


def test_shared_navigation_is_compact_and_message_is_not_duplicated(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get("/?message=Операция+выполнена")

    assert response.status_code == 200
    assert response.text.count("Операция выполнена") == 1
    assert 'aria-current="page">Работа</a>' in response.text
    assert ">История писем</a>" in response.text
    assert ">Настройки</a>" in response.text
    assert ">Документы</a>" not in response.text
    assert "Сверка ОсОО.KG" not in response.text
    assert "Все обращения" not in response.text
    assert ">Журнал</a>" not in response.text
    assert "Локальное приложение" not in response.text
    assert "app-footer" not in response.text


@pytest.mark.parametrize(
    ("url", "anchor"),
    [
        ("/?tab=incoming", "incoming"),
        ("/?tab=review", "review"),
        ("/?tab=responses&response_view=prepare", "prepare"),
        ("/?tab=responses&response_view=created", "created"),
        ("/?tab=responses&response_view=manual", "manual"),
        ("/history", "history"),
    ],
)
def test_guided_tour_has_stable_anchor_on_every_server_screen(
    workflow,
    monkeypatch,
    url,
    anchor,
):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get(url)

    assert response.status_code == 200
    assert f'data-tour-anchor="{anchor}"' in response.text
    assert 'data-tour-start' in response.text
    assert 'data-tour-card' in response.text
    assert "Пропустить" in response.text


def test_guided_tour_uses_safe_get_navigation_and_versioned_storage():
    package_dir = Path(main.__file__).parent
    script = (package_dir / "static" / "app.js").read_text(encoding="utf-8")
    template = (package_dir / "templates" / "index.html").read_text(
        encoding="utf-8"
    )

    for step_id in (
        "work-stages",
        "receive-mail",
        "review",
        "prepare",
        "created",
        "manual",
        "history",
    ):
        assert f'id: "{step_id}"' in script
    assert 'data-tour-anchor="work-tabs"' in template
    assert "window.sessionStorage" in script
    assert "window.localStorage" in script
    assert 'const storagePrefix = "gns-guided-tour-v1"' in script
    assert "window.location.assign(step.path)" in script
    assert ".requestSubmit()" not in script[: script.index("const fileInput")]
    assert ".click()" not in script[: script.index("const fileInput")]


def test_created_response_card_reveals_all_saved_taxpayers(workflow, monkeypatch):
    from html import unescape

    from starlette.testclient import TestClient

    from test_daily_batch import _insert_ready_case

    stylesheet = (Path(main.__file__).parent / "static" / "styles.css").read_text(
        encoding="utf-8"
    )
    template = (
        Path(main.__file__).parent / "templates" / "_work_responses.html"
    ).read_text(encoding="utf-8")

    assert ".response-row" in stylesheet
    assert ".created-letter-line" in stylesheet
    assert ".response-taxpayer-disclosure" in stylesheet
    assert ".response-taxpayer-list" in stylesheet
    assert 'class="response-list"' in template
    assert 'data-testid="created-response-taxpayer-disclosure"' in template
    assert 'class="action-menu"' in template
    assert "Подробнее" not in template
    assert "data-page-size" not in template

    first_name = 'ОсОО "Первое лицо готового ответа"'
    second_name = 'ОсОО "Второе лицо готового ответа"'
    _insert_ready_case(workflow, "12345678901234", first_name)
    _insert_ready_case(workflow, "23456789012345", second_name)
    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    workflow.generate_daily_response(group["group_key"], taxpayers_per_page=1)

    generated_group = workflow.today_overview(view="created")["generated_groups"][0]
    assert len(generated_group["letters"]) == 2
    taxpayer_names = [item["name"] for item in generated_group["taxpayers"]]
    assert len(taxpayer_names) == 2

    monkeypatch.setattr(main, "workflow", workflow)
    response = TestClient(main.app).get("/?tab=responses&response_view=created")

    assert response.status_code == 200
    assert response.text.count(
        'data-testid="created-response-taxpayer-disclosure"'
    ) == 1
    rendered = unescape(response.text)
    assert all(name in rendered for name in taxpayer_names)


def test_response_menu_opens_one_combined_letter_page():
    template = (
        Path(main.__file__).parent / "templates" / "_work_responses.html"
    ).read_text(encoding="utf-8")

    assert template.count('data-testid="open-case-with-document"') == 1
    assert "Открыть письмо и данные" in template
    assert "Открыть исходное письмо" not in template
    assert 'target="_blank"' not in template


def test_case_page_embeds_the_linked_letter_pdf(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from test_daily_batch import _insert_ready_case

    case_id = _insert_ready_case(
        workflow,
        "12345678901234",
        'ОсОО "Встроенный PDF"',
    )
    case = workflow.get_case(case_id)
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, case_id, page_number, page_type,
            status, created_at, updated_at
        ) VALUES ('case-decision-page', ?, ?, 1, 'decision',
                  'completed', ?, ?)
        """,
        (case["upload_id"], case_id, now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, case_id, page_number, page_type,
            status, created_at, updated_at
        ) VALUES ('case-letter-page', ?, ?, 2, 'letter',
                  'completed', ?, ?)
        """,
        (case["upload_id"], case_id, now, now),
    )
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get(
        f"/cases/{case_id}?return_to=responses"
    )

    assert response.status_code == 200
    assert 'data-testid="case-workspace"' in response.text
    assert 'data-testid="case-source-pdf"' in response.text
    assert (
        'src="/pages/case-letter-page/pdf#page=1&zoom=page-fit"'
        in response.text
    )
    assert "/pages/case-decision-page/pdf" not in response.text
    assert 'data-testid="case-sidebar"' in response.text
    assert "← Ответы" in response.text
    assert "← Проверка" not in response.text
    assert '/?tab=responses&amp;response_view=prepare' in response.text


def test_case_page_without_source_keeps_data_on_the_same_page(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    from test_daily_batch import _insert_ready_case

    case_id = _insert_ready_case(
        workflow,
        "23456789012345",
        'ОсОО "Без исходного листа"',
    )
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get(f"/cases/{case_id}")

    assert response.status_code == 200
    assert "Исходный лист не найден" in response.text
    assert 'data-testid="case-source-pdf"' not in response.text
    assert 'data-testid="case-details"' in response.text


def test_created_response_actions_are_reachable_without_card_expansion():
    package_dir = Path(main.__file__).parent
    template = (package_dir / "templates" / "_work_responses.html").read_text(encoding="utf-8")
    drawer = (package_dir / "templates" / "_response_letter_drawer.html").read_text(encoding="utf-8")
    script = (package_dir / "static" / "app.js").read_text(encoding="utf-8")

    assert 'class="created-letter-line"' in template
    assert "Исх. №" in template
    assert "Открыть письмо" in template
    assert "Сканировать" in template
    assert "Подготовить письмо в Outlook" in template
    assert "Открыть письмо повторно" in template
    assert "Скачать копию Word" in template
    assert "Подробнее" not in template
    assert "/scan-session/start" in template
    assert 'class="number-sequence-bar"' not in template
    assert "data-number-sequence-start" not in script
    assert "data-inline-number-form" in template
    assert "data-inline-number-input" in template
    assert "data-inline-number-form" in script
    assert "letter_ids" in script
    assert "Подтвердить скан" in drawer
    assert "Добавить лист" in drawer
    assert "Открыть PDF" not in drawer
    assert 'type="checkbox"' not in drawer
    assert 'class="scan-pdf-preview"' in drawer


def test_taxpayer_counter_uses_person_not_record_terminology():
    assert main.taxpayer_word(1) == "лицо"
    assert main.taxpayer_word(2) == "лица"
    assert main.taxpayer_word(5) == "лиц"


def test_identity_conflict_is_only_shown_in_manual_review(workflow, monkeypatch):
    from starlette.testclient import TestClient

    monkeypatch.setattr(main, "workflow", workflow)
    from test_daily_batch import _insert_ready_case

    inn = "12345678901234"
    _insert_ready_case(workflow, inn, 'ОсОО "Первое название"')
    _insert_ready_case(
        workflow,
        inn,
        'ОсОО "Второе название"',
        source_kind="manual",
    )
    workflow.check_abs_today("batch-user", "one-time-secret")

    client = TestClient(main.app)
    responses = client.get("/?tab=responses&response_view=prepare")
    queue = client.get("/?tab=review")
    review_id = workflow.list_case_match_reviews()[0]["id"]
    comparison = client.get(f"/review/matches/{review_id}")

    assert responses.status_code == 200
    assert "Исправить данные" not in responses.text
    assert "Не включены в ответы" in responses.text
    assert queue.status_code == 200
    assert "Это одно и то же обращение?" in queue.text
    assert "Первое название" in queue.text
    assert "Второе название" in queue.text
    assert comparison.status_code == 200
    assert comparison.text.count('name="name_choice"') == 3
    assert comparison.text.count('name="inn_choice"') == 3
    assert "Да, применить итоговые данные" in comparison.text
    assert "Нет, это разные обращения" in comparison.text
    assert "Значения не объединяются автоматически" in comparison.text
    assert comparison.text.count("Открыть карточку") == 2
    assert comparison.text.count('class="match-field-difference"') == 2
    assert comparison.text.count("Отличается") == 2
    assert comparison.text.count('class="match-resolution-difference"') == 1


def test_history_search_combines_incoming_word_number_and_scan(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    from test_daily_batch import _insert_ready_case, _png_scan

    _insert_ready_case(workflow, "12345678901234", 'ОсОО "Первый"')
    workflow.check_abs_today("batch-user", "one-time-secret")
    group = workflow.today_overview()["not_found_groups"][0]
    group_id, _ = workflow.generate_daily_response(group["group_key"])
    letter = workflow.db.fetch_one(
        "SELECT * FROM response_letters WHERE response_group_id = ?",
        (group_id,),
    )
    workflow.set_outgoing_number(letter["id"], "9544")
    scan = workflow.register_signed_response_scan(
        letter["id"], "signed.png", _png_scan()
    )
    workflow.confirm_signed_response_scan(
        scan["id"],
        correct_letter=True,
        signature_present=True,
        bank_seal_present=True,
    )
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get("/history?query=9544")

    assert response.status_code == 200
    assert "Первый" in response.text
    assert "Исх. № 9544" in response.text
    assert "Есть подписанный скан" in response.text
    assert f'href="/response-groups/{group_id}"' in response.text
    assert f'href="/signed-response-scans/{scan["id"]}"' in response.text
    assert 'name="sort_order"' in response.text
    assert "data-history-search" in response.text


def test_history_search_accepts_prefix_typo_and_sorts(workflow):
    from test_daily_batch import _insert_ready_case

    older_id = _insert_ready_case(
        workflow,
        "12345678901234",
        'ОсОО "Первый Альфа"',
        district_place="по Ленинскому району города Бишкек",
    )
    newer_id = _insert_ready_case(
        workflow,
        "22345678901234",
        'ОсОО "Второй Бета"',
        district_place="по Аламудунскому району Чуйской области",
    )
    workflow.db.execute(
        "UPDATE uploads SET created_at = '2026-09-01T08:00:00+00:00' "
        "WHERE id = (SELECT upload_id FROM cases WHERE id = ?)",
        (older_id,),
    )
    workflow.db.execute(
        "UPDATE uploads SET created_at = '2026-09-02T08:00:00+00:00' "
        "WHERE id = (SELECT upload_id FROM cases WHERE id = ?)",
        (newer_id,),
    )

    assert workflow.list_letter_history("первы")["items"][0]["id"] == older_id
    assert workflow.list_letter_history("первыи")["items"][0]["id"] == older_id
    district_items = workflow.list_letter_history(
        "", sort_order="district_asc"
    )["items"]
    date_items = workflow.list_letter_history("", sort_order="date_asc")["items"]

    assert [item["id"] for item in district_items] == [newer_id, older_id]
    assert [item["id"] for item in date_items] == [older_id, newer_id]
