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
    assert path == "/review"
    assert "Проблемных страниц больше нет" in unquote(query)


def test_create_all_route_reports_zero_when_nothing_is_ready(
    workflow, monkeypatch
):
    monkeypatch.setattr(main, "workflow", workflow)

    response = main.create_all_grouped_responses()

    assert response.status_code == 303
    location = response.headers["location"]
    path, _, query = location.partition("?")
    assert path == "/today"
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
