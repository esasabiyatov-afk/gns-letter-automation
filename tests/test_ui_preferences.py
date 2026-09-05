from dataclasses import replace

from gns_app.text_cleanup import clean_taxpayer_name


def test_review_preferences_are_simple_by_default(workflow):
    assert workflow.get_ui_preferences() == {
        "allow_multiple_taxpayers": False,
        "show_recipient_salutation": True,
        "require_review_checkbox": False,
    }


def test_abs_insecure_tls_state_is_exposed_for_global_warning(workflow):
    workflow.abs.tls_verification_disabled = True

    assert workflow.abs_tls_verification_disabled()


def test_abs_insecure_tls_warning_is_not_rendered(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main

    workflow.abs.tls_verification_disabled = True
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get("/cases")

    assert response.status_code == 200
    assert "Проверка сертификата АБС отключена" not in response.text
    assert "Подлинность сервера не подтверждается" not in response.text


def test_real_tolubay_mode_is_named_correctly_in_login_dialog(
    workflow,
    monkeypatch,
):
    from starlette.testclient import TestClient

    from gns_app import main

    workflow.abs.is_fake = False
    overview = workflow.today_overview(view="prepare")
    overview["ready_abs"] = 1
    monkeypatch.setattr(
        workflow,
        "today_overview",
        lambda *, view="all": overview,
    )
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get(
        "/?tab=responses&response_view=prepare"
    )

    assert response.status_code == 200
    assert "АБС TOLUBAY" in response.text
    assert "ТЕСТОВАЯ АБС" not in response.text


def test_abs_login_dialog_auto_opens_after_session_error(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main

    overview = workflow.today_overview(view="prepare")
    overview["ready_abs"] = 1
    monkeypatch.setattr(
        workflow,
        "today_overview",
        lambda *, view="all": overview,
    )
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get(
        "/?tab=responses&response_view=prepare&abs_login=1"
    )

    assert response.status_code == 200
    assert 'id="today-abs-dialog" class="modal"' in response.text
    assert "data-auto-open" in response.text
    assert "Войдите один раз" in response.text


def test_outlook_test_mode_is_shown_without_certificate_warning(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main
    from gns_app.services.outlook_service import (
        OutlookOutgoingService,
        OutlookService,
    )

    test_settings = replace(
        workflow.settings,
        outlook_test_email="esensabiyatov@gmail.com",
        outlook_allow_test_send=True,
        outlook_allow_insecure_certificate=True,
    )
    monkeypatch.setattr(main, "workflow", workflow)
    monkeypatch.setattr(main, "settings", test_settings)
    monkeypatch.setattr(
        main,
        "outlook_outgoing",
        OutlookOutgoingService(
            workflow.db,
            test_settings,
            OutlookService(object()),
        ),
    )

    response = TestClient(main.app).get("/cases")

    assert response.status_code == 200
    assert "Outlook автоматически подтверждает" not in response.text
    assert "esensabiyatov@gmail.com" in response.text
    assert "отправка разрешена только" in response.text


def test_confirmed_recipient_is_suggested_by_partial_name(workflow):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('recipient-history-upload', 'scan.pdf', 'scan.pdf',
                  'recipient-history-hash', 1, 'ready',
                  '2026-08-04T00:00:00+00:00')
        """
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, district_place, recipient_position,
            recipient_full_name, recipient_display_name, fields_confirmed,
            created_at, updated_at
        ) VALUES ('recipient-history-case', 'recipient-history-upload',
                  'ready_for_abs', 'по Ленинскому району г. Бишкек',
                  'Зам. начальника управления',
                  'Мураканов Улан Муратович', 'Мураканову У. М.', 1,
                  '2026-08-04T00:00:00+00:00',
                  '2026-08-04T00:00:00+00:00')
        """
    )

    suggestions = workflow.recipient_suggestions("мура")

    assert suggestions[0]["full_name"] == "Мураканов Улан Муратович"
    assert suggestions[0]["district_place"] == (
        "по Ленинскому району г.Бишкек"
    )


def test_review_preferences_are_persisted(workflow):
    workflow.update_ui_preferences(
        allow_multiple_taxpayers=True,
        show_recipient_salutation=True,
        require_review_checkbox=True,
    )

    assert all(workflow.get_ui_preferences().values())


def test_period_threshold_and_inbox_path_are_persisted(workflow, tmp_path):
    inbox = tmp_path / "incoming-pdf"

    workflow.update_operational_settings(
        period_threshold="2020-01-01",
        inbox_dir=str(inbox),
    )

    assert workflow.get_period_threshold().isoformat() == "2020-01-01"
    assert workflow.get_inbox_dir() == inbox.resolve()
    assert inbox.is_dir()


def test_scanner_settings_are_persisted_and_rendered(
    workflow,
    tmp_path,
    monkeypatch,
):
    from starlette.testclient import TestClient

    from gns_app import main

    assert workflow.get_scanner_settings() == {
        "dpi": 150,
        "color_mode": "grayscale",
    }
    workflow.update_operational_settings(
        period_threshold="2020-01-01",
        inbox_dir=str(tmp_path / "incoming-pdf"),
        scanner_dpi=300,
        scanner_color_mode="color",
    )
    assert workflow.get_scanner_settings() == {
        "dpi": 300,
        "color_mode": "color",
    }

    monkeypatch.setattr(main, "workflow", workflow)
    response = TestClient(main.app).get("/settings")

    assert response.status_code == 200
    assert 'data-testid="scanner-settings"' in response.text
    assert '<option value="300" selected>' in response.text
    assert 'value="color"' in response.text


def test_office_directory_is_editable_from_settings(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main

    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()
    monkeypatch.setattr(main, "workflow", workflow)
    client = TestClient(main.app)

    page = client.get("/settings#offices")
    assert page.status_code == 200
    assert 'data-settings-tab="offices"' in page.text
    assert 'data-office-filter' in page.text
    assert "Резервные почты" in page.text
    assert 'name="aliases"' in page.text
    assert "УГНС по Иссык-Атинскому району" in page.text

    created = client.post(
        "/settings/offices/save",
        data={
            "district_place": "по Новому району Нарынской области",
            "email_address": "new@example.kg",
            "backup_emails": "reserve@example.kg",
            "aliases": "Новый район, Новая налоговая",
        },
        follow_redirects=False,
    )
    assert created.status_code == 303
    assert created.headers["location"].endswith("#offices")
    office = next(
        item
        for item in workflow.list_gns_offices()
        if item["email_address"] == "new@example.kg"
    )
    assert office["aliases"] == ["Новый район", "Новая налоговая"]

    deleted = client.post(
        f"/settings/offices/{office['id']}/delete",
        follow_redirects=False,
    )
    assert deleted.status_code == 303
    assert all(
        item["id"] != office["id"] for item in workflow.list_gns_offices()
    )


def test_registry_priority_defaults_to_osoo(workflow):
    assert workflow.get_registry_priority() == "osoo"


def test_registry_priority_is_persisted(workflow, tmp_path):
    workflow.update_operational_settings(
        period_threshold="2020-01-01",
        inbox_dir=str(tmp_path / "incoming-pdf"),
        registry_priority="reestr_kg",
    )

    assert workflow.get_registry_priority() == "reestr_kg"


def test_processing_worker_count_is_persisted_and_bounded(workflow, tmp_path):
    workflow.update_operational_settings(
        period_threshold="2020-01-01",
        inbox_dir=str(tmp_path / "incoming-pdf"),
        processing_workers=1,
    )
    assert workflow.get_processing_workers() == 1

    workflow.update_operational_settings(
        period_threshold="2020-01-01",
        inbox_dir=str(tmp_path / "incoming-pdf"),
        processing_workers=8,
    )
    assert workflow.get_processing_workers() == 2


def test_registry_priority_falls_back_to_osoo_for_unknown_value(
    workflow, tmp_path
):
    workflow.update_operational_settings(
        period_threshold="2020-01-01",
        inbox_dir=str(tmp_path / "incoming-pdf"),
        registry_priority="not_a_real_registry",
    )

    assert workflow.get_registry_priority() == "osoo"


def test_long_legal_form_is_abbreviated_without_changing_name():
    assert clean_taxpayer_name(
        'Общество с ограниченной ответственностью "Жер Компани"'
    ) == 'ОсОО "Жер Компани"'
