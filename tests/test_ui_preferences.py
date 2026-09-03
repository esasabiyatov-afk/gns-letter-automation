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


def test_abs_insecure_tls_warning_is_rendered(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main

    workflow.abs.tls_verification_disabled = True
    monkeypatch.setattr(main, "workflow", workflow)

    response = TestClient(main.app).get("/cases")

    assert response.status_code == 200
    assert "Проверка сертификата АБС отключена" in response.text
    assert "Подлинность сервера не подтверждается" in response.text


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


def test_outlook_office_test_warnings_are_rendered(workflow, monkeypatch):
    from starlette.testclient import TestClient

    from gns_app import main
    from gns_app.services.outlook_service import (
        OutlookOutgoingService,
        OutlookService,
    )

    test_settings = replace(
        workflow.settings,
        outlook_test_email="esasabiyatov@gmail.com",
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
    assert "Outlook автоматически подтверждает" in response.text
    assert "esasabiyatov@gmail.com" in response.text
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
        "по Ленинскому району г. Бишкек"
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
