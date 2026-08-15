from gns_app.text_cleanup import clean_taxpayer_name


def test_review_preferences_are_simple_by_default(workflow):
    assert workflow.get_ui_preferences() == {
        "allow_multiple_taxpayers": False,
        "show_recipient_salutation": True,
        "require_review_checkbox": False,
    }


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


def test_long_legal_form_is_abbreviated_without_changing_name():
    assert clean_taxpayer_name(
        'Общество с ограниченной ответственностью "Жер Компани"'
    ) == 'ОсОО "Жер Компани"'
