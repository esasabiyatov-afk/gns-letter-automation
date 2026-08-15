from gns_app.domain import ExtractedFields, ExtractedTaxpayer
from gns_app.services.workflow import WorkflowService


def _offices():
    return [
        {
            "office_name": "УГНС по Демо-району",
            "district_place": "по Демо-району города Бишкек",
            "postal_address": "720000, г. Бишкек, ул. Примерная, 1",
            "aliases": ["Демо райондук салык башкармалыгы"],
        },
        {
            "office_name": "УГНС по Тестовому району",
            "district_place": "по Тестовому району Чуйской области",
            "postal_address": "720001, Чуйская область, с. Тестовое",
            "aliases": [],
        },
    ]


def test_gns_offices_are_stored_and_matched_locally(workflow):
    assert workflow.replace_gns_offices(_offices()) == 2

    offices = workflow.list_gns_offices()
    match = workflow.match_gns_office(
        "УГНС по Демо-району, адрес: ул. Примерная, 1"
    )

    assert len(offices) == 2
    assert match is not None
    assert match["district_place"] == "по Демо-району города Бишкек"


def test_replacing_directory_expands_legacy_case_location(workflow):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('legacy-office-upload', 'scan.pdf', 'scan.pdf',
                  'legacy-office-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("legacy-office-upload", "page-1")
    workflow.db.execute(
        "UPDATE cases SET district_place = ? WHERE id = ?",
        ("по Демо-району", case_id),
    )

    workflow.replace_gns_offices(_offices())

    assert workflow.get_case(case_id)["district_place"] == (
        "по Демо-району города Бишкек"
    )


def test_ambiguous_office_text_is_not_guessed(workflow):
    workflow.replace_gns_offices(_offices())

    match = workflow.match_gns_office(
        "УГНС по Демо-району и УГНС по Тестовому району"
    )

    assert match is None


def test_exact_office_match_prefills_only_unconfirmed_scan_case(workflow):
    workflow.replace_gns_offices(_offices())
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('office-upload', 'scan.pdf', 'scan.pdf',
                  'office-hash', 1, 'ready', '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("office-upload", "page-1")
    suggestion = workflow.match_gns_office("УГНС по Демо-району")

    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(),
        critical_fields_agree=False,
        office_suggestion=suggestion,
    )

    case = workflow.get_case(case_id)
    assert case["district_place"] == "по Демо-району города Бишкек"
    assert not case["fields_confirmed"]


def test_two_ocr_passes_prefill_recipient_only_as_unconfirmed_hint(workflow):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('recipient-upload', 'scan.pdf', 'scan.pdf',
                  'recipient-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("recipient-upload", "page-1")

    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(
            recipient_position="Зам. начальника управления",
            recipient_full_name="Омошев Максат Тологонович",
        ),
        critical_fields_agree=False,
        recipient_fields_agree=True,
    )

    case = workflow.get_case(case_id)
    assert case["recipient_position"] == "Зам. начальника управления"
    assert case["recipient_full_name"] == "Омошев Максат Тологонович"
    assert not case["fields_confirmed"]


def test_structured_ocr_recipient_is_visible_without_model_agreement(
    workflow,
):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('recipient-hint-upload', 'scan.pdf', 'scan.pdf',
                  'recipient-hint-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case(
        "recipient-hint-upload", "page-1"
    )

    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(
            recipient_position="Зам. начальника управления",
            recipient_full_name="Омошгсав Максат Тологонович",
        ),
        recipient_fields_agree=False,
        allow_ocr_suggestions=True,
    )

    case = workflow.get_case(case_id)
    assert case["recipient_position"] == "Зам. начальника управления"
    assert case["recipient_full_name"] == (
        "Омошгсав Максат Тологонович"
    )
    assert not case["fields_confirmed"]


def test_good_ocr_prefills_period_and_taxpayer_as_unconfirmed_hints(workflow):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('good-ocr-upload', 'scan.pdf', 'scan.pdf',
                  'good-ocr-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("good-ocr-upload", "page-1")

    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(
            period_start="2018-01-01",
            period_end="2026-01-01",
            taxpayers=[
                ExtractedTaxpayer(
                    name="ИП Иванов Иван",
                    inn="20101019900011",
                    confidence=0.5,
                )
            ],
        ),
        critical_fields_agree=False,
        allow_ocr_suggestions=True,
    )

    case = workflow.get_case(case_id)
    taxpayers = workflow.get_taxpayers(case_id)
    assert case["period_start"] == "2018-01-01"
    assert case["period_end"] == "2026-01-01"
    assert not case["fields_confirmed"]
    assert taxpayers[0]["name"] == "ИП Иванов Иван"
    assert taxpayers[0]["inn"] == "20101019900011"
    assert not taxpayers[0]["manually_confirmed"]


def test_low_confidence_ocr_does_not_prefill_critical_hints(workflow):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('weak-ocr-upload', 'scan.pdf', 'scan.pdf',
                  'weak-ocr-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("weak-ocr-upload", "page-1")

    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(
            period_start="2018-01-01",
            period_end="2026-01-01",
            taxpayers=[
                ExtractedTaxpayer(
                    name="Сомнительное имя",
                    inn="20101019900011",
                    confidence=0.2,
                )
            ],
        ),
        critical_fields_agree=False,
        allow_ocr_suggestions=False,
    )

    case = workflow.get_case(case_id)
    assert case["period_start"] is None
    assert case["period_end"] is None
    assert workflow.get_taxpayers(case_id) == []


def test_ocr_does_not_duplicate_one_inn_with_conflicting_names(workflow):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('duplicate-ocr-upload', 'scan.pdf', 'scan.pdf',
                  'duplicate-ocr-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case(
        "duplicate-ocr-upload", "page-1"
    )

    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(
            taxpayers=[
                ExtractedTaxpayer(
                    name="Мамарасулова Мшюжатхон Мамиржаиовна",
                    inn="12904195800139",
                    confidence=0.5,
                ),
                ExtractedTaxpayer(
                    name="Мамарасулова Миножатхон Мамиржановна",
                    inn="12904195800139",
                    confidence=0.5,
                ),
            ]
        ),
        allow_ocr_suggestions=True,
    )

    taxpayers = workflow.get_taxpayers(case_id)
    assert len(taxpayers) == 1
    assert taxpayers[0]["inn"] == "12904195800139"
    assert taxpayers[0]["name"] == ""
    assert not taxpayers[0]["manually_confirmed"]


def test_bundled_office_csv_contains_all_records(project_root):
    records = WorkflowService.read_gns_offices_csv(
        project_root / "src" / "gns_app" / "data" / "ugns_addresses.csv"
    )

    assert len(records) == 64
    assert records[0]["office_name"] == "УГНС по Октябрьскому району"
    assert records[0]["district_place"] == (
        "по Октябрьскому району г. Бишкек"
    )
    assert records[0]["postal_address"] == (
        "г. Бишкек, 10-й микрорайон, 29"
    )
    assert all(
        "области" in record["district_place"]
        or "г. Бишкек" in record["district_place"]
        for record in records
    )


def test_noisy_ocr_office_is_offered_as_suggestion(workflow):
    workflow.initialize_gns_offices()

    suggestions = workflow.suggest_gns_offices(
        "УГНС по Иссык Атинскону району, г. Кант"
    )

    assert suggestions
    assert suggestions[0]["office_name"] == (
        "УГНС по Иссык-Атинскому району"
    )
    assert len(suggestions) == 1


def test_generic_city_words_do_not_offer_unrelated_offices(workflow):
    workflow.initialize_gns_offices()

    assert workflow.suggest_gns_offices(
        "Кыргызская Республика, город Бишкек, управление"
    ) == []


def test_unique_fuzzy_ocr_match_replaces_unconfirmed_raw_location(workflow):
    workflow.initialize_gns_offices()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('fuzzy-office-upload', 'scan.pdf', 'scan.pdf',
                  'fuzzy-office-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("fuzzy-office-upload", "fuzzy-page")
    workflow.db.execute(
        "UPDATE cases SET district_place = ? WHERE id = ?",
        ("по Иссык Атинскону району", case_id),
    )
    workflow.db.execute(
        """
        INSERT INTO pages(
            id, upload_id, case_id, page_number, page_type,
            extracted_text, status, created_at, updated_at
        ) VALUES ('fuzzy-page', 'fuzzy-office-upload', ?, 1, 'letter',
                  'УГНС по Иссык Атинскону району, г. Кант', 'needs_review',
                  '2026-01-01T00:00:00+00:00',
                  '2026-01-01T00:00:00+00:00')
        """,
        (case_id,),
    )

    assert workflow.reconcile_gns_office_hints() == 1
    assert workflow.get_case(case_id)["district_place"] == (
        "по Иссык-Атинскому району Чуйской области"
    )
