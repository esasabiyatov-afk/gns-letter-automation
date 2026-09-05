import pytest

from gns_app.domain import ExtractedFields, ExtractedTaxpayer
from gns_app.services.workflow import WorkflowService, WorkflowValidationError


def test_official_gns_emails_are_loaded_by_exact_office_name(workflow):
    workflow.initialize_gns_offices()

    summary = workflow.initialize_gns_office_emails()

    assert summary["updated"] >= 55
    pervomaisky = next(
        office
        for office in workflow.list_gns_offices()
        if office["office_name"] == "УГНС по Первомайскому району"
    )
    balykchy = next(
        office
        for office in workflow.list_gns_offices()
        if office["office_name"] == "УГНС по г. Балыкчы"
    )
    assert pervomaisky["email_address"] == "004pervom@sti.gov.kg"
    assert balykchy["email_address"] == "020balykchy@sti.gov.kg"
    manas_city = next(
        office
        for office in workflow.list_gns_offices()
        if office["office_name"] == "УГНС по г. Манас"
    )
    manas_district = next(
        office
        for office in workflow.list_gns_offices()
        if office["office_name"] == "УГНС по Манасскому району"
    )
    assert manas_city["email_address"] == "048manas@sti.gov.kg"
    assert manas_district["email_address"] == "056manas@sti.gov.kg"
    assert "УГНС по г. Манас" not in summary["unmatched"]

    repeated = workflow.initialize_gns_office_emails()
    assert repeated == {"updated": 0, "unmatched": []}


def test_old_jalal_abad_city_names_match_manas_but_not_manas_district(workflow):
    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()

    current_city = workflow.match_gns_office(
        "УГНС по г. Манас Джалал-Абадской области"
    )
    old_city_d = workflow.match_gns_office(
        "УГНС по городу Джалал-Абад Джалал-Абадской области"
    )
    old_city_zh = workflow.match_gns_office(
        "УГНС по г. Жалал-Абад"
    )
    manas_district = workflow.match_gns_office(
        "УГНС по Манасскому району Таласской области"
    )

    assert current_city is not None
    assert old_city_d is not None
    assert old_city_zh is not None
    assert manas_district is not None
    assert current_city["office_name"] == "УГНС по г. Манас"
    assert old_city_d["office_key"] == current_city["office_key"]
    assert old_city_zh["office_key"] == current_city["office_key"]
    assert current_city["email_address"] == "048manas@sti.gov.kg"
    assert manas_district["office_key"] != current_city["office_key"]
    assert manas_district["email_address"] == "056manas@sti.gov.kg"


def test_osh_city_email_is_not_mixed_with_osh_region_offices(workflow):
    """Only the city УГНС may receive the city Ош address.

    A response group can use the short OCR/manual value ``по городу Ош``.
    It must resolve to the city УГНС with the required regional suffix, while
    names of specialised offices must never silently receive that address.
    """
    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()

    city_matches = [
        workflow.match_gns_office("УГНС по г. Ош"),
        workflow.match_gns_office("УГНС по городу Ош"),
        workflow.match_gns_office("по городу Ош"),
    ]

    assert all(city_matches)
    city = city_matches[0]
    assert city is not None
    assert {
        office["office_key"]
        for office in city_matches
        if office is not None
    } == {city["office_key"]}
    assert city["office_name"] == "УГНС по г. Ош"
    assert city["district_place"] == "по г.Ош Ошской области"
    assert city["email_address"] == "032oshg@sti.gov.kg"

    for other_office_text in (
        "ЦОП по г. Ош",
        "УККН по г. Ош и Южному региону",
        "УГНС по Алайскому району Ошской области",
    ):
        other = workflow.match_gns_office(other_office_text)
        assert other is not None
        assert other["office_key"] != city["office_key"]
        assert other.get("email_address") != city["email_address"]

    regional_wording = workflow.match_gns_office(
        "УГНС по городу Ош и Ошской области"
    )
    assert regional_wording is not None
    assert regional_wording["office_key"] == city["office_key"]


def test_existing_jalal_abad_case_is_migrated_to_current_city_name(workflow):
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('jalal-abad-upload', 'scan.pdf', 'scan.pdf',
                  'jalal-abad-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("jalal-abad-upload", "page-1")
    workflow.db.execute(
        "UPDATE cases SET district_place = ? WHERE id = ?",
        ("по г. Жалал-Абад Джалал-Абадской области", case_id),
    )

    workflow.initialize_gns_offices()

    assert workflow.get_case(case_id)["district_place"] == (
        "по г.Манас Джалал-Абадской области"
    )


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
    assert match["district_place"] == "по Демо-району г.Бишкек"


def test_employee_can_add_update_and_delete_office_email(workflow):
    office_id = workflow.save_gns_office(
        "по Новому району Чуйской области",
        "main@example.kg",
        "reserve1@example.kg, reserve2@example.kg",
        "Новый район, Новая налоговая",
    )

    office = next(
        item for item in workflow.list_gns_offices() if item["id"] == office_id
    )
    assert office["email_address"] == "main@example.kg"
    assert office["backup_emails"] == [
        "reserve1@example.kg",
        "reserve2@example.kg",
    ]
    assert office["aliases"] == ["Новый район", "Новая налоговая"]
    assert workflow.match_gns_office("Новая налоговая")["id"] == office_id
    assert workflow.office_delivery_email(office) == "main@example.kg"

    workflow.save_gns_office(
        "по Новому району Иссык-Кульской области",
        "",
        "reserve2@example.kg",
        "Обновлённый район",
        office_id=office_id,
    )
    updated = next(
        item for item in workflow.list_gns_offices() if item["id"] == office_id
    )
    assert updated["email_address"] is None
    assert workflow.office_delivery_email(updated) == "reserve2@example.kg"
    assert updated["aliases"] == [
        "Обновлённый район",
        "по Новому району Чуйской области",
    ]
    assert workflow.match_gns_office(
        "по Новому району Чуйской области"
    )["id"] == office_id

    workflow.delete_gns_office(office_id)
    assert all(
        item["id"] != office_id for item in workflow.list_gns_offices()
    )
    stored = workflow.db.fetch_one(
        "SELECT active, user_modified FROM gns_offices WHERE id = ?",
        (office_id,),
    )
    assert stored == {"active": 0, "user_modified": 1}


def test_office_settings_reject_duplicate_or_invalid_email(workflow):
    workflow.save_gns_office(
        "по Первому тестовому району",
        "one@example.kg",
    )

    with pytest.raises(WorkflowValidationError, match="Некорректная почта"):
        workflow.save_gns_office(
            "по Второму тестовому району",
            "не-почта",
        )
    with pytest.raises(WorkflowValidationError, match="другого района"):
        workflow.save_gns_office(
            "по Второму тестовому району",
            "two@example.kg",
            "one@example.kg",
        )

    workflow.save_gns_office(
        "по Третьему тестовому району",
        aliases="Общий алиас",
    )
    with pytest.raises(WorkflowValidationError, match="другим районом"):
        workflow.save_gns_office(
            "по Четвёртому тестовому району",
            aliases="Общий алиас",
        )


def test_user_office_changes_survive_default_directory_refresh(workflow):
    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()
    office = workflow.list_gns_offices()[0]
    workflow.save_gns_office(
        office["district_place"],
        "changed@example.kg",
        "reserve@example.kg",
        "Ручной алиас",
        office_id=office["id"],
    )
    workflow.db.execute(
        "DELETE FROM settings WHERE key IN (?, ?)",
        (
            workflow.GNS_OFFICES_SOURCE_SETTING,
            workflow.GNS_EMAILS_SOURCE_SETTING,
        ),
    )

    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()

    refreshed = next(
        item
        for item in workflow.list_gns_offices()
        if item["office_key"] == office["office_key"]
    )
    assert refreshed["email_address"] == "changed@example.kg"
    assert refreshed["backup_emails"] == ["reserve@example.kg"]
    assert refreshed["aliases"] == ["Ручной алиас"]


def test_existing_office_aliases_can_be_saved_unchanged(workflow):
    workflow.initialize_gns_offices()
    workflow.initialize_gns_office_emails()

    for office in workflow.list_gns_offices():
        workflow.save_gns_office(
            office["district_place"],
            office["email_address"] or "",
            office["backup_emails"],
            office["aliases"],
            office_id=office["id"],
        )


def test_gns_office_directory_is_read_once_for_repeated_matches(
    workflow,
    monkeypatch,
):
    workflow.replace_gns_offices(_offices())
    original_fetch_all = workflow.db.fetch_all
    directory_reads = 0

    def counted_fetch_all(sql, parameters=()):
        nonlocal directory_reads
        if "FROM gns_offices" in sql:
            directory_reads += 1
        return original_fetch_all(sql, parameters)

    monkeypatch.setattr(workflow.db, "fetch_all", counted_fetch_all)

    for _ in range(20):
        office = workflow.canonical_gns_office(
            "по Демо-району города Бишкек"
        )
        assert office is not None

    assert directory_reads == 1


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
        "по Демо-району г.Бишкек"
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
    assert case["district_place"] == "по Демо-району г.Бишкек"
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
    assert all(record["district_place"] for record in records)
    osh_city = next(
        record
        for record in records
        if record["office_name"] == "УГНС по г. Ош"
    )
    assert osh_city["district_place"] == "по г.Ош Ошской области"


def test_noisy_ocr_office_is_offered_as_suggestion(workflow):
    workflow.initialize_gns_offices()

    suggestions = workflow.suggest_gns_offices(
        "УГНС по Иссык Атинскону району, г. Кант"
    )

    assert suggestions
    assert suggestions[0]["office_name"] == (
        "УГНС по Ысык-Атинскому району"
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
        "по Ысык-Атинскому району Чуйской области"
    )


def test_saved_district_typo_is_canonicalized_from_directory(workflow):
    workflow.initialize_gns_offices()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('panfilov-upload', 'scan.pdf', 'scan.pdf',
                  'panfilov-hash', 1, 'ready',
                  '2026-01-01T00:00:00+00:00')
        """
    )
    case_id = workflow._create_scan_case("panfilov-upload", "panfilov-page")
    workflow.db.execute(
        "UPDATE cases SET district_place = ? WHERE id = ?",
        ("по Панфиловсому району Чуйской области", case_id),
    )

    result = workflow.reconcile_gns_office_districts()

    assert result == {"cases": 1, "response_groups": 0}
    assert workflow.get_case(case_id)["district_place"] == (
        "по Панфиловскому району Чуйской области"
    )
