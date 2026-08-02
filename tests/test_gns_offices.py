from gns_app.domain import ExtractedFields


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
