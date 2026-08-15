from datetime import datetime

from gns_app.domain import AbsStatus, CaseStatus, OdbStatus


def test_manual_odb_route_works_without_fabricated_dates(workflow):
    timestamp = "2026-01-01T00:00:00+00:00"
    workflow.db.execute(
        """
        INSERT INTO uploads(id, original_filename, stored_path, sha256,
                            page_count, status, created_at)
        VALUES ('route-upload', 'route.pdf', 'route.pdf', 'route-hash',
                1, 'ready', ?)
        """,
        (timestamp,),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(id, upload_id, status, source_kind, period_route,
                          employee_name, fields_confirmed, created_at, updated_at)
        VALUES ('route-case', 'route-upload', 'ready_for_abs', 'manual', 'odb',
                'Исполнитель', 1, ?, ?)
        """,
        (timestamp, timestamp),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(id, case_id, display_order, name, inn,
                              name_source, inn_source, manually_confirmed,
                              created_at, updated_at)
        VALUES ('route-taxpayer', 'route-case', 1, 'ОсОО Тест',
                '99999999999999', 'manual', 'manual', 1, ?, ?)
        """,
        (timestamp, timestamp),
    )

    workflow.check_abs("route-case", "user", "password")

    assert workflow.get_case("route-case")["status"] == CaseStatus.MANUAL_PERIOD_RULE
    assert workflow.get_case("route-case")["period_start"] is None

    saved = workflow.confirm_odb_taxpayer(
        "route-case",
        "99999999999999",
        OdbStatus.NOT_FOUND,
    )
    assert saved["remaining"] == 0
    assert saved["next_status"] == CaseStatus.READY_FOR_RESPONSE
    assert not saved["manual_response"]


def test_period_before_2019_goes_to_manual_rule(workflow):
    timestamp = "2026-01-01T00:00:00+00:00"
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('old-upload', 'old.pdf', 'old.pdf', 'old-hash',
                  1, 'ready', ?)
        """,
        (timestamp,),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind, period_start, period_end,
            employee_name, fields_confirmed, created_at, updated_at
        ) VALUES ('old-case', 'old-upload', 'ready_for_abs', 'manual',
                  '2018-12-31', '2020-01-01', 'Исполнитель', 1, ?, ?)
        """,
        (timestamp, timestamp),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn, name_source, inn_source,
            manually_confirmed, created_at, updated_at
        ) VALUES ('old-taxpayer', 'old-case', 1, 'ОсОО Тест',
                  '99999999999999', 'manual', 'manual', 1, ?, ?)
        """,
        (timestamp, timestamp),
    )

    result = workflow.check_abs("old-case", "user", "one-time-password")

    assert result.status == AbsStatus.NOT_FOUND
    assert workflow.get_case("old-case")["status"] == (
        CaseStatus.MANUAL_PERIOD_RULE
    )
    assert workflow.get_taxpayers("old-case")[0]["odb_result"] is None
    overview = workflow.today_overview()
    assert overview["odb_pending"] == 1
    assert [item["id"] for item in overview["odb_cases"]] == ["old-case"]
    assert overview["odb_cases"][0]["taxpayers"] == [
        {
            "case_id": "old-case",
            "name": "ОсОО Тест",
            "inn": "99999999999999",
            "odb_result": None,
        }
    ]

    next_status = workflow.confirm_odb(
        "old-case",
        [
            {
                "inn": "99999999999999",
                "result": OdbStatus.NOT_FOUND,
            }
        ],
    )

    assert next_status == CaseStatus.READY_FOR_RESPONSE
    assert workflow.get_taxpayers("old-case")[0]["odb_result"] == (
        OdbStatus.NOT_FOUND
    )
    overview = workflow.today_overview()
    assert overview["odb_pending"] == 0
    assert not overview["odb_cases"]


def test_old_period_odb_match_appears_only_after_odb_is_complete(workflow):
    business_date = datetime.now(workflow.BUSINESS_TIMEZONE).date()
    timestamp = workflow._business_day_bounds(business_date)[0]
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('odb-upload', 'odb.pdf', 'odb.pdf', 'odb-hash',
                  1, 'ready', ?)
        """,
        (timestamp,),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind, period_start, period_end,
            employee_name, fields_confirmed, created_at, updated_at
        ) VALUES ('odb-case', 'odb-upload', 'ready_for_abs', 'qr_official',
                  '2018-01-01', '2020-01-01', 'Employee', 1, ?, ?)
        """,
        (timestamp, timestamp),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn, name_source, inn_source,
            manually_confirmed, created_at, updated_at
        ) VALUES ('odb-taxpayer', 'odb-case', 1, 'Company',
                  '99999999999999', 'qr', 'qr', 1, ?, ?)
        """,
        (timestamp, timestamp),
    )

    workflow.check_abs("odb-case", "user", "password")
    assert not workflow.today_overview()["found_groups"]

    saved = workflow.confirm_odb_taxpayer(
        "odb-case",
        "99999999999999",
        OdbStatus.FOUND,
    )
    assert saved["manual_response"]
    assert saved["remaining"] == 0
    overview = workflow.today_overview()

    assert overview["odb_pending"] == 0
    assert len(overview["found_groups"]) == 1
    assert overview["found_groups"][0]["taxpayer_count"] == 1
