from gns_app.domain import AbsStatus, CaseStatus


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
