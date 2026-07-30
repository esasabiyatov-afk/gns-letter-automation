from __future__ import annotations

from dataclasses import replace

from gns_app.database import Database
from gns_app.database import utc_now
from gns_app.domain import CaseStatus, ExtractedFields, ExtractedTaxpayer
from gns_app.services.registry_service import (
    RegistryCompany,
    RegistryLookupResult,
)
from gns_app.services.workflow import WorkflowService


def test_registry_check_compares_name_and_stores_director(
    workflow,
    monkeypatch,
):
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('registry-upload', 'scan.pdf', 'scan.pdf',
                  'registry-hash', 1, 'ready', ?)
        """,
        (now,),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind,
            fields_confirmed, created_at, updated_at
        ) VALUES ('registry-case', 'registry-upload', 'needs_review',
                  'ocr_scan', 0, ?, ?)
        """,
        (now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn,
            name_source, inn_source, manually_confirmed,
            created_at, updated_at
        ) VALUES ('registry-taxpayer', 'registry-case', 1,
                  'Общество с ограниченной ответственностью "Альфа"',
                  '12345678901234', 'ocr_scan', 'ocr_scan', 0, ?, ?)
        """,
        (now, now),
    )
    company = RegistryCompany(
        inn="12345678901234",
        name='ОсОО "Альфа"',
        director="Асанов Асан",
    )
    monkeypatch.setattr(
        workflow.registry,
        "lookup_by_inn",
        lambda _inn: RegistryLookupResult(
            status="found",
            query=company.inn,
            search_mode="inn",
            matches=(company,),
            inn=company.inn,
            official_name=company.name,
            director=company.director,
        ),
    )

    summary = workflow.check_registry_case("registry-case")
    taxpayer = workflow.get_taxpayers("registry-case")[0]

    assert summary == {"match": 1}
    assert taxpayer["registry_status"] == "match"
    assert taxpayer["registry_name"] == 'ОсОО "Альфа"'
    assert taxpayer["registry_director"] == "Асанов Асан"


def test_ocr_taxpayer_is_checked_automatically(
    test_settings,
    monkeypatch,
):
    settings = replace(test_settings, auto_registry_check=True)
    database = Database(settings.database_path)
    database.initialize()
    workflow = WorkflowService(database, settings)
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('ocr-registry-upload', 'scan.pdf', 'scan.pdf',
                  'ocr-registry-hash', 1, 'ready', ?)
        """,
        (now,),
    )
    case_id = workflow._create_scan_case(
        "ocr-registry-upload",
        "ocr-registry-page",
    )
    company = RegistryCompany(
        inn="12345678901234",
        name='ОсОО "Альфа"',
        director="Асанов Асан",
    )
    calls: list[str] = []

    def fake_lookup(inn: str) -> RegistryLookupResult:
        calls.append(inn)
        return RegistryLookupResult(
            status="found",
            query=inn,
            search_mode="inn",
            matches=(company,),
            inn=inn,
            official_name=company.name,
            director=company.director,
        )

    monkeypatch.setattr(workflow.registry, "lookup_by_inn", fake_lookup)
    workflow._prefill_scan_case(
        case_id,
        ExtractedFields(
            period_start="2020-01-01",
            period_end="2026-01-01",
            taxpayers=[
                ExtractedTaxpayer(
                    name='ОсОО "Альфа"',
                    inn="12345678901234",
                    confidence=0.98,
                )
            ],
        ),
        critical_fields_agree=True,
    )

    taxpayer = workflow.get_taxpayers(case_id)[0]
    assert calls == ["12345678901234"]
    assert taxpayer["registry_status"] == "match"


def test_automatic_registry_mismatch_blocks_until_employee_accepts(
    workflow,
    monkeypatch,
):
    workflow.settings = replace(
        workflow.settings,
        auto_registry_check=True,
    )
    now = utc_now()
    workflow.db.execute(
        """
        INSERT INTO uploads(
            id, original_filename, stored_path, sha256,
            page_count, status, created_at
        ) VALUES ('mismatch-upload', 'scan.pdf', 'scan.pdf',
                  'mismatch-hash', 1, 'ready', ?)
        """,
        (now,),
    )
    workflow.db.execute(
        """
        INSERT INTO cases(
            id, upload_id, status, source_kind,
            fields_confirmed, created_at, updated_at
        ) VALUES ('mismatch-case', 'mismatch-upload', 'ready_for_abs',
                  'manual', 1, ?, ?)
        """,
        (now, now),
    )
    workflow.db.execute(
        """
        INSERT INTO taxpayers(
            id, case_id, display_order, name, inn,
            name_source, inn_source, manually_confirmed,
            created_at, updated_at
        ) VALUES ('mismatch-taxpayer', 'mismatch-case', 1,
                  'OCR Company', '12345678901234',
                  'ocr_scan', 'ocr_scan', 1, ?, ?)
        """,
        (now, now),
    )
    monkeypatch.setattr(
        workflow.registry,
        "lookup_by_inn",
        lambda inn: RegistryLookupResult(
            status="found",
            query=inn,
            search_mode="inn",
            inn=inn,
            official_name="Official Company",
            director="Director",
        ),
    )

    workflow.check_registry_case("mismatch-case", automatic=True)

    assert workflow.get_case("mismatch-case")["status"] == (
        CaseStatus.NEEDS_REVIEW
    )
    assert workflow.get_taxpayers("mismatch-case")[0]["registry_status"] == (
        "mismatch"
    )

    workflow.accept_registry_variance("mismatch-case")

    assert workflow.get_case("mismatch-case")["status"] == (
        CaseStatus.READY_FOR_ABS
    )
