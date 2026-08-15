from __future__ import annotations

from gns_app import main
from gns_app.services.registry_service import RegistryLookupResult


def test_live_suggestion_returns_official_name_and_director(monkeypatch):
    calls: list[str] = []

    def lookup(inn: str) -> RegistryLookupResult:
        calls.append(inn)
        return RegistryLookupResult(
            status="found",
            query=inn,
            search_mode="inn",
            inn=inn,
            official_name='ОсОО "Жер Компани"',
            director="Асанов Асан",
        )

    monkeypatch.setattr(main.workflow, "lookup_registry", lookup)
    result = main.registry_suggestion("01412201510188", "")

    assert calls == ["01412201510188"]
    assert result["status"] == "found"
    assert result["official_name"] == 'ОсОО "Жер Компани"'
    assert result["director"] == "Асанов Асан"


def test_live_suggestion_does_not_query_personal_pin(monkeypatch):
    monkeypatch.setattr(
        main.workflow,
        "lookup_registry",
        lambda _inn: (_ for _ in ()).throw(
            AssertionError("Физлицо нельзя отправлять в ОсОО.KG")
        ),
    )
    assert main.registry_suggestion("11412201510188", "Иванов Иван")[
        "status"
    ] == "not_applicable"


def test_live_suggestion_does_not_guess_ambiguous_prefix_four(monkeypatch):
    monkeypatch.setattr(
        main.workflow,
        "lookup_registry",
        lambda _inn: (_ for _ in ()).throw(
            AssertionError("Неоднозначный ИНН нельзя отправлять")
        ),
    )
    assert main.registry_suggestion("41412201510188", "")["status"] == (
        "classification_uncertain"
    )


def test_live_suggestion_queries_branch_with_prefix_four(monkeypatch):
    monkeypatch.setattr(
        main.workflow,
        "lookup_registry",
        lambda inn: RegistryLookupResult(
            status="not_found", query=inn, search_mode="inn"
        ),
    )
    result = main.registry_suggestion(
        "41412201510188", 'Филиал иностранной компании "Альфа"'
    )
    assert result["status"] == "not_found"
