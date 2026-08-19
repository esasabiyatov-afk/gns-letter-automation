from __future__ import annotations

import pytest

from gns_app.services.registry_service import (
    CompositeRegistryClient,
    ReestrKgClient,
    RegistryLookupError,
    RegistryLookupResult,
)


class _StubClient:
    """Простая заглушка реестра для проверки составного клиента."""

    def __init__(self, *, name: str, fail: bool = False):
        self.name = name
        self.fail = fail
        self.inn_calls: list[str] = []
        self.name_calls: list[str] = []

    def lookup_by_inn(self, inn: str) -> RegistryLookupResult:
        self.inn_calls.append(inn)
        if self.fail:
            raise RegistryLookupError(f"{self.name} недоступен")
        return RegistryLookupResult(
            status="found",
            query=inn,
            search_mode="inn",
            inn=inn,
            official_name=f'Компания от {self.name}',
            provider=self.name,
        )

    def search_by_name(self, name: str) -> RegistryLookupResult:
        self.name_calls.append(name)
        if self.fail:
            raise RegistryLookupError(f"{self.name} недоступен")
        return RegistryLookupResult(
            status="found",
            query=name,
            search_mode="name",
            official_name=name,
            provider=self.name,
        )


def test_composite_client_uses_primary_when_priority_is_osoo():
    osoo = _StubClient(name="ОсОО.KG")
    reestr = _StubClient(name="reestr.kg")
    client = CompositeRegistryClient(
        osoo, reestr, priority_getter=lambda: "osoo"
    )

    result = client.lookup_by_inn("02312201410117")

    assert result.provider == "ОсОО.KG"
    assert osoo.inn_calls == ["02312201410117"]
    assert reestr.inn_calls == []


def test_composite_client_uses_primary_when_priority_is_reestr_kg():
    osoo = _StubClient(name="ОсОО.KG")
    reestr = _StubClient(name="reestr.kg")
    client = CompositeRegistryClient(
        osoo, reestr, priority_getter=lambda: "reestr_kg"
    )

    result = client.lookup_by_inn("02312201410117")

    assert result.provider == "reestr.kg"
    assert reestr.inn_calls == ["02312201410117"]
    assert osoo.inn_calls == []


def test_composite_client_falls_back_to_secondary_when_primary_fails():
    osoo = _StubClient(name="ОсОО.KG", fail=True)
    reestr = _StubClient(name="reestr.kg")
    client = CompositeRegistryClient(
        osoo, reestr, priority_getter=lambda: "osoo"
    )

    result = client.lookup_by_inn("02312201410117")

    assert result.provider == "reestr.kg"
    assert osoo.inn_calls == ["02312201410117"]
    assert reestr.inn_calls == ["02312201410117"]


def test_composite_client_raises_secondary_error_when_both_fail():
    osoo = _StubClient(name="ОсОО.KG", fail=True)
    reestr = _StubClient(name="reestr.kg", fail=True)
    client = CompositeRegistryClient(
        osoo, reestr, priority_getter=lambda: "osoo"
    )

    with pytest.raises(RegistryLookupError):
        client.lookup_by_inn("02312201410117")


def test_composite_client_falls_back_for_search_by_name_too():
    osoo = _StubClient(name="ОсОО.KG", fail=True)
    reestr = _StubClient(name="reestr.kg")
    client = CompositeRegistryClient(
        osoo, reestr, priority_getter=lambda: "osoo"
    )

    result = client.search_by_name('ОсОО "Альфа"')

    assert result.provider == "reestr.kg"
    assert reestr.name_calls == ['ОсОО "Альфа"']


def test_composite_client_falls_back_to_osoo_on_unknown_priority_value():
    osoo = _StubClient(name="ОсОО.KG")
    reestr = _StubClient(name="reestr.kg")
    client = CompositeRegistryClient(
        osoo, reestr, priority_getter=lambda: "something_unexpected"
    )

    result = client.lookup_by_inn("02312201410117")

    assert result.provider == "ОсОО.KG"


REESTR_SEARCH_HTML = """
<html><body>
<div class="card">
  <a href="/companies/02312201410117" aria-label='Перейти к компании ОсОО "Альфа"'>
    <h3>ОсОО "Альфа"</h3>
  </a>
</div>
<div class="card">
  <a href="/companies/00000000000000" aria-label='Перейти к компании ОсОО "Другая"'>
    <h3>ОсОО "Другая"</h3>
  </a>
</div>
</body></html>
"""


class _FakeResponse:
    def __init__(self, text: str, *, status_code: int = 200,
                 url: str = "https://reestr.kg/search"):
        self.text = text
        self.content = text.encode("utf-8")
        self.status_code = status_code
        self.url = url


class _FakeScraper:
    def __init__(self, response: _FakeResponse):
        self._response = response
        self.calls: list[dict] = []

    def get(self, url, params=None, timeout=None):
        self.calls.append({"url": url, "params": params})
        return self._response


def test_reestr_kg_client_finds_exact_inn_match():
    scraper = _FakeScraper(_FakeResponse(REESTR_SEARCH_HTML))
    client = ReestrKgClient(scraper=scraper, sleeper=lambda _s: None)

    result = client.lookup_by_inn("02312201410117")

    assert result.status == "found"
    assert result.official_name == 'ОсОО "Альфа"'
    assert result.provider == "reestr.kg"


def test_reestr_kg_client_reports_not_found_when_inn_does_not_match():
    scraper = _FakeScraper(_FakeResponse(REESTR_SEARCH_HTML))
    client = ReestrKgClient(scraper=scraper, sleeper=lambda _s: None)

    result = client.lookup_by_inn("99999999999999")

    assert result.status == "not_found"


def test_reestr_kg_client_raises_on_http_error():
    scraper = _FakeScraper(_FakeResponse("oops", status_code=503))
    client = ReestrKgClient(scraper=scraper, sleeper=lambda _s: None)

    with pytest.raises(RegistryLookupError):
        client.lookup_by_inn("02312201410117")


def test_reestr_kg_client_rejects_foreign_redirect():
    scraper = _FakeScraper(
        _FakeResponse(REESTR_SEARCH_HTML, url="https://evil.example/search")
    )
    client = ReestrKgClient(scraper=scraper, sleeper=lambda _s: None)

    with pytest.raises(RegistryLookupError):
        client.lookup_by_inn("02312201410117")
