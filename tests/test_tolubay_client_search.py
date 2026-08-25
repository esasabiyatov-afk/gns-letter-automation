from __future__ import annotations

from email.message import Message
from urllib.parse import parse_qs

import pytest

from gns_app.vendor.tolubay_hub import ProtocolError, TolubayClient, TolubayConfig


class _Response:
    def __init__(self, body: str, url: str):
        self._body = body.encode("utf-8")
        self._url = url
        self.headers = Message()

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self) -> bytes:
        return self._body

    def geturl(self) -> str:
        return self._url


class _Opener:
    def __init__(self, body: str, base_url: str):
        self.body = body
        self.base_url = base_url
        self.request = None

    def open(self, request, timeout):
        self.request = request
        return _Response(self.body, request.full_url)


def _client(html: str) -> tuple[TolubayClient, _Opener]:
    base_url = "https://abs.example.test"
    client = TolubayClient(TolubayConfig(base_url=base_url))
    opener = _Opener(html, base_url)
    client._opener = opener
    client._authenticated = True
    return client, opener


def test_customer_search_posts_full_form_and_reads_exact_inn():
    inn = "12345678901234"
    client, opener = _client(
        f"""
        <table><thead><tr><th>ID</th><th>ФИО / Наименование</th>
        <th>Паспорт / ИНН</th><th>Дата рождения</th>
        <th>Фактический адрес</th><th>Номера телефонов</th></tr></thead>
        <tbody><tr><td>42</td><td>Тестовый клиент</td><td>{inn}</td>
        <td></td><td></td><td></td><td>
        <a href="/OnlineBank.Management.MVC/Customers/Details?customerID=42">view</a>
        </td></tr></tbody></table>
        """
    )

    results = client.search_customers({"SearchIdentificationNo": inn})

    assert opener.request is not None
    assert opener.request.get_method() == "POST"
    assert opener.request.get_header("X-requested-with") == "XMLHttpRequest"
    fields = parse_qs(opener.request.data.decode("utf-8"), keep_blank_values=True)
    assert fields["SearchIdentificationNo"] == [inn]
    assert fields["ShowLinks"] == ["True"]
    assert fields["SearchSurname"] == [""]
    assert len(results) == 1
    assert results[0].customer_id == "42"
    assert results[0].identity == inn


def test_customer_search_accepts_recognizable_empty_table():
    client, _ = _client(
        "<table><thead><tr><th>ID</th><th>ФИО / Наименование</th>"
        "</tr></thead><tbody></tbody></table>"
    )

    assert client.search_customers({"SearchIdentificationNo": "12345678901234"}) == []


def test_customer_search_rejects_unknown_html_instead_of_reporting_not_found():
    client, _ = _client("<html><body>unexpected proxy page</body></html>")

    with pytest.raises(ProtocolError, match="recognizable result table"):
        client.search_customers({"SearchIdentificationNo": "12345678901234"})
