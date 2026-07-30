from __future__ import annotations

from urllib.parse import parse_qs

import httpx
import pytest

from gns_app.services.registry_service import (
    OsooRegistryClient,
    RegistryLookupError,
)


HOME_HTML = """
<html><body>
  <form action="/search/" method="post">
    <input type="hidden" name="csrfmiddlewaretoken" value="token-1">
  </form>
</body></html>
"""


def _client(handler, *, sleeper=lambda _seconds: None) -> OsooRegistryClient:
    return OsooRegistryClient(
        transport=httpx.MockTransport(handler),
        crawl_delay_seconds=10,
        sleeper=sleeper,
    )


def test_exact_html_lookup_sends_only_inn_and_uses_cache():
    requests: list[httpx.Request] = []

    def handler(request: httpx.Request) -> httpx.Response:
        requests.append(request)
        if request.method == "GET":
            return httpx.Response(
                200,
                text=HOME_HTML,
                headers={"content-type": "text/html; charset=utf-8"},
            )
        form = parse_qs(request.content.decode())
        assert form == {
            "csrfmiddlewaretoken": ["token-1"],
            "text": ["12345678901234"],
            "exact": ["on"],
        }
        return httpx.Response(
            200,
            text="""
            <html><body><div id="results"><table>
              <tr>
                <td>ОсОО «Кыргыз Тест»</td>
                <td><a href="/inn/12345678901234/">12345678901234</a></td>
                <td>Руководитель</td>
              </tr>
            </table></div></body></html>
            """,
            headers={"content-type": "text/html; charset=utf-8"},
        )

    client = _client(handler)
    first = client.lookup_by_inn("12345678901234")
    second = client.lookup_by_inn("12345678901234")

    assert first.status == "found"
    assert first.official_name == "ОсОО «Кыргыз Тест»"
    assert not first.from_cache
    assert second.from_cache
    assert [request.method for request in requests] == ["GET", "POST"]


def test_html_lookup_respects_osoo_crawl_delay():
    sleeps: list[float] = []

    def handler(request: httpx.Request) -> httpx.Response:
        text = HOME_HTML if request.method == "GET" else "<html></html>"
        return httpx.Response(
            200,
            text=text,
            headers={"content-type": "text/html"},
        )

    client = _client(handler, sleeper=sleeps.append)
    client.lookup_by_inn("12345678901234")

    assert len(sleeps) == 1
    assert sleeps[0] > 9


def test_html_lookup_rejects_non_14_digit_inn():
    client = _client(lambda _request: pytest.fail("Запрос не должен уйти"))

    with pytest.raises(RegistryLookupError, match="14 цифр"):
        client.lookup_by_inn("123")


def test_html_lookup_wraps_http_status_error():
    def handler(request: httpx.Request) -> httpx.Response:
        if request.method == "GET":
            return httpx.Response(
                200,
                text=HOME_HTML,
                headers={"content-type": "text/html"},
            )
        return httpx.Response(
            503,
            text="<html>unavailable</html>",
            headers={"content-type": "text/html"},
        )

    client = _client(handler)

    with pytest.raises(RegistryLookupError, match="ошибку HTTP"):
        client.lookup_by_inn("12345678901234")
