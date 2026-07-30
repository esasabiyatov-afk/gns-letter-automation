from __future__ import annotations

import re
import threading
import time
from dataclasses import dataclass
from html.parser import HTMLParser
from typing import Callable
from urllib.parse import urlparse

import httpx


class RegistryLookupError(ValueError):
    pass


@dataclass(frozen=True, slots=True)
class RegistryLookupResult:
    status: str
    inn: str
    official_name: str | None = None
    provider: str = "ОсОО.KG"
    from_cache: bool = False


class _CompanySearchParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[dict[str, object]]] = []
        self._row: list[dict[str, object]] | None = None
        self._cell: dict[str, object] | None = None

    def handle_starttag(
        self,
        tag: str,
        attrs: list[tuple[str, str | None]],
    ) -> None:
        attributes = dict(attrs)
        if tag == "tr":
            self._row = []
        elif tag == "td" and self._row is not None:
            self._cell = {"text": [], "hrefs": []}
        elif tag == "a" and self._cell is not None:
            href = attributes.get("href")
            if href:
                self._cell["hrefs"].append(href)

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell["text"].append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "td" and self._row is not None and self._cell is not None:
            self._cell["text"] = " ".join(
                " ".join(self._cell["text"]).split()
            )
            self._row.append(self._cell)
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None
            self._cell = None

    def companies(self) -> list[tuple[str, str]]:
        companies: list[tuple[str, str]] = []
        for row in self.rows:
            if len(row) < 2:
                continue
            name = str(row[0]["text"]).strip()
            inn_candidates: list[str] = []
            inn_candidates.extend(
                re.findall(r"(?<!\d)\d{14}(?!\d)", str(row[1]["text"]))
            )
            for href in row[1]["hrefs"]:
                match = re.fullmatch(r"/inn/(\d{14})/", str(href))
                if match:
                    inn_candidates.append(match.group(1))
            for inn in dict.fromkeys(inn_candidates):
                if name:
                    companies.append((inn, name))
        return companies


class OsooRegistryClient:
    """Ручная HTML-сверка одного ИНН через публичную форму ОсОО.KG.

    Клиент не обходит авторизацию, CAPTCHA или robots.txt. Между запросами
    выдерживается Crawl-delay: 10. PDF и OCR-текст наружу не передаются.
    """

    BASE_URL = "https://www.osoo.kg"
    ALLOWED_HOSTS = frozenset({"osoo.kg", "www.osoo.kg"})

    def __init__(
        self,
        *,
        transport: httpx.BaseTransport | None = None,
        crawl_delay_seconds: float = 10.0,
        cache_seconds: float = 3600.0,
        clock: Callable[[], float] = time.monotonic,
        sleeper: Callable[[float], None] = time.sleep,
    ):
        self.crawl_delay_seconds = crawl_delay_seconds
        self.cache_seconds = cache_seconds
        self.clock = clock
        self.sleeper = sleeper
        self._last_request_at: float | None = None
        self._csrf_token: str | None = None
        self._cache: dict[str, tuple[float, RegistryLookupResult]] = {}
        self._lock = threading.Lock()
        self._client = httpx.Client(
            base_url=self.BASE_URL,
            follow_redirects=True,
            timeout=httpx.Timeout(20.0, connect=10.0),
            transport=transport,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 "
                    "(compatible; GNSLetterAutomationRegistryCheck/0.1)"
                )
            },
        )

    def lookup_by_inn(self, inn: str) -> RegistryLookupResult:
        clean_inn = re.sub(r"\D", "", inn)
        if len(clean_inn) != 14:
            raise RegistryLookupError("Для сверки нужны ровно 14 цифр ИНН")

        with self._lock:
            cached = self._cache.get(clean_inn)
            now = self.clock()
            if cached and now - cached[0] <= self.cache_seconds:
                value = cached[1]
                return RegistryLookupResult(
                    status=value.status,
                    inn=value.inn,
                    official_name=value.official_name,
                    provider=value.provider,
                    from_cache=True,
                )

            try:
                self._ensure_csrf()
                response = self._request(
                    "POST",
                    "/search/",
                    data={
                        "csrfmiddlewaretoken": self._csrf_token or "",
                        "text": clean_inn,
                        "exact": "on",
                    },
                    headers={"Referer": f"{self.BASE_URL}/"},
                )
            except httpx.TimeoutException as exc:
                raise RegistryLookupError(
                    "ОсОО.KG не ответил вовремя. Повторите позже."
                ) from exc
            except httpx.HTTPError as exc:
                raise RegistryLookupError(
                    "Не удалось выполнить HTML-запрос к ОсОО.KG."
                ) from exc

            self._validate_response(response)
            self._update_csrf(response.text)
            parser = _CompanySearchParser()
            parser.feed(response.text)
            exact = [
                (found_inn, name)
                for found_inn, name in parser.companies()
                if found_inn == clean_inn
            ]
            unique_names = list(dict.fromkeys(name for _, name in exact))
            if len(unique_names) == 1:
                result = RegistryLookupResult(
                    status="found",
                    inn=clean_inn,
                    official_name=unique_names[0],
                )
            elif len(unique_names) > 1:
                result = RegistryLookupResult(
                    status="multiple",
                    inn=clean_inn,
                )
            else:
                result = RegistryLookupResult(
                    status="not_found",
                    inn=clean_inn,
                )
            self._cache[clean_inn] = (self.clock(), result)
            return result

    def _ensure_csrf(self) -> None:
        if self._csrf_token:
            return
        try:
            response = self._request("GET", "/")
        except httpx.TimeoutException as exc:
            raise RegistryLookupError(
                "ОсОО.KG не ответил вовремя. Повторите позже."
            ) from exc
        except httpx.HTTPError as exc:
            raise RegistryLookupError(
                "Не удалось открыть HTML-форму ОсОО.KG."
            ) from exc
        self._validate_response(response)
        self._update_csrf(response.text)
        if not self._csrf_token:
            raise RegistryLookupError(
                "ОсОО.KG изменил форму поиска: CSRF-токен не найден."
            )

    def _request(self, method: str, path: str, **kwargs) -> httpx.Response:
        if self._last_request_at is not None:
            elapsed = self.clock() - self._last_request_at
            remaining = self.crawl_delay_seconds - elapsed
            if remaining > 0:
                self.sleeper(remaining)
        response = self._client.request(method, path, **kwargs)
        self._last_request_at = self.clock()
        return response

    def _validate_response(self, response: httpx.Response) -> None:
        try:
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            raise RegistryLookupError(
                "ОсОО.KG вернул ошибку HTTP. Повторите проверку позже."
            ) from exc
        parsed = urlparse(str(response.url))
        if parsed.scheme != "https" or parsed.hostname not in self.ALLOWED_HOSTS:
            raise RegistryLookupError(
                "ОсОО.KG перенаправил запрос на посторонний адрес."
            )
        content_type = response.headers.get("content-type", "").casefold()
        if "text/html" not in content_type:
            raise RegistryLookupError(
                "ОсОО.KG вернул неожиданный формат ответа."
            )
        if len(response.content) > 2 * 1024 * 1024:
            raise RegistryLookupError(
                "HTML-ответ ОсОО.KG превышает допустимый размер."
            )

    def _update_csrf(self, html: str) -> None:
        match = re.search(
            r'name=["\']csrfmiddlewaretoken["\']\s+'
            r'value=["\']([^"\']+)["\']',
            html,
            re.IGNORECASE,
        )
        if match:
            self._csrf_token = match.group(1)
