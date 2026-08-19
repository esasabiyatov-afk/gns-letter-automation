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
class RegistryCompany:
    inn: str
    name: str
    director: str | None = None


@dataclass(frozen=True, slots=True)
class RegistryLookupResult:
    status: str
    query: str
    search_mode: str
    matches: tuple[RegistryCompany, ...] = ()
    inn: str | None = None
    official_name: str | None = None
    director: str | None = None
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

    def companies(self) -> list[RegistryCompany]:
        companies: list[RegistryCompany] = []
        for row in self.rows:
            if len(row) < 2:
                continue
            name = str(row[0]["text"]).strip()
            director = (
                str(row[2]["text"]).strip()
                if len(row) >= 3 and str(row[2]["text"]).strip()
                else None
            )
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
                    companies.append(
                        RegistryCompany(
                            inn=inn,
                            name=name,
                            director=director,
                        )
                    )
        return companies


class ReestrKgClient:
    """Сверка компаний через публичный поиск reestr.kg.

    reestr.kg отдаёт страницу поиска за Cloudflare-проверкой, поэтому вместо
    обычного httpx-клиента используется cloudscraper (обходит только
    браузерную JS-проверку Cloudflare, не CAPTCHA и не авторизацию).
    Между запросами выдерживается небольшая пауза, как и для ОсОО.KG.
    """

    BASE_URL = "https://reestr.kg"
    ALLOWED_HOSTS = frozenset({"reestr.kg", "www.reestr.kg"})

    def __init__(
        self,
        *,
        crawl_delay_seconds: float = 3.0,
        cache_seconds: float = 3600.0,
        clock: Callable[[], float] = time.monotonic,
        sleeper: Callable[[float], None] = time.sleep,
        scraper: object | None = None,
    ):
        self.crawl_delay_seconds = crawl_delay_seconds
        self.cache_seconds = cache_seconds
        self.clock = clock
        self.sleeper = sleeper
        self._last_request_at: float | None = None
        self._cache: dict[str, tuple[float, RegistryLookupResult]] = {}
        self._lock = threading.Lock()
        self._scraper = scraper

    def _client(self):
        if self._scraper is not None:
            return self._scraper
        try:
            import cloudscraper
        except ImportError as exc:  # pragma: no cover - зависит от окружения
            raise RegistryLookupError(
                "Библиотека cloudscraper не установлена. Переустановите "
                "зависимости приложения."
            ) from exc
        self._scraper = cloudscraper.create_scraper()
        return self._scraper

    def lookup_by_inn(self, inn: str) -> RegistryLookupResult:
        clean_inn = re.sub(r"\D", "", inn)
        if len(clean_inn) != 14:
            raise RegistryLookupError("Для сверки нужны ровно 14 цифр ИНН")
        return self._search(clean_inn, search_mode="inn", exact=True)

    def search_by_name(self, name: str) -> RegistryLookupResult:
        clean_name = " ".join(name.split())
        if len(clean_name) < 2:
            raise RegistryLookupError(
                "Для поиска по названию введите минимум 2 символа"
            )
        if len(clean_name) > 160:
            raise RegistryLookupError(
                "Название для поиска не должно быть длиннее 160 символов"
            )
        return self._search(clean_name, search_mode="name", exact=False)

    def _search(
        self,
        query: str,
        *,
        search_mode: str,
        exact: bool,
    ) -> RegistryLookupResult:
        cache_key = f"{search_mode}:{query.casefold()}"
        with self._lock:
            cached = self._cache.get(cache_key)
            now = self.clock()
            if cached and now - cached[0] <= self.cache_seconds:
                value = cached[1]
                return RegistryLookupResult(
                    status=value.status,
                    query=value.query,
                    search_mode=value.search_mode,
                    matches=value.matches,
                    inn=value.inn,
                    official_name=value.official_name,
                    director=value.director,
                    provider=value.provider,
                    from_cache=True,
                )

            if self._last_request_at is not None:
                elapsed = self.clock() - self._last_request_at
                remaining = self.crawl_delay_seconds - elapsed
                if remaining > 0:
                    self.sleeper(remaining)

            try:
                response = self._client().get(
                    f"{self.BASE_URL}/search",
                    params={"q": query, "page": 1},
                    timeout=15,
                )
            except Exception as exc:  # noqa: BLE001 - cloudscraper/requests
                self._last_request_at = self.clock()
                raise RegistryLookupError(
                    "reestr.kg не ответил вовремя или недоступен. "
                    "Повторите позже."
                ) from exc
            self._last_request_at = self.clock()

            self._validate_response(response)
            matches = self._parse(response.text)
            if search_mode == "inn":
                matches = [
                    company for company in matches if company.inn == query
                ]
            unique_matches = tuple(
                {
                    (company.inn, company.name, company.director): company
                    for company in matches
                }.values()
            )
            if unique_matches:
                result = RegistryLookupResult(
                    status=(
                        "found" if len(unique_matches) == 1 else "multiple"
                    ),
                    query=query,
                    search_mode=search_mode,
                    matches=unique_matches,
                    inn=(
                        unique_matches[0].inn
                        if len(unique_matches) == 1
                        else None
                    ),
                    official_name=(
                        unique_matches[0].name
                        if len(unique_matches) == 1
                        else None
                    ),
                    director=(
                        unique_matches[0].director
                        if len(unique_matches) == 1
                        else None
                    ),
                    provider="reestr.kg",
                )
            else:
                result = RegistryLookupResult(
                    status="not_found",
                    query=query,
                    search_mode=search_mode,
                    provider="reestr.kg",
                )
            self._cache[cache_key] = (self.clock(), result)
            return result

    def _validate_response(self, response) -> None:
        status_code = getattr(response, "status_code", None)
        if status_code is not None and status_code >= 400:
            raise RegistryLookupError(
                f"reestr.kg вернул ошибку HTTP {status_code}."
            )
        final_url = str(getattr(response, "url", ""))
        parsed = urlparse(final_url)
        if parsed.hostname and parsed.hostname not in self.ALLOWED_HOSTS:
            raise RegistryLookupError(
                "reestr.kg перенаправил запрос на посторонний адрес."
            )
        content = getattr(response, "content", b"") or b""
        if len(content) > 4 * 1024 * 1024:
            raise RegistryLookupError(
                "HTML-ответ reestr.kg превышает допустимый размер."
            )

    @staticmethod
    def _parse(html: str) -> list[RegistryCompany]:
        try:
            from bs4 import BeautifulSoup
        except ImportError as exc:  # pragma: no cover - зависит от окружения
            raise RegistryLookupError(
                "Библиотека beautifulsoup4 не установлена. Переустановите "
                "зависимости приложения."
            ) from exc

        soup = BeautifulSoup(html, "html.parser")
        companies: list[RegistryCompany] = []
        seen_inns: set[str] = set()
        for a_tag in soup.find_all("a", href=True):
            href = a_tag["href"]
            if "/companies/" not in href:
                continue
            inn = href.rstrip("/").split("/companies/")[-1].strip()
            if not re.fullmatch(r"\d{14}", inn) or inn in seen_inns:
                continue

            parent_card = a_tag.parent
            h3 = parent_card.find("h3") if parent_card else None
            name = h3.get_text(strip=True) if h3 else ""
            if not name:
                aria = a_tag.get("aria-label", "") or ""
                name = aria.replace("Перейти к компании ", "").strip()
            if not name:
                continue

            seen_inns.add(inn)
            companies.append(RegistryCompany(inn=inn, name=name))
        return companies


class CompositeRegistryClient:
    """Основной + резервный реестр с настраиваемым приоритетом.

    Сначала пробуем реестр, выбранный как основной в настройках. Если он
    выбросил ошибку (сеть недоступна, сайт не отвечает, изменилась
    разметка) — автоматически пробуем второй. Если оба недоступны,
    поднимается ошибка первого из них, чтобы сообщение оставалось понятным.
    """

    PRIMARY_OSOO = "osoo"
    PRIMARY_REESTR_KG = "reestr_kg"

    def __init__(
        self,
        osoo_client: "OsooRegistryClient",
        reestr_kg_client: "ReestrKgClient",
        priority_getter: Callable[[], str],
    ):
        self._clients = {
            self.PRIMARY_OSOO: osoo_client,
            self.PRIMARY_REESTR_KG: reestr_kg_client,
        }
        self._priority_getter = priority_getter

    def _ordered_clients(self) -> tuple[object, object]:
        try:
            primary_key = self._priority_getter()
        except Exception:  # noqa: BLE001 - настройки не должны валить сверку
            primary_key = self.PRIMARY_OSOO
        if primary_key not in self._clients:
            primary_key = self.PRIMARY_OSOO
        secondary_key = (
            self.PRIMARY_REESTR_KG
            if primary_key == self.PRIMARY_OSOO
            else self.PRIMARY_OSOO
        )
        return self._clients[primary_key], self._clients[secondary_key]

    def lookup_by_inn(self, inn: str) -> RegistryLookupResult:
        primary, secondary = self._ordered_clients()
        try:
            return primary.lookup_by_inn(inn)
        except RegistryLookupError:
            return secondary.lookup_by_inn(inn)

    def search_by_name(self, name: str) -> RegistryLookupResult:
        primary, secondary = self._ordered_clients()
        try:
            return primary.search_by_name(name)
        except RegistryLookupError:
            return secondary.search_by_name(name)


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
        return self._search(clean_inn, search_mode="inn", exact=True)

    def search_by_name(self, name: str) -> RegistryLookupResult:
        clean_name = " ".join(name.split())
        if len(clean_name) < 2:
            raise RegistryLookupError(
                "Для поиска по названию введите минимум 2 символа"
            )
        if len(clean_name) > 160:
            raise RegistryLookupError(
                "Название для поиска не должно быть длиннее 160 символов"
            )
        return self._search(clean_name, search_mode="name", exact=False)

    def _search(
        self,
        query: str,
        *,
        search_mode: str,
        exact: bool,
    ) -> RegistryLookupResult:
        cache_key = f"{search_mode}:{query.casefold()}"
        with self._lock:
            cached = self._cache.get(cache_key)
            now = self.clock()
            if cached and now - cached[0] <= self.cache_seconds:
                value = cached[1]
                return RegistryLookupResult(
                    status=value.status,
                    query=value.query,
                    search_mode=value.search_mode,
                    matches=value.matches,
                    inn=value.inn,
                    official_name=value.official_name,
                    director=value.director,
                    provider=value.provider,
                    from_cache=True,
                )

            try:
                self._ensure_csrf()
                form_data = {
                    "csrfmiddlewaretoken": self._csrf_token or "",
                    "text": query,
                }
                if exact:
                    form_data["exact"] = "on"
                response = self._request(
                    "POST",
                    "/search/",
                    data=form_data,
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
            parsed_matches = parser.companies()
            if search_mode == "inn":
                parsed_matches = [
                    company
                    for company in parsed_matches
                    if company.inn == query
                ]
            unique_matches = tuple(
                {
                    (company.inn, company.name, company.director): company
                    for company in parsed_matches
                }.values()
            )
            if search_mode == "inn" and len(unique_matches) == 1:
                company = unique_matches[0]
                result = RegistryLookupResult(
                    status="found",
                    query=query,
                    search_mode=search_mode,
                    matches=unique_matches,
                    inn=company.inn,
                    official_name=company.name,
                    director=company.director,
                )
            elif unique_matches:
                result = RegistryLookupResult(
                    status=(
                        "multiple"
                        if len(unique_matches) > 1
                        else "found"
                    ),
                    query=query,
                    search_mode=search_mode,
                    matches=unique_matches,
                    inn=(
                        unique_matches[0].inn
                        if len(unique_matches) == 1
                        else None
                    ),
                    official_name=(
                        unique_matches[0].name
                        if len(unique_matches) == 1
                        else None
                    ),
                    director=(
                        unique_matches[0].director
                        if len(unique_matches) == 1
                        else None
                    ),
                )
            else:
                result = RegistryLookupResult(
                    status="not_found",
                    query=query,
                    search_mode=search_mode,
                )
            self._cache[cache_key] = (self.clock(), result)
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
