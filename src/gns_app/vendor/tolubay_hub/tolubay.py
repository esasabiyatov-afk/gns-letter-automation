from __future__ import annotations

import base64
import json
import re
import ssl
import time
from dataclasses import dataclass
from datetime import date, datetime
from email.message import Message
from http.cookiejar import CookieJar
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Callable, Mapping
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlencode, urljoin, urlparse
from urllib.request import HTTPCookieProcessor, HTTPSHandler, Request, build_opener

from .readonly import ReadOnlyPolicy


ROOT = "/OnlineBank.Management.MVC"
SIGN_IN = f"{ROOT}/Account/SignIn"
TEMPLATE_REPORT = f"{ROOT}/Management/ExcelTemplatesReport"
TEMPLATE_JOB = f"{ROOT}/Management/ExcelTemplatesReportJob"
CHECK_SERVICE = f"{TEMPLATE_JOB}/CheckService"
RUN_JOB = f"{TEMPLATE_JOB}/Run"
JOB_INDEX = f"{ROOT}/Job"
JOB_LOAD = f"{ROOT}/Job/Load"
JOB_DOWNLOAD = f"{ROOT}/Job/Download"
CUSTOMER_SEARCH_RESULT = f"{ROOT}/Customers/SearchResult"
CUSTOMER_DETAILS = f"{ROOT}/Customers/Details"
CUSTOMER_VIEW = f"{ROOT}/Customers/Edit"
CUSTOMER_ACCOUNTS = f"{ROOT}/Deposits/Customer/GetDepositsByCustomerId"
ADDITIONAL_REPORT = f"{ROOT}/Management/AdditionalReport"
ADDITIONAL_REPORT_LOAD = f"{ADDITIONAL_REPORT}/Load"
ADDITIONAL_REPORT_JOB = f"{ROOT}/Management/AdditionalReportJob"
ADDITIONAL_CHECK_SERVICE = f"{ADDITIONAL_REPORT_JOB}/CheckService"
ADDITIONAL_RUN_JOB = f"{ADDITIONAL_REPORT_JOB}/Run"

CUSTOMER_SEARCH_FIELDS = frozenset(
    {
        "SearchCustomerID",
        "SearchSurname",
        "SearchCustomerName",
        "SearchOtchestvo",
        "SearchAgreementNo",
        "SearchAccountNo",
        "SearchCompanyName",
        "SearchIdentificationNo",
        "SearchGroupName",
        "SearchStreetName",
        "SearchHouseNo",
        "SearchFlatNo",
        "SearchPhoneNumber",
    }
)

REPORT_EXECUTE_PATHS = frozenset(
    {
        f"{ROOT}/CashIncomeJournalReport/Execute",
        f"{ROOT}/CashWithdrawalJournalReport/Execute",
        f"{ROOT}/MemorialOrderReport/Execute",
        f"{ROOT}/BalanceGroupsTurnoversReport/Execute",
        f"{ROOT}/TransactionsJournalReport/Execute",
        f"{ROOT}/AccountsStatementReport/Execute",
        f"{ROOT}/TransactionsByBalanceGroupsReport/Execute",
        f"{ROOT}/BalanceGroupsStatementReport/Execute",
        f"{ROOT}/BalanceSheetStatementReport/Execute",
        f"{ROOT}/TurnoverBalanceSheetByAccountsReport/Execute",
        f"{ROOT}/TurnoverBalanceSheetByGroupsReport/Execute",
        f"{ROOT}/Loans/OverduesDepartmentsReport/Execute",
        f"{ROOT}/Deposits/MainAccountStatementCommand/Execute",
        f"{ROOT}/Deposits/DetailedStatementCommand/Execute",
        f"{ROOT}/Deposits/PercentAccountStatementCommand/Execute",
    }
)


class TolubayError(RuntimeError):
    pass


class AuthenticationError(TolubayError):
    pass


class ProtocolError(TolubayError):
    pass


class JobTimeoutError(TolubayError):
    pass


class JobFailedError(TolubayError):
    pass


class _LoginFormParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.token: str | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "input":
            return
        values = dict(attrs)
        if values.get("name") == "__RequestVerificationToken":
            self.token = values.get("value")


class _HtmlTablesParser(HTMLParser):
    """Small dependency-free table parser used for customer search results."""

    def __init__(self) -> None:
        super().__init__()
        self.tables: list[dict[str, Any]] = []
        self._table: dict[str, Any] | None = None
        self._section = ""
        self._row: list[str] | None = None
        self._row_links: list[str] = []
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = dict(attrs)
        tag = tag.lower()
        if tag == "table" and self._table is None:
            self._table = {"headers": [], "rows": []}
        elif self._table is not None and tag in {"thead", "tbody"}:
            self._section = tag
        elif self._table is not None and tag == "tr":
            self._row = []
            self._row_links = []
        elif self._row is not None and tag in {"th", "td"}:
            self._cell = []
        elif self._row is not None and tag == "a" and values.get("href"):
            self._row_links.append(str(values["href"]))

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if self._table is None:
            return
        if tag in {"th", "td"} and self._cell is not None and self._row is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if self._section == "thead":
                self._table["headers"] = self._row
            elif self._section == "tbody" and any(self._row):
                self._table["rows"].append((self._row, tuple(self._row_links)))
            self._row = None
            self._row_links = []
        elif tag in {"thead", "tbody"}:
            self._section = ""
        elif tag == "table":
            self.tables.append(self._table)
            self._table = None


class _FormValuesParser(HTMLParser):
    """Extracts current form values without ever submitting the form."""

    def __init__(self) -> None:
        super().__init__()
        self.fields: dict[str, Any] = {}
        self.action = ""
        self._in_target_form = False
        self._select: dict[str, Any] | None = None
        self._option: dict[str, Any] | None = None
        self._textarea_name: str | None = None
        self._textarea: list[str] = []

    def _add(self, name: str, value: Any) -> None:
        if not name or name == "__RequestVerificationToken":
            return
        if name not in self.fields:
            self.fields[name] = value
        elif not isinstance(self.fields[name], list):
            self.fields[name] = [self.fields[name], value]
        else:
            self.fields[name].append(value)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        values = dict(attrs)
        tag = tag.lower()
        if tag == "form" and not self._in_target_form:
            action = values.get("action") or ""
            if "/Customers/Edit" in action:
                self._in_target_form = True
                self.action = action
            return
        if not self._in_target_form:
            return
        if tag == "input":
            name = values.get("name") or ""
            input_type = (values.get("type") or "text").lower()
            if input_type in {"submit", "button", "file", "password", "image", "reset"}:
                return
            if input_type in {"checkbox", "radio"} and "checked" not in values:
                return
            self._add(name, values.get("value", "on" if input_type in {"checkbox", "radio"} else ""))
        elif tag == "select":
            self._select = {"name": values.get("name") or "", "options": []}
        elif tag == "option" and self._select is not None:
            self._option = {
                "value": values.get("value", ""),
                "selected": "selected" in values,
                "text": [],
            }
        elif tag == "textarea":
            self._textarea_name = values.get("name") or ""
            self._textarea = []

    def handle_data(self, data: str) -> None:
        if self._option is not None:
            self._option["text"].append(data)
        if self._textarea_name is not None:
            self._textarea.append(data)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "option" and self._option is not None and self._select is not None:
            self._option["text"] = " ".join("".join(self._option["text"]).split())
            self._select["options"].append(self._option)
            self._option = None
        elif tag == "select" and self._select is not None:
            options = self._select["options"]
            chosen = [option for option in options if option["selected"]]
            if not chosen and options:
                chosen = [options[0]]
            values = [{"value": option["value"], "text": option["text"]} for option in chosen]
            if values:
                self._add(self._select["name"], values[0] if len(values) == 1 else values)
            self._select = None
        elif tag == "textarea" and self._textarea_name is not None:
            self._add(self._textarea_name, "".join(self._textarea))
            self._textarea_name = None
            self._textarea = []
        elif tag == "form" and self._in_target_form:
            self._in_target_form = False


class _LinksParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        href = dict(attrs).get("href")
        if href:
            self.links.append(href)


@dataclass(frozen=True)
class TolubayConfig:
    base_url: str = "https://ob.tolubay.kg"
    ca_file: str | Path | None = None
    verify_tls: bool = True
    timeout_seconds: float = 30.0
    poll_interval_seconds: float = 1.0


@dataclass(frozen=True)
class TemplateReportResult:
    job_id: int | str
    result_id: int | str
    file_name: str
    path: Path
    completed_with_messages: bool


@dataclass(frozen=True)
class CustomerSummary:
    customer_id: str
    name: str
    identity: str
    birth_date: str
    address: str
    phones: str
    details_url: str | None = None
    questionnaire_url: str | None = None


@dataclass(frozen=True)
class CustomerQuestionnaire:
    customer_id: str
    customer_type: str
    fields: Mapping[str, Any]
    source_url: str


@dataclass(frozen=True)
class AccountRecord:
    customer_id: str
    account_no: str
    currency_id: int | str | None
    status_id: int | str | None
    is_closed: bool
    data: Mapping[str, Any]


@dataclass(frozen=True)
class ReportDownloadResult:
    file_name: str
    path: Path
    content_type: str
    source_url: str


@dataclass(frozen=True)
class AdditionalReportItem:
    report_name: str
    report_type: str
    report_group: str
    data: Mapping[str, Any]


class TolubayClient:
    """HTTP-only, fail-closed adapter for Tolubay read/report operations."""

    def __init__(
        self,
        config: TolubayConfig | None = None,
        *,
        sleeper: Callable[[float], None] = time.sleep,
    ) -> None:
        self.config = config or TolubayConfig()
        parsed = urlparse(self.config.base_url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError("base_url must be an absolute HTTP(S) URL")

        self.cookies = CookieJar()
        handlers: list[Any] = [HTTPCookieProcessor(self.cookies)]
        if parsed.scheme == "https":
            if self.config.verify_tls:
                context = ssl.create_default_context(
                    cafile=str(self.config.ca_file) if self.config.ca_file else None
                )
            else:
                context = ssl._create_unverified_context()  # noqa: SLF001 - explicit opt-in only
            handlers.append(HTTPSHandler(context=context))
        self._opener = build_opener(*handlers)
        self._sleep = sleeper
        self._authenticated = False
        self._policy = ReadOnlyPolicy(
            self.config.base_url,
            allowed_post_paths=frozenset(
                {
                    SIGN_IN,
                    # Поиск клиента изменяет только серверный фильтр и
                    # возвращает HTML-таблицу. В интерфейсе Tolubay этот
                    # read-only endpoint вызывается POST-формой.
                    CUSTOMER_SEARCH_RESULT,
                    TEMPLATE_JOB,
                    CHECK_SERVICE,
                    RUN_JOB,
                    JOB_LOAD,
                    CUSTOMER_ACCOUNTS,
                    ADDITIONAL_REPORT_LOAD,
                    ADDITIONAL_REPORT_JOB,
                    ADDITIONAL_CHECK_SERVICE,
                    ADDITIONAL_RUN_JOB,
                    *REPORT_EXECUTE_PATHS,
                }
            ),
            view_only_get_routes=(
                (CUSTOMER_VIEW, (("isShow", "true"),)),
                (ADDITIONAL_REPORT, ()),
            ),
        )

    def _request(
        self,
        method: str,
        path: str,
        *,
        data: bytes | None = None,
        headers: dict[str, str] | None = None,
        field_names: tuple[str, ...] = (),
    ) -> tuple[bytes, Message, str]:
        url = self._policy.validate(method, path, field_names=field_names)
        request = Request(url, data=data, headers=headers or {}, method=method.upper())
        try:
            with self._opener.open(request, timeout=self.config.timeout_seconds) as response:
                body = response.read()
                final_url = response.geturl()
                response_headers = response.headers
        except HTTPError as exc:
            payload = exc.read(1024).decode("utf-8", errors="replace")
            raise ProtocolError(f"HTTP {exc.code} for {url}: {payload[:300]}") from exc
        except URLError as exc:
            raise ProtocolError(f"Cannot reach {url}: {exc.reason}") from exc

        final_path = urlparse(final_url).path
        if path != SIGN_IN and final_path == SIGN_IN:
            self._authenticated = False
            raise AuthenticationError("Tolubay session expired or authentication is required")
        return body, response_headers, final_url

    def _get_text(self, path: str) -> str:
        body, _, _ = self._request("GET", path)
        return body.decode("utf-8", errors="replace")

    def _post_form(
        self,
        path: str,
        fields: dict[str, Any],
        *,
        headers: dict[str, str] | None = None,
    ) -> tuple[bytes, Message, str]:
        normalized = {
            key: "" if value is None else str(value).lower() if isinstance(value, bool) else str(value)
            for key, value in fields.items()
        }
        encoded = urlencode(normalized).encode("utf-8")
        request_headers = {
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            **(headers or {}),
        }
        return self._request(
            "POST",
            path,
            data=encoded,
            headers=request_headers,
            field_names=tuple(fields),
        )

    def _post_json(self, path: str, payload: dict[str, Any]) -> dict[str, Any]:
        encoded = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        body, _, _ = self._request(
            "POST",
            path,
            data=encoded,
            headers={"Content-Type": "application/json; charset=UTF-8"},
            field_names=tuple(payload),
        )
        try:
            decoded = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ProtocolError(f"Expected JSON from {path}") from exc
        if not isinstance(decoded, dict):
            raise ProtocolError(f"Expected a JSON object from {path}")
        return decoded

    @staticmethod
    def _extract_js_json(html: str, variable: str) -> dict[str, Any]:
        marker = re.search(rf"\bvar\s+{re.escape(variable)}\s*=\s*", html)
        if marker is None:
            raise ProtocolError(f"JavaScript model {variable} was not found")
        try:
            value, _ = json.JSONDecoder().raw_decode(html[marker.end() :])
        except json.JSONDecodeError as exc:
            raise ProtocolError(f"JavaScript model {variable} is invalid") from exc
        if not isinstance(value, dict):
            raise ProtocolError(f"JavaScript model {variable} must be an object")
        return value

    @staticmethod
    def _status_ok(payload: dict[str, Any], operation: str) -> None:
        if str(payload.get("status", "")).lower() != "ok":
            message = str(payload.get("message") or f"{operation} failed")
            raise ProtocolError(message)

    def login(self, username: str, password: str) -> None:
        if not username or not password:
            raise ValueError("username and password are required")
        parser = _LoginFormParser()
        parser.feed(self._get_text(SIGN_IN))
        if not parser.token:
            raise AuthenticationError("Login CSRF token was not found")
        _, _, final_url = self._post_form(
            SIGN_IN,
            {
                "__RequestVerificationToken": parser.token,
                "UserName": username,
                "Password": password,
            },
        )
        if urlparse(final_url).path == SIGN_IN:
            raise AuthenticationError("Login was rejected")
        self._authenticated = True

    def _require_authenticated(self) -> None:
        if not self._authenticated:
            raise AuthenticationError("Call login() before accessing Tolubay data")

    @staticmethod
    def _customer_id(value: int | str) -> str:
        normalized = str(value).strip()
        if not normalized or not normalized.isdigit():
            raise ValueError("customer_id must contain digits only")
        return normalized

    def search_customers(
        self,
        criteria: Mapping[str, Any],
        *,
        page: int = 1,
        page_size: int = 50,
    ) -> list[CustomerSummary]:
        """Search customers through the same read-only AJAX endpoint as the UI."""
        self._require_authenticated()
        unknown = set(criteria) - CUSTOMER_SEARCH_FIELDS
        if unknown:
            raise ValueError(f"Unsupported customer search fields: {', '.join(sorted(unknown))}")
        normalized = {
            key: str(value).strip()
            for key, value in criteria.items()
            if value is not None and str(value).strip()
        }
        if not normalized:
            raise ValueError("At least one customer search criterion is required")
        if page < 1 or not 1 <= page_size <= 500:
            raise ValueError("page must be >= 1 and page_size must be between 1 and 500")
        payload = {
            "ShowLinks": "True",
            **{field: "" for field in CUSTOMER_SEARCH_FIELDS},
            **normalized,
        }
        body, _, _ = self._post_form(
            CUSTOMER_SEARCH_RESULT,
            payload,
            headers={"X-Requested-With": "XMLHttpRequest"},
        )
        html = body.decode("utf-8", errors="replace")
        parser = _HtmlTablesParser()
        parser.feed(html)
        table = next(
            (
                item
                for item in parser.tables
                if (
                    any(header.strip().casefold() == "id" for header in item["headers"])
                    and any(
                        token in header.casefold()
                        for header in item["headers"]
                        for token in ("фио", "наименование")
                    )
                )
                or any(
                    "/Customers/Details" in link or "/Customers/Edit" in link
                    for _, links in item["rows"]
                    for link in links
                )
            ),
            None,
        )
        if table is None:
            explicit_empty = re.search(
                r"(?:ничего\s+не\s+найден|запис(?:ей|и)\s+не\s+найден|"
                r"данн(?:ые|ых)\s+отсутств|нет\s+данных|no\s+(?:records|data))",
                html,
                flags=re.IGNORECASE,
            )
            if explicit_empty:
                return []
            raise ProtocolError(
                "Customer search response did not contain a recognizable result table"
            )

        headers = [header.strip().casefold() for header in table["headers"]]

        def column_index(*needles: str) -> int | None:
            for index, header in enumerate(headers):
                if any(needle in header for needle in needles):
                    return index
            return None

        id_index = column_index("id")
        name_index = column_index("фио", "наименование")
        identity_index = column_index("инн", "пин", "паспорт", "идентификац")
        birth_index = column_index("дата рождения")
        address_index = column_index("адрес")
        phone_index = column_index("телефон")
        requested_identity = re.sub(
            r"\D",
            "",
            normalized.get("SearchIdentificationNo", ""),
        )

        def cell(cells: list[str], index: int | None) -> str:
            return cells[index] if index is not None and index < len(cells) else ""

        results: list[CustomerSummary] = []
        for cells, links in table["rows"]:
            details_url = next((urljoin(self.config.base_url, link) for link in links if "/Customers/Details" in link), None)
            view_url = next((urljoin(self.config.base_url, link) for link in links if "/Customers/Edit" in link), None)
            linked_id = ""
            for link in links:
                parsed_link = urlparse(link)
                values = parse_qs(parsed_link.query)
                linked_id = next(
                    (
                        str(value[0]).strip()
                        for key, value in values.items()
                        if key.casefold() == "customerid" and value
                    ),
                    "",
                )
                if linked_id:
                    break
            identity = cell(cells, identity_index)
            if requested_identity:
                exact_cell = next(
                    (
                        value
                        for value in cells
                        if re.sub(r"\D", "", value) == requested_identity
                    ),
                    "",
                )
                if exact_cell:
                    identity = exact_cell
            customer_id = linked_id or cell(cells, id_index)
            if not customer_id:
                raise ProtocolError("Customer search row did not contain a customer ID")
            results.append(
                CustomerSummary(
                    customer_id=customer_id,
                    name=cell(cells, name_index),
                    identity=identity,
                    birth_date=cell(cells, birth_index),
                    address=cell(cells, address_index),
                    phones=cell(cells, phone_index),
                    details_url=details_url,
                    questionnaire_url=view_url,
                )
            )
        return results

    def search_customers_batch(
        self,
        criteria_items: list[Mapping[str, Any]],
        *,
        page_size: int = 50,
    ) -> list[CustomerSummary]:
        """Run multiple searches and return a de-duplicated customer list."""
        combined: dict[str, CustomerSummary] = {}
        for criteria in criteria_items:
            for customer in self.search_customers(criteria, page_size=page_size):
                combined[customer.customer_id] = customer
        return list(combined.values())

    def get_customer_questionnaire(self, customer_id: int | str) -> CustomerQuestionnaire:
        """Read the full questionnaire through the legacy view-mode GET route."""
        self._require_authenticated()
        normalized_id = self._customer_id(customer_id)
        path = f"{CUSTOMER_VIEW}?{urlencode({'customerID': normalized_id, 'isShow': 'True'})}"
        body, _, final_url = self._request("GET", path)
        parser = _FormValuesParser()
        parser.feed(body.decode("utf-8", errors="replace"))
        if not parser.fields:
            raise ProtocolError("Customer questionnaire form was not found")
        final_path = urlparse(final_url).path
        customer_type = final_path.rsplit("/", 1)[-1].removeprefix("Edit").removesuffix("Customer")
        return CustomerQuestionnaire(
            customer_id=normalized_id,
            customer_type=customer_type or "Unknown",
            fields=parser.fields,
            source_url=final_url,
        )

    @staticmethod
    def _account_records(customer_id: str, model: Mapping[str, Any]) -> list[AccountRecord]:
        deposits = model.get("Deposits") or []
        if not isinstance(deposits, list):
            raise ProtocolError("Accounts response does not contain a Deposits list")
        records: list[AccountRecord] = []
        for raw in deposits:
            if not isinstance(raw, dict):
                continue
            close_date = raw.get("CloseDate")
            records.append(
                AccountRecord(
                    customer_id=customer_id,
                    account_no=str(raw.get("MainAccountNo") or raw.get("AccountNo") or ""),
                    currency_id=raw.get("CurrencyID"),
                    status_id=raw.get("DepositAccountStatusID"),
                    is_closed=close_date not in (None, "", False),
                    data=raw,
                )
            )
        return records

    def get_accounts(
        self,
        customer_id: int | str,
        *,
        include_closed: bool = False,
    ) -> list[AccountRecord]:
        """Read active accounts, optionally including closed accounts."""
        self._require_authenticated()
        normalized_id = self._customer_id(customer_id)
        if include_closed:
            payload = self._post_json(
                CUSTOMER_ACCOUNTS,
                {"customerId": int(normalized_id), "showAll": False},
            )
            if payload.get("error"):
                raise ProtocolError(str(payload["error"]))
            model = payload.get("data", payload)
            if not isinstance(model, dict):
                raise ProtocolError("Accounts endpoint returned invalid data")
        else:
            html = self._get_text(
                f"{CUSTOMER_DETAILS}?{urlencode({'customerID': normalized_id})}"
            )
            model = self._extract_js_json(html, "vmDepositsJs")
        return self._account_records(normalized_id, model)

    def list_additional_reports(self) -> list[AdditionalReportItem]:
        """Return the additional-report catalogue loaded by the management UI."""
        self._require_authenticated()
        html = self._get_text(ADDITIONAL_REPORT)
        model = self._extract_js_json(html, "vmUniReferenceJs")
        payload = self._unwrap_load_response(
            self._post_json(ADDITIONAL_REPORT_LOAD, self._job_load_payload(model))
        )
        items = payload.get("Items") or []
        if not isinstance(items, list):
            raise ProtocolError("Additional report list does not contain Items")
        result: list[AdditionalReportItem] = []
        for item in items:
            if not isinstance(item, dict):
                continue
            result.append(
                AdditionalReportItem(
                    report_name=str(item.get("ReportName") or ""),
                    report_type=str(item.get("ReportType") or ""),
                    report_group=str(item.get("ReportGroup") or ""),
                    data=item,
                )
            )
        return result

    def _create_additional_report_job(
        self,
        report_name: str,
        report_type: str,
        start_date: str,
        end_date: str,
    ) -> dict[str, Any]:
        if not report_name.strip() or not report_type.strip():
            raise ValueError("report_name and report_type are required")
        body, _, _ = self._post_form(
            ADDITIONAL_REPORT_JOB,
            {
                "ReportName": report_name,
                "ReportType": report_type,
                "DateFrom": start_date,
                "DateTo": end_date,
                "returnUrl": urljoin(self.config.base_url, ADDITIONAL_REPORT),
            },
        )
        return self._extract_js_json(body.decode("utf-8", errors="replace"), "vmJobJs")

    def generate_additional_report(
        self,
        *,
        report_name: str,
        report_type: str,
        start_date: str,
        end_date: str,
        output_dir: str | Path,
        timeout_seconds: float = 180.0,
    ) -> ReportDownloadResult:
        """Generate a configured additional report through the shared job queue."""
        self._require_authenticated()
        model = self._create_additional_report_job(
            report_name,
            report_type,
            start_date,
            end_date,
        )
        context_key = self._run_job(
            model,
            check_service_path=ADDITIONAL_CHECK_SERVICE,
            run_path=ADDITIONAL_RUN_JOB,
        )
        job = self._wait_for_job(context_key, timeout_seconds)
        result = (job.get("JobResults") or [])[0]
        result_id, result_name = self._result_identity(result)
        body, headers, source_url = self._request(
            "GET", f"{JOB_DOWNLOAD}?{urlencode({'id': result_id})}"
        )
        safe_stem = re.sub(r'[<>:"/\\|?*]+', "_", report_name).strip(" .") or "report"
        return self._save_download(
            body,
            headers,
            source_url,
            Path(output_dir).expanduser().resolve(),
            result_name or safe_stem,
        )

    @staticmethod
    def _report_date(value: str | date | datetime) -> str:
        if isinstance(value, datetime):
            return value.replace(microsecond=0).isoformat()
        if isinstance(value, date):
            return datetime.combine(value, datetime.min.time()).isoformat()
        text = value.strip()
        if not text:
            raise ValueError("report_date is required")
        return text

    @staticmethod
    def _data_url(template_path: Path) -> str:
        if template_path.suffix.lower() != ".xlsx":
            raise ValueError("Template must be an .xlsx file")
        encoded = base64.b64encode(template_path.read_bytes()).decode("ascii")
        mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        return f"data:{mime};base64,{encoded}"

    def _create_job_model(
        self,
        template_path: Path,
        report_date: str | date | datetime,
        *,
        use_cache: bool,
        include_formula_as_comment: bool,
    ) -> dict[str, Any]:
        fields = {
            "fileName": template_path.name,
            "data": self._data_url(template_path),
            "reportDate": self._report_date(report_date),
            "includeFormulaAsComment": include_formula_as_comment,
            "useCache": use_cache,
            "returnUrl": urljoin(self.config.base_url.rstrip("/") + "/", TEMPLATE_REPORT.lstrip("/")),
        }
        body, _, _ = self._post_form(TEMPLATE_JOB, fields)
        return self._extract_js_json(body.decode("utf-8", errors="replace"), "vmJobJs")

    def _run_job(
        self,
        model: dict[str, Any],
        *,
        check_service_path: str = CHECK_SERVICE,
        run_path: str = RUN_JOB,
    ) -> str:
        context = model.get("Context")
        data_model = model.get("Data")
        if not isinstance(context, dict):
            raise ProtocolError("Job Context is missing")
        context_key = str(context.get("ContextKey") or "")
        if not context_key:
            raise ProtocolError("Job ContextKey is missing")

        check_body, _, _ = self._post_form(check_service_path, {})
        try:
            self._status_ok(json.loads(check_body.decode("utf-8")), "CheckService")
        except json.JSONDecodeError as exc:
            raise ProtocolError("CheckService returned invalid JSON") from exc

        fields = {
            "contextData": json.dumps(context, ensure_ascii=False, separators=(",", ":")),
            "contextDataType": model.get("ContextDataType"),
            "dataViewModel": json.dumps(data_model, ensure_ascii=False, separators=(",", ":")),
            "dataViewModelType": model.get("DataType"),
            "tryReconnectToJob": False,
        }
        run_body, _, _ = self._post_form(run_path, fields)
        try:
            run_result = json.loads(run_body.decode("utf-8"))
        except json.JSONDecodeError as exc:
            raise ProtocolError("Run returned invalid JSON") from exc
        self._status_ok(run_result, "Run")
        return context_key

    @staticmethod
    def _job_load_payload(model: dict[str, Any]) -> dict[str, Any]:
        filters = model.get("Filter") or {}
        if isinstance(filters, dict):
            filters = list(filters.values())
        context = model.get("Context") or []
        if isinstance(context, dict):
            context = list(context.values())
        return {
            "Filter": filters,
            "ReferenceType": model.get("ReferenceType"),
            "context": context,
            "descriptorType": model.get("DescriptorType"),
            "Pagination": model.get("Pagination"),
        }

    @staticmethod
    def _unwrap_load_response(payload: dict[str, Any]) -> dict[str, Any]:
        if payload.get("error"):
            raise ProtocolError(str(payload["error"]))
        data = payload.get("data", payload)
        if not isinstance(data, dict):
            raise ProtocolError("Job/Load response does not contain data")
        return data

    def _wait_for_job(self, context_key: str, timeout_seconds: float) -> dict[str, Any]:
        index_html = self._get_text(JOB_INDEX)
        grid_model = self._extract_js_json(index_html, "vmUniReferenceJs")
        deadline = time.monotonic() + timeout_seconds
        while True:
            data = self._unwrap_load_response(
                self._post_json(JOB_LOAD, self._job_load_payload(grid_model))
            )
            items = data.get("Items") or []
            job = next(
                (item for item in items if str(item.get("ContextKey") or "") == context_key),
                None,
            )
            if job is not None:
                state = int(job.get("JobState", 0))
                if state == 1:
                    results = job.get("JobResults") or []
                    if not results:
                        raise JobFailedError("Job completed without downloadable results")
                    return job
                if state == 2:
                    raise JobFailedError("Job was cancelled")
            if time.monotonic() >= deadline:
                raise JobTimeoutError(f"Job did not finish within {timeout_seconds:g} seconds")
            self._sleep(self.config.poll_interval_seconds)

    @staticmethod
    def _result_identity(result: dict[str, Any]) -> tuple[int | str, str | None]:
        result_id = result.get("ID", result.get("Id", result.get("id")))
        if result_id in (None, ""):
            raise ProtocolError("Job result does not contain an ID")
        name = result.get("Name") or result.get("FileName") or result.get("File")
        return result_id, str(name) if name else None

    @staticmethod
    def _content_disposition_name(headers: Message) -> str | None:
        value = headers.get("Content-Disposition")
        if not value:
            return None
        message = Message()
        message["content-disposition"] = value
        name = message.get_filename()
        return Path(name).name if name else None

    @staticmethod
    def _unique_destination(directory: Path, file_name: str) -> Path:
        safe_name = Path(file_name).name or "report.xlsx"
        candidate = directory / safe_name
        counter = 1
        while candidate.exists() or candidate.with_suffix(candidate.suffix + ".part").exists():
            candidate = directory / f"{Path(safe_name).stem} ({counter}){Path(safe_name).suffix}"
            counter += 1
        return candidate

    @staticmethod
    def _report_link(html: str, base_url: str) -> str | None:
        parser = _LinksParser()
        parser.feed(html)
        candidates = [urljoin(base_url, link) for link in parser.links]
        return next(
            (
                link
                for link in candidates
                if "reportID=" in link
                or "/Job/Download" in link
                or urlparse(link).path.lower().endswith((".xls", ".xlsx", ".pdf", ".csv"))
            ),
            None,
        )

    @staticmethod
    def _default_report_extension(content_type: str) -> str:
        return {
            "application/pdf": ".pdf",
            "application/vnd.ms-excel": ".xls",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
            "text/csv": ".csv",
        }.get(content_type, ".bin")

    def _save_download(
        self,
        body: bytes,
        headers: Message,
        source_url: str,
        output_dir: Path,
        fallback_name: str,
    ) -> ReportDownloadResult:
        content_type = headers.get_content_type()
        if content_type == "text/html" or body.lstrip().lower().startswith(b"<!doctype html"):
            raise ProtocolError("Report download returned HTML instead of a file")
        name = self._content_disposition_name(headers)
        if not name:
            url_name = Path(urlparse(source_url).path).name
            name = url_name if "." in url_name else fallback_name
        if not Path(name).suffix:
            name += self._default_report_extension(content_type)
        output_dir.mkdir(parents=True, exist_ok=True)
        destination = self._unique_destination(output_dir, Path(name).name)
        temporary = destination.with_suffix(destination.suffix + ".part")
        temporary.write_bytes(body)
        temporary.replace(destination)
        return ReportDownloadResult(
            file_name=destination.name,
            path=destination,
            content_type=content_type,
            source_url=source_url,
        )

    def execute_report(
        self,
        execute_path: str,
        fields: Mapping[str, Any],
        output_dir: str | Path,
        *,
        fallback_name: str = "report",
    ) -> ReportDownloadResult:
        """Generate and download a catalogued read-only report form."""
        self._require_authenticated()
        normalized_path = execute_path
        if not normalized_path.startswith("/"):
            normalized_path = f"{ROOT}/{normalized_path.lstrip('/')}"
        if not normalized_path.endswith("/Execute"):
            normalized_path += "/Execute"
        if normalized_path not in REPORT_EXECUTE_PATHS:
            raise PermissionError(f"Report route is not in the read-only allowlist: {normalized_path}")

        body, headers, final_url = self._post_form(normalized_path, dict(fields))
        content_type = headers.get_content_type()
        destination_dir = Path(output_dir).expanduser().resolve()
        if content_type != "text/html":
            return self._save_download(body, headers, final_url, destination_dir, fallback_name)

        current_html = body.decode("utf-8", errors="replace")
        current_url = final_url
        for _ in range(3):
            link = self._report_link(current_html, current_url)
            if not link:
                raise ProtocolError("Generated report response does not contain a download link")
            next_body, next_headers, next_url = self._request("GET", link)
            if next_headers.get_content_type() != "text/html":
                return self._save_download(
                    next_body,
                    next_headers,
                    next_url,
                    destination_dir,
                    fallback_name,
                )
            current_html = next_body.decode("utf-8", errors="replace")
            current_url = next_url
        raise ProtocolError("Report download link did not resolve to a file")

    def generate_main_account_statement(
        self,
        *,
        customer_id: int | str,
        account_no: str,
        currency_id: int | str,
        start_date: str,
        end_date: str,
        output_dir: str | Path,
        output_format: str = "XLS",
        currency_name: str = "",
    ) -> ReportDownloadResult:
        output_format = output_format.upper()
        if output_format not in {"PDF", "XLS"}:
            raise ValueError("output_format must be PDF or XLS")
        if not account_no.strip():
            raise ValueError("account_no is required")
        fields = {
            "Deposit.AccountNo": account_no.strip(),
            "Deposit.CurrencyName": currency_name,
            "Deposit.CurrencyID": currency_id,
            "Period.StartDate": start_date,
            "Period.EndDate": end_date,
            "FileFormat.Value": output_format,
        }
        return self.execute_report(
            f"{ROOT}/Deposits/MainAccountStatementCommand/Execute",
            fields,
            output_dir,
            fallback_name=f"statement-{self._customer_id(customer_id)}-{account_no}",
        )

    def _download_result(
        self,
        result: dict[str, Any],
        output_dir: Path,
        fallback_name: str,
    ) -> tuple[int | str, str, Path]:
        result_id, result_name = self._result_identity(result)
        body, headers, _ = self._request("GET", f"{JOB_DOWNLOAD}?{urlencode({'id': result_id})}")
        content_type = headers.get_content_type()
        if content_type == "text/html" or not body.startswith(b"PK"):
            raise ProtocolError("Download did not return an XLSX file")
        name = self._content_disposition_name(headers) or result_name or fallback_name
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        output_dir.mkdir(parents=True, exist_ok=True)
        destination = self._unique_destination(output_dir, name)
        temporary = destination.with_suffix(destination.suffix + ".part")
        temporary.write_bytes(body)
        temporary.replace(destination)
        return result_id, destination.name, destination

    def generate_template_report(
        self,
        template_path: str | Path,
        report_date: str | date | datetime,
        output_dir: str | Path,
        *,
        use_cache: bool = False,
        include_formula_as_comment: bool = False,
        timeout_seconds: float = 180.0,
    ) -> TemplateReportResult:
        if not self._authenticated:
            raise AuthenticationError("Call login() before generating a report")
        template = Path(template_path).expanduser().resolve()
        if not template.is_file():
            raise FileNotFoundError(template)
        model = self._create_job_model(
            template,
            report_date,
            use_cache=use_cache,
            include_formula_as_comment=include_formula_as_comment,
        )
        context_key = self._run_job(model)
        job = self._wait_for_job(context_key, timeout_seconds)
        result = (job.get("JobResults") or [])[0]
        result_id, file_name, destination = self._download_result(
            result,
            Path(output_dir).expanduser().resolve(),
            template.name,
        )
        messages = bool(job.get("Result"))
        return TemplateReportResult(
            job_id=job.get("ID", ""),
            result_id=result_id,
            file_name=file_name,
            path=destination,
            completed_with_messages=messages,
        )
