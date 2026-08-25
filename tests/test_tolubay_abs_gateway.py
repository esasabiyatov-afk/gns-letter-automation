from __future__ import annotations

from urllib.error import URLError

import pytest

from gns_app.domain import AbsStatus
from gns_app.services.abs_service import TolubayAbsGateway
from gns_app.vendor.tolubay_hub import (
    AuthenticationError,
    CustomerSummary,
    ProtocolError,
    TolubayConfig,
)


def customer(customer_id: str, identity: str) -> CustomerSummary:
    return CustomerSummary(
        customer_id=customer_id,
        name="Обезличенный клиент",
        identity=identity,
        birth_date="",
        address="",
        phones="",
    )


class StubTolubayClient:
    def __init__(self, searches=None, accounts=None):
        self.searches = searches or {}
        self.accounts = accounts or {}
        self.login_count = 0
        self.account_requests: list[tuple[str, bool]] = []

    def login(self, username: str, password: str) -> None:
        assert username == "employee"
        assert password == "one-time-secret"
        self.login_count += 1

    def search_customers(self, criteria, *, page_size=50):
        assert page_size == 100
        return self.searches.get(criteria["SearchIdentificationNo"], [])

    def get_accounts(self, customer_id: str, *, include_closed=False):
        self.account_requests.append((customer_id, include_closed))
        return self.accounts.get(customer_id, [])


def gateway(client) -> TolubayAbsGateway:
    return TolubayAbsGateway(
        TolubayConfig(base_url="https://abs.example.test"),
        client_factory=lambda config: client,
    )


def test_tolubay_gateway_distinguishes_questionnaire_from_absent_customer():
    first_inn = "12345678901234"
    second_inn = "23456789012345"
    client = StubTolubayClient(
        searches={
            first_inn: [customer("42", first_inn)],
            second_inn: [],
        },
    )

    result = gateway(client).check(
        "employee",
        "one-time-secret",
        [
            {"inn": first_inn, "name": "Первый"},
            {"inn": second_inn, "name": "Второй"},
        ],
    )

    assert result.status == AbsStatus.FOUND
    assert not result.is_fake
    assert [item["result"] for item in result.taxpayers] == [
        AbsStatus.FOUND,
        AbsStatus.NOT_FOUND,
    ]
    assert not client.account_requests


def test_tolubay_gateway_requires_one_exact_identity_match():
    inn = "12345678901234"
    client = StubTolubayClient(
        searches={
            inn: [
                customer("42", inn),
                customer("43", "99999999999999"),
            ]
        }
    )

    result = gateway(client).check(
        "employee",
        "one-time-secret",
        [{"inn": inn, "name": "Клиент"}],
    )

    assert result.status == AbsStatus.MULTIPLE
    assert result.taxpayers[0]["result"] == AbsStatus.MULTIPLE
    assert not client.account_requests


def test_tolubay_gateway_routes_exact_questionnaire_to_manual_processing():
    inn = "12345678901234"
    client = StubTolubayClient(
        searches={inn: [customer("42", inn)]},
    )

    result = gateway(client).check(
        "employee",
        "one-time-secret",
        [{"inn": inn, "name": "Клиент"}],
    )

    assert result.status == AbsStatus.FOUND
    assert result.taxpayers[0]["result"] == AbsStatus.FOUND
    assert not client.account_requests


def test_tolubay_gateway_does_not_treat_transport_error_as_not_found():
    class UnavailableClient(StubTolubayClient):
        def login(self, username: str, password: str) -> None:
            try:
                raise URLError(TimeoutError("offline"))
            except URLError as exc:
                raise ProtocolError("transport failed") from exc

    result = gateway(UnavailableClient()).check(
        "employee",
        "one-time-secret",
        [{"inn": "12345678901234", "name": "Клиент"}],
    )

    assert result.status == AbsStatus.UNAVAILABLE
    assert not result.taxpayers


def test_tolubay_gateway_maps_rejected_login_without_echoing_credentials():
    class RejectedClient(StubTolubayClient):
        def login(self, username: str, password: str) -> None:
            raise AuthenticationError("rejected")

    result = gateway(RejectedClient()).check(
        "employee",
        "one-time-secret",
        [{"inn": "12345678901234", "name": "Клиент"}],
    )

    assert result.status == AbsStatus.AUTH_ERROR
    assert "employee" not in result.message
    assert "one-time-secret" not in result.message


def test_tolubay_gateway_rejects_unconfirmed_inn_before_network_call():
    created = False

    def create_client(config):
        nonlocal created
        created = True
        return StubTolubayClient()

    adapter = TolubayAbsGateway(
        TolubayConfig(base_url="https://abs.example.test"),
        client_factory=create_client,
    )

    result = adapter.check(
        "employee",
        "one-time-secret",
        [{"inn": "123", "name": "Клиент"}],
    )

    assert result.status == AbsStatus.TECHNICAL_ERROR
    assert not created


def test_tolubay_gateway_forbids_disabled_tls_verification():
    with pytest.raises(ValueError, match="TLS"):
        TolubayAbsGateway(
            TolubayConfig(
                base_url="https://abs.example.test",
                verify_tls=False,
            )
        )
