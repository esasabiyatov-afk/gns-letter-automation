from gns_app.domain import QrStatus
from gns_app.services.qr_service import QrService


def test_accepts_only_expected_qr_endpoint():
    service = QrService(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
    )
    result = service._validate(
        "https://qr.salyk.kg/getsti010decission?encodedText=secret",
        "test",
    )
    assert result.status == QrStatus.FOUND
    assert result.safe_url == "https://qr.salyk.kg/getsti010decission"
    assert "secret" not in result.safe_url


def test_rejects_untrusted_host():
    service = QrService(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
    )
    result = service._validate(
        "https://example.com/getsti010decission?encodedText=x",
        "test",
    )
    assert result.status == QrStatus.INVALID_URL

