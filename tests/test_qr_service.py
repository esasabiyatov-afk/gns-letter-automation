from pathlib import Path

import pytest
from PIL import Image

from gns_app.domain import QrStatus
from gns_app.services.pdf_service import PdfService
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


def test_zbar_fallback_accepts_a_valid_official_qr(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    image_path = tmp_path / "small.png"
    Image.new("L", (32, 32), 255).save(image_path)
    payload = "https://qr.salyk.kg/getsti010decission?encodedText=secret"
    service = QrService(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
    )
    monkeypatch.setattr(service, "_try_decode", lambda *args: None)
    monkeypatch.setattr(
        service,
        "_try_decode_zbar",
        lambda *args: (payload, "zbar-original"),
    )

    result = service.decode(image_path)

    assert result.status == QrStatus.FOUND
    assert result.method == "zbar-original"


def test_conflicting_decoders_require_manual_review(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    image_path = tmp_path / "conflict.png"
    Image.new("L", (32, 32), 255).save(image_path)
    service = QrService(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
    )
    monkeypatch.setattr(
        service,
        "_try_decode",
        lambda *args: (
            "https://qr.salyk.kg/getsti010decission?encodedText=one",
            "zxing",
        ),
    )
    monkeypatch.setattr(
        service,
        "_try_decode_zbar",
        lambda *args: (
            "https://qr.salyk.kg/getsti010decission?encodedText=two",
            "zbar",
        ),
    )

    result = service.decode(image_path)

    assert result.status == QrStatus.DECODE_ERROR
    assert result.method == "zxing-zbar-conflict"


def test_high_resolution_retry_uses_zbar_fallback(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    image_path = tmp_path / "high-resolution.png"
    Image.new("L", (64, 64), 255).save(image_path)
    payload = "https://qr.salyk.kg/getsti010decission?encodedText=secret"
    service = QrService(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
    )
    monkeypatch.setattr(service, "_try_decode", lambda *args: None)
    monkeypatch.setattr(
        service,
        "_try_decode_zbar",
        lambda *args: (payload, "zbar-highres-original"),
    )

    result = service.decode_high_resolution(image_path)

    assert result.status == QrStatus.FOUND
    assert result.method == "zbar-highres-original"


def test_high_resolution_retry_reads_difficult_qr(
    project_root: Path, tmp_path: Path
):
    source = project_root / "УГНС 1611.pdf"
    if not source.is_file():
        pytest.skip("Локальный регрессионный PDF отсутствует")
    image = PdfService().render_page_for_qr(
        source,
        65,
        tmp_path / "page-65-qr.png",
    )
    service = QrService(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
    )

    result = service.decode_high_resolution(image)

    assert result.status == QrStatus.FOUND
    assert result.safe_url == "https://qr.salyk.kg/getsti010decission"
