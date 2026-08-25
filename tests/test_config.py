from __future__ import annotations

import pytest

from gns_app.config import Settings


def test_tolubay_tls_verification_is_enabled_by_default(monkeypatch):
    monkeypatch.delenv("GNS_TOLUBAY_VERIFY_TLS", raising=False)

    settings = Settings.load()

    assert settings.tolubay_verify_tls


def test_tolubay_tls_verification_can_be_explicitly_disabled(monkeypatch):
    monkeypatch.setenv("GNS_TOLUBAY_VERIFY_TLS", "false")

    settings = Settings.load()

    assert not settings.tolubay_verify_tls


def test_tolubay_tls_verification_rejects_unknown_value(monkeypatch):
    monkeypatch.setenv("GNS_TOLUBAY_VERIFY_TLS", "sometimes")

    with pytest.raises(ValueError, match="true или false"):
        Settings.load()


def test_outlook_office_test_flags_are_disabled_by_default(monkeypatch):
    monkeypatch.delenv("GNS_OUTLOOK_TEST_EMAIL", raising=False)
    monkeypatch.delenv("GNS_OUTLOOK_ALLOW_TEST_SEND", raising=False)
    monkeypatch.delenv(
        "GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE", raising=False
    )

    settings = Settings.load()

    assert settings.outlook_test_email == ""
    assert not settings.outlook_allow_test_send
    assert not settings.outlook_allow_insecure_certificate


def test_outlook_office_test_flags_can_be_explicitly_enabled(monkeypatch):
    monkeypatch.setenv("GNS_OUTLOOK_TEST_EMAIL", "ESASABIYATOV@GMAIL.COM")
    monkeypatch.setenv("GNS_OUTLOOK_ALLOW_TEST_SEND", "true")
    monkeypatch.setenv("GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE", "true")

    settings = Settings.load()

    assert settings.outlook_test_email == "esasabiyatov@gmail.com"
    assert settings.outlook_allow_test_send
    assert settings.outlook_allow_insecure_certificate


def test_outlook_test_email_rejects_every_other_recipient(monkeypatch):
    monkeypatch.setenv("GNS_OUTLOOK_TEST_EMAIL", "real.person@example.com")

    with pytest.raises(ValueError, match="согласованного тестового адреса"):
        Settings.load()
