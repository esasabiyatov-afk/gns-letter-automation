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
