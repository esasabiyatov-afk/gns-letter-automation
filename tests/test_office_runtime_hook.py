from __future__ import annotations

import os
import runpy
from pathlib import Path


HOOK = (
    Path(__file__).resolve().parents[1]
    / "packaging"
    / "office_runtime_hook.py"
)


def test_office_runtime_hook_enables_confirmed_office_integrations(
    monkeypatch,
):
    for name in (
        "GNS_ABS_MODE",
        "GNS_TOLUBAY_VERIFY_TLS",
        "GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE",
        "GNS_PROCESSING_WORKERS",
    ):
        monkeypatch.delenv(name, raising=False)

    runpy.run_path(str(HOOK))

    assert os.environ["GNS_ABS_MODE"] == "tolubay"
    assert os.environ["GNS_TOLUBAY_VERIFY_TLS"] == "false"
    assert os.environ["GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE"] == "true"
    assert os.environ["GNS_PROCESSING_WORKERS"] == "1"


def test_office_runtime_hook_does_not_override_explicit_configuration(
    monkeypatch,
):
    monkeypatch.setenv("GNS_TOLUBAY_VERIFY_TLS", "true")
    monkeypatch.setenv("GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE", "false")

    runpy.run_path(str(HOOK))

    assert os.environ["GNS_TOLUBAY_VERIFY_TLS"] == "true"
    assert os.environ["GNS_OUTLOOK_ALLOW_INSECURE_CERTIFICATE"] == "false"
