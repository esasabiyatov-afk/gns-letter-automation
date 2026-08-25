from __future__ import annotations

import sys

from gns_app.launcher import _server_command, supervise
from gns_app.runtime_commands import module_command


def test_supervisor_restarts_after_crash_and_stops_after_clean_exit():
    exit_codes = iter([3, 0])
    waits: list[float] = []

    result = supervise(
        lambda: next(exit_codes),
        wait=waits.append,
        restart_delay_seconds=2,
    )

    assert result == 0
    assert waits == [2]


def test_source_module_command_uses_python_module(monkeypatch):
    monkeypatch.delattr(sys, "frozen", raising=False)

    command = module_command("gns_app.services.pdf_service", "--flag")

    assert command == [
        sys.executable,
        "-m",
        "gns_app.services.pdf_service",
        "--flag",
    ]


def test_frozen_commands_reenter_the_executable(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)

    assert _server_command() == [sys.executable, "--server"]
    assert module_command("gns_app.services.scanner_service", "--wia-worker") == [
        sys.executable,
        "--worker-module",
        "gns_app.services.scanner_service",
        "--wia-worker",
    ]
