from __future__ import annotations

import sys
from contextlib import nullcontext

from gns_app import launcher
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


def test_frozen_commands_use_hidden_app_and_separate_worker(
    monkeypatch,
    tmp_path,
):
    app_executable = tmp_path / "GNS-Portable.exe"
    worker_executable = tmp_path / "GNS-Worker.exe"
    worker_executable.write_bytes(b"worker")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(app_executable))

    assert _server_command() == [str(app_executable), "--server"]
    assert module_command("gns_app.services.scanner_service", "--wia-worker") == [
        str(worker_executable),
        "--worker-module",
        "gns_app.services.scanner_service",
        "--wia-worker",
    ]


def test_background_start_does_not_open_browser(monkeypatch):
    calls = []
    monkeypatch.setattr(
        launcher,
        "WindowsSingleInstance",
        lambda: nullcontext(),
    )
    monkeypatch.setattr(
        launcher,
        "_run_server_process",
        lambda command, *, open_browser: (
            calls.append((command, open_browser)) or 0
        ),
    )

    assert launcher.main(["--background"]) == 0
    assert calls == [(_server_command(), False)]
