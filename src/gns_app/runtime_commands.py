from __future__ import annotations

import sys
from pathlib import Path


def module_command(module_name: str, *arguments: object) -> list[str]:
    """Build a module command that also works inside a frozen executable."""

    tail = [str(argument) for argument in arguments]
    if getattr(sys, "frozen", False):
        worker_executable = Path(sys.executable).with_name("GNS-Worker.exe")
        executable = (
            str(worker_executable)
            if worker_executable.is_file()
            else sys.executable
        )
        return [
            executable,
            "--worker-module",
            module_name,
            *tail,
        ]
    return [sys.executable, "-m", module_name, *tail]
