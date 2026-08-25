from __future__ import annotations

import sys


def module_command(module_name: str, *arguments: object) -> list[str]:
    """Build a module command that also works inside a frozen executable."""

    tail = [str(argument) for argument in arguments]
    if getattr(sys, "frozen", False):
        return [
            sys.executable,
            "--worker-module",
            module_name,
            *tail,
        ]
    return [sys.executable, "-m", module_name, *tail]
