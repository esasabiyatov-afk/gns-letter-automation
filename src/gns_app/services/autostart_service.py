from __future__ import annotations

import sys
from pathlib import Path
from typing import Any


class AutostartError(RuntimeError):
    """Safe user-facing error for the per-user Windows startup entry."""


class WindowsAutostartService:
    RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
    VALUE_NAME = "GNSLetterAutomation"

    def __init__(
        self,
        *,
        executable_path: Path | None = None,
        frozen: bool | None = None,
        platform: str | None = None,
        registry: Any | None = None,
    ) -> None:
        self.executable_path = Path(
            executable_path or sys.executable
        ).resolve()
        self.frozen = (
            bool(getattr(sys, "frozen", False))
            if frozen is None
            else frozen
        )
        self.platform = sys.platform if platform is None else platform
        self._registry_override = registry

    def supported(self) -> bool:
        return (
            self.platform == "win32"
            and self.frozen
            and self.executable_path.suffix.casefold() == ".exe"
        )

    def command(self) -> str:
        return f'"{self.executable_path}" --background'

    def _registry(self):
        if self._registry_override is not None:
            return self._registry_override
        import winreg

        return winreg

    def _read_command(self) -> str:
        registry = self._registry()
        try:
            with registry.OpenKey(
                registry.HKEY_CURRENT_USER,
                self.RUN_KEY,
                0,
                registry.KEY_READ,
            ) as key:
                value, _value_type = registry.QueryValueEx(
                    key,
                    self.VALUE_NAME,
                )
        except FileNotFoundError:
            return ""
        return str(value).strip()

    def _write_command(self) -> None:
        registry = self._registry()
        with registry.CreateKeyEx(
            registry.HKEY_CURRENT_USER,
            self.RUN_KEY,
            0,
            registry.KEY_SET_VALUE,
        ) as key:
            registry.SetValueEx(
                key,
                self.VALUE_NAME,
                0,
                registry.REG_SZ,
                self.command(),
            )

    def _delete_command(self) -> None:
        registry = self._registry()
        try:
            with registry.OpenKey(
                registry.HKEY_CURRENT_USER,
                self.RUN_KEY,
                0,
                registry.KEY_SET_VALUE,
            ) as key:
                registry.DeleteValue(key, self.VALUE_NAME)
        except FileNotFoundError:
            return

    def status(self) -> dict[str, bool | str]:
        if not self.supported():
            return {
                "supported": False,
                "enabled": False,
                "needs_update": False,
                "error": "",
            }
        try:
            stored = self._read_command()
        except OSError:
            return {
                "supported": True,
                "enabled": False,
                "needs_update": False,
                "error": "Не удалось прочитать автозапуск Windows.",
            }
        enabled = stored.casefold() == self.command().casefold()
        return {
            "supported": True,
            "enabled": enabled,
            "needs_update": bool(stored and not enabled),
            "error": "",
        }

    def set_enabled(self, enabled: bool) -> dict[str, bool | str]:
        if not self.supported():
            raise AutostartError(
                "Автозапуск доступен в портативной EXE-сборке Windows."
            )
        try:
            if enabled:
                self._write_command()
            else:
                self._delete_command()
        except OSError as exc:
            raise AutostartError(
                "Windows не разрешила изменить автозапуск."
            ) from exc
        return self.status()
