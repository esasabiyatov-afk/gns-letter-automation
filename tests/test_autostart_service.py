from gns_app.services.autostart_service import (
    AutostartError,
    WindowsAutostartService,
)


def test_portable_autostart_can_be_enabled_updated_and_disabled(
    tmp_path,
    monkeypatch,
):
    executable = tmp_path / "GNS-Portable.exe"
    executable.write_bytes(b"exe")
    service = WindowsAutostartService(
        executable_path=executable,
        frozen=True,
        platform="win32",
    )
    stored = {"command": ""}
    monkeypatch.setattr(
        service,
        "_read_command",
        lambda: stored["command"],
    )
    monkeypatch.setattr(
        service,
        "_write_command",
        lambda: stored.update(command=service.command()),
    )
    monkeypatch.setattr(
        service,
        "_delete_command",
        lambda: stored.update(command=""),
    )

    assert not service.status()["enabled"]
    assert service.set_enabled(True)["enabled"]
    assert stored["command"] == f'"{executable.resolve()}" --background'

    stored["command"] = '"C:\\Old\\GNS-Portable.exe" --background'
    assert service.status()["needs_update"]
    assert service.set_enabled(True)["enabled"]

    assert not service.set_enabled(False)["enabled"]
    assert stored["command"] == ""


def test_autostart_is_rejected_outside_portable_windows(tmp_path):
    service = WindowsAutostartService(
        executable_path=tmp_path / "python.exe",
        frozen=False,
        platform="win32",
    )

    assert not service.status()["supported"]
    try:
        service.set_enabled(True)
    except AutostartError:
        pass
    else:
        raise AssertionError("source mode must not change Windows autostart")
