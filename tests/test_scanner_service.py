from __future__ import annotations

import json
import sys
from types import ModuleType, SimpleNamespace

from gns_app.services import scanner_service


def _install_wia_modules(monkeypatch, dialog) -> None:
    client = ModuleType("win32com.client")
    client.Dispatch = lambda name: dialog
    win32com = ModuleType("win32com")
    win32com.__path__ = []
    win32com.client = client
    pythoncom = SimpleNamespace(
        CoInitialize=lambda: None,
        CoUninitialize=lambda: None,
    )
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)
    monkeypatch.setattr(
        scanner_service,
        "focus_next_dialog_for_current_process",
        lambda: None,
    )


def test_wia_worker_applies_saved_profile_and_prefers_smaller_transfer(
    monkeypatch,
    tmp_path,
):
    calls: list[tuple[str, tuple[object, ...]]] = []

    class Property:
        def __init__(self) -> None:
            self.Value = None

    properties = {
        scanner_service.WIA_IPS_CUR_INTENT: Property(),
        scanner_service.WIA_IPS_XRES: Property(),
        scanner_service.WIA_IPS_YRES: Property(),
    }

    class Properties:
        def Item(self, key):
            return properties[int(key)]

    item = SimpleNamespace(Properties=Properties())

    class Items:
        def Item(self, index: int):
            assert index == 1
            return item

    device = SimpleNamespace(Items=Items())

    class Image:
        def SaveFile(self, destination: str) -> None:
            with open(destination, "wb") as output:
                output.write(b"png")

    class Dialog:
        def ShowSelectDevice(self, *args):
            calls.append(("select", args))
            return device

        def ShowTransfer(self, *args):
            calls.append(("transfer", args))
            return Image()

    _install_wia_modules(monkeypatch, Dialog())
    destination = tmp_path / "signed-response.png"

    assert scanner_service._wia_worker(destination, 200, "color") == 0
    assert destination.read_bytes() == b"png"
    assert properties[scanner_service.WIA_IPS_CUR_INTENT].Value == (
        scanner_service.SCANNER_COLOR_INTENTS["color"]
        | scanner_service.WIA_MINIMIZE_SIZE
    )
    assert properties[scanner_service.WIA_IPS_XRES].Value == 200
    assert properties[scanner_service.WIA_IPS_YRES].Value == 200
    assert calls == [
        (
            "select",
            (scanner_service.WIA_SCANNER_DEVICE, False, False),
        ),
        (
            "transfer",
            (item, scanner_service.WIA_FORMAT_PNG, False),
        ),
    ]


def test_wia_worker_reports_unsupported_profile_property(
    monkeypatch,
    tmp_path,
    capsys,
):
    class Property:
        Value = None

    class Properties:
        def Item(self, key):
            property_id = int(key)
            if property_id == scanner_service.WIA_IPS_YRES:
                raise KeyError(property_id)
            return Property()

        def __getitem__(self, key):
            raise KeyError(key)

        def __iter__(self):
            return iter(())

    item = SimpleNamespace(Properties=Properties())
    device = SimpleNamespace(
        Items=SimpleNamespace(Item=lambda index: item),
    )

    class Dialog:
        def ShowSelectDevice(self, *args):
            return device

        def ShowTransfer(self, *args):
            raise AssertionError("transfer must not start")

    _install_wia_modules(monkeypatch, Dialog())

    assert scanner_service._wia_worker(
        tmp_path / "scan.png",
        150,
        "grayscale",
    ) == 1
    message = json.loads(capsys.readouterr().out)["message"]
    assert "разрешение по вертикали" in message
    assert "загрузите готовый файл" in message


def test_scanner_service_passes_profile_to_worker(monkeypatch, tmp_path):
    commands: list[list[str]] = []
    destination = tmp_path / "scan.png"

    monkeypatch.setattr(scanner_service.os, "name", "nt")
    monkeypatch.setattr(
        scanner_service,
        "module_command",
        lambda module, *args: [module, *args],
    )

    def fake_run(command, **kwargs):
        commands.append(command)
        destination.write_bytes(b"png")
        return SimpleNamespace(returncode=0, stdout='{"message": "ok"}')

    monkeypatch.setattr(scanner_service.subprocess, "run", fake_run)

    result = scanner_service.ScannerService().acquire_a4(
        destination,
        dpi=300,
        color_mode="black_white",
    )

    assert result == destination
    assert commands == [[
        "gns_app.services.scanner_service",
        "--wia-worker",
        str(destination),
        "--dpi",
        "300",
        "--color-mode",
        "black_white",
    ]]
