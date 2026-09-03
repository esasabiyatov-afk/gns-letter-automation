from __future__ import annotations

import sys
from types import ModuleType, SimpleNamespace

from gns_app.services import scanner_service


def test_wia_worker_keeps_driver_profile_and_prefers_smaller_transfer(
    monkeypatch,
    tmp_path,
):
    calls: list[tuple[object, ...]] = []

    class Image:
        def SaveFile(self, destination: str) -> None:
            with open(destination, "wb") as output:
                output.write(b"png")

    class Dialog:
        def ShowAcquireImage(self, *args):
            calls.append(args)
            return Image()

    client = ModuleType("win32com.client")
    client.Dispatch = lambda name: Dialog()
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

    destination = tmp_path / "signed-response.png"

    assert scanner_service._wia_worker(destination) == 0
    assert destination.read_bytes() == b"png"
    assert calls == [
        (
            scanner_service.WIA_SCANNER_DEVICE,
            scanner_service.WIA_UNSPECIFIED_INTENT,
            scanner_service.WIA_MINIMIZE_SIZE,
            scanner_service.WIA_FORMAT_PNG,
            False,
            True,
            False,
        )
    ]
