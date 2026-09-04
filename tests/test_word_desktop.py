from __future__ import annotations

import sys
from types import SimpleNamespace

from gns_app.services import word_desktop


def test_opens_document_with_word_and_restores_minimized_window(
    monkeypatch,
    tmp_path,
):
    source = tmp_path / "response.docx"
    source.write_bytes(b"docx")
    events: list[object] = []

    class Document:
        FullName = str(source)

        def Activate(self):
            events.append("document.activate")

    document = Document()

    class Documents:
        Count = 0

        def Open(self, path):
            events.append(("documents.open", path))
            return document

    class Application:
        Visible = False
        WindowState = word_desktop.WD_WINDOW_STATE_MINIMIZE

        def Activate(self):
            events.append("application.activate")

    application = Application()
    application.Documents = Documents()
    client = SimpleNamespace(
        GetActiveObject=lambda name: (_ for _ in ()).throw(OSError()),
        Dispatch=lambda name: application,
    )
    pythoncom = SimpleNamespace(
        CoInitialize=lambda: events.append("co.init"),
        CoUninitialize=lambda: events.append("co.uninit"),
    )
    monkeypatch.setitem(sys.modules, "pythoncom", pythoncom)
    monkeypatch.setitem(
        sys.modules,
        "win32com",
        SimpleNamespace(client=client),
    )
    monkeypatch.setattr(word_desktop.os, "name", "nt")

    word_desktop.open_word_document(source)

    assert application.Visible is True
    assert application.WindowState == word_desktop.WD_WINDOW_STATE_NORMAL
    assert events == [
        "co.init",
        ("documents.open", str(source.resolve())),
        "document.activate",
        "application.activate",
        "application.activate",
        "co.uninit",
    ]
