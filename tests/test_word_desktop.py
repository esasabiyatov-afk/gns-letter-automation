from __future__ import annotations

from gns_app.services import word_desktop


def test_opens_document_with_windows_file_association(monkeypatch, tmp_path):
    source = tmp_path / "response.docx"
    source.write_bytes(b"docx")
    opened: list[str] = []
    monkeypatch.setattr(word_desktop.os, "name", "nt")
    monkeypatch.setattr(
        word_desktop.os,
        "startfile",
        lambda path: opened.append(path),
        raising=False,
    )

    word_desktop.open_word_document(source)

    assert opened == [str(source.resolve())]
