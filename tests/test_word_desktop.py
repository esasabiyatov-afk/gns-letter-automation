from __future__ import annotations

from types import SimpleNamespace

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


def test_prints_document_with_windows_file_association(monkeypatch, tmp_path):
    source = tmp_path / "response.docx"
    source.write_bytes(b"docx")
    printed: list[tuple[str, str]] = []
    monkeypatch.setattr(word_desktop.os, "name", "nt")
    monkeypatch.setattr(
        word_desktop.os,
        "startfile",
        lambda path, verb: printed.append((path, verb)),
        raising=False,
    )

    word_desktop.print_word_document(source)

    assert printed == [(str(source.resolve()), "print")]


def test_exact_page_count_uses_isolated_word_worker(monkeypatch, tmp_path):
    source = tmp_path / "response.docx"
    source.write_bytes(b"docx")
    captured = {}
    monkeypatch.setattr(word_desktop.os, "name", "nt")
    monkeypatch.setattr(
        word_desktop,
        "module_command",
        lambda *parts: [str(part) for part in parts],
    )

    def run(command, **options):
        captured.update(command=command, options=options)
        return SimpleNamespace(returncode=0, stdout='{"page_count": 3}')

    monkeypatch.setattr(word_desktop.subprocess, "run", run)

    assert word_desktop.exact_word_page_count(source) == 3
    assert captured["command"][-1] == str(source.resolve())
