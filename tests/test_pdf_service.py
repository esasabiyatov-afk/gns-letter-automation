from pathlib import Path
from subprocess import CompletedProcess

import pytest
from pypdf import PdfReader, PdfWriter

from gns_app.services.pdf_service import PdfProcessingError, PdfService


def test_extract_page_pdf_keeps_only_requested_page(tmp_path):
    source = tmp_path / "source.pdf"
    output = tmp_path / "page-0002.pdf"
    writer = PdfWriter()
    writer.add_blank_page(width=100, height=140)
    writer.add_blank_page(width=220, height=310)
    with source.open("wb") as stream:
        writer.write(stream)

    PdfService().extract_page_pdf(source, 2, output)

    reader = PdfReader(output)
    assert len(reader.pages) == 1
    assert float(reader.pages[0].mediabox.width) == 220
    assert float(reader.pages[0].mediabox.height) == 310


def test_native_pdfium_crash_does_not_escape_renderer_process(
    tmp_path: Path,
    monkeypatch,
):
    source = tmp_path / "source.pdf"
    output = tmp_path / "preview.jpg"
    enhanced = tmp_path / "enhanced.jpg"
    source.write_bytes(b"%PDF-test")

    monkeypatch.setattr(
        "gns_app.services.pdf_service.subprocess.run",
        lambda *args, **kwargs: CompletedProcess(args[0], -1073741819, "", ""),
    )

    with pytest.raises(PdfProcessingError, match="сервер продолжает работу"):
        PdfService().render_page(source, 1, output, enhanced)

    assert not output.exists()
    assert not enhanced.exists()
