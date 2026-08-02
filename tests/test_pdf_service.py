from pypdf import PdfReader, PdfWriter

from gns_app.services.pdf_service import PdfService


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
