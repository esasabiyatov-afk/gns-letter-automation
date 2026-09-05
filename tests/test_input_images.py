from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest
from PIL import Image
from pypdf import PdfReader

from gns_app.domain import (
    ClassificationResult,
    OcrResult,
    OcrStatus,
    PageType,
    QrDecodeResult,
    QrStatus,
)
from gns_app.services.storage import StorageError


@pytest.mark.parametrize(
    ("extension", "image_format"),
    (("bmp", "BMP"), ("png", "PNG"), ("jpg", "JPEG"), ("jpeg", "JPEG")),
)
def test_image_upload_preserves_original_and_creates_working_pdf(
    workflow,
    extension: str,
    image_format: str,
):
    payload = BytesIO()
    Image.new("RGB", (240, 320), "white").save(payload, format=image_format)
    original = payload.getvalue()
    payload.seek(0)

    upload_id = workflow.create_upload(f"Письмо.{extension}", payload)

    upload = workflow.get_upload(upload_id)
    stored_path = Path(upload["stored_path"])
    assert upload["page_count"] == 1
    assert len(workflow.get_upload_pages(upload_id)) == 1
    assert (stored_path.parent / f"source.{extension}").read_bytes() == original
    assert len(PdfReader(stored_path).pages) == 1


def test_image_extension_must_match_content(workflow, sample_pdf):
    with sample_pdf.open("rb") as stream:
        with pytest.raises(StorageError, match="входящее изображение"):
            workflow.create_upload("Подмена.png", stream)


def test_image_processing_reads_qr_and_ocr_from_original_pixels(
    workflow,
    monkeypatch,
):
    payload = BytesIO()
    Image.new("RGB", (240, 320), "white").save(payload, format="PNG")
    payload.seek(0)
    upload_id = workflow.create_upload("Письмо.png", payload)
    ocr_images: list[Path] = []

    def decode(path: Path) -> QrDecodeResult:
        if path.name == "source.png":
            return QrDecodeResult(
                status=QrStatus.FOUND,
                payload="https://qr.salyk.kg/getsti010decission?id=image",
                payload_hash="image-qr",
                safe_url="https://qr.salyk.kg/getsti010decission",
                method="test-original",
            )
        return QrDecodeResult(status=QrStatus.NOT_FOUND)

    def recognize(_pdf, _page, image_path, **_kwargs):
        ocr_images.append(Path(image_path))
        return OcrResult(
            status=OcrStatus.COMPLETED,
            text="Письмо",
            confidence=0.99,
            language="rus",
        )

    monkeypatch.setattr(workflow.qr, "decode", decode)
    monkeypatch.setattr(workflow.ocr, "recognize", recognize)
    monkeypatch.setattr(
        workflow.classifier,
        "classify",
        lambda *_args, **_kwargs: ClassificationResult(
            page_type=PageType.LETTER,
            confidence=0.99,
        ),
    )

    workflow.process_upload(upload_id)

    page = workflow.get_upload_pages(upload_id)[0]
    assert page["qr_status"] == QrStatus.FOUND
    assert page["qr_method"] == "original-image-test-original"
    assert [path.name for path in ocr_images] == ["source.png"]
    assert ocr_images[0].is_file()
