from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from threading import Barrier, Lock

from gns_app.database import Database, utc_now
from gns_app.domain import (
    ClassificationResult,
    OcrResult,
    OcrStatus,
    PageStatus,
    PageType,
    QrDecodeResult,
    QrStatus,
)
from gns_app.services.workflow import WorkflowService


def test_complete_official_qr_skips_ocr_and_high_resolution_render(
    test_settings,
    sample_pdf: Path,
    monkeypatch,
):
    settings = replace(
        test_settings,
        auto_download_official=True,
        processing_workers=1,
    )
    database = Database(settings.database_path)
    database.initialize()
    workflow = WorkflowService(database, settings)
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
    upload = workflow.get_upload(upload_id)
    page = workflow.get_upload_pages(upload_id)[0]

    monkeypatch.setattr(
        workflow.qr,
        "decode",
        lambda _path: QrDecodeResult(
            status=QrStatus.FOUND,
            payload="https://qr.salyk.kg/getsti010decission?id=test",
            payload_hash="complete-official",
            safe_url="https://qr.salyk.kg/getsti010decission",
            method="test",
        ),
    )
    monkeypatch.setattr(
        workflow.official,
        "download",
        lambda *_args: sample_pdf,
    )
    monkeypatch.setattr(
        workflow,
        "_apply_official_document",
        lambda *_args: True,
    )
    monkeypatch.setattr(
        workflow.ocr,
        "recognize",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("OCR не должен запускаться")
        ),
    )
    monkeypatch.setattr(
        workflow.pdf,
        "render_page_high_resolution",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            AssertionError("Высокий рендер не должен запускаться")
        ),
    )

    workflow._process_page(upload, page)

    processed = workflow.get_page(page["id"])
    assert processed["status"] == PageStatus.COMPLETED
    assert processed["ocr_status"] == OcrStatus.SKIPPED_OFFICIAL
    assert processed["extracted_text"] == ""


def test_failed_qr_and_ocr_share_one_high_resolution_color_render(
    workflow,
    sample_pdf: Path,
    monkeypatch,
):
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
    upload = workflow.get_upload(upload_id)
    page = workflow.get_upload_pages(upload_id)[0]
    high_resolution_calls: list[Path] = []
    consumed_paths: list[Path] = []
    original_render = workflow.pdf.render_page_high_resolution

    def render_high_resolution(pdf_path, page_number, output_path):
        high_resolution_calls.append(output_path)
        return original_render(pdf_path, page_number, output_path)

    monkeypatch.setattr(
        workflow.qr,
        "decode",
        lambda _path: QrDecodeResult(status=QrStatus.NOT_FOUND),
    )

    def decode_high_resolution(path):
        consumed_paths.append(path)
        return QrDecodeResult(status=QrStatus.NOT_FOUND)

    monkeypatch.setattr(
        workflow.qr,
        "decode_high_resolution",
        decode_high_resolution,
    )
    monkeypatch.setattr(
        workflow.pdf,
        "render_page_high_resolution",
        render_high_resolution,
    )

    def recognize(_pdf, _page, image_path, **_kwargs):
        consumed_paths.append(image_path)
        return OcrResult(
            status=OcrStatus.LOW_CONFIDENCE,
            text="",
            confidence=0.0,
            language="rus+kir fast",
        )

    monkeypatch.setattr(workflow.ocr, "recognize", recognize)

    workflow._process_page(upload, page)

    assert len(high_resolution_calls) == 1
    assert consumed_paths == [
        high_resolution_calls[0],
        high_resolution_calls[0],
    ]
    assert not high_resolution_calls[0].exists()


def test_unknown_page_keeps_structured_ocr_fields_as_manual_hints(
    workflow,
    sample_pdf: Path,
    monkeypatch,
):
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
    upload = workflow.get_upload(upload_id)
    page = workflow.get_upload_pages(upload_id)[0]
    inn = "12345678901234"
    text = (
        "о налогоплательщике:\n"
        "Наименование: Тестовый налогоплательщик\n"
        f"ИНН: {inn}\n"
        "Период: с 01.01.2020 по 01.01.2026"
    )

    monkeypatch.setattr(
        workflow.qr,
        "decode",
        lambda _path: QrDecodeResult(status=QrStatus.NOT_FOUND),
    )
    monkeypatch.setattr(
        workflow.qr,
        "decode_high_resolution",
        lambda _path: QrDecodeResult(status=QrStatus.NOT_FOUND),
    )
    monkeypatch.setattr(
        workflow.ocr,
        "recognize",
        lambda *_args, **_kwargs: OcrResult(
            status=OcrStatus.COMPLETED,
            text=text,
            confidence=0.72,
            language="rus+kir fast",
        ),
    )
    monkeypatch.setattr(
        workflow.classifier,
        "classify",
        lambda *_args, **_kwargs: ClassificationResult(
            page_type=PageType.UNKNOWN,
            confidence=0.0,
        ),
    )

    workflow._process_page(upload, page)

    processed = workflow.get_page(page["id"])
    assert processed["page_type"] == PageType.UNKNOWN
    assert processed["status"] == PageStatus.NEEDS_REVIEW
    assert processed["case_id"]
    taxpayers = workflow.get_taxpayers(processed["case_id"])
    assert len(taxpayers) == 1
    assert taxpayers[0]["inn"] == inn
    assert not taxpayers[0]["manually_confirmed"]
    case = workflow.get_case(processed["case_id"])
    assert case["period_start"] == "2020-01-01"
    assert case["period_end"] == "2026-01-01"


def test_two_page_workers_run_concurrently_and_keep_terminal_statuses(
    test_settings,
    sample_pdf: Path,
    monkeypatch,
):
    settings = replace(test_settings, processing_workers=2)
    database = Database(settings.database_path)
    database.initialize()
    workflow = WorkflowService(database, settings)
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)

    barrier = Barrier(2, timeout=5)
    state_lock = Lock()
    active = 0
    maximum_active = 0

    def process_page(_upload, page):
        nonlocal active, maximum_active
        with state_lock:
            active += 1
            maximum_active = max(maximum_active, active)
        barrier.wait()
        workflow.db.execute(
            "UPDATE pages SET status = ?, updated_at = ? WHERE id = ?",
            (PageStatus.COMPLETED, utc_now(), page["id"]),
        )
        with state_lock:
            active -= 1

    monkeypatch.setattr(workflow, "_process_page", process_page)

    workflow.process_upload(upload_id)

    pages = workflow.get_upload_pages(upload_id)
    assert maximum_active == 2
    assert {page["status"] for page in pages} == {PageStatus.COMPLETED}


def test_precise_ocr_is_explicit_and_keeps_page_in_manual_review(
    workflow,
    sample_pdf: Path,
    monkeypatch,
):
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
    page = workflow.get_upload_pages(upload_id)[0]
    workflow.db.execute(
        "UPDATE pages SET status = ? WHERE id = ?",
        (PageStatus.NEEDS_REVIEW, page["id"]),
    )
    called_models: list[str] = []

    monkeypatch.setattr(workflow.ocr, "_model_available", lambda _path: True)

    def render(_pdf, _page_number, output_path):
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(b"temporary")
        return output_path

    monkeypatch.setattr(workflow.pdf, "render_page_high_resolution", render)

    def recognize(_pdf, _page_number, _image, model_name="fast"):
        called_models.append(model_name)
        return OcrResult(
            status=OcrStatus.COMPLETED,
            text="Точный локальный текст",
            confidence=0.81,
            language="rus+kir best",
        )

    monkeypatch.setattr(workflow.ocr, "recognize", recognize)

    requested_upload = workflow.request_precise_ocr(page["id"])
    assert requested_upload == upload_id
    assert workflow.get_page(page["id"])["status"] == PageStatus.PROCESSING

    workflow.process_precise_ocr(page["id"])

    processed = workflow.get_page(page["id"])
    assert called_models == ["best"]
    assert processed["status"] == PageStatus.NEEDS_REVIEW
    assert processed["extracted_text"] == "Точный локальный текст"
    assert processed["manual_confirmed"] == 0
