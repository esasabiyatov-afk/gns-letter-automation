from __future__ import annotations

from pathlib import Path

import httpx
import pytest

from gns_app.domain import PageStatus
from gns_app.services.official_document import (
    OfficialDocumentClient,
    OfficialDocumentError,
)


def test_official_timeout_becomes_reviewable_error(tmp_path: Path):
    def handler(request: httpx.Request) -> httpx.Response:
        raise httpx.ReadTimeout("timed out", request=request)

    client = OfficialDocumentClient(
        frozenset({"qr.salyk.kg"}),
        frozenset({"/getsti010decission"}),
        transport=httpx.MockTransport(handler),
    )

    with pytest.raises(OfficialDocumentError, match="не ответил вовремя"):
        client.download(
            "https://qr.salyk.kg/getsti010decission",
            tmp_path,
        )

    assert not (tmp_path / "official.download").exists()


def test_preview_survives_later_page_processing_error(
    workflow,
    sample_pdf: Path,
    monkeypatch,
):
    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)

    def fail_after_render(_preview_path):
        raise RuntimeError("ошибка после сохранения превью")

    monkeypatch.setattr(workflow.qr, "decode", fail_after_render)
    workflow.process_upload(upload_id)

    pages = workflow.get_upload_pages(upload_id)
    assert pages
    assert all(page["status"] == PageStatus.TECHNICAL_ERROR for page in pages)
    assert all(page["preview_path"] for page in pages)
    assert all(Path(page["preview_path"]).exists() for page in pages)

    first, second = pages
    returned_upload_id = workflow.request_page_reprocess(first["id"])
    refreshed = workflow.get_upload_pages(upload_id)

    assert returned_upload_id == upload_id
    assert refreshed[0]["status"] == PageStatus.REGISTERED
    assert refreshed[1]["status"] == PageStatus.TECHNICAL_ERROR
