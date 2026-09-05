from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image


def test_inbox_imports_new_pdf_and_skips_duplicate(
    workflow,
    sample_pdf: Path,
):
    workflow.initialize_employee_profiles()
    inbox_file = workflow.settings.inbox_dir / "Новое письмо.PDF"
    shutil.copyfile(sample_pdf, inbox_file)

    first = workflow.import_inbox()
    second = workflow.import_inbox()

    assert len(first["imported"]) == 1
    assert not first["errors"]
    assert not second["imported"]
    assert second["skipped"] == ["Новое письмо.PDF"]

    old_response = workflow.settings.responses_dir / "old-response.docx"
    old_response.write_bytes(b"old")
    reset = workflow.reset_processing_data()

    assert reset["uploads"] == 1
    assert not reset["cleanup_errors"]
    assert inbox_file.exists()
    assert not old_response.exists()
    assert not workflow.list_uploads()

    third = workflow.import_inbox()
    assert len(third["imported"]) == 1
    assert not third["skipped"]


def test_inbox_imports_supported_image(workflow):
    workflow.initialize_employee_profiles()
    inbox_file = workflow.settings.inbox_dir / "Фото письма.PNG"
    Image.new("RGB", (240, 320), "white").save(inbox_file)

    result = workflow.import_inbox()

    assert len(result["imported"]) == 1
    upload = workflow.get_upload(result["imported"][0])
    assert upload["original_filename"] == "Фото письма.PNG"
    assert upload["page_count"] == 1
    assert Path(upload["stored_path"]).suffix == ".pdf"
    assert (Path(upload["stored_path"]).parent / "source.png").is_file()
