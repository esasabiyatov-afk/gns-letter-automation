from __future__ import annotations

import shutil
from pathlib import Path


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
