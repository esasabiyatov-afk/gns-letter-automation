from __future__ import annotations

import shutil
from pathlib import Path


def test_inbox_imports_new_pdf_and_skips_duplicate(
    workflow,
    project_root: Path,
):
    workflow.initialize_employee_profiles()
    sample = project_root / "УГНС" / "пример письма.pdf"
    inbox_file = workflow.settings.inbox_dir / "Новое письмо.PDF"
    shutil.copyfile(sample, inbox_file)

    first = workflow.import_inbox()
    second = workflow.import_inbox()

    assert len(first["imported"]) == 1
    assert not first["errors"]
    assert not second["imported"]
    assert second["skipped"] == ["Новое письмо.PDF"]
