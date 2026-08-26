from pathlib import Path

import pytest

from gns_app.services.workflow import WorkflowValidationError


def test_employee_can_be_added_selected_and_applied_to_open_cases(
    workflow,
    sample_pdf: Path,
):
    workflow.initialize_employee_profiles()
    assert workflow.get_active_employee() == "Гапарова Э."
    assert [item["name"] for item in workflow.list_employee_profiles()] == [
        "Гапарова Э."
    ]

    with sample_pdf.open("rb") as stream:
        upload_id = workflow.create_upload(sample_pdf.name, stream)
    case_id, _ = workflow._ensure_qr_case(upload_id, "employee-test-qr")
    assert workflow.get_case(case_id)["employee_name"] == "Гапарова Э."

    selected = workflow.add_employee("  Абдылдаева   Айгуль  ")
    assert selected == "Абдылдаева Айгуль"
    assert workflow.get_active_employee() == "Абдылдаева Айгуль"
    assert workflow.get_case(case_id)["employee_name"] == (
        "Абдылдаева Айгуль"
    )

    duplicate = workflow.add_employee("абдылдаева айгуль")
    assert duplicate == "Абдылдаева Айгуль"
    assert len(workflow.list_employee_profiles()) == 2

    selected = workflow.select_employee("гапарова э.")
    assert selected == "Гапарова Э."
    assert workflow.get_active_employee() == "Гапарова Э."
    assert workflow.get_case(case_id)["employee_name"] == "Гапарова Э."


@pytest.mark.parametrize("name", ["", "   ", "12345", "А" * 121])
def test_invalid_employee_name_is_rejected(workflow, name: str):
    with pytest.raises(WorkflowValidationError):
        workflow.add_employee(name)
