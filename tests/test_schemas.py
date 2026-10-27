"""Tests for the typed contracts agents exchange (schemas and graph state)."""

import pytest
from pydantic import ValidationError

from spreadsheet_agents.graph.state import SpreadsheetState
from spreadsheet_agents.schemas.task import (
    EditPlan,
    SpreadsheetOperation,
    TaskPlan,
    TaskStep,
    ValidationReport,
)
from spreadsheet_agents.schemas.workbook import SheetInfo, WorkbookInfo


def test_operation_rejects_unknown_type():
    with pytest.raises(ValidationError):
        SpreadsheetOperation(type="run_python", sheet="Sheet1", cell="A1")


@pytest.mark.parametrize(
    "op_type",
    ["set_cell", "set_formula", "create_sheet", "copy_range", "clear_cell"],
)
def test_operation_accepts_every_supported_type(op_type):
    assert SpreadsheetOperation(type=op_type).type == op_type


def test_operation_rejects_non_scalar_value():
    with pytest.raises(ValidationError):
        SpreadsheetOperation(type="set_cell", sheet="S", cell="A1", value=["x"])


def test_task_plan_requires_reasoning_summary():
    with pytest.raises(ValidationError):
        TaskPlan()


def test_task_step_requires_core_fields():
    with pytest.raises(ValidationError):
        TaskStep(id="s1")


def test_plan_defaults_are_not_shared_between_instances():
    first = TaskPlan(reasoning_summary="a")
    second = TaskPlan(reasoning_summary="b")
    first.steps.append(TaskStep(id="s1", description="d", operation_type="set_cell"))
    assert second.steps == []


def test_edit_plan_round_trips_through_json():
    plan = EditPlan(
        explanation="Fill total",
        operations=[
            SpreadsheetOperation(type="set_formula", sheet="Sales", cell="D2", value="=B2*C2")
        ],
    )
    assert EditPlan.model_validate_json(plan.model_dump_json()) == plan


def test_validation_report_requires_passed_flag():
    with pytest.raises(ValidationError):
        ValidationReport()


def test_validation_report_defaults():
    report = ValidationReport(passed=True)
    assert report.errors == []
    assert report.warnings == []
    assert report.checked_operations == 0


def test_workbook_info_lists_sheet_names():
    info = WorkbookInfo(
        path="a.xlsx",
        sheets=[SheetInfo(name="One", rows=1, columns=1), SheetInfo(name="Two", rows=2, columns=2)],
    )
    assert info.sheet_names == ["One", "Two"]


def test_state_can_hold_agent_outputs_as_plain_dicts():
    # LangGraph state stores plain dicts, so every contract must serialise to one.
    plan = TaskPlan(reasoning_summary="x")
    state: SpreadsheetState = {
        "user_request": "sum column B",
        "task_plan": plan.model_dump(),
        "retry_count": 0,
    }
    assert TaskPlan.model_validate(state["task_plan"]) == plan
