"""Tests for the Inspector and Orchestrator agents with a fake LLM (no network)."""

from pathlib import Path

import pytest
from openpyxl import Workbook

from spreadsheet_agents.agents import inspector, orchestrator
from spreadsheet_agents.agents.inspector import InspectionReport, run_inspector
from spreadsheet_agents.agents.orchestrator import run_orchestrator
from spreadsheet_agents.schemas.task import TaskPlan
from spreadsheet_agents.schemas.workbook import SheetInfo, WorkbookInfo


class FakeLLM:
    """Stands in for ChatOpenAI and records how the agent called it."""

    instances: list["FakeLLM"] = []
    response = None  # set per test with monkeypatch

    def __init__(self, model: str, temperature: float):
        self.model = model
        self.temperature = temperature
        self.schema = None
        self.messages = None
        FakeLLM.instances.append(self)

    def with_structured_output(self, schema):
        self.schema = schema
        return self

    def invoke(self, messages):
        self.messages = messages
        return self.response


@pytest.fixture
def fake_llm(monkeypatch):
    FakeLLM.instances = []
    monkeypatch.setattr(inspector, "ChatOpenAI", FakeLLM)
    monkeypatch.setattr(orchestrator, "ChatOpenAI", FakeLLM)
    return FakeLLM


@pytest.fixture
def workbook_path(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Region", "Units"])
    ws.append(["North", 10])
    path = tmp_path / "in.xlsx"
    wb.save(path)
    return path


def sent_text(llm: FakeLLM) -> str:
    return "\n".join(content for _, content in llm.messages)


def test_inspector_returns_workbook_info_and_report(fake_llm, workbook_path, monkeypatch):
    report = InspectionReport(relevant_sheets=["Sales"], observations=["has Units"])
    monkeypatch.setattr(FakeLLM, "response", report)

    info, result = run_inspector(str(workbook_path), "sum units", model="test-model")

    assert info.sheet_names == ["Sales"]
    assert result == report


def test_inspector_sends_request_and_real_workbook_schema(fake_llm, workbook_path, monkeypatch):
    monkeypatch.setattr(FakeLLM, "response", InspectionReport())

    run_inspector(str(workbook_path), "sum units", model="test-model")

    llm = fake_llm.instances[0]
    text = sent_text(llm)
    assert "sum units" in text
    assert "Region" in text and "Units" in text
    assert llm.messages[0][0] == "system"
    assert "Inspection Agent" in llm.messages[0][1]


def test_inspector_is_deterministic_and_structured(fake_llm, workbook_path, monkeypatch):
    monkeypatch.setattr(FakeLLM, "response", InspectionReport())

    run_inspector(str(workbook_path), "x", model="test-model")

    llm = fake_llm.instances[0]
    assert llm.model == "test-model"
    assert llm.temperature == 0
    assert llm.schema is InspectionReport


def test_inspector_fails_before_calling_llm_when_file_missing(fake_llm, tmp_path):
    with pytest.raises(FileNotFoundError):
        run_inspector(str(tmp_path / "nope.xlsx"), "x", model="test-model")
    assert fake_llm.instances == []


def test_orchestrator_returns_structured_plan(fake_llm, monkeypatch):
    plan = TaskPlan(reasoning_summary="add totals")
    monkeypatch.setattr(FakeLLM, "response", plan)

    result = run_orchestrator(
        "add totals",
        WorkbookInfo(path="a.xlsx", sheets=[SheetInfo(name="Sales", rows=2, columns=2)]),
        InspectionReport(relevant_sheets=["Sales"]),
        model="test-model",
    )

    assert result == plan
    assert fake_llm.instances[0].schema is TaskPlan


def test_orchestrator_sends_request_workbook_and_inspection(fake_llm, monkeypatch):
    monkeypatch.setattr(FakeLLM, "response", TaskPlan(reasoning_summary="x"))

    run_orchestrator(
        "add totals",
        WorkbookInfo(path="a.xlsx", sheets=[SheetInfo(name="Sales", rows=2, columns=2)]),
        InspectionReport(observations=["units column found"]),
        model="test-model",
    )

    llm = fake_llm.instances[0]
    text = sent_text(llm)
    assert "add totals" in text
    assert "Sales" in text
    assert "units column found" in text
    assert "Orchestrator Agent" in llm.messages[0][1]
    assert llm.temperature == 0
