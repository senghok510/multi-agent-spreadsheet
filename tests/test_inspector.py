"""Tests for deterministic workbook inspection (no LLM involved)."""

from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.formula import ArrayFormula

from spreadsheet_agents.tools.workbook_reader import inspect_workbook


@pytest.fixture
def workbook_path(tmp_path: Path) -> Path:
    wb = Workbook()
    sales = wb.active
    sales.title = "Sales"
    sales.append(["Region", "Units", "Price", "Total"])
    for i, region in enumerate(["North", "South", "East"], start=2):
        sales.append([region, i * 10, 2.5, f"=B{i}*C{i}"])

    notes = wb.create_sheet("Notes")
    notes["A1"] = "Remember to check the totals"

    wb.create_sheet("Empty")

    path = tmp_path / "sample.xlsx"
    wb.save(path)
    return path


def test_reports_sheet_names_in_order(workbook_path):
    info = inspect_workbook(workbook_path)
    assert info.sheet_names == ["Sales", "Notes", "Empty"]


def test_reports_dimensions_and_headers(workbook_path):
    sales = inspect_workbook(workbook_path).sheets[0]
    assert (sales.rows, sales.columns) == (4, 4)
    assert sales.headers == ["Region", "Units", "Price", "Total"]


def test_sample_rows_respect_limit(workbook_path):
    sales = inspect_workbook(workbook_path, max_preview_rows=2).sheets[0]
    assert len(sales.sample_rows) == 2
    assert sales.sample_rows[1][0] == "North"


def test_formulas_are_reported_not_evaluated(workbook_path):
    sales = inspect_workbook(workbook_path).sheets[0]
    assert {"cell": "D2", "formula": "=B2*C2"} in sales.formulas
    assert len(sales.formulas) == 3


def test_formula_count_is_capped(workbook_path):
    sales = inspect_workbook(workbook_path, max_formulas=1).sheets[0]
    assert len(sales.formulas) == 1


def test_empty_sheet_does_not_crash(workbook_path):
    empty = inspect_workbook(workbook_path).sheets[2]
    assert empty.formulas == []


def test_input_file_is_not_modified(workbook_path):
    before = workbook_path.read_bytes()
    inspect_workbook(workbook_path)
    assert workbook_path.read_bytes() == before


def test_missing_file_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        inspect_workbook(tmp_path / "nope.xlsx")


@pytest.fixture
def array_formula_path(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Calc"
    ws.append([3, 9, 5])
    ws["E1"] = ArrayFormula("E1", "=MAX(IF(A1:C1<>0,COLUMN(A1:C1)))")
    ws["E2"] = "=SUM(A1:C1)"
    path = tmp_path / "array.xlsx"
    wb.save(path)
    return path


def test_array_formula_text_is_reported_not_an_object_repr(array_formula_path):
    calc = inspect_workbook(array_formula_path).sheets[0]
    formulas = {item["cell"]: item["formula"] for item in calc.formulas}
    assert formulas["E1"] == "=MAX(IF(A1:C1<>0,COLUMN(A1:C1)))"
    assert formulas["E2"] == "=SUM(A1:C1)"
    assert "object at 0x" not in str(calc.formulas)


def test_array_formula_in_sample_rows_is_plain_text(array_formula_path):
    calc = inspect_workbook(array_formula_path).sheets[0]
    assert calc.sample_rows[0][4] == "=MAX(IF(A1:C1<>0,COLUMN(A1:C1)))"


def test_workbook_info_with_array_formula_serializes_to_json(array_formula_path):
    info = inspect_workbook(array_formula_path)
    assert "MAX(IF(" in info.model_dump_json()


def make_sheet(path: Path, rows: int, columns: int) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in range(1, rows + 1):
        ws.append([f"r{r}c{c}" for c in range(1, columns + 1)])
    wb.save(path)
    return path


def test_default_view_shows_a_whole_small_sheet(tmp_path):
    # The first real runs saw only 5 rows; guard against going back to that.
    sheet = inspect_workbook(make_sheet(tmp_path / "a.xlsx", rows=40, columns=8)).sheets[0]
    assert len(sheet.sample_rows) == 40
    assert sheet.sample_rows[-1][0] == "r40c1"
    assert sheet.truncated is False


def test_sheet_exactly_at_the_cap_is_not_truncated(tmp_path):
    sheet = inspect_workbook(make_sheet(tmp_path / "a.xlsx", rows=50, columns=30)).sheets[0]
    assert len(sheet.sample_rows) == 50
    assert len(sheet.sample_rows[0]) == 30
    assert sheet.truncated is False


def test_extra_rows_are_cut_and_flagged_with_the_true_size(tmp_path):
    sheet = inspect_workbook(make_sheet(tmp_path / "a.xlsx", rows=51, columns=4)).sheets[0]
    assert len(sheet.sample_rows) == 50
    assert sheet.rows == 51
    assert sheet.truncated is True


def test_extra_columns_are_cut_and_flagged(tmp_path):
    sheet = inspect_workbook(make_sheet(tmp_path / "a.xlsx", rows=3, columns=31)).sheets[0]
    assert len(sheet.sample_rows[0]) == 30
    assert sheet.columns == 31
    assert sheet.truncated is True


def test_custom_limits_still_control_truncation(tmp_path):
    path = make_sheet(tmp_path / "a.xlsx", rows=10, columns=3)
    assert inspect_workbook(path, max_preview_rows=10).sheets[0].truncated is False
    assert inspect_workbook(path, max_preview_rows=9).sheets[0].truncated is True


def test_truncated_flag_reaches_the_model_input(tmp_path):
    info = inspect_workbook(make_sheet(tmp_path / "a.xlsx", rows=60, columns=2))
    assert '"truncated": true' in info.model_dump_json(indent=2)


@pytest.mark.parametrize("agent", ["inspector", "orchestrator", "editor"])
def test_prompts_explain_truncated_sheets(agent):
    prompt = Path(f"src/spreadsheet_agents/prompts/baseline/{agent}.md").read_text()
    assert '"truncated": true' in prompt
    assert "never invent" in prompt.lower()
