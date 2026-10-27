"""Tests for the deterministic workbook editing tool (no LLM involved)."""

from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from spreadsheet_agents.schemas.task import SpreadsheetOperation
from spreadsheet_agents.tools.workbook_editor import apply_operations


def op(type_: str, **fields) -> SpreadsheetOperation:
    return SpreadsheetOperation(type=type_, **fields)


@pytest.fixture
def source(tmp_path: Path) -> Path:
    wb = Workbook()
    sales = wb.active
    sales.title = "Sales"
    sales.append(["Region", "Units", "Price"])
    sales.append(["North", 10, 2.5])
    sales.append(["South", 20, 3.0])
    path = tmp_path / "in.xlsx"
    wb.save(path)
    return path


@pytest.fixture
def out(tmp_path: Path) -> Path:
    return tmp_path / "result" / "out.xlsx"


def read(path: Path):
    return load_workbook(path)


def test_set_cell_writes_to_new_copy(source, out):
    apply_operations(source, out, [op("set_cell", sheet="Sales", cell="D1", value="Total")])
    assert read(out)["Sales"]["D1"].value == "Total"


def test_input_workbook_is_never_modified(source, out):
    before = source.read_bytes()
    apply_operations(source, out, [op("set_cell", sheet="Sales", cell="A1", value="X")])
    assert source.read_bytes() == before


def test_output_parent_directory_is_created(source, out):
    assert not out.parent.exists()
    apply_operations(source, out, [])
    assert out.exists()


def test_set_formula_is_stored_as_formula(source, out):
    apply_operations(source, out, [op("set_formula", sheet="Sales", cell="D2", value="=B2*C2")])
    assert read(out)["Sales"]["D2"].value == "=B2*C2"


def test_set_formula_requires_leading_equals(source, out):
    with pytest.raises(RuntimeError, match="must start with"):
        apply_operations(source, out, [op("set_formula", sheet="Sales", cell="D2", value="B2*C2")])


def test_set_formula_requires_string(source, out):
    with pytest.raises(RuntimeError, match="must be a string"):
        apply_operations(source, out, [op("set_formula", sheet="Sales", cell="D2", value=5)])


def test_clear_cell(source, out):
    apply_operations(source, out, [op("clear_cell", sheet="Sales", cell="A2")])
    assert read(out)["Sales"]["A2"].value is None


def test_create_sheet(source, out):
    apply_operations(source, out, [op("create_sheet", sheet="Summary")])
    assert read(out).sheetnames == ["Sales", "Summary"]


def test_create_sheet_rejects_duplicate(source, out):
    with pytest.raises(RuntimeError, match="already exists"):
        apply_operations(source, out, [op("create_sheet", sheet="Sales")])


def test_copy_range_copies_values(source, out):
    apply_operations(
        source,
        out,
        [
            op(
                "copy_range",
                source_sheet="Sales",
                source_range="A1:B2",
                destination_sheet="Sales",
                destination_cell="E1",
            )
        ],
    )
    sales = read(out)["Sales"]
    assert [sales["E1"].value, sales["F1"].value] == ["Region", "Units"]
    assert [sales["E2"].value, sales["F2"].value] == ["North", 10]


def test_copy_range_handles_overlap(source, out):
    apply_operations(
        source,
        out,
        [
            op(
                "copy_range",
                source_sheet="Sales",
                source_range="A1:A3",
                destination_sheet="Sales",
                destination_cell="A2",
            )
        ],
    )
    sales = read(out)["Sales"]
    assert [sales["A2"].value, sales["A3"].value, sales["A4"].value] == [
        "Region",
        "North",
        "South",
    ]


def _copy(source_range, destination_cell, sheet="Sales"):
    return op(
        "copy_range",
        source_sheet=sheet,
        source_range=source_range,
        destination_sheet=sheet,
        destination_cell=destination_cell,
    )


def test_copy_range_shifts_relative_references_like_fill_down(source, out):
    apply_operations(
        source,
        out,
        [
            op("set_formula", sheet="Sales", cell="D2", value="=B2*C2"),
            _copy("D2", "D3"),
            _copy("D2", "D4"),
        ],
    )
    sales = read(out)["Sales"]
    assert sales["D3"].value == "=B3*C3"
    assert sales["D4"].value == "=B4*C4"


def test_copy_range_keeps_absolute_references_fixed(source, out):
    apply_operations(
        source,
        out,
        [op("set_formula", sheet="Sales", cell="D2", value="=B2*$C$2"), _copy("D2", "D3")],
    )
    assert read(out)["Sales"]["D3"].value == "=B3*$C$2"


def test_copy_range_shifts_references_across_columns_too(source, out):
    apply_operations(
        source,
        out,
        [op("set_formula", sheet="Sales", cell="D2", value="=B2+C2"), _copy("D2", "E2")],
    )
    assert read(out)["Sales"]["E2"].value == "=C2+D2"


def test_copy_range_translates_every_formula_in_a_block(source, out):
    apply_operations(
        source,
        out,
        [
            op("set_formula", sheet="Sales", cell="D2", value="=B2"),
            op("set_formula", sheet="Sales", cell="E2", value="=C2"),
            _copy("D2:E2", "D5"),
        ],
    )
    sales = read(out)["Sales"]
    assert (sales["D5"].value, sales["E5"].value) == ("=B5", "=C5")


def test_copy_range_does_not_touch_plain_text_that_looks_like_a_formula(source, out):
    apply_operations(source, out, [_copy("A2", "A5")])  # "North": ordinary value
    assert read(out)["Sales"]["A5"].value == "North"


def test_copy_range_requires_all_fields(source, out):
    with pytest.raises(RuntimeError, match="requires source and destination"):
        apply_operations(source, out, [op("copy_range", source_sheet="Sales")])


def test_unknown_sheet_is_rejected(source, out):
    with pytest.raises(RuntimeError, match="Unknown worksheet"):
        apply_operations(source, out, [op("set_cell", sheet="Nope", cell="A1", value=1)])


def test_missing_sheet_or_cell_is_rejected(source, out):
    with pytest.raises(RuntimeError, match="requires sheet and cell"):
        apply_operations(source, out, [op("set_cell", value=1)])


@pytest.mark.parametrize("cell", ["1A", "A", "A1:B2", ""])
def test_invalid_cell_reference_is_rejected(source, out, cell):
    with pytest.raises(RuntimeError):
        apply_operations(source, out, [op("set_cell", sheet="Sales", cell=cell, value=1)])


def test_error_names_the_failing_operation_index(source, out):
    ops = [
        op("set_cell", sheet="Sales", cell="A1", value="ok"),
        op("set_cell", sheet="Nope", cell="A1", value=1),
    ]
    with pytest.raises(RuntimeError, match="Operation 1 failed"):
        apply_operations(source, out, ops)


def test_failed_run_leaves_no_output_file(source, out):
    with pytest.raises(RuntimeError):
        apply_operations(source, out, [op("set_cell", sheet="Nope", cell="A1", value=1)])
    assert not out.exists()


def test_operation_limit_is_enforced(source, out):
    ops = [op("set_cell", sheet="Sales", cell="A1", value=1)] * 3
    with pytest.raises(ValueError, match="Too many operations"):
        apply_operations(source, out, ops, max_operations=2)


def test_missing_input_raises(tmp_path, out):
    with pytest.raises(FileNotFoundError):
        apply_operations(tmp_path / "nope.xlsx", out, [])


def test_output_must_not_overwrite_input(source):
    with pytest.raises(ValueError, match="overwrite"):
        apply_operations(source, source, [op("set_cell", sheet="Sales", cell="A1", value="X")])
