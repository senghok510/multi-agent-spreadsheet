"""Tests for the SpreadsheetBench V1 adapter, using a tiny synthetic dataset."""

import json
from pathlib import Path

import pytest
from openpyxl import Workbook

from spreadsheet_agents.benchmark.spreadsheetbench_v1 import (
    BenchmarkTask,
    find_answer_path,
    find_task_instruction,
    load_metadata,
    load_task,
)

REAL_DATASET = Path("data/raw/spreadsheetbench_v1/sample_data_200")


def make_workbook(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.active["A1"] = path.name
    wb.save(path)


@pytest.fixture
def dataset(tmp_path: Path) -> Path:
    root = tmp_path / "bench"
    records = [
        {"id": "101", "instruction": "Sum column A", "spreadsheet_path": "spreadsheet/101"},
        {"id": "CF_7", "instruction": "Highlight", "spreadsheet_path": "spreadsheet/CF_7"},
    ]
    root.mkdir()
    (root / "dataset.json").write_text(json.dumps(records))
    for variant in (1, 2):
        make_workbook(root / "spreadsheet" / "101" / f"{variant}_101_input.xlsx")
        make_workbook(root / "spreadsheet" / "101" / f"{variant}_101_answer.xlsx")
    make_workbook(root / "spreadsheet" / "CF_7" / "1_CF_7_input.xlsx")  # no answer file
    return root


# --- metadata ---------------------------------------------------------------


def test_load_metadata_reads_dataset_json_list(dataset):
    assert [r["id"] for r in load_metadata(dataset)] == ["101", "CF_7"]


def test_load_metadata_reads_jsonl(tmp_path):
    (tmp_path / "tasks.jsonl").write_text(
        '{"id": "1", "instruction": "a"}\n\n{"id": "2", "instruction": "b"}\n'
    )
    assert [r["id"] for r in load_metadata(tmp_path)] == ["1", "2"]


def test_load_metadata_reads_tasks_wrapper(tmp_path):
    (tmp_path / "dataset.json").write_text(json.dumps({"tasks": [{"id": "1", "instruction": "a"}]}))
    assert len(load_metadata(tmp_path)) == 1


def test_load_metadata_rejects_unknown_format(tmp_path):
    (tmp_path / "dataset.json").write_text(json.dumps({"something": []}))
    with pytest.raises(ValueError, match="Unknown metadata format"):
        load_metadata(tmp_path)


def test_load_metadata_fails_clearly_when_nothing_found(tmp_path):
    with pytest.raises(FileNotFoundError, match="No JSONL or dataset.json"):
        load_metadata(tmp_path)


# --- instructions -----------------------------------------------------------


def test_find_instruction_accepts_string_or_numeric_id(dataset):
    assert find_task_instruction(dataset, "101") == "Sum column A"
    assert find_task_instruction(dataset, 101) == "Sum column A"


def test_find_instruction_unknown_task(dataset):
    with pytest.raises(KeyError, match="999"):
        find_task_instruction(dataset, "999")


def test_find_instruction_rejects_empty_instruction(tmp_path):
    (tmp_path / "dataset.json").write_text(json.dumps([{"id": "1", "instruction": ""}]))
    with pytest.raises(ValueError, match="no instruction"):
        find_task_instruction(tmp_path, "1")


# --- load_task --------------------------------------------------------------


def test_load_task_returns_instruction_and_input_workbook(dataset):
    task = load_task(dataset, "101")
    assert task.task_id == "101"
    assert task.variant == 1
    assert task.instruction == "Sum column A"
    assert task.input_path == dataset / "spreadsheet/101/1_101_input.xlsx"


def test_load_task_selects_the_requested_variant(dataset):
    assert load_task(dataset, "101", variant=2).input_path.name == "2_101_input.xlsx"


def test_load_task_accepts_numeric_id_as_stored_in_real_metadata(dataset):
    task = load_task(dataset, 101)
    assert task.task_id == "101"


def test_load_task_works_for_ids_with_prefix(dataset):
    assert load_task(dataset, "CF_7").input_path.name == "1_CF_7_input.xlsx"


def test_load_task_missing_workbook(dataset):
    with pytest.raises(FileNotFoundError, match="found 0"):
        load_task(dataset, "101", variant=3)


def test_load_task_ambiguous_workbook(dataset):
    make_workbook(dataset / "spreadsheet" / "copy" / "1_101_input.xlsx")
    with pytest.raises(FileNotFoundError, match="found 2"):
        load_task(dataset, "101")


# --- no answer leakage ------------------------------------------------------


def test_agent_facing_task_has_no_answer_information(dataset):
    task = load_task(dataset, "101")
    assert "answer" not in set(BenchmarkTask.model_fields)
    assert "answer" not in task.model_dump_json().lower()


def test_answer_lookup_is_a_separate_evaluator_function(dataset):
    assert find_answer_path(dataset, "101") == dataset / "spreadsheet/101/1_101_answer.xlsx"
    assert find_answer_path(dataset, "101", variant=2).name == "2_101_answer.xlsx"


def test_answer_lookup_returns_none_when_absent(dataset):
    assert find_answer_path(dataset, "CF_7") is None


# --- real dataset (skipped in CI, where data/ is not committed) ------------


@pytest.mark.skipif(not REAL_DATASET.exists(), reason="SpreadsheetBench not downloaded")
def test_real_dataset_task_loads():
    record = load_metadata(REAL_DATASET)[0]
    task = load_task(REAL_DATASET, record["id"])
    assert task.instruction
    assert task.input_path.exists()
    assert "answer" not in task.model_dump_json().lower()
