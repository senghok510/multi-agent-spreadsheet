"""Load SpreadsheetBench v1 tasks and evaluate outputs.

"""
import json
from pathlib import Path
from typing import Any

from pydantic import BaseModel


class BenchmarkTask(BaseModel):
    task_id: str
    variant: int
    instruction: str
    input_path: Path
    answer_path: Path | None = None


def load_metadata(root: str | Path) -> list[dict[str, Any]]:
    root = Path(root)

    metadata_files = sorted(root.rglob("*.jsonl"))

    if not metadata_files:
        metadata_files = sorted(root.rglob("dataset.json"))

    if not metadata_files:
        raise FileNotFoundError(
            f"No JSONL or dataset.json found under {root}"
        )

    records = []

    for path in metadata_files:
        if path.suffix == ".jsonl":
            with path.open(encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        records.append(json.loads(line))

        else:
            with path.open(encoding="utf-8") as f:
                data = json.load(f)

            if isinstance(data, list):
                records.extend(data)
            elif isinstance(data, dict):
                if "tasks" in data:
                    records.extend(data["tasks"])
                else:
                    raise ValueError(
                        f"Unknown metadata format: {path}"
                    )

    return records


def find_task_instruction(
    root: str | Path,
    task_id: str,
) -> str:

    records = load_metadata(root)

    matches = [
        record
        for record in records
        if str(record.get("id", record.get("task_id")))
        == str(task_id)
    ]

    if not matches:
        raise KeyError(
            f"Task {task_id} not found in metadata"
        )

    instruction = matches[0].get("instruction")

    if not instruction:
        raise ValueError(
            f"Task {task_id} has no instruction"
        )

    return str(instruction)


def load_task(
    root: str | Path,
    task_id: str,
    variant: int = 1,
) -> BenchmarkTask:

    root = Path(root)

    instruction = find_task_instruction(root, task_id)

    filename = f"{variant}_{task_id}_input.xlsx"

    candidates = list(
        (root / "spreadsheet").rglob(filename)
    )

    if len(candidates) != 1:
        raise FileNotFoundError(
            f"Expected one input workbook named {filename}, "
            f"found {len(candidates)}"
        )

    input_path = candidates[0]

    answer_path = input_path.with_name(
        f"{variant}_{task_id}_answer.xlsx"
    )

    return BenchmarkTask(
        task_id=task_id,
        variant=variant,
        instruction=instruction,
        input_path=input_path,
        answer_path=(
            answer_path if answer_path.exists() else None
        ),
    )
