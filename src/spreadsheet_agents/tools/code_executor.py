"""Provide the future interface for executing workbook editing code.
"""

from pathlib import Path

from spreadsheet_agents.schemas.task import EditPlan
from spreadsheet_agents.tools.workbook_editor import (
    apply_operations,
)


def execute_edit_plan(
    input_path: str | Path,
    output_path: str | Path,
    edit_plan: EditPlan,
    max_operations: int = 1000,
) -> Path:
    return apply_operations(
        input_path=input_path,
        output_path=output_path,
        operations=edit_plan.operations,
        max_operations=max_operations,
    )