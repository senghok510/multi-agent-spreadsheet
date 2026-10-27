"""Define shared workflow state.
"""
from typing import Any, TypedDict


class SpreadsheetState(TypedDict, total=False):

    # Inputs
    ## prompt given by the user
    user_request: str
    ## input 
    input_path: str
    # output
    output_path: str
    model: str

    # Agent outputs
    workbook_info: dict[str, Any]
    inspection_report: dict[str, Any]
    task_plan: dict[str, Any]
    edit_plan: dict[str, Any]
    validation_report: dict[str, Any]

    # Execution status
    status: str
    error: str | None
    retry_count: int

    # Observability
    trace: list[dict[str, Any]]