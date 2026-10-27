"""Define spreadsheet task and result models.

"""
from typing import Literal
from pydantic import BaseModel, Field


class TaskStep(BaseModel):
    id: str
    description: str
    operation_type: str
    target_sheet: str | None = None
    depends_on: list[str] = Field(default_factory=list)


class TaskPlan(BaseModel):
    reasoning_summary: str
    steps: list[TaskStep] = Field(default_factory=list)


class SpreadsheetOperation(BaseModel):
    type: Literal[
        "set_cell",
        "set_formula",
        "create_sheet",
        "copy_range",
        "clear_cell",
    ]

    sheet: str | None = None
    cell: str | None = None
    value: str | int | float | bool | None = None

    source_sheet: str | None = None
    source_range: str | None = None
    destination_sheet: str | None = None
    destination_cell: str | None = None


class EditPlan(BaseModel):
    explanation: str
    operations: list[SpreadsheetOperation] = Field(
        default_factory=list
    )


class ValidationReport(BaseModel):
    passed: bool
    errors: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    checked_operations: int = 0