"""Define workbook and worksheet metadata models.

"""
from typing import Any

from pydantic import BaseModel, Field


class SheetInfo(BaseModel):
    name: str
    rows: int
    columns: int
    headers: list[str] = Field(default_factory=list)
    sample_rows: list[list[Any]] = Field(default_factory=list)
    formulas: list[dict[str, str]] = Field(default_factory=list)


class WorkbookInfo(BaseModel):
    path: str
    sheets: list[SheetInfo]

    @property
    def sheet_names(self) -> list[str]:
        return [sheet.name for sheet in self.sheets]