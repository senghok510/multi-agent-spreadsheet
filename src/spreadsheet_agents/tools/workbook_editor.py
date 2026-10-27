"""Apply workbook edits while preserving source files.
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import (
    coordinate_from_string,
    range_boundaries,
)

from spreadsheet_agents.schemas.task import (
    SpreadsheetOperation,
)


def _validate_cell(cell: str) -> None:
    coordinate_from_string(cell)


def _apply_operation(wb, operation: SpreadsheetOperation):
    op = operation

    if op.type == "create_sheet":
        if not op.sheet:
            raise ValueError("create_sheet requires sheet")

        if op.sheet in wb.sheetnames:
            raise ValueError(
                f"Sheet already exists: {op.sheet}"
            )

        wb.create_sheet(op.sheet)
        return

    if op.type in {"set_cell", "set_formula", "clear_cell"}:
        if not op.sheet or not op.cell:
            raise ValueError(
                f"{op.type} requires sheet and cell"
            )

        if op.sheet not in wb.sheetnames:
            raise ValueError(
                f"Unknown worksheet: {op.sheet}"
            )

        _validate_cell(op.cell)
        ws = wb[op.sheet]

        if op.type == "clear_cell":
            ws[op.cell] = None

        elif op.type == "set_formula":
            if not isinstance(op.value, str):
                raise ValueError("Formula must be a string")

            if not op.value.startswith("="):
                raise ValueError(
                    "Excel formula must start with '='"
                )

            ws[op.cell] = op.value

        else:
            ws[op.cell] = op.value

        return

    if op.type == "copy_range":
        required = [
            op.source_sheet,
            op.source_range,
            op.destination_sheet,
            op.destination_cell,
        ]

        if any(value is None for value in required):
            raise ValueError(
                "copy_range requires source and destination"
            )

        if op.source_sheet not in wb.sheetnames:
            raise ValueError("Unknown source sheet")

        if op.destination_sheet not in wb.sheetnames:
            raise ValueError("Unknown destination sheet")

        src = wb[op.source_sheet]
        dst = wb[op.destination_sheet]

        min_col, min_row, max_col, max_row = (
            range_boundaries(op.source_range)
        )

        dst_col, dst_row, _, _ = range_boundaries(
            op.destination_cell
        )

        # Snapshot first to support overlapping ranges.
        values = [
            [
                src.cell(row=r, column=c).value
                for c in range(min_col, max_col + 1)
            ]
            for r in range(min_row, max_row + 1)
        ]

        for row_offset, row_values in enumerate(values):
            for col_offset, value in enumerate(row_values):
                dst.cell(
                    row=dst_row + row_offset,
                    column=dst_col + col_offset,
                    value=value,
                )

        return

    raise ValueError(
        f"Unsupported operation: {op.type}"
    )


def apply_operations(
    input_path: str | Path,
    output_path: str | Path,
    operations: list[SpreadsheetOperation],
    max_operations: int = 1000,
) -> Path:
    input_path = Path(input_path)
    output_path = Path(output_path)

    if len(operations) > max_operations:
        raise ValueError(
            f"Too many operations: {len(operations)}"
        )

    if not input_path.exists():
        raise FileNotFoundError(input_path)

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    wb = load_workbook(input_path)

    try:
        for index, operation in enumerate(operations):
            try:
                _apply_operation(wb, operation)
            except Exception as exc:
                raise RuntimeError(
                    f"Operation {index} failed: {exc}"
                ) from exc

        wb.save(output_path)

    finally:
        wb.close()

    return output_path