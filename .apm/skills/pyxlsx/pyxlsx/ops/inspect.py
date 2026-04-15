"""Inspect operations: workbook-level info and sheet listing."""

from __future__ import annotations

import datetime
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


def info(file: str | Path) -> Dict[str, Any]:
    """Return workbook-level metadata: file path, sheet names, and named range names.

    Args:
        file: Path to the .xlsx file.

    Returns:
        dict with keys:
            file (str): The file path as given.
            sheets (list[str]): Sheet names in workbook order.
            named_ranges (list[str]): Named range names defined in the workbook.
    """
    path = Path(file)
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)

    sheets: List[str] = wb.sheetnames
    named_ranges: List[str] = list(wb.defined_names)

    wb.close()

    return {
        "file": str(file),
        "sheets": sheets,
        "named_ranges": named_ranges,
    }


def list_sheets(file: str | Path) -> Dict[str, Any]:
    """Return per-sheet metadata: name, row count, column count, and visibility.

    Args:
        file: Path to the .xlsx file.

    Returns:
        dict with key:
            sheets (list[dict]): Each dict has keys name, rows, cols, visible.
    """
    path = Path(file)
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)

    result: List[Dict[str, Any]] = []
    for ws in wb.worksheets:
        rows = ws.max_row or 0
        cols = ws.max_column or 0
        # SheetState values: "visible", "hidden", "veryHidden"
        visible = ws.sheet_state == "visible"
        result.append(
            {
                "name": ws.title,
                "rows": rows,
                "cols": cols,
                "visible": visible,
            }
        )

    wb.close()

    return {"sheets": result}


def _convert_cell(value: Any) -> Any:
    """Convert an openpyxl cell value to a JSON-serialisable Python type.

    Conversion rules:
        None         → None
        bool         → bool  (must be checked before int since bool subclasses int)
        int          → int
        float        → float
        str          → str
        datetime     → ISO-8601 string with time component (checked before date)
        date         → ISO-8601 date string (YYYY-MM-DD)
        anything else → str(value)
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return value
    if isinstance(value, str):
        return value
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%dT%H:%M:%S")
    if isinstance(value, datetime.date):
        return value.strftime("%Y-%m-%d")
    return str(value)


def read_sheet(
    file: str | Path,
    sheet: str,
    range_str: Optional[str] = None,
) -> Dict[str, Any]:
    """Return a 2D list of typed cell values for a sheet or a specified sub-range.

    Args:
        file: Path to the .xlsx file.
        sheet: Name of the sheet to read.
        range_str: Optional range in A1:H50 notation. If omitted, the entire
                   used range is returned.

    Returns:
        dict with keys:
            sheet (str): The sheet name as given.
            range (str): The actual range read, in A1:XY notation.
            rows (list[list]): 2-D list; each inner list is one row, left to right.
                Cell values are int, float, str, bool, ISO-8601 str, or None.
    """
    path = Path(file)
    try:
        wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        sys.exit(1)

    if sheet not in wb.sheetnames:
        print(f"error: sheet '{sheet}' not found", file=sys.stderr)
        wb.close()
        sys.exit(1)

    ws = wb[sheet]

    if range_str is not None:
        # range_boundaries returns (min_col, min_row, max_col, max_row) — all 1-based
        min_col, min_row, max_col, max_row = range_boundaries(range_str.upper())
        used_range = range_str.upper()
    else:
        min_row = 1
        max_row = ws.max_row or 0
        min_col = 1
        max_col = ws.max_column or 0
        if max_row == 0 or max_col == 0:
            wb.close()
            return {"sheet": sheet, "range": "A1:A1", "rows": []}
        used_range = f"A1:{get_column_letter(max_col)}{max_row}"

    rows: List[List[Any]] = []
    for row in ws.iter_rows(
        min_row=min_row,
        max_row=max_row,
        min_col=min_col,
        max_col=max_col,
    ):
        rows.append([_convert_cell(cell.value) for cell in row])

    wb.close()

    return {
        "sheet": sheet,
        "range": used_range,
        "rows": rows,
    }
