"""Inspect operations: workbook-level info and sheet listing."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Dict, Any

import openpyxl


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
