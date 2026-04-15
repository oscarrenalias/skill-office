import json
import sys
from typing import Callable

import click

import pyxlsx
from pyxlsx.ops.inspect import info as _info, list_sheets as _list_sheets, read_sheet as _read_sheet


def output_result(data: dict, plain: bool, plain_fn: Callable[[dict], str]) -> None:
    """Write data to stdout as JSON or as plain text via plain_fn."""
    if plain:
        sys.stdout.write(plain_fn(data) + "\n")
    else:
        sys.stdout.write(json.dumps(data) + "\n")


@click.group()
@click.version_option(version=pyxlsx.__version__, prog_name="pyxlsx")
@click.option("--plain", is_flag=True, default=False, help="Output plain text instead of JSON.")
@click.pass_context
def cli(ctx: click.Context, plain: bool) -> None:
    """pyxlsx — Excel manipulation toolkit."""
    ctx.ensure_object(dict)
    ctx.obj["plain"] = plain


@cli.command("info")
@click.argument("file")
@click.pass_context
def info_cmd(ctx: click.Context, file: str) -> None:
    """Show workbook-level metadata: sheet names and named ranges."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        lines = [data["file"]]
        lines.append("Sheets: " + ", ".join(data["sheets"]))
        nr = data["named_ranges"]
        lines.append("Named ranges: " + (", ".join(nr) if nr else "(none)"))
        return "\n".join(lines)

    try:
        result = _info(file)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@cli.group()
def sheet() -> None:
    """Commands for working with sheets."""


@sheet.command("list")
@click.argument("file")
@click.pass_context
def sheet_list_cmd(ctx: click.Context, file: str) -> None:
    """List all sheets with row/column counts and visibility."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        lines = []
        for s in data["sheets"]:
            line = f"{s['name']:<20} {s['rows']:>5} rows  {s['cols']:>3} cols"
            if not s["visible"]:
                line += "  [hidden]"
            lines.append(line)
        return "\n".join(lines)

    try:
        result = _list_sheets(file)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@sheet.command("read")
@click.argument("file")
@click.argument("sheet")
@click.option("--range", "range_str", default=None, help="Cell range in A1:H50 notation.")
@click.pass_context
def sheet_read_cmd(ctx: click.Context, file: str, sheet: str, range_str: str | None) -> None:
    """Read a sheet as a 2D array of typed cell values."""
    plain: bool = ctx.obj["plain"]

    def plain_fn(data: dict) -> str:
        rows = data.get("rows", [])
        if not rows:
            return ""
        col_count = max(len(row) for row in rows)
        # Convert all values to strings for display
        str_rows = []
        for row in rows:
            str_row = [str(v) if v is not None else "" for v in row]
            # Pad short rows to col_count
            while len(str_row) < col_count:
                str_row.append("")
            str_rows.append(str_row)
        # Compute per-column widths
        widths = [0] * col_count
        for str_row in str_rows:
            for i, v in enumerate(str_row):
                widths[i] = max(widths[i], len(v))
        lines = []
        for str_row in str_rows:
            padded = [str_row[i].ljust(widths[i]) for i in range(col_count)]
            lines.append("  ".join(padded).rstrip())
        return "\n".join(lines)

    try:
        result = _read_sheet(file, sheet, range_str)
    except SystemExit:
        sys.exit(1)
    output_result(result, plain, plain_fn)


@cli.group()
def table() -> None:
    """Commands for working with tables."""


@cli.group()
def cell() -> None:
    """Commands for working with cells."""
