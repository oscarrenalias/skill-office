import json
import sys
from typing import Callable

import click

import pyxlsx


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


@cli.group()
def sheet() -> None:
    """Commands for working with sheets."""


@cli.group()
def table() -> None:
    """Commands for working with tables."""


@cli.group()
def cell() -> None:
    """Commands for working with cells."""
