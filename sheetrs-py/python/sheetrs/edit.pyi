"""Workbook modification (`sheetedit`)."""

import os

def modify(
    input: str | os.PathLike[str],
    output: str | os.PathLike[str],
    *,
    remove_sheets: list[str] | None = None,
    remove_ranges: list[str] | None = None,
    convert_to_values: list[str] | None = None,
    delete_rows: list[str] | None = None,
    delete_cols: list[str] | None = None,
    recalc_on_open: bool = False,
) -> None:
    """Write a modified copy of `input` to `output` (it never overwrites its input)."""
