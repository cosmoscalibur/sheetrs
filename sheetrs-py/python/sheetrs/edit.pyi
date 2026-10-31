"""Workbook modification (`sheetedit`).

Two modes, with the same result for the same operations:

- **Lazy** (`Plan`): operations are collected, resolved into the order that gives the right
  result and applied in a single pass over the file on `save`. Every operation is written
  against the *input* workbook, in any order: the plan merges overlapping ranges and sequences
  the operations so that none sees the positions another one moved (formulas are converted
  before the cells they read are deleted, rows and columns are deleted from the last to the
  first, sheets are removed last). `schedule` shows the order.
- **Eager** (the functions of this module): one operation applied at once, one pass each. To
  chain them, pass the output of one as the input of the next; each then addresses the result of
  the previous ones, so the positions are yours to translate.

Scopes are `"Sheet"`, `"Sheet!B:D"` (columns) and `"Sheet!3:10"` (rows); quote a sheet name that
contains `!`. Rows are deleted with a row range and columns with a column range (`ValueError`
otherwise).
"""

import os
from typing import Self

Path = str | os.PathLike[str]

class Plan:
    """A plan: operations on one workbook, resolved and applied in one pass on `save`."""

    def __init__(self, input: Path) -> None: ...
    def remove_sheets(self, *names: str) -> Self: ...
    def remove_ranges(self, *names: str) -> Self: ...
    def convert_to_values(self, *scopes: str) -> Self: ...
    def delete_rows(self, *scopes: str) -> Self: ...
    def delete_columns(self, *scopes: str) -> Self: ...
    def recalc_on_open(self) -> Self: ...
    @property
    def operations(self) -> list[str]:
        """The operations as they were added, as text."""
    @property
    def schedule(self) -> list[str]:
        """The plan as it is applied: normalized and in canonical order, as text."""
    def save(self, output: Path) -> None:
        """Apply the operations in one pass and write `output` (never the input itself)."""
    def __len__(self) -> int: ...

def remove_sheets(input: Path, output: Path, *names: str) -> None: ...
def remove_ranges(input: Path, output: Path, *names: str) -> None: ...
def convert_to_values(input: Path, output: Path, *scopes: str) -> None: ...
def delete_rows(input: Path, output: Path, *scopes: str) -> None: ...
def delete_columns(input: Path, output: Path, *scopes: str) -> None: ...
def recalc_on_open(input: Path, output: Path) -> None: ...
