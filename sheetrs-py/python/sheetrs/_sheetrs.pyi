import os
from typing import Any, Literal, Self, TypedDict

__version__: str

Source = str | os.PathLike[str] | bytes
Path = str | os.PathLike[str]

class SheetrsError(Exception):
    """A workbook could not be read, analyzed or written."""

class Violation:
    """One finding of the linter."""

    rule_id: str
    rule_name: str
    severity: Literal["error", "warning", "info"]
    message: str
    scope: Literal["book", "sheet", "cell"]
    sheet: str | None
    cell: str | None
    row: int | None
    col: int | None
    def to_dict(self) -> dict[str, Any]: ...

class Rule(TypedDict):
    id: str
    name: str
    category: str
    default: bool
    enabled: bool

class Linter:
    """The linter with a configuration validated at construction.

    `config` is the TOML text of a `sheetlint.toml`, `config_path` the path of one (not both);
    an unknown rule, category or parameter raises `ValueError`.
    """

    def __init__(
        self, *, config: str | None = None, config_path: Path | None = None
    ) -> None: ...
    def lint(self, source: Source, *, extension: str | None = None) -> list[Violation]:
        """Lint a workbook given by path or as bytes (then `extension` is required)."""
    def rules(self) -> list[Rule]:
        """The rules ordered by id; `enabled` tells whether this configuration runs each one."""

def lint(
    source: Source,
    *,
    config: str | None = None,
    config_path: Path | None = None,
    extension: str | None = None,
) -> list[Violation]:
    """Lint with a one-off `Linter(config=..., config_path=...)`."""

class SheetStats(TypedDict):
    name: str
    used_cells: int
    formulas: int
    errors: int
    empty_cells: int
    end_of_sheet: str | None
    comments: int
    charts: int
    tables: int
    xml_size: int

class WorkbookStats(TypedDict):
    total_sheets: int
    total_named_ranges: int
    total_file_size: int
    total_used_cells: int
    total_formulas: int
    total_errors: int
    total_empty_cells: int
    total_comments: int
    total_charts: int
    total_tables: int
    total_xml_size: int
    sheets: list[SheetStats]

def stats(source: Source, *, extension: str | None = None) -> WorkbookStats:
    """Statistics of a workbook given by path or as bytes (`total_file_size` is 0 for bytes)."""

class EditPlan:
    """Operations on one workbook, written against the input and applied in one pass on `save`.

    Scopes are `"Sheet"`, `"Sheet!B:D"` (columns) and `"Sheet!3:10"` (rows); quote a sheet name
    that contains `!`. Rows are deleted with a row range and columns with a column range
    (`ValueError` otherwise). Operations address the input workbook in any order; the plan
    merges overlapping ranges and applies them in the order that gives the right result
    (see `operations`). A formula with a 3D reference makes `save` fail when rows, columns or
    sheets are deleted or removed (issue #75).
    """

    def __init__(self, input: Path) -> None: ...
    def remove_sheets(self, *names: str) -> Self: ...
    def remove_ranges(self, *names: str) -> Self: ...
    def convert_to_values(self, *scopes: str) -> Self: ...
    def delete_rows(self, *scopes: str) -> Self: ...
    def delete_columns(self, *scopes: str) -> Self: ...
    def recalc_on_open(self) -> Self: ...
    @property
    def operations(self) -> list[str]:
        """The operations as they will be applied (resolved, in order), as text."""
    def save(self, output: Path) -> None:
        """Apply the operations in one pass and write `output` (never the input itself)."""
    def __len__(self) -> int: ...

def edit(
    input: Path,
    output: Path,
    *,
    remove_sheets: list[str] | None = None,
    remove_ranges: list[str] | None = None,
    convert_to_values: list[str] | None = None,
    delete_rows: list[str] | None = None,
    delete_columns: list[str] | None = None,
    recalc_on_open: bool = False,
) -> None:
    """Eager editing: `EditPlan(input)` with these operations, saved to `output` at once."""
