import os
from typing import Any, Literal, TypedDict

__version__: str

Source = str | os.PathLike[str] | bytes

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

def lint(
    source: Source,
    *,
    config: str | None = None,
    config_path: str | os.PathLike[str] | None = None,
    extension: str | None = None,
) -> list[Violation]: ...
def rules() -> list[Rule]: ...
def stats(source: Source, *, extension: str | None = None) -> WorkbookStats: ...
def edit(
    input: str | os.PathLike[str],
    output: str | os.PathLike[str],
    *,
    remove_sheets: list[str] | None = None,
    remove_ranges: list[str] | None = None,
    convert_to_values: list[str] | None = None,
    delete_rows: list[str] | None = None,
    delete_cols: list[str] | None = None,
    recalc_on_open: bool = False,
) -> None: ...
