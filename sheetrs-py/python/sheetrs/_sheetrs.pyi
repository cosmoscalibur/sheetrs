import os
from typing import Any, Literal, TypedDict

__version__: str

Path = str | os.PathLike[str]
Source = Path | bytes | Workbook

class SheetrsError(Exception):
    """A workbook could not be read, analyzed or written."""

class Violation:
    """One finding of the linter: the type of the items `Linter.lint` returns. It is exposed so
    that results can be annotated and checked; it is not built by the user."""

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
        """Lint a workbook given by path, as bytes (then `extension` is required) or loaded."""
    def rules(self) -> list[Rule]:
        """The rules ordered by id; `enabled` tells whether this configuration runs each one."""

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

class Workbook:
    """A workbook loaded for editing: eager mode.

    It has a working copy of the file, and every operation runs when it is called, on the result
    of the previous ones, and returns the `Workbook` that results (the one it was called on does
    not change). The intermediate workbook can be linted or measured (`Linter.lint(workbook)`,
    `stats(workbook)`), and a position is the one the earlier operations left.

    Scopes are `"Sheet"`, `"Sheet!B:D"` (columns) and `"Sheet!3:10"` (rows); quote a sheet name
    that contains `!`. The scopes of one call are simultaneous. Rows are deleted with a row range
    and columns with a column range (`ValueError` otherwise).
    """

    def remove_sheets(self, *names: str) -> Workbook: ...
    def remove_ranges(self, *names: str) -> Workbook: ...
    def convert_to_values(self, *scopes: str) -> Workbook: ...
    def delete_rows(self, *scopes: str) -> Workbook: ...
    def delete_columns(self, *scopes: str) -> Workbook: ...
    def recalc_on_open(self) -> Workbook: ...
    def lazy(self) -> LazyWorkbook:
        """The same workbook, lazily: what follows is recorded, not run."""
    def write(self, path: Path) -> None:
        """Save it to `path`, which must not be the file it was read from."""

class LazyWorkbook:
    """A workbook scanned for editing: lazy mode. Nothing is read.

    The methods of `Workbook`, with the same meaning (each operation addresses the workbook as the
    previous ones leave it), that record the operation and return the plan that results. The plan
    runs on `sink` or `collect`, in the fewest passes over the file that give what running the
    operations one at a time gives; `explain` shows them. A formula with a 3D reference
    (`Sheet1:Sheet3!A1`) makes the run fail when rows, columns or sheets are deleted or removed
    (issue #75).
    """

    def remove_sheets(self, *names: str) -> LazyWorkbook: ...
    def remove_ranges(self, *names: str) -> LazyWorkbook: ...
    def convert_to_values(self, *scopes: str) -> LazyWorkbook: ...
    def delete_rows(self, *scopes: str) -> LazyWorkbook: ...
    def delete_columns(self, *scopes: str) -> LazyWorkbook: ...
    def recalc_on_open(self) -> LazyWorkbook: ...
    @property
    def operations(self) -> list[str]:
        """The operations as they were written, as text."""
    def explain(self) -> str:
        """The passes over the file that run the plan, as text."""
    def sink(self, path: Path) -> None:
        """Run the plan and write the result to `path` (never the file it reads)."""
    def collect(self) -> Workbook:
        """Run the plan and load the result."""

def read(source: Path | bytes, *, extension: str | None = None) -> Workbook:
    """Load a workbook for editing, eagerly. Bytes need `extension`. The file is never modified."""

def scan(path: Path) -> LazyWorkbook:
    """Scan a workbook for editing, lazily. Only the path and the extension are checked."""
