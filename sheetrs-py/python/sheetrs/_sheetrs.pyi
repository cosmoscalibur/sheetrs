import os
from typing import Any, Literal, TypedDict

__version__: str

Path = str | os.PathLike[str]

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
    def lint(self, source: Source) -> list[Violation]:
        """Lint the workbook of a `Source` (the result of `sheetrs.open`)."""
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

class Source:
    """The input of every tool: a reference to a workbook (a file, or its bytes), made by `open`.

    It reads nothing and is not a navigable model of the workbook: the tools read it when they
    run, and the same `Source` can be given to several. It has no behavior of its own:
    `Linter.lint(source)`, `stats.collect(source)` and `Editor(source)` take it.
    """

def collect(source: Source) -> WorkbookStats:
    """Statistics of the workbook (`total_file_size` is 0 for a workbook given as bytes)."""

class Editor:
    """The changes to make to a workbook, and the way to save them.

    Made for a `Source`; every method adds a change and returns the editor, so calls chain. Each
    change addresses the workbook as the previous ones leave it: after `delete_rows("T!3")`,
    `"T!3"` is what was row 4. The scopes of one call are simultaneous: `delete_rows("T!3",
    "T!4")` deletes two rows. `save` runs the changes in the fewest passes over the file that
    give what running them one at a time gives; each pass reads and writes the file once, as a
    stream. `explain` says what `save` will do.

    Scopes are `"Sheet"`, `"Sheet!B:D"` (columns) and `"Sheet!3:10"` (rows); quote a sheet name
    that contains `!`. A malformed scope, or rows deleted with a column range, raises
    `ValueError` where it is written; an error of the workbook (a missing sheet, a formula with a
    3D reference, issue #75) shows at `save`.
    """

    def __init__(self, source: Source) -> None: ...
    def remove_sheets(self, *names: str) -> Editor: ...
    def remove_ranges(self, *names: str) -> Editor: ...
    def convert_to_values(self, *scopes: str) -> Editor: ...
    def delete_rows(self, *rows: str) -> Editor: ...
    def delete_columns(self, *columns: str) -> Editor: ...
    def explain(self) -> str:
        """What `save` will do: the passes over the file, and the mark to recalculate."""
    def save(self, path: Path) -> None:
        """Run the changes and write the result to `path` (never the file it was opened from).

        The editor decides whether the saved file is marked so that the application that opens it
        recalculates every formula, and puts the mark last: it does when a change leaves the
        cached results of formulas stale (rows, columns, sheets or names removed). There is no
        option for it.
        """

def open(source: Path | bytes, *, extension: str | None = None) -> Source:
    """Open a workbook: make the `Source` that every tool takes. Bytes need `extension`."""
