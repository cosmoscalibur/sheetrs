import datetime
import os
from collections.abc import Iterable, Iterator, Mapping
from typing import Any, BinaryIO, Literal, Self, TypedDict, overload

__version__: str

Path = str | os.PathLike[str]
Source = Path | bytes | bytearray | BinaryIO
"""A workbook: a path, its bytes or a binary stream that reads and seeks (`io.BytesIO`)."""
Destination = Path | BinaryIO
"""Where a workbook or its value ranges are written: a path or a binary stream."""
Targets = str | Iterable[str]
"""One target (a `str`) or an iterable of them: every operation takes the array of its targets."""
Properties = str | Iterable[Literal["value", "formula", "number_format", "display", "comment"]]
Cell = float | int | str | bool | None | CellError | datetime.date | datetime.time | datetime.datetime

class CellError:
    """A cell that holds an error: `#DIV/0!`, `#N/A`, or the code only LibreOffice has (`Err:502`).

    It is not a `str`: a cell with the text `#N/A` is a `str`, and one that calculated to `#N/A` is a
    `CellError`. Two errors are equal when their codes are; an error is never equal to a `str`.
    """

    def __init__(self, code: str) -> None: ...
    @property
    def code(self) -> str: ...

class Block:
    """The cells of one range: one list of rows per property that was asked for.

    A bounded range is as big as the range, with `None` where a cell is empty; a range open at one
    end (`A:C`, `3:10`) is cut at its last non-empty cell. It iterates and indexes the rows of
    `values` (of the first property asked for when `value` was not). A property that was not asked
    for raises `AttributeError`.
    """

    @property
    def range(self) -> str: ...
    @property
    def major(self) -> Literal["rows", "columns"]: ...
    @property
    def values(self) -> list[list[Cell]]: ...
    @property
    def formulas(self) -> list[list[str | None]]: ...
    @property
    def number_formats(self) -> list[list[str | None]]: ...
    @property
    def displays(self) -> list[list[str | None]]: ...
    @property
    def comments(self) -> list[list[str | None]]: ...
    def __len__(self) -> int: ...
    def __iter__(self) -> Iterator[list[Any]]: ...
    @overload
    def __getitem__(self, index: int) -> list[Any]: ...
    @overload
    def __getitem__(self, index: slice) -> list[list[Any]]: ...
    def to_dicts(self) -> list[dict[str, Cell]]:
        """The rows as dicts, the first row being the header (as `str` keys)."""

class Match:
    """One cell a search found, with the properties that were asked for (`None` for the others)."""

    sheet: str
    cell: str
    row: int
    col: int
    value: Cell
    formula: str | None
    number_format: str | None
    display: str | None
    comment: str | None

class Violation:
    """One finding of the linter: the type of the items `Book.lint` returns."""

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

class Diagnostic:
    """One thing a run reports besides its result.

    The same records go to the `sheetrs` logger of `logging` (the `code` and the context fields in
    `extra`) and, the warnings, to `warnings.warn`.
    """

    @property
    def level(self) -> Literal["error", "warning", "info", "debug"]: ...
    @property
    def code(self) -> str:
        """The stable name of what happened (`"RECALC_NOT_SUPPORTED"`): it does not change with the
        wording of the message."""
    @property
    def message(self) -> str: ...
    @property
    def context(self) -> dict[str, str]:
        """The fields the message was made from (`operation`, `reason`, `sheet`, `cell`...)."""
    def to_dict(self) -> dict[str, Any]: ...
    def to_json(self) -> str: ...

class Skipped:
    """A cell a `replace` found and left as it was."""

    sheet: str
    cell: str
    reason: str
    def to_dict(self) -> dict[str, str]: ...

class Report:
    """What a run did besides writing the workbook."""

    @property
    def replaced(self) -> int:
        """The cells a `replace` changed."""
    @property
    def written(self) -> int: ...
    @property
    def skipped(self) -> list[Skipped]: ...
    @property
    def diagnostics(self) -> list[Diagnostic]: ...
    def to_dict(self) -> dict[str, Any]: ...
    def to_json(self) -> str: ...

class ValueRanges:
    """The value ranges of a calculation: a mapping from the name of each one (the range it was asked
    with, or the name of the declared output) to a `Block` with its values."""

    def __len__(self) -> int: ...
    def __contains__(self, key: str) -> bool: ...
    def __getitem__(self, key: str) -> Block: ...
    def __iter__(self) -> Iterator[str]: ...
    def keys(self) -> list[str]: ...
    def values(self) -> list[Block]: ...
    def items(self) -> list[tuple[str, Block]]: ...
    def to_dict(self) -> dict[str, list[list[Cell]]]:
        """`{name: rows}`: the values as the lists Python has."""
    def to_json(self) -> str:
        """The JSON answer of `sheetc calculate` (`{"valueRanges": [...]}`), as text."""

class Sheet:
    """A sheet of the workbook and what the description of the workbook found out about it."""

    @property
    def name(self) -> str: ...
    @property
    def used_range(self) -> str | None:
        """The rectangle of the cells that hold a value or a formula, from the first used row and
        column to the last (`"D10:GV5439"`); `None` for a sheet with none."""
    @property
    def state(self) -> Literal["visible", "hidden", "very_hidden"]: ...
    @property
    def used_cells(self) -> int: ...
    @property
    def formulas(self) -> int: ...
    @property
    def errors(self) -> int: ...
    @property
    def empty_cells(self) -> int: ...
    @property
    def comments(self) -> int: ...
    @property
    def charts(self) -> int: ...
    @property
    def tables(self) -> int: ...
    @property
    def xml_size(self) -> int: ...
    def to_dict(self) -> dict[str, Any]: ...
    def to_json(self) -> str: ...

class Sheets:
    """The sheets of a workbook, by name or by position, in order."""

    def __len__(self) -> int: ...
    def __contains__(self, name: str) -> bool: ...
    def __getitem__(self, key: str | int) -> Sheet: ...
    def __iter__(self) -> Iterator[Sheet]: ...
    def names(self) -> list[str]: ...

class Stats:
    """The statistics of the workbook: the eleven totals, the sheets, `has_macros` and `defined_names`, as
    `sheetrs stats` reports them (`to_dict()` is its `--format json`; it holds the totals and the sheets)."""

    @property
    def total_sheets(self) -> int: ...
    @property
    def total_named_ranges(self) -> int: ...
    @property
    def total_file_size(self) -> int: ...
    @property
    def total_used_cells(self) -> int: ...
    @property
    def total_formulas(self) -> int: ...
    @property
    def total_errors(self) -> int: ...
    @property
    def total_empty_cells(self) -> int: ...
    @property
    def total_comments(self) -> int: ...
    @property
    def total_charts(self) -> int: ...
    @property
    def total_tables(self) -> int: ...
    @property
    def total_xml_size(self) -> int: ...
    @property
    def sheets(self) -> Sheets: ...
    @property
    def has_macros(self) -> bool:
        """Whether the workbook carries macros (found, never run)."""
    @property
    def defined_names(self) -> dict[str, str]:
        """Name to what it refers to, as the file writes it, ordered by name."""
    def to_dict(self) -> dict[str, Any]: ...
    def to_json(self) -> str: ...

class LintReport:
    """The result of a lint run: a sequence of `Violation` that writes itself as `sheetrs lint` does.
    `file` is the name the documents give the workbook."""

    @property
    def violations(self) -> list[Violation]: ...
    def __len__(self) -> int: ...
    def __iter__(self) -> Iterator[Violation]: ...
    @overload
    def __getitem__(self, index: int) -> Violation: ...
    @overload
    def __getitem__(self, index: slice) -> list[Violation]: ...
    def to_text(self, file: str | None = None) -> str:
        """The report for a person, as plain text."""
    def to_json(self, file: str | None = None) -> str:
        """The JSON document of `sheetrs lint --format json`."""
    def to_sarif(self, file: str | None = None) -> str:
        """The SARIF 2.1.0 log of `sheetrs lint --format sarif`."""
    def to_dict(self, file: str | None = None) -> dict[str, Any]:
        """The JSON document as dicts and lists."""

class FunctionUse:
    """A function a workbook uses."""

    name: str
    calls: int
    implemented: bool

class Inspection:
    """What `sheetc inspect` says about a workbook."""

    @property
    def sheets(self) -> list[str]: ...
    @property
    def date_system(self) -> str:
        """The null date (`"1899-12-30"`, `"1900-01-01"` or `"1904-01-01"`)."""
    @property
    def cells(self) -> int: ...
    @property
    def formulas(self) -> int: ...
    @property
    def circular_groups(self) -> int: ...
    @property
    def run_time_references(self) -> bool: ...
    @property
    def unsupported_functions(self) -> list[str]: ...
    @property
    def functions(self) -> list[FunctionUse]: ...
    @property
    def diagnostics(self) -> list[str]: ...
    def to_dict(self) -> dict[str, Any]: ...
    def to_json(self) -> str: ...

class Book:
    """A workbook: a file, its bytes or a binary stream, with the operations to make to it.

    It reads nothing when it is made and keeps nothing of the workbook. The operations
    (`delete_sheets`, `delete_names`, `formula_to_value`, `delete_rows`, `delete_columns`, `replace`, `calculate`) accumulate on the same `Book` and return
    it, so the calls chain; each is written against the workbook as it is, the order they are given
    in changes nothing, and nothing runs until `save`, the cut. The queries (`range`, `find`, `lint`,
    `inspect`, `stats`) raise `PendingOperationsError` while operations are pending.
    """

    def __init__(self, source: Source) -> None: ...
    @property
    def format(self) -> Literal["xlsx", "xlsm", "ods"]: ...
    def stats(self) -> Stats:
        """The statistics of the workbook, as `sheetrs stats` reports them: the totals, the sheets, the
        macros and the defined names. It reads the workbook each time it is called; the `Book` keeps
        nothing of what it found. Raises `PendingOperationsError` with operations pending."""
    @property
    def report(self) -> Report | None:
        """What the last `save` or `to_bytes` did besides writing the workbook; `None` for a `Book`
        nothing wrote."""
    def delete_sheets(self, names: Targets) -> Self:
        """Delete sheets by name: one `str` or an iterable of them (a sweep gives its array).
        An empty iterable does nothing."""
    def delete_names(self, names: Targets) -> Self:
        """Delete defined names: one `str` or an iterable of them."""
    def formula_to_value(self, scopes: Targets) -> Self:
        """Replace the formulas of the scopes (`"Sheet"`, `"Sheet!B:D"`, `"Sheet!3:10"`) by their
        values: one `str` or an iterable of them."""
    def delete_rows(self, rows: Targets) -> Self:
        """Delete rows (`"Sheet!3:10"`, `"Sheet!3"`): one `str` or an iterable of them, simultaneous
        within the call."""
    def delete_columns(self, columns: Targets) -> Self:
        """Delete columns (`"Sheet!B:D"`, `"Sheet!B"`): one `str` or an iterable of them,
        simultaneous within the call."""
    def hide_rows(self, rows: Targets) -> Self:
        """Hide rows (`"Sheet!3:10"`, `"Sheet!3"`): one `str` or an iterable of them. The rows stay in
        the file; the ranges are those of the workbook as it is, before any deletion."""
    def show_rows(self, rows: Targets) -> Self:
        """Show rows, whatever hid them (by hand or by a filter)."""
    def hide_columns(self, columns: Targets) -> Self:
        """Hide columns (`"Sheet!B:D"`, `"Sheet!B"`): one `str` or an iterable of them."""
    def show_columns(self, columns: Targets) -> Self:
        """Show columns."""
    def hide_sheets(self, names: Targets) -> Self:
        """Hide sheets by name; one sheet must stay visible. A very hidden XLSX sheet stays so."""
    def show_sheets(self, names: Targets) -> Self:
        """Show sheets by name (hidden or very hidden)."""
    @overload
    def replace(
        self,
        text: str,
        replacement: str,
        *,
        regex: bool = False,
        wildcards: bool = False,
        match_case: bool = False,
        entire_cells: bool = False,
        search_in: Literal["values", "formulas", "comments"] = "values",
        formatted: bool = False,
        filtered: bool = False,
        scope: Targets | None = None,
    ) -> Self:
        """Find `text` and replace it with `replacement`."""
    @overload
    def replace(
        self,
        text: Mapping[str, str] | Iterable[tuple[str, str]],
        replacement: None = None,
        *,
        regex: bool = False,
        wildcards: bool = False,
        match_case: bool = False,
        entire_cells: bool = False,
        search_in: Literal["values", "formulas", "comments"] = "values",
        formatted: bool = False,
        filtered: bool = False,
        scope: Targets | None = None,
    ) -> Self:
        """A sweep: each `(text, replacement)` pair is an operation of the plan, in order, that
        sees what the pair before it left (as if `replace` were called once per pair); the options
        and the scope are shared. The order matters: `{"a": "b", "b": "c"}` turns `a` into `c`."""
    def calculate(
        self,
        inputs: Mapping[str, Any] | None = None,
        ranges: Targets | None = None,
        *,
        major: Literal["rows", "columns"] = "rows",
    ) -> Self:
        """An operation: the engine writes `inputs` and calculates every formula of the workbook as it
        is, before the conversions and the deletions of the same `save` (the operations are not a
        sequence: each is written against the workbook as it is, and the planner orders them);
        `save` writes the results and no mark to recalculate unless a deletion after leaves some stale.
        `ranges` are the ranges `value_ranges()` gives and `save(json=...)` writes, listed as `major` says
        (`"rows"`, the default, or `"columns"`, the `majorDimension` of the request). Where the engine
        cannot calculate the workbook it is written marked to be recalculated, with a
        `RecalculationWarning`."""
    def value_ranges(self) -> ValueRanges:
        """The value ranges of the calculation, in memory: runs the operations and writes the workbook
        nowhere. Raises `Error` without a `calculate` step."""
    def explain(self) -> str: ...
    def save(self, dest: Destination | None = None, *, json: Destination | None = None) -> Book:
        """The cut: run the operations and write the workbook.

        `save(dest)` writes to a path or to a binary stream (`io.BytesIO()`) and returns a `Book` on what
        it wrote. `save()` replaces the original atomically, the file or the stream the `Book` was made
        from (rewritten from its start), and returns this `Book`, which then has no operations. `json`
        is a path or a stream where the value ranges of the `calculate` step are written as JSON besides the
        workbook.
        """
    def to_bytes(self) -> bytes:
        """The workbook the operations make, as bytes in the format of the source: the cut like `save`,
        but to memory, with no file. The bytes are the ones `save(path)` writes for the same plan. Like
        `save(path)` it keeps the operations of this `Book`; `report` says what the run did and a
        `RecalculationWarning` is issued where `save` would issue it. With no operations it returns
        the workbook as it is (the file, or the bytes the `Book` was made from)."""
    def range(
        self,
        ranges: Targets,
        properties: Properties | None = None,
        *,
        major: Literal["rows", "columns"] = "rows",
        dates: bool = False,
    ) -> dict[str, Block]: ...
    def find(
        self,
        text: str,
        *,
        regex: bool = False,
        wildcards: bool = False,
        match_case: bool = False,
        entire_cells: bool = False,
        search_in: Literal["values", "formulas", "comments"] = "values",
        include_formulas: bool = False,
        formatted: bool = False,
        filtered: bool = False,
        scope: Targets | None = None,
        properties: Properties | None = None,
        by: Literal["rows", "columns"] = "rows",
        dates: bool = False,
    ) -> list[Match]: ...
    def lint(self, config_path: Path | None = None) -> LintReport:
        """The violations of the rules in the configuration file `config_path` (the default rules
        without it), as a `LintReport`."""
    def inspect(self) -> Inspection:
        """What `sheetc inspect` says: the sheets, the cells and formulas, the functions used and
        which of them the engine lacks. Compiles the workbook in memory; calculates nothing."""

def evaluate(formula: str, *, locale: str | None = None) -> Cell:
    """The value of one formula, as `sheetc evaluate` gives it."""

def lint_rules(config_path: Path | None = None) -> list[Rule]:
    """The rules ordered by id; `enabled` tells whether the configuration runs each one."""
