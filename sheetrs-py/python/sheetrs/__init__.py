"""Lint, inspect, query, edit and calculate XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite: one object, ``Book``, a workbook (a file, its bytes or a binary
stream). It reads the workbook once when it is opened, to count what it holds, and keeps only the
numbers; the queries read what they need.

* The **properties** say what the opening found out: ``format``, ``sheet_names``, ``sheets[name]``
  (with ``used_range``, ``state`` and counts), ``defined_names``, ``date_system``, ``has_macros``, the
  eleven ``total_*`` and ``stats``.
* The **operations** (``delete_sheets``, ``delete_names``, ``formula_to_value``, ``delete_rows``,
  ``delete_columns``, ``replace`` and ``calculate``) accumulate on the ``Book`` and return it, so the
  calls chain. Each is written against the workbook as it is: the order they are given in changes
  nothing, and nothing runs until ``save``, the cut, which plans all of them in the fewest passes over
  the file and writes the result.
* The **queries** (``range``, ``find``, ``lint``, ``inspect``) read what they need. A query on a
  ``Book`` with operations that nothing has run raises ``PendingOperationsError``.
* ``save(dest)`` writes a new workbook to a path or a binary stream (``io.BytesIO``); ``save()``
  replaces the original, the file or the stream the ``Book`` was made from; ``json=`` writes the
  value ranges of a ``calculate`` step as JSON; ``value_ranges()`` gives them in memory.
* Every result is typed and writes itself explicitly: ``to_dict()`` (dicts and lists), ``to_json()``
  (text) and, for the lint report, ``to_sarif()`` and ``to_text()``.
* What a run reports besides its result is a ``Diagnostic`` (level, stable code, message, context
  fields): in ``book.report.diagnostics``, in the ``sheetrs`` logger of ``logging`` (code and fields in
  ``extra``) and, for the warnings, through ``warnings.warn``.

::

    import io
    import sheetrs

    book = sheetrs.Book("book.xlsx")
    book.sheet_names, book.sheets["Data"].used_range, book.total_formulas

    r = book.range(["Data!A1:C10", "Summary!B2"], properties=("value", "formula"))
    r["Data!A1:C10"].values            # list of rows; an error cell is a sheetrs.CellError

    (book.delete_rows("Data!3:10")
         .formula_to_value("Report!B:D")
         .replace("2024", "2025")
         .calculate())                  # the engine recalculates the edited workbook
    answer = book.value_ranges()       # a ValueRanges, written nowhere; answer.to_dict() for the lists
    print(book.explain())              # the passes over the file
    clean = book.save("clean.xlsx")    # the cut; a Book on the written file
    clean.save(io.BytesIO())           # to a stream
    print(book.lint().to_sarif())      # the lint report as SARIF
"""

from ._errors import (
    Error,
    FormatError,
    InvalidRangeError,
    PendingOperationsError,
    RecalculationWarning,
    UnavailableFunctionError,
    UnsupportedDateSystemError,
)
from ._sheetrs import (
    Block,
    Book,
    CellError,
    Diagnostic,
    FunctionUse,
    Inspection,
    LintReport,
    Match,
    Report,
    Sheet,
    Sheets,
    Skipped,
    Stats,
    ValueRanges,
    Violation,
    __version__,
    evaluate,
    lint_rules,
)

__all__ = [
    "Block",
    "Book",
    "CellError",
    "Diagnostic",
    "Error",
    "FormatError",
    "FunctionUse",
    "InvalidRangeError",
    "Inspection",
    "LintReport",
    "Match",
    "PendingOperationsError",
    "RecalculationWarning",
    "Report",
    "Sheet",
    "Sheets",
    "Skipped",
    "Stats",
    "UnavailableFunctionError",
    "UnsupportedDateSystemError",
    "ValueRanges",
    "Violation",
    "__version__",
    "evaluate",
    "lint_rules",
]
