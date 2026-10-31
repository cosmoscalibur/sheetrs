"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite (``sheetlint``, ``sheetstats``, ``sheetedit``). A workbook
is a path, its bytes, or a ``Workbook`` already loaded::

    import sheetrs

    linter = sheetrs.Linter(config_path="sheetlint.toml")   # the configuration is validated
    for violation in linter.lint("book.xlsx"):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)

    sheetrs.stats("book.xlsx")["total_formulas"]

Editing has an eager and a lazy mode with the same methods; what differs is what is loaded and
when it runs, as in dataframe libraries::

    # eager: read() loads a working copy; each operation runs when it is called
    wb = sheetrs.read("book.xlsx").remove_sheets("Secrets").delete_rows("Data!3:10")
    linter.lint(wb)                      # the result can be inspected as it is
    wb.write("clean.xlsx")

    # lazy: scan() reads nothing; operations are recorded and run, in as few passes over the
    # file as possible, by sink() or collect()
    (sheetrs.scan("book.xlsx").remove_sheets("Secrets").delete_rows("Data!3:10")
        .sink("clean.xlsx"))
"""

from ._sheetrs import (
    LazyWorkbook,
    Linter,
    SheetrsError,
    Violation,
    Workbook,
    __version__,
    read,
    scan,
    stats,
)

__all__ = [
    "LazyWorkbook",
    "Linter",
    "SheetrsError",
    "Violation",
    "Workbook",
    "__version__",
    "read",
    "scan",
    "stats",
]
