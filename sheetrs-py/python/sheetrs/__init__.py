"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite (``sheetlint``, ``sheetstats``, ``sheetedit``). A workbook
is a path or its bytes::

    import sheetrs

    linter = sheetrs.Linter(config_path="sheetlint.toml")   # the configuration is validated
    for violation in linter.lint("book.xlsx"):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)

    sheetrs.stats("book.xlsx")["total_formulas"]

    # open, record operations, save: nothing is read until save(), which runs them over the
    # file in as few passes as it can
    (sheetrs.open("book.xlsx")
        .remove_sheets("Secrets")
        .delete_rows("Data!3:10")
        .save("clean.xlsx"))
"""

from ._sheetrs import (
    Linter,
    SheetrsError,
    Violation,
    Workbook,
    __version__,
    open,
    stats,
)

__all__ = [
    "Linter",
    "SheetrsError",
    "Violation",
    "Workbook",
    "__version__",
    "open",
    "stats",
]
