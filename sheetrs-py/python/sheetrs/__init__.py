"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite (``sheetlint``, ``sheetstats``, ``sheetedit``)::

    import sheetrs

    for violation in sheetrs.lint("book.xlsx"):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)

    print(sheetrs.stats("book.xlsx")["total_formulas"])
    sheetrs.edit("book.xlsx", "clean.xlsx", remove_sheets=["Secrets"])
"""

from ._sheetrs import (
    SheetrsError,
    Violation,
    __version__,
    edit,
    lint,
    rules,
    stats,
)

__all__ = [
    "SheetrsError",
    "Violation",
    "__version__",
    "edit",
    "lint",
    "rules",
    "stats",
]
