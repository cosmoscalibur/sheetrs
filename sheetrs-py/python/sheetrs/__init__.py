"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite (``sheetlint``, ``sheetstats``, ``sheetedit``). A workbook
is a path or its bytes::

    import sheetrs

    linter = sheetrs.Linter(config_path="sheetlint.toml")   # the configuration is validated
    for violation in linter.lint("book.xlsx"):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)
    linter.rules()        # the rules, and which of them this configuration runs

    sheetrs.stats("book.xlsx")["total_formulas"]

    # lazy: collect operations (against the input workbook); one pass over the file on save
    (sheetrs.EditPlan("book.xlsx").remove_sheets("Secrets").convert_to_values("Report!B:D")
        .save("clean.xlsx"))
    # eager: the same operations given at once and applied at once
    sheetrs.edit("book.xlsx", "clean.xlsx", remove_sheets=["Secrets"])
"""

from ._sheetrs import (
    EditPlan,
    Linter,
    SheetrsError,
    Violation,
    __version__,
    edit,
    lint,
    stats,
)

__all__ = [
    "EditPlan",
    "Linter",
    "SheetrsError",
    "Violation",
    "__version__",
    "edit",
    "lint",
    "stats",
]
