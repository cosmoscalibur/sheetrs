"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite (``sheetlint``, ``sheetstats``, ``sheetedit``).

``open`` makes the input of every tool: a ``Source``, a reference to a workbook that reads nothing
and is not a navigable model of it. Each tool takes it::

    import sheetrs

    source = sheetrs.open("book.xlsx")

    linter = sheetrs.Linter(config_path="sheetlint.toml")    # the configuration is validated
    for violation in linter.lint(source):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)

    source.stats()["total_formulas"]

    # Editing: the operations on a source make an EditPlan, which is what has save().
    # Nothing is read until save(), which runs them in as few passes over the file as it can.
    (source.remove_sheets("Secrets")
        .delete_rows("Data!3:10")
        .save("clean.xlsx"))
"""

from ._sheetrs import (
    EditPlan,
    Linter,
    SheetrsError,
    Source,
    Violation,
    __version__,
    open,
)

__all__ = [
    "EditPlan",
    "Linter",
    "SheetrsError",
    "Source",
    "Violation",
    "__version__",
    "open",
]
