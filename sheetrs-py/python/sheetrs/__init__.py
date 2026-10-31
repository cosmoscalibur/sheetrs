"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite, one submodule per tool::

    import sheetrs

    linter = sheetrs.linter.Linter(config_path="sheetlint.toml")   # configuration is validated
    for violation in linter.lint("book.xlsx"):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)
    linter.rules()                               # the rules and which of them this config runs

    sheetrs.stats.compute("book.xlsx")["total_formulas"]

    # lazy: collect operations, one pass over the file on save
    (sheetrs.edit.Plan("book.xlsx").remove_sheets("Secrets").convert_to_values("Report!B:D")
        .save("clean.xlsx"))
    # eager: one operation now
    sheetrs.edit.remove_sheets("book.xlsx", "clean.xlsx", "Secrets")
"""

import sys as _sys

from ._sheetrs import SheetrsError, __version__, edit, linter, stats

# The submodules are native: register them so that `from sheetrs.linter import lint` works.
for _module in (edit, linter, stats):
    _sys.modules[f"{__name__}.{_module.__name__.rsplit('.', 1)[-1]}"] = _module

__all__ = ["SheetrsError", "__version__", "edit", "linter", "stats"]
