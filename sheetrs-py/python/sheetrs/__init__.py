"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite, one submodule per tool::

    import sheetrs

    for violation in sheetrs.linter.lint("book.xlsx"):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)
    sheetrs.linter.rules()                       # the catalog of the linter's rules

    sheetrs.stats.compute("book.xlsx")["total_formulas"]
    sheetrs.edit.modify("book.xlsx", "clean.xlsx", remove_sheets=["Secrets"])
"""

import sys as _sys

from ._sheetrs import SheetrsError, __version__, edit, linter, stats

# The submodules are native: register them so that `from sheetrs.linter import lint` works.
for _module in (edit, linter, stats):
    _sys.modules[f"{__name__}.{_module.__name__.rsplit('.', 1)[-1]}"] = _module

__all__ = ["SheetrsError", "__version__", "edit", "linter", "stats"]
