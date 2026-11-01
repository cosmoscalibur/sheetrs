"""Lint, inspect and edit XLSX and ODS spreadsheets.

Python bindings of the SheetRS suite. The package is laid out as the crates are:

* ``sheetrs``: ``open`` and ``Source`` (the input of every tool), ``SheetrsError``.
* ``sheetrs.linter``: ``Linter`` (``sheetlint``).
* ``sheetrs.stats``: ``collect`` (``sheetstats``).
* ``sheetrs.edit``: ``Editor`` (``sheetedit``).

``open`` makes a ``Source``, a reference to a workbook that reads nothing and is not a navigable
model of it. The tool is the subject and takes it::

    import sheetrs
    from sheetrs import edit, linter, stats

    source = sheetrs.open("book.xlsx")

    for violation in linter.Linter(config_path="sheetlint.toml").lint(source):
        print(violation.rule_id, violation.sheet, violation.cell, violation.message)

    stats.collect(source)["total_formulas"]

    # Nothing is read until save(), which runs the changes in as few passes over the file as it can.
    (edit.Editor(source)
        .remove_sheets("Secrets")
        .delete_rows("Data!3:10")
        .save("clean.xlsx"))

An object when the tool has a configuration or accumulates state (``Linter``, ``Editor``), a
function when it has neither (``stats.collect``).
"""

from . import edit, linter, stats
from ._sheetrs import SheetrsError, Source, __version__, open

__all__ = [
    "SheetrsError",
    "Source",
    "__version__",
    "edit",
    "linter",
    "open",
    "stats",
]
