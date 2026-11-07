# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs
from sheetrs import edit, linter, stats

source = sheetrs.open("book.xlsx")        # a reference to the workbook; nothing is read

# Linter (sheetlint). The configuration is validated when it is built: an unknown rule,
# category or parameter raises ValueError instead of being ignored.
lint = linter.Linter(config_path="sheetlint.toml")
for v in lint.lint(source):
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)
lint.rules()        # [{"id", "name", "category", "default", "enabled"}, ...]

# Statistics (sheetstats): a dict, as `sheetstats --format json`
stats.collect(source)["total_formulas"]

# Editing (sheetedit): changes are collected and saved
editor = (edit.Editor(source)
          .remove_sheets("Secrets")
          .convert_to_values("Report!B:D")
          .delete_rows("Data!3:10"))
print(editor.explain())                     # the passes over the file
editor.save("clean.xlsx")
```

The package is laid out as the crates are, so each tool has one home:

| Python | Crate | Executable |
| ------ | ----- | ---------- |
| `sheetrs` (`open`, `Source`, `SheetrsError`) | `sheetrs-io` | |
| `sheetrs.linter` (`Linter`, `Violation`) | `sheetrs-linter` | `sheetlint` |
| `sheetrs.stats` (`collect`) | `sheetrs-stats` | `sheetstats` |
| `sheetrs.edit` (`Editor`) | `sheetrs-edit` | `sheetedit` |

**The convention: the tool is the subject and takes the `Source`.** It is an object when it has a
configuration or accumulates state (`Linter`, `Editor`), and a function when it has neither
(`stats.collect`). `Source` has no behavior of its own, the same as in Rust, where the core cannot
depend on the tools.

What each name is:

- **`open` / `Source`**: `open` makes the input of every tool, the way a file or bytes are given to
  Rust's `Source::from_file`. A `Source` is a *reference* to a workbook for lazy processing: it
  reads nothing and is not a navigable model of the workbook. The same `Source` is given to a
  `Linter`, to `stats.collect` or to an `Editor`. It has **no `save`** and no other method:
  opening is not editing.
- **`Linter`** is the linter with its configuration, built once and used on many sources; there
  is one way to lint. `Linter.rules()` lists the rules with the ids and categories that
  `disabled_rules`/`enabled_rules` accept and whether *this* linter's configuration runs each one;
  rules are configured through the TOML, not through that list. **`Violation`** is the type of the
  items `lint` returns (rule, severity, sheet, cell, message); it is exposed so results can be
  annotated and checked, and is not built by the user.
- **`Editor`** collects changes for one source (`remove_sheets`, `remove_ranges`,
  `convert_to_values`, `delete_rows`, `delete_columns`); each returns the editor, so calls chain.
  Editing has one mode: a workbook is streamed, never loaded, so nothing is read until
  `save(path)`, which never overwrites the file it was opened from. `explain()` says what `save`
  will do.
- **Changes are sequential**: each addresses the workbook as the previous ones left it (after
  `delete_rows("T!3")`, `"T!3"` is the old row 4). The scopes of *one* call are simultaneous:
  `delete_rows("T!3", "T!4")` deletes two rows. `save` plans the sequence into the fewest passes
  over the file that give what running the changes one at a time gives (a conversion after
  another change, and deletions that touch, need their own pass); each pass reads and writes the
  file once as a stream, so memory does not grow with the workbook.
- **Recalculation is the editor's decision, and there is no option for it.** When a change leaves
  the cached results of formulas stale (rows, columns, sheets or names removed), the editor marks
  the saved file so that the application that opens it recalculates every formula, and puts the
  mark last, after every change. An editor that only converts formulas to values does not mark. It
  is not asked for and it cannot be avoided: whoever does not know what a stale cached result is
  would get a wrong file by choosing.
- A scope that is malformed, or of the wrong kind (rows deleted with a column range), raises
  `ValueError` where it is written; an error of the workbook shows at `save`.
- **3D references** (`Sheet1:Sheet3!A1`) are not supported: deleting rows or columns or removing
  sheets in a workbook with a formula that holds one raises `SheetrsError`, naming issue #75.

Errors: `ValueError` (invalid configuration or scope), `OSError` (a file cannot
be read), `sheetrs.SheetrsError` (a workbook cannot be parsed or written).

## Development

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop          # builds and installs the extension in the venv
pytest
```
