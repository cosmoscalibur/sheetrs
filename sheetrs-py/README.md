# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs

source = sheetrs.open("book.xlsx")        # a reference to the workbook; nothing is read

# Linter (sheetlint). The configuration is validated when it is built: an unknown rule,
# category or parameter raises ValueError instead of being ignored.
linter = sheetrs.Linter(config_path="sheetlint.toml")
for v in linter.lint(source):
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)
linter.rules()      # [{"id", "name", "category", "default", "enabled"}, ...]

# Statistics (sheetstats): a dict, as `sheetstats --format json`
source.stats()["total_formulas"]

# Editing (sheetedit): the operations on a source make an EditPlan, which has save()
plan = (source.remove_sheets("Secrets")
              .convert_to_values("Report!B:D")
              .delete_rows("Data!3:10"))
print(plan.explain())                       # the passes over the file
plan.save("clean.xlsx")
```

What each name is:

- **`open` / `Source`**: `open` makes the input of every tool, the way a file or bytes are given to
  Rust's `Source::from_file`. A `Source` is a *reference* to a workbook for lazy processing: it
  reads nothing and is not a navigable model of the workbook. The same `Source` is given to
  `Linter.lint`, measured with `Source.stats()`, or edited. It has **no `save`**: opening is not
  editing, so `open(path).save(...)` is an `AttributeError`, not a copy.
- **`Linter`** is the linter with its configuration, built once and used on many sources; there
  is one way to lint. `Linter.rules()` lists the rules with the ids and categories that
  `disabled_rules`/`enabled_rules` accept and whether *this* linter's configuration runs each one;
  rules are configured through the TOML, not through that list. **`Violation`** is the type of the
  items `lint` returns (rule, severity, sheet, cell, message); it is exposed so results can be
  annotated and checked, and is not built by the user.
- **`EditPlan`**: the operations of a source (`remove_sheets`, `remove_ranges`,
  `convert_to_values`, `delete_rows`, `delete_columns`) make one, and each further operation
  returns the plan that results (the one it was called on is not changed). Editing has one mode: a
  workbook is streamed, never loaded, so nothing is read until `save(path)`, which never overwrites
  the file it was opened from.
- **Operations are sequential**: each addresses the workbook as the previous ones left it (after
  `delete_rows("T!3")`, `"T!3"` is the old row 4). The scopes of *one* call are simultaneous:
  `delete_rows("T!3", "T!4")` deletes two rows. `save` plans the sequence into the fewest passes
  over the file that give what running the operations one at a time gives (a conversion after
  another change, and deletions that touch, need their own pass); each pass reads and writes the
  file once as a stream, so memory does not grow with the workbook. `explain()` shows the passes.
- **Recalculation is the plan's decision.** When an operation leaves the cached results of
  formulas stale (rows, columns, sheets or names removed), the plan marks the saved file so that
  the application that opens it recalculates every formula, and puts the mark last, after every
  operation. You do not ask for it; you can avoid it (`save(path, recalc_on_open=False)`, to keep
  the cached results) or force it (`True`).
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
