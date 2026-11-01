# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs

# Linter (sheetlint). The configuration is validated when it is built: an unknown rule,
# category or parameter raises ValueError instead of being ignored.
linter = sheetrs.Linter(config_path="sheetlint.toml")
for v in linter.lint("book.xlsx"):          # or bytes, with extension="ods"
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)
linter.rules()      # [{"id", "name", "category", "default", "enabled"}, ...]

# Statistics (sheetstats): a dict, as `sheetstats --format json`
sheetrs.stats("book.xlsx")["total_formulas"]

# Editing (sheetedit): open, record operations, save
wb = (sheetrs.open("book.xlsx")
      .remove_sheets("Secrets")
      .convert_to_values("Report!B:D")
      .delete_rows("Data!3:10"))
print(wb.explain())                          # the passes over the file
wb.save("clean.xlsx", recalc_on_open=True)
```

What each name is:

- **`Linter`** is the linter with its configuration, built once and used on many workbooks; there
  is one way to lint. `Linter.rules()` lists the rules with the ids and categories that
  `disabled_rules`/`enabled_rules` accept and whether *this* linter's configuration runs each one;
  rules are configured through the TOML, not through that list. **`Violation`** is the type of the
  items `lint` returns (rule, severity, sheet, cell, message); it is exposed so results can be
  annotated and checked, and is not built by the user.
- **`open` / `Workbook`**: editing has one mode. A workbook is streamed, never loaded, so there is
  no model to apply an operation to before `save`: `open` checks the path and the extension,
  each method records an operation and returns the `Workbook` that results (the one it was called
  on is not changed), and nothing is read until `save(path)`, which runs the operations and never
  overwrites the file it was opened from.
- **Operations are sequential**: each addresses the workbook as the previous ones left it (after
  `delete_rows("T!3")`, `"T!3"` is the old row 4). The scopes of *one* call are simultaneous:
  `delete_rows("T!3", "T!4")` deletes two rows. `save` plans the sequence into the fewest passes
  over the file that give what running the operations one at a time gives (a conversion after
  another change, and deletions that touch, need their own pass); each pass reads and writes the
  file once as a stream, so memory does not grow with the workbook. `explain()` shows the passes.
- **`recalc_on_open`** is an argument of `save`, not an operation: it marks the *saved file* so
  that the application that opens it recalculates every formula. It changes how the file is opened,
  not what it holds, so it is applied once, at the end, and cannot change what a conversion to
  values reads.
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
