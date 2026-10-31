# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs

# Linter (sheetlint). The configuration is validated when it is built: an unknown rule,
# category or parameter raises ValueError instead of being ignored.
linter = sheetrs.Linter(config_path="sheetlint.toml")
for v in linter.lint("book.xlsx"):          # or bytes (with extension="ods"), or a Workbook
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)
linter.rules()      # [{"id", "name", "category", "default", "enabled"}, ...]

# Statistics (sheetstats): a dict, as `sheetstats --format json`
sheetrs.stats("book.xlsx")["total_formulas"]

# Editing, eager: read() loads a working copy; each operation runs when it is called
wb = sheetrs.read("book.xlsx").remove_sheets("Secrets").delete_rows("Data!3:10")
linter.lint(wb)                             # the intermediate result can be inspected
wb.write("clean.xlsx")

# Editing, lazy: scan() reads nothing; operations are recorded, and run by sink() or collect()
plan = sheetrs.scan("book.xlsx").remove_sheets("Secrets").delete_rows("Data!3:10")
print(plan.explain())                       # the passes over the file
plan.sink("clean.xlsx")
```

What each name is:

- **`Linter`** is the linter with its configuration, built once and used on many workbooks. There
  is no `lint()` function next to it: one way to lint. `Linter.rules()` lists the rules with the
  ids and categories that `disabled_rules`/`enabled_rules` accept and whether *this* linter's
  configuration runs each one; rules are configured through the TOML, not through that list.
  **`Violation`** is the type of the items `lint` returns (rule, severity, sheet, cell, message);
  it is exposed so results can be annotated and checked, and is not built by the user.
- **`read` / `scan`** follow the eager/lazy convention of dataframe libraries (`read_csv` /
  `scan_csv`): the two modes have the **same methods with the same meaning**; they differ in what
  is loaded and when it runs. `read` gives a `Workbook` with a working copy of the file (the
  original is never modified) whose operations run when called; `scan` gives a `LazyWorkbook`
  where nothing is read, operations are recorded, and `sink(path)` or `collect()` (a `Workbook`)
  runs them. `Workbook.lazy()` goes from the first to the second.
- **Operations are sequential in both modes**: each addresses the workbook as the previous ones
  left it (after `delete_rows("T!3")`, `"T!3"` is the old row 4). The scopes of *one* call are
  simultaneous: `delete_rows("T!3", "T!4")` deletes two rows. The lazy mode plans the sequence
  into the fewest passes over the file that give what the eager one gives (a conversion after
  another change, and deletions that touch, need their own pass); each pass reads and writes the
  file once as a stream, so memory does not grow with the workbook.
- Eager errors are raised where the operation is called; lazy ones when the plan runs (`sink`,
  `collect`, `explain` for an operation on a sheet an earlier one removed).
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
