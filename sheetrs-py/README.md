# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs

# Linter (sheetlint). The configuration is validated when it is parsed: an unknown rule,
# category or parameter raises ValueError instead of being ignored.
linter = sheetrs.Linter(config_path="sheetlint.toml")
for v in linter.lint("book.xlsx"):          # or bytes, with extension="ods"
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)
linter.rules()      # [{"id", "name", "category", "default", "enabled"}, ...]
sheetrs.lint("book.xlsx")                   # one-off, default configuration

# Statistics (sheetstats): a dict, as `sheetstats --format json`
stats = sheetrs.stats("book.xlsx")

# Editing (sheetedit), lazy: collect operations, apply them in ONE pass over the file on save
(sheetrs.EditPlan("book.xlsx")
    .remove_sheets("Secrets")
    .convert_to_values("Report!B:D")
    .delete_rows("Data!3:10")
    .save("clean.xlsx"))

# Editing, eager: the same operations given at once and applied at once
sheetrs.edit("book.xlsx", "clean.xlsx", remove_sheets=["Secrets"], delete_rows=["Data!3:10"])
```

- **Rules** belong to the linter. `Linter.rules()` is read-only: it lists the rules with the ids
  and categories that `disabled_rules`/`enabled_rules` of the configuration accept, and whether
  *this* linter's configuration runs each one. Rules are configured through the TOML, not through
  that list.
- **Editing is planned, in both modes.** Every operation is written against the *input* workbook,
  in any order, and is resolved into the order that gives the right result: formulas converted
  before the cells they read are deleted, overlapping ranges merged, rows and columns deleted from
  the last to the first, sheets removed after. `plan.operations` lists that order (the same for
  any order the operations were added in), and the file is read and written once, as a stream:
  memory does not grow with the size of the workbook. `EditPlan` is lazy (nothing happens until
  `save`, so a plan can be built up, inspected and saved more than once); `edit` is the same
  with everything given at once. To make a second edit depend on the result of the first, call
  them one after the other and write the second against the first's output.
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
