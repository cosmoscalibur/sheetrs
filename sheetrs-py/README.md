# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs

# Linter (sheetlint). The configuration is validated when the linter is built: an unknown
# rule, category or parameter raises ValueError instead of being ignored.
linter = sheetrs.linter.Linter(config_path="sheetlint.toml")
for v in linter.lint("book.xlsx"):          # or bytes, with extension="ods"
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)
linter.rules()      # [{"id", "name", "category", "default", "enabled"}, ...]
sheetrs.linter.lint("book.xlsx")            # one-off, default configuration

# Statistics (sheetstats): a dict, as `sheetstats --format json`
stats = sheetrs.stats.compute("book.xlsx")

# Editing (sheetedit), lazy: collect operations, apply them in ONE pass over the file on save
(sheetrs.edit.Plan("book.xlsx")
    .remove_sheets("Secrets")
    .convert_to_values("Report!B:D")
    .delete_rows("Data!3:10")
    .save("clean.xlsx"))

# Editing, eager: one operation now, one pass each
sheetrs.edit.remove_sheets("book.xlsx", "clean.xlsx", "Secrets")
```

There is one submodule per tool, as there is one Rust crate per tool.

- **Rules** belong to the linter. `Linter.rules()` is read-only: it lists the rules with the ids
  and categories that `disabled_rules`/`enabled_rules` of the configuration accept, and whether
  *this* linter's configuration runs each one. Rules are configured through the TOML, not through
  that list.
- **Lazy vs eager editing.** Both give the same result for the same operations; they differ in
  cost and in whose job the positions are.
  - A `Plan` is a plan: every operation is written against the *input* workbook, in any order,
    and the plan resolves their interactions (overlapping ranges merged; formulas converted
    before the cells they read are deleted; rows and columns deleted from the last to the first,
    across sheets, so no deletion shifts the positions of another; sheets removed last, so a
    deletion in a sheet that goes still shapes the references that span it). `plan.schedule`
    lists the order it applies. One read and one write of the file, and the same result as the
    schedule one operation per pass.
  - Eager functions each cost a pass and address the output of the previous one: to chain them,
    pass the output of one as the input of the next and translate the positions yourself (rows
    deleted by one shift the rows the next addresses).

Errors: `ValueError` (invalid configuration or scope), `OSError` (a file cannot
be read), `sheetrs.SheetrsError` (a workbook cannot be parsed or written).

## Development

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop          # builds and installs the extension in the venv
pytest
```
