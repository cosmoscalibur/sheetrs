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
- **Lazy vs eager editing.** A `Plan` addresses the *input* workbook for every operation, however
  the operations were added, and costs one read and one write of the file. Eager functions each
  cost a pass; to chain them, pass the output of one as the input of the next, and each then sees
  the result of the previous ones (rows deleted by one shift the rows the next addresses). Prefer
  `Plan` unless an operation depends on the outcome of the one before it.

Errors: `ValueError` (invalid configuration or scope), `OSError` (a file cannot
be read), `sheetrs.SheetrsError` (a workbook cannot be parsed or written).

## Development

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop          # builds and installs the extension in the venv
pytest
```
