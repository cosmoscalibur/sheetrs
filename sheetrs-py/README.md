# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs

# Linter (sheetlint): violations ordered by book, sheet, cell and rule
for v in sheetrs.linter.lint("book.xlsx", config_path="sheetlint.toml"):
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)
sheetrs.linter.rules()   # the linter's rules

# A workbook can also be given as bytes (the extension tells the format)
sheetrs.linter.lint(data, extension="ods")

# Statistics (sheetstats): a dict, as `sheetstats --format json`
stats = sheetrs.stats.compute("book.xlsx")

# Editing (sheetedit): write a modified copy
sheetrs.edit.modify("book.xlsx", "clean.xlsx",
                    remove_sheets=["Secrets"], convert_to_values=["Report!B:D"])
```

There is one submodule per tool, as there is one Rust crate per tool; the rules belong to the
linter, so they are listed there.

Errors: `ValueError` (invalid configuration or scope), `OSError` (a file cannot
be read), `sheetrs.SheetrsError` (a workbook cannot be parsed or written).

## Development

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop          # builds and installs the extension in the venv
pytest
```
