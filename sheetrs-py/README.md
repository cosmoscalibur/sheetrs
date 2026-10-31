# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: lint, inspect and
edit XLSX and ODS workbooks with the same engine as the `sheetlint`,
`sheetstats` and `sheetedit` executables. Wheels are `abi3` (Python ≥ 3.10).

```python
import sheetrs

# sheetlint: violations ordered by book, sheet, cell and rule
for v in sheetrs.lint("book.xlsx", config_path="sheetlint.toml"):
    print(v.rule_id, v.severity, v.sheet, v.cell, v.message)

# A workbook can also be given as bytes (the extension tells the format)
sheetrs.lint(data, extension="ods")

# sheetstats: a dict, as `sheetstats --format json`
stats = sheetrs.stats("book.xlsx")

# sheetedit: write a modified copy
sheetrs.edit("book.xlsx", "clean.xlsx",
             remove_sheets=["Secrets"], convert_to_values=["Report!B:D"])

sheetrs.rules()  # the catalog of rules
```

Errors: `ValueError` (invalid configuration or scope), `OSError` (a file cannot
be read), `sheetrs.SheetrsError` (a workbook cannot be parsed or written).

## Development

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop          # builds and installs the extension in the venv
pytest
```
