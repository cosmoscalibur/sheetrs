### Changed

- The Python package is one object, `sheetrs.Book` (a reference to a workbook that reads nothing), replacing `open`/`Source`, `Linter`, `stats.collect` and `Editor`. Its operations (`delete_sheets`, `delete_names`, `formula_to_value`, `delete_rows`, `delete_columns`, `replace`, `calculate`) accumulate and `save` is the cut that plans and runs them; its queries are `range` (several ranges in one pass, as `Block`s), `find`, `calculate(outputs=...)`, `lint` (the path of the configuration file) and `stats`. A query with operations pending raises `PendingOperationsError`; an error cell is a `CellError`; dates are numbers unless `dates="datetime"`.
- `sheetrs.evaluate(formula)` and `sheetrs.lint_rules()`.
