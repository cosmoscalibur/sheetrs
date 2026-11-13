### Changed

- The `Book` of the Python and WebAssembly packages reads nothing when it is made and keeps nothing of the workbook, and the properties that were derived from its statistics are gone: `sheet_names`, `sheets`, `defined_names`, `date_system`, `has_macros` and the eleven `total_*` in Python, and `sheetNames`, `sheets`, `definedNames`, `dateSystem`, `hasMacros` and the `total*` in WebAssembly. What a `Book` says about the workbook is `stats()`, a method that reads it each time it is called and gives a `Stats` (the totals, `sheets`, `has_macros` and `defined_names`; in WebAssembly `toJSON()`, `hasMacros` and `definedNames`). The date system is in `inspect()`, and a workbook of an unsupported one raises in `stats()` and in every query, not when the `Book` is made. `Book.stats` is a method now, not a property.
- The demo of `docs/` reads the statistics with `stats()`.

### Fixed

- Remove the mentions of the flags `--recalculate` and `--recalculate-on-open`, which do not exist, from the documentation and the comments of `sheetrs-rt-wasm`.
