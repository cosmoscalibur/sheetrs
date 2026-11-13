### Changed

- A `Book` of the Python and WebAssembly packages reads nothing when it is made: the statistics (and the shortcut properties derived from them: `sheet_names`, `sheets`, `defined_names`, `date_system`, `has_macros`, `stats` and the `total_*`) are read once, the first time one of them is asked. A workbook of an unsupported date system or one that cannot be read raises there, and in any query, and no longer when the `Book` is made. In WebAssembly those getters throw instead of returning a value that was read at construction.

### Fixed

- Remove the mentions of the flags `--recalculate` and `--recalculate-on-open`, which do not exist, from the documentation and the comments of `sheetrs-rt-wasm`.
