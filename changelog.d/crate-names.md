### Changed

- The crates of the workspace are named alike, `sheetrs-<role>`: `sheetrs` is now `sheetrs-core` (the library crate is `sheetrs_core`), and the tools are `sheetrs-linter`, `sheetrs-stats` and `sheetrs-edit`, in directories of the same names. Install with `cargo install --path sheetrs-linter` (and the others).
- The executable `sheetcli` is now `sheetedit`: it edits workbooks (remove sheets and names, convert formulas to values, delete rows and columns) and the old name did not say so. `sheetlint` and `sheetstats` keep their names.
