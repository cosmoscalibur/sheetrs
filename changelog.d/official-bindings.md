### Added

- Add `sheetrs-py`, the official Python bindings (PyO3, abi3 wheels for Python ≥ 3.10), flat:
  `Linter` (validated configuration, its `lint` and `rules`), `stats`, and `open` for editing:
  `open(path)`, operations recorded on the `Workbook` that results, and `save(path)`, which runs
  them. There is no eager mode: a workbook is streamed, never loaded. A workbook is a path or
  bytes; type stubs and a pytest suite included.
- Add `sheetrs_edit::EditPlan` and `Operation`: a sequence of operations on a workbook, each on the
  result of the previous one (rows and columns are typed `RowRange`/`ColumnRange`; the plan is a
  collection; the mark to recalculate when opened is a property of the output, applied once), planned by `passes()` into the fewest passes over the file that
  give what running them one at a time gives (`explain()` shows them). `EditPlan::against_input`
  is the declarative case `sheetedit` uses for its flags (one pass, in no order); `--dry-run`
  prints the passes. Tests run generated sequences both ways in XLSX and ODS.
- Add `lint_workbook_records` and `workbook_stats` to `sheetrs-wasm`: structured results besides
  the text reports (`get_rules_definition` is now `lint_rules`). Add `scripts/package-wasm.sh`,
  which builds the npm package.
- Add a release workflow that attaches executables (Linux x86_64/aarch64, macOS arm64, Windows
  x86_64), Python wheels and sdist, and the npm tarball to each `vX.Y.Z` release; CI builds the
  Python extension and the npm package on every pull request.
- `rust-version` (1.99) is declared by every crate.

### Removed

- `sheetedit` no longer adjusts references across sheets (`Sheet1:Sheet3!A1`): a command that
  deletes rows or columns or removes sheets in a workbook with a formula that holds one stops with
  an error that names issue #75, instead of rewriting it. It is not implemented; the code that did
  it, in `sheetrs-formula` (`Workbook`, `ThreeD`, `apply_three_d`) and in both writers, is gone.
  `Formula::has_three_d` tells whether a formula holds one.

### Changed

- **Breaking:** `sheetrs-core` is the shared kernel, the workbook model and the readers. The rules,
  `LinterConfig`, `Linter` and the violation types moved to `sheetrs-linter`
  (`sheetrs_linter::{rules, config, violation, Linter}`), and the `writer` to `sheetrs-edit`
  (`sheetrs_edit::writer`). `ExcelError` is `sheetrs_core::reader::ExcelError`. The `link-validation`
  feature is now a feature of `sheetrs-linter`.
- `sheetrs-linter` and `sheetrs-edit` are libraries as well as executables, and `sheetrs-stats`
  too (`stats::WorkbookStats`, `StatsSink`).
- Path dependencies between the publishable crates carry a version, and the crates have keywords and
  categories, so they can be published to crates.io. `sheetrs-wasm` and `sheetrs-py` are not
  published.
- All crates are at version 0.4.0.

### Fixed

- `sheetedit` gave different workbooks for the same command: the formula edits of each sheet were
  applied in hash order, which changes the result when deletions in several sheets meet a
  reference across sheets. They are applied in one canonical order, and references across sheets
  are no longer rewritten (see Removed).
- A worksheet that holds a reference across sheets written with a quoted name that has a space
  (`'Sheet 1:Sheet 3'!A1`) was copied without being read; it is now read, and refused.
- A configuration naming a rule or category the linter does not know is rejected when it is
  parsed, in every interface (the executable checked it after parsing; the bindings did not). A
  `LinterConfig` is valid by construction: its fields are private to the crate and its
  `Deserialize` validates.
- The WASM text report printed the `Sheet:` header twice for a sheet with sheet-level and
  cell-level violations; both renderers now group with `group_by_scope` (#70).
- A cell error value that could not be parsed was reported as an "unknown rule ID".
