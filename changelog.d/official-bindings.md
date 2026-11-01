### Added

- Add `sheetrs-py`, the official Python bindings (PyO3, abi3 wheels for Python ≥ 3.10), flat:
  `open(path or bytes)` makes a `Source`, a reference to a workbook that reads nothing and has no
  `save`; the tools take it: `Linter.lint(source)` (validated configuration, and `rules`),
  `Source.stats()`, and the operations on it, which make an `EditPlan` (the only object with
  `save`). There is no eager mode: a workbook is streamed, never loaded. Type stubs and a pytest
  suite included.
- Add `sheetrs_core::reader::{Source, Format}`: the reference to a workbook (a file or bytes) that
  the linter, the statistics (`WorkbookStats::from_source`) and the editor
  (`EditPlan::execute(&Source, output)`) take as input.
- Add `sheetrs_edit::EditPlan` and `Operation`: a sequence of operations on a workbook, each on the
  result of the previous one (rows and columns are typed `RowRange`/`ColumnRange`; the plan is a
  collection), planned by `passes()` into the fewest passes over the file that
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

- **Breaking:** `sheetedit` marks the output to be recalculated when it is opened whenever rows,
  columns, sheets or names are removed (their cached results are stale), as the last thing the
  plan does; `--no-recalc-on-open` avoids it and `--recalc-on-open` forces it. In Rust,
  `EditPlan::recalculation(Recalculation)`.
- **Breaking:** public names lose their `get_` prefix: `LinterConfig::param_int`, `param_str`,
  `param_array`, `param_float_array`, `param_bool`; `registry::rule_metadata`,
  `registry::valid_tokens`. Public types implement `Debug`; `RuleId`, `Operation`,
  `Recalculation` and `Format` are `#[non_exhaustive]`; `WalkerRule` requires `Debug`.

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

- `sheetedit` on an ODS converted to an empty string a formula cell that had lost its result and
  carried a shape or a note, because the paragraph of the shape counted as the value of the cell;
  now such a formula stays a formula, as in the XLSX.
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
