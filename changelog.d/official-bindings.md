### Added

- Add `sheetrs-py`, the official Python bindings (PyO3, abi3 wheels for Python ≥ 3.10), organized
  by tool: `sheetrs.linter` (`Linter` with a validated configuration, its `lint` and `rules`, and
  `lint`), `sheetrs.stats` (`compute`) and `sheetrs.edit`, which offers a lazy mode (`Plan`:
  operations are collected and applied in one pass on `save`) and an eager one (a function per
  operation). A workbook is a path or bytes; type stubs and a pytest suite included.
- Add `sheetrs_edit::EditPlan` and `Operation`: the operations on a workbook, collected and
  applied in one pass. `sheetedit` builds its run with it.

- Add `sheetrs_linter::report` (`group_by_scope`, `ViolationRecord`, `violation_records`): the
  findings grouped by book, sheet and cell and in a flat serializable form, shared by the
  executable and the bindings.
- Add `lint_workbook_records` and `workbook_stats` to `sheetrs-wasm`: structured results besides
  the text reports (`get_rules_definition` is now `lint_rules`). Add `scripts/package-wasm.sh`,
  which builds the npm package.
- Add a release workflow that attaches executables (Linux x86_64/aarch64, macOS arm64, Windows
  x86_64), Python wheels and sdist, and the npm tarball to each `vX.Y.Z` release; CI builds the
  Python extension and the npm package on every pull request.
- `rust-version` (1.99) is declared by every crate.

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

- A configuration naming a rule or category the linter does not know is rejected when it is
  parsed, in every interface (the executable checked it after parsing; the bindings did not). A
  `LinterConfig` is valid by construction: its fields are private to the crate and its
  `Deserialize` validates.
- The WASM text report printed the `Sheet:` header twice for a sheet with sheet-level and
  cell-level violations; both renderers now group with `group_by_scope` (#70).
- A cell error value that could not be parsed was reported as an "unknown rule ID".
