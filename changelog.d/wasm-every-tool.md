### Added

- `sheetrs-wasm`: every tool in WebAssembly, in memory, from the bytes of a workbook: `lint_workbook_sarif`, `query_workbook`, `edit_workbook`, `explain_edit`, `calculate_workbook` (values, `embed_template`, `recalculate`), `inspect_workbook`, `verify_workbook` and `evaluate_formula`, besides the linter and the statistics. The Cargo features `query`, `edit` and `calculate` choose what the module contains. A smoke test (`sheetrs-wasm/tests/smoke.mjs`) runs every tool in Node on real workbooks and compares each answer with the command-line tool's.
- `sheetrs_compiler::memory`: `sheetc calculate`, `inspect`, `verify` and `evaluate` over bytes and JSON, with no file, process or thread; `Editor::write_to` in `sheetrs-edit` and `fill_template_to` in `sheetrs-answer`: the editor and the answer of a template written to any `io::Write`, as the bytes of a workbook; `sheetrs_edit::parse_scope`; `sheetrs_engine::clock::set_wall_clock`, the clock a host without an operating system provides.

### Changed

- The compiler lowers the formulas on the calling thread when there is one worker, and reads the clock for its stage timing only when `SHEETC_TRACE` is set (reading it panics on WebAssembly).
- `lint_rules`, `lint_workbook_records` and `workbook_stats` of the WASM package return plain objects, with `null` for a missing value, as the JSON of the command-line tools does.

### Fixed

- `cargo doc` with warnings as errors passes again (links left unresolved by the renames).
