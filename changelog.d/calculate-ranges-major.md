### Changed

- **Breaking:** the request of `calculate` is `inputs`, `ranges` and `major` in every layer (as the JSON request `{"inputs", "ranges", "majorDimension"}`): `Calculation::outputs` is `ranges` and has a `major` (`MajorDimension`); Python's `calculate(inputs=None, ranges=None, *, major="rows")` replaces `calculate(inputs, outputs)`, and `major="columns"` lists each range as columns. The layout of the answer is done once, by the planner (`Outputs::with_major` in `sheetrs-protocol`), so that Rust, Python and WebAssembly give the same lists.

### Also in this change

- The engine answers with the types of `sheetrs-protocol`: `Program::handle*` return `Outputs` (not a JSON value that the caller parsed again), `fill_template_*_with_answer` return `Outputs` and `Calculated::answer` is an `Option<Outputs>`; `Outputs::to_json()` is the JSON answer, byte for byte the one `sheetc` printed.
- **Breaking:** the module `sheetrs_formula::r1c1` (no consumer) is gone: only the A1 notation is supported, recorded in `docs/approved-differences.md`. `sheetrs_compiler::graph` is private.
- `#[non_exhaustive]` on `FromCalcError`, `ExcelError` and `EntryFormat`.
- A manual page break that a deletion leaves before the first row or column is no longer written (XLSX and ODS): LibreOffice reads none from a file, and the structure comparison with it no longer ignores it.
