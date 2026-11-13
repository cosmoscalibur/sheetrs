# sheetrs-engine

The calculation engine of the SheetRS suite: the intermediate representation of a workbook program, the
function library that reproduces `LibreOffice` Calc 26.8, and the run-time evaluation with JSON in and
out. It is the engine of `sheetc` (`sheetrs-compiler`), of the `calculate` step of `sheetrs-edit` and of
the compiled executables and WebAssembly modules (`sheetrs-rt-wasm`).

The clock of `NOW()` and `TODAY()` is the system clock; a host without one (WebAssembly) installs it
once with `clock::set_wall_clock`, and a run can fix `now` and `seed` in its options.

The API is experimental: names and shapes may change.
