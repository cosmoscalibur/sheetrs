### Added

- `sheetrs-formula` crate: lossless formula AST with a single-pass parser, a canonical printer, reference rewriting for inserted/deleted rows and columns and renamed/deleted sheets, and an R1C1 form. Verified by parsing and round-tripping every formula of the repository's workbooks.
- `sheetrs-engine` crate: format-independent IR, function library and demand-driven evaluator with JSON in and JSON out (`sheetrs-rt`, the runtime stub of compiled executables).
- `sheetrs-compiler` crate and `sheetc` binary: compiles XLSX/ODS workbooks into standalone executables with hybrid compile/run-time evaluation, input/output pruning, constant folding, iterative calculation, and dialects for Excel and LibreOffice behavior (`sheetrs-compiler/docs/FIDELITY.md`). It is the former `cosmoscalibur/xlslang` repository.
