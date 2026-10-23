### Added

- `sheetrs-formula` crate: lossless formula AST with a single-pass parser, a canonical printer, reference rewriting for inserted/deleted rows and columns and renamed/deleted sheets, and an R1C1 form. Verified by parsing and round-tripping every formula of the repository's workbooks.
- `xlslang` crate: compiler from XLSX/ODS to standalone executables (JSON in, JSON out) with a format-independent IR, hybrid compile/run-time evaluation, input/output pruning, constant folding, iterative calculation, and dialects for Excel and LibreOffice behavior (`xlslang/docs/FIDELITY.md`). It is the former `cosmoscalibur/xlslang` repository.
