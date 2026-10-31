### Added

- `sheetc build` warns when the document is in manual calculation mode (XLSX `calcMode="manual"`, ODS `AutoCalculate` false); the program calculates automatically and `INFO("recalc")` is `Automatic`, with no warning at run time.
