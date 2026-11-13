### Changed

- One convention for the flags of every tool: the input workbook is `-i/--input FILE` (`--file` is accepted; the positional `FILE` of `sheetrs lint`, `sheetrs stats`, `sheetrs query`, `sheetrs edit` and `sheetc` is gone), the output is `-o/--output`, and the request of a compiled program and of `sheetc calculate` is `--request` (`--input` of the executables is gone; `--range` adds a range to `ranges`). `cinp` and `cout` are reserved for the declaration of which cells are inputs and outputs when a program is built.
- `sheetc calculate` is one form: `calculate -i FILE|PROGRAM.json [--request REQUEST|@FILE|-] [--range RANGE]… [-o OUT | --in-place] [--json]`. Without `-o` or `--in-place` it answers JSON; with one it writes the workbook, and with nothing declared or given every formula is calculated again. There is no `--recalculate` (it is `calculate -i FILE -o OUT`) and no `--embed-template` in `calculate` (it is only for `build`); both say so when used.
- The standard input is read only when it is asked for (`--request -`) or under `--serve`; a program run with nothing piped no longer waits for it to end.
- `sheetc verify --reference cache|libreoffice` says what the formulas are compared with (replaces `--recalculate`). The time it reports now includes the run of LibreOffice, which it left out.
- The WebAssembly `calculate_workbook` loses the `recalculate` option: `embed_template: true` with an empty request does the same.

### Added

- `sheetc calculate --in-place` and `sheetrs edit --in-place` replace the input through a temporary file and a rename, with its permissions, so a failure leaves the original as it was.

### Fixed

- A failed write no longer removes an output that was there before the run, nor a device (`/dev/full`, `/dev/null`); only a file the run created is removed. An archive that failed to finish no longer prints `ZipWriter::drop: failed to finalize archive`.
