### Changed

- Performance documentation: `docs/performance.md` holds the current time and memory of every tool, measured on generated workbooks, with the contrast with LibreOffice 26.8.2 and the known limits; the Comparisons log of `scripts/README.md` (history of changes, which is the history of the commits) is gone. `scripts/bench.py` is the harness (`fixtures` generates the workbooks, `run` measures with dist-profile binaries); it replaces `scripts/benchmark.sh` and `scripts/benchmark_edit.sh`, which needed private production files and `/usr/bin/time`. `tests/production/` is no longer referenced anywhere (`.gitignore`, `.worktreeinclude`, `CLAUDE.md`, scripts and docs).
- `sheetrs-linter/docs/PROFILING.md`: the profile is taken on the generated `formulas.xlsx`.
- `sheetrs-compiler/scripts/measure_perf.py`: the stages are read again (the format of the trace had changed and every stage was reported as 0), the requests name outputs the workbooks declare, and the private workbook is gone. `gen_large_fixture.py` defines a name for the name-deletion benchmark. `measure_libreoffice_flow.py` also measures loading and searching in LibreOffice.

### Added

- `sheetrs-wasm/tests/module_run.mjs`: runs the WebAssembly module of a compiled workbook as one process, to time it against its native executable.
- `scripts/measure_libreoffice_flow.py` has the model flow (open once, set the input, calculate, read the outputs, 100 calls), and `scripts/bench.py` the `model` group (`sheetc calculate --request`, the compiled executable with `--request` and `--serve`, the Python `Book` and LibreOffice); `sheetrs-wasm/tests/call_run.mjs` gives the peak memory of one call of the WebAssembly package.
