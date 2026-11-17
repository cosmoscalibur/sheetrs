# Performance

Time and memory of every tool of SheetRS at the state of the repository, on workbooks that any checkout can
generate, and the same flows in LibreOffice 26.8.2 as the reference. The document holds the current numbers only:
what a change did to them is in the history of the commit that made it. A change that moves a number updates
the table of its flow here, and `sheetrs-linter/docs/PROFILING.md` when it targets the cost of a rule or a phase.

- [Method](#method)
- [Fixtures](#fixtures)
- [Results](#results): [lint](#sheetrs-lint), [stats](#sheetrs-stats), [query](#sheetrs-query),
  [edit](#sheetrs-edit), [the cache and calculate_all](#the-cache-and-calculate_all),
  [calculate and the compiler](#sheetc-calculate-and-sheetc-build),
  [Python](#python), [WebAssembly](#webassembly), [sizes](#sizes)
- [Against LibreOffice](#against-libreoffice)
- [Known limits](#known-limits)
- [Pending: third-party libraries](#pending-benchmark-against-third-party-libraries)

## Method

Everything below comes from runs of `scripts/bench.py` on one machine, nothing else running. A table is updated by the groups and rows that a change moves
(`run DIR GROUP --scenarios ROWS`, and `--bin DIR` for another build); a value in parentheses is the one of the build before, measured in the same session with the same rows.

| | |
|---|---|
| Machine | Linux 6.18 virtual machine, Intel Xeon @ 2.10 GHz, 4 cores, 16 GB, no hardware performance counters |
| Binaries | `cargo build --profile dist -p sheetrs-cli -p sheetrs-compiler` (thin LTO, one codegen unit, stripped); the wasm package by `scripts/package-wasm.sh` (`release`); the Python wheel by `maturin build --profile dist` |
| Runs | Every command runs 3 times; a cell gives the minimum and the maximum |
| Wall time | `time.perf_counter()` around the process |
| Peak memory | `ru_maxrss` of the child (`wait4`); for LibreOffice, `VmHWM` of `soffice.bin` |
| Reference | LibreOffice 26.8.2 headless, one process and one profile per run, driven through UNO (`scripts/measure_libreoffice_flow.py`) |

Every figure of peak memory includes the floor of the harness: a child starts with the resident size of the Python
process that launched it, so a process that only prints its version peaks at 16-17 MB (`sheetrs --version`,
`sheetc --version`; `python -c "import sheetrs"` 17 MB; `node -e ""` 42 MB). Times of a shared machine move 10-20 %
between sessions: compare the rows of one table, not tables of different runs.

To measure again:

```bash
cargo build --profile dist -p sheetrs-cli -p sheetrs-compiler
python3 scripts/bench.py fixtures DIR                  # about 4 minutes; needs `xlsxwriter`, `openpyxl` and LibreOffice
python3 scripts/bench.py run DIR                       # every group; or `run DIR lint,edit` for some (`cache`: `edit` by default against `--calculate-all`; `modes`: `sheetc build` in its three modes)
python3 scripts/bench.py run DIR python --python VENV/bin/python   # a venv with `maturin develop --profile dist`
scripts/package-wasm.sh && scripts/build-wasm-rt.sh && python3 scripts/bench.py run DIR wasm
```

The run is about an hour and a half (the ODS column deletions alone take two minutes each). The instruction counts that the
earlier scripts took from `perf stat` are not available in a virtual machine; for the profile of a tool,
`perf record -e cpu-clock` samples (see [PROFILING.md](../sheetrs-linter/docs/PROFILING.md)).

## Fixtures

Nothing private is needed: the benchmarks use generated workbooks and two files of the repository. `scripts/bench.py
fixtures DIR` writes them into `DIR`, outside git; the generators are deterministic in content.

| File | Made by | What it is |
|---|---|---|
| `formulas.xlsx` | `sheetrs-compiler/scripts/gen_large_fixture.py DIR/formulas.xlsx 300000` | 5 sheets, 3.6 M used cells, **2.4 M formulas**, 29 MB, 210 MB of sheet XML: `Calc` has eight formula columns per row (arithmetic, `IF`, exact `VLOOKUP`, running total, text assembly, `IFERROR`, `MAX`/`MIN`), `Data` the inputs, `Lookup` a 2000-row table, `Summary` five aggregates, `settings` the inputs and outputs, and one defined name (`base_rate`). The file stores `0` as the cached result of every formula (xlsxwriter), so a command that needs results calculates them first |
| `formulas.ods` | `soffice --convert-to ods` of the file above | The same workbook in ODS, 25 MB, with the results LibreOffice computed |
| `formulas_cached.xlsx` | `soffice --convert-to xlsx` of `formulas.ods` | The same workbook in XLSX, 50 MB, with the true results cached (and no `fullCalcOnLoad`): a plan on it can find what a deletion leaves stale, which it cannot on `formulas.xlsx` |
| `data.xlsx`, `data_declared.xlsx` | `sheetrs-compiler/scripts/gen_data_fixture.py` | 600 000 rows of numbers, a label and a row `SUM` (49 MB, 600 005 formulas, no cached values); the second adds a `settings` sheet that declares one input and its outputs |
| `small.xlsx` | `sheetrs-compiler/tests/assets/loan.xlsx` | A loan schedule: a few formulas, three declared outputs |

`formulas.xlsx` has the shape of a financial model of about two million formulas; it replaces the private production
workbooks that earlier measurements used, which nobody else could obtain.

## Results

Wall time and peak memory, min-max of 3 runs, on `formulas.xlsx` / `formulas.ods` unless said.

### `sheetrs lint`

With `scripts/benchmark.toml`: every default-active rule, thresholds fixed so that every run does the same rule work. The cost by phase and by rule
is in [PROFILING.md](../sheetrs-linter/docs/PROFILING.md).

| Workbook | Wall time and peak memory |
|---|---|
| `formulas.xlsx` | 14.14-14.78 s (15.07 s), 661-664 MB (660 MB) |
| `formulas.ods` | 18.65-19.66 s (18.94 s), 697-699 MB (701 MB) |

### `sheetrs stats`

The counts are taken while the reader goes: the workbook is never built, and the formulas of an ODS are not converted to Excel's notation (the counts do not read them).

| Workbook | Wall time and peak memory |
|---|---|
| `formulas.xlsx` | 1.31-1.32 s (1.80 s), 17-18 MB (20 MB) |
| `formulas.ods` | 2.38-2.48 s (3.07 s), 18 MB (20 MB) |

### `sheetrs query`

The readers skip the sheets that are not asked for and stop after the last row a window reaches; an ODS has to read its content part to know where the sheets are.

| Query | `formulas.xlsx` | `formulas.ods` |
|---|---|---|
| `--list-sheets` | 0.026-0.030 s (0.032 s), 18 MB (20 MB) | 0.76-0.78 s (1.33 s), 20 MB (20 MB) |
| `--range Calc!A1:H1000` | 0.032-0.036 s (0.041 s), 18 MB (20 MB) | 0.28-0.31 s (0.49 s), 20 MB (20 MB) |
| `--sheet Summary` | 0.025-0.032 s (0.030 s), 18 MB (20 MB) | 0.77-0.79 s (1.34 s), 20 MB (20 MB) |
| `--find K00` (values, every sheet; 14 930 cells) | 1.72-1.84 s (2.36 s), 18-19 MB (20 MB) | 2.85-2.95 s (3.61 s), 20 MB (20 MB) |
| `--find VLOOKUP --in formula` (every formula) | 5.30-5.55 s (5.90 s), 104 MB (103 MB) | 4.95-5.20 s (6.28 s), 110 MB (108 MB) |
| `--find K00 --include-formulas` (values and the results of 2.4 M formulas) | 15.36 s (15.51 s), 623 MB (622 MB) | 17.17 s (19.98 s), 648 MB (647 MB) |

### `sheetrs edit`

The calculation mode decides what the engine does in a plan that leaves results stale (rows, columns, sheets or names deleted, text replaced) or that converts
formulas to values (see the README of `sheetrs-edit`, *The cache is the default*): by default, as `LibreOffice` does, the engine calculates the formulas the plan
reaches and what depends on them and every other formula keeps the result the file caches; with `--calculate-all` it calculates every formula. The workbooks
are the referent size of every table here, **300 000 rows** of `Calc` (`scripts/bench.py fixtures DIR`: 2.4 M formulas, 3.6 M cells).
`formulas.xlsx` stores `0` as every cached result and says so (`fullCalcOnLoad`), so a run on it calculates as `--calculate-all` does; the comparisons of the
modes use `formulas_cached.xlsx` (the same workbook with the results LibreOffice computed) and `formulas.ods` (which stores them). An ODS starts every plan
with the calculation that `LibreOffice` does when it opens it (`sheetrs-compiler/docs/FIDELITY.md`, *What LibreOffice calculates when it opens a file*).

`scripts/bench.py run DIR edit`, default scenario:

| Scenario (flags) | `formulas.xlsx` | `formulas.ods` |
|---|---|---|
| `--delete-sheets Lookup` | 14.61-15.43 s (16.11 s), 1819-1821 MB (1819 MB) | 23.19-23.46 s (24.21 s), 1901 MB (1900 MB) |
| `--delete-names base_rate` | 11.14-11.33 s (12.47 s), 1270 MB (1269 MB) | 20.91-21.52 s (22.58 s), 1291 MB (1289 MB) |
| `--formula-to-value Calc` | 11.93-12.11 s (12.81 s), 1590 MB (1589 MB) | 20.75-21.11 s (22.32 s), 1633 MB (1632 MB) |
| `--formula-to-value Calc!B:D Calc!1:1000` | 11.33-11.83 s (12.54 s), 1613 MB (1612 MB) | 20.74-21.14 s (22.37 s), 1658-1665 MB (1657 MB) |
| delete sheet and name, convert `Calc` (one command) | 14.34-14.69 s (15.71 s), 1590 MB (1589 MB) | 26.06-26.49 s (28.18 s), 1633-1634 MB (1632 MB) |
| `--delete-rows Data!10:20` | 16.79-16.97 s (17.82 s), 1409 MB (1408 MB) | 28.70-28.96 s (29.62 s), 1416 MB (1415 MB) |
| `--delete-rows Data!10:20 --formula-to-value Calc` | 14.82-15.33 s (16.76 s), 1586-1587 MB (1585 MB) | 29.86-30.25 s (32.04 s), 1633 MB (1632 MB) |
| `--delete-rows Calc!3:10 --formula-to-value Summary` | 26.47-26.54 s (27.35 s), 1625-1664 MB (1623 MB) | 46.47-47.60 s (47.80 s), 1672 MB (1671 MB) |
| `--delete-columns Data!B:D` | 16.04-16.25 s (17.20 s), 2170 MB (2169 MB) | 25.77-26.18 s (26.72 s), 2186 MB (2185 MB) |
| rows and columns together | 16.18-17.51 s (17.17 s), 2170 MB (2169 MB) | 26.10-26.67 s (28.28 s), 2186 MB (2185 MB) |
| sheet, name, rows and columns (one command) | 16.86-17.22 s (17.32 s), 2164-2168 MB (2166 MB) | 24.62-24.93 s (25.59 s), 2186 MB (2185 MB) |
| `--hide-columns Data!B` | 1.36-1.48 s (1.46 s), 18 MB (18 MB) | 22.37-23.58 s (23.63 s), 1287 MB (1286 MB) |
| `--hide-rows Data!10:20` | 11.86-12.37 s (12.60 s), 1271-1272 MB (1270 MB) | 27.24-27.65 s (28.89 s), 1300-1309 MB (1304 MB) |

A conversion of formulas to values in its two scenarios (`scripts/bench.py run DIR edit --workbooks formulas_cached.xlsx,formulas.ods --scenarios ...`): by default it converts the
results the file caches for the formulas the plan does not reach; with `--calculate-all` the engine calculates what it converts.

| Scenario (flags) | Default (the cache) | `--calculate-all` |
|---|---|---|
| XLSX `formulas_cached.xlsx`, `--formula-to-value Calc` | 5.50 s (5.29 s), 17 MB (18 MB) | 14.32 s (14.71 s), 1607 MB (1606 MB) |
| XLSX `formulas_cached.xlsx`, `--formula-to-value Calc!B:D Calc!1:1000` | 4.79 s (4.94 s), 18 MB (18 MB) | 10.86 s (10.97 s), 1173 MB (1171 MB) |
| XLSX `formulas_cached.xlsx`, delete sheet and name, convert `Calc` | 8.40 s (8.66 s), 299 MB (297 MB) | 16.84 s (18.14 s), 1607 MB (1605 MB) |
| XLSX `formulas_cached.xlsx`, `--delete-rows Data!10:20 --formula-to-value Calc` | 9.13 s (9.41 s), 265 MB (264 MB) | 17.94 s (19.16 s), 1600 MB (1599 MB) |
| ODS `formulas.ods`, `--formula-to-value Calc` | 20.75-21.11 s (22.32 s), 1633 MB (1632 MB) | 19.78-19.98 s (20.70 s), 1627 MB (1628 MB) |
| ODS `formulas.ods`, `--formula-to-value Calc!B:D Calc!1:1000` | 20.74-21.14 s (22.37 s), 1658-1665 MB (1657 MB) | 20.34-20.40 s (21.11 s), 1651-1653 MB (1651 MB) |
| ODS `formulas.ods`, delete sheet and name, convert `Calc` | 26.06-26.49 s (28.18 s), 1633-1634 MB (1632 MB) | 25.29-25.51 s (26.72 s), 1627 MB (1626 MB) |
| ODS `formulas.ods`, `--delete-rows Data!10:20 --formula-to-value Calc` | 29.86-30.25 s (32.04 s), 1633 MB (1632 MB) | 28.41-28.92 s (30.03 s), 1627 MB (1626 MB) |

The cache against `--calculate-all` for the plans that leave results stale (`scripts/bench.py run DIR cache`); the last column is the share of the 2.4 M formulas that the recalculation reaches.

| Workbook and scenario | Default (the cache) | `--calculate-all` | Formulas calculated (cone) |
|---|---|---|---|
| XLSX `--delete-names base_rate` (nothing reads it) | 6.12 s (6.70 s), 1236 MB (1234 MB) | 13.86 s (13.89 s), 1272 MB (1270 MB) | 0 |
| XLSX `--delete-rows Data!299990:299995` | 16.11 s (17.47 s), 1257 MB (1255 MB) | 19.59 s (20.05 s), 1273 MB (1272 MB) | 59 |
| XLSX `--delete-rows Data!10:20` (the running total reads on) | 17.93 s (17.76 s), 1276 MB (1275 MB) | 20.71 s (20.41 s), 1406 MB (1405 MB) | 300 074 (12.5 %) |
| XLSX `--delete-sheets Lookup` (a sheet goes: `LibreOffice` calculates every formula) | 18.26 s (18.67 s), 1821 MB (1820 MB) | 18.47 s (18.53 s), 1821 MB (1820 MB) | all |
| XLSX `--delete-columns Data!B:D` (every formula reads them) | 20.60 s (20.72 s), 2168 MB (2167 MB) | 19.38 s (19.11 s), 2170 MB (2169 MB) | 37.5 % that read the columns and what depends on them |
| ODS `--delete-names base_rate` | 21.48 s (21.49 s), 1291 MB (1289 MB) | 19.75 s (19.93 s), 1280 MB (1278 MB) | all (every formula is a root of the startup) |
| ODS `--delete-rows Data!299990:299995` | 27.37 s (27.07 s), 1286 MB (1281 MB) | 25.29 s (24.88 s), 1279 MB (1278 MB) | all (every formula is a root of the startup) |
| ODS `--delete-rows Data!10:20` | 28.99 s (27.66 s), 1416 MB (1415 MB) | 27.38 s (26.58 s), 1417 MB (1416 MB) | all (every formula is a root of the startup) |
| ODS `--delete-sheets Lookup` | 23.38 s (22.23 s), 1902 MB (1901 MB) | 23.57 s (23.16 s), 1901 MB (1900 MB) | all (every formula is a root of the startup) |
| ODS `--delete-columns Data!B:D` | 25.62 s (25.44 s), 2186 MB (2185 MB) | 24.12 s (23.33 s), 2186 MB (2186 MB) | all (every formula is a root of the startup) |

### `sheetc calculate` and `sheetc build`

`sheetc calculate` has the calculation mode of the run (see the README of `sheetrs-compiler`, *The calculation mode*): by default what the inputs it is given reach is
calculated and the rest of the workbook is the result it caches, and `--calculate-all` calculates the whole workbook. A calculation is built and run in one go, so the
cells that can change are the inputs the request gives, not every input the workbook declares (a build is run later with any of them and keeps the cone of all). The workbook
written is the answer: it carries what was calculated and, for the rest, what the template caches. A request read from the standard input (`--request -`) or from a file is not known
when the program is built, so the cone of every declared input is built then. `formulas.xlsx` stores `0` as every cached result and says so with `fullCalcOnLoad`, so the
comparisons use `formulas_cached.xlsx`, which declares the input `rate` and its outputs (`scripts/bench.py run DIR calculate`).

| Flow | Default (the cache) | `--calculate-all` |
|---|---|---|
| `sheetc calculate -i formulas_cached.xlsx -o OUT` (the workbook written, nothing given) | 5.02-5.47 s (6.73 s), 1233 MB (1231 MB) | 17.53-17.89 s (18.88 s), 1432-1439 MB (1432 MB) |
| `sheetc calculate -i formulas_cached.xlsx --range Summary!A1` (JSON) | 4.90-5.07 s (6.08 s), 1233 MB (1232 MB) | 10.29-10.66 s (11.54 s), 1432-1433 MB (1432 MB) |
| `sheetc calculate -i formulas_cached.xlsx --request '{"inputs": {"rate": 0.07}}' -o OUT` (the input given: its cone is calculated and written) | 14.11-14.65 s (15.71 s), 1227 MB (1247 MB) | 21.25-22.15 s (22.59 s), 1432 MB (1431 MB) |

**The artifact** (`scripts/bench.py run DIR modes`, on `formulas_cached.xlsx`): a program built by default carries the cached result of the formulas it keeps; built with
`--force-calculate-all` it calculates all of them in every run and is lighter. A program built by default and one built with `--calculate-all` give the same answers.

| Flow | Wall time and peak memory | Artifact |
|---|---|---|
| build default (cache) | 8.23-8.63 s (9.74 s), 1281 MB (1287 MB) | 108.18 MB (108.18 MB) |
| run of the default (cache) artifact (the input changed) | 1.49-1.59 s (1.56 s), 519-520 MB (519 MB) | |
| run of the default (cache) artifact (nothing given) | 1.42-1.54 s (1.62 s), 519 MB (519 MB) | |
| run of the default (cache) artifact with `--calculate-all` (the input changed) | 1.23-1.29 s (1.33 s), 519 MB (519 MB) | |
| run of the default (cache) artifact with `--calculate-all` (nothing given) | 1.25-1.33 s (1.35 s), 519 MB (519 MB) | |
| build `--calculate-all` | 13.23-13.50 s (14.84 s), 1432-1439 MB (1432 MB) | 108.18 MB (108.18 MB) |
| run of the `--calculate-all` artifact (the input changed) | 1.45-1.63 s (1.50 s), 519-520 MB (519 MB) | |
| run of the `--calculate-all` artifact (nothing given) | 1.45-1.50 s (1.66 s), 519 MB (519 MB) | |
| run of the `--calculate-all` artifact with `--calculate-all` (the input changed) | 1.20-1.26 s (1.41 s), 519 MB (519 MB) | |
| run of the `--calculate-all` artifact with `--calculate-all` (nothing given) | 1.26-1.29 s (1.30 s), 519 MB (519 MB) | |
| build `--force-calculate-all` | 8.00-8.28 s (9.30 s), 1235-1252 MB (1239 MB) | 89.57 MB (89.57 MB) |
| run of the `--force-calculate-all` artifact (the input changed) | 1.20-1.21 s (1.23 s), 473-475 MB (474 MB) | |
| run of the `--force-calculate-all` artifact (nothing given) | 1.15-1.19 s (1.26 s), 473-474 MB (474 MB) | |
| run of the `--force-calculate-all` artifact with `--calculate-all` (the input changed) | 1.18-1.19 s (1.27 s), 473-474 MB (474 MB) | |
| run of the `--force-calculate-all` artifact with `--calculate-all` (nothing given) | 1.12-1.22 s (1.25 s), 473-474 MB (474 MB) | |

**The embedded build** (`scripts/bench.py run DIR embedded`): `sheetc build --embed-template` makes an executable that carries the workbook, so it can write a workbook as its answer (`-o OUT`).

| Flow | Wall time and peak memory |
|---|---|
| build `--embed-template` | 8.39-8.94 s (9.60 s), 1281 MB (1287 MB) |
| run of the embedded executable (the input changed): JSON | 1.78-1.85 s (1.83 s), 616-617 MB (616 MB) |
| run of the embedded executable (the input changed): workbook written (`-o OUT`) | 2.57-2.68 s (2.60 s), 616 MB (617 MB) |

`sheetc build` and the executable it writes, by stage (`scripts/bench.py run DIR compiler`: `sheetrs-compiler/scripts/measure_perf.py`; the stages come from `SHEETC_TRACE`; median of 3). The
stage rows are parts of the build total of their workbook (the peak memory is that of the process; the stages do not add up to the total, which also holds the start and the end of the process).
A workbook that declares its inputs and outputs keeps only what the outputs need, so its executable is small; one that does not carries every cell. The payload of the executables is `postcard` (see
[the executable format](../sheetrs-compiler/docs/DESIGN.md#executable-format)).

| Workbook | Stage | Wall time, peak memory |
|---|---|---|
| `small.xlsx` | build total (executable 5.07 MB) | 0.01 s (0.04 s), 14 MB (14 MB) |
| | read workbook / lower formulas / dependency graph / fold (engine at build) / assemble / write executable / other (compile remainder, exit) | 0.00 / 0.00 / 0.00 / 0.00 / 0.00 / 0.01 / 0.00 (0.00 / 0.00 / 0.00 / 0.00 / 0.00 / 0.01 / 0.02) |
| | engine start-up (`--meta`) / execute a request | 0.003 / 0.003 s (0.004 / 0.004 s), 14 / 14 MB (14 / 14 MB) |
| `data.xlsx` (no settings: inputs deduced) | build total (executable 118.39 MB) | 10.76 s (12.07 s), 1563 MB (1566 MB) |
| | read workbook / lower formulas / dependency graph / fold (engine at build) / assemble / write executable / other (compile remainder, exit) | 2.51 / 0.34 / 2.81 / 0.08 / 0.04 / 0.49 / 4.49 (3.57 / 0.38 / 2.87 / 0.08 / 0.04 / 0.56 / 4.57) |
| | engine start-up (`--meta`) / execute a request | 0.628 / 1.105 s (0.644 / 1.160 s), 478 / 591 MB (478 / 591 MB) |
| `data_declared.xlsx` (inputs and outputs declared) | build total (executable 5.03 MB) | 6.78 s (8.10 s), 1474 MB (1473 MB) |
| | read workbook / lower formulas / dependency graph / fold (engine at build) / assemble / write executable / other (compile remainder, exit) | 2.50 / 0.66 / 0.09 / 0.00 / 0.00 / 0.01 / 3.52 (3.47 / 0.70 / 0.09 / 0.00 / 0.00 / 0.02 / 3.82) |
| | engine start-up (`--meta`) / execute a request | 0.002 / 0.002 s (0.003 / 0.003 s), 14 / 14 MB (14 / 14 MB) |
| `formulas.xlsx` | build total (executable 108.18 MB) | 12.88 s (13.87 s), 1424 MB (1423 MB) |
| | read workbook / lower formulas / dependency graph / fold (engine at build) / assemble / write executable / other (compile remainder, exit) | 1.52 / 2.98 / 0.60 / 0.01 / 1.41 / 0.67 / 5.68 (2.20 / 2.96 / 0.60 / 0.01 / 1.29 / 0.68 / 6.14) |
| | engine start-up (`--meta`) / execute a request | 0.748 / 1.486 s (0.656 / 1.575 s), 486 / 519 MB (486 / 519 MB) |

### Python

The `Book` of the wheel (`maturin develop --profile dist`) on `formulas.xlsx`; every step runs in its own process, which
opens the `Book` first, so the peak of a step is that of its process:

| Step | Wall time of the step | Wall time of the process | Peak memory of the process |
|---|---|---|---|
| `Book(path)` (opens the file) | 0.01-0.02 s | 0.03-0.05 s | 18 MB |
| `range(["Calc!A1:H1000"])` | 0.02 s | 0.05-0.06 s | 21 MB |
| `find("K00")` | 2.17-2.41 s | 2.21-2.44 s | 24 MB |
| `lint()` | 10.63-12.51 s | 10.75-12.64 s | 592-601 MB |
| `delete_rows("Calc!3:10").save(out)` (with the recalculation) | 16.23-17.33 s | 16.30-17.39 s | 1417 MB |
| `calculate({"rate": 0.07}, ["Summary!A1", "Summary!A3", "Summary!A4"]).value_ranges()` | 8.23-9.13 s | 8.30-9.21 s | 1266 MB |

### WebAssembly

One call of each tool of the package in Node 22 against the command on the same small workbook. The times are those of
`sheetrs-wasm/tests/timing.mjs` (median of 20 calls after 3 warm-ups, `Uint8Array` in; the command's median of 20 runs
includes the start of the process, the call does not). The memory is the peak of one process per scenario
(`sheetrs-wasm/tests/call_run.mjs`: Node loads the package, makes the call once and exits; Node with the package loaded
and no call peaks at 90 MB in 0.071-0.086 s); the workbooks are small, so these are the cost of a call:

| Call | WebAssembly time (ms) | WebAssembly peak (Node process) | Command time (ms) | Command peak |
|---|---:|---:|---:|---:|
| lint, SARIF (`tests/lint_book.xlsx`) | 9.1 | 108-110 MB | 9.8 | 28 MB |
| statistics (`tests/lint_book.xlsx`) | 3.3 | 103 MB | 5.6 | 28 MB |
| query, 9 rows (`book.xlsx`) | 2.0 | 103-106 MB | 4.3 | 28 MB |
| edit, 3 operations (`book.xlsx`) | 5.4 | 115 MB | 10.3 | 28 MB |
| calculate (`loan.xlsx`) | 13.6 | 121-130 MB | 9.0 | 28 MB |

A compiled workbook as a native executable and as a WebAssembly module (`sheetc build --target wasm`; whole process,
the module run by `sheetrs-wasm/tests/module_run.mjs`; same answer, the executable adds a final newline):

| Program and request | Wall time | Peak memory |
|---|---|---|
| `small.xlsx`, one output: native executable | 0.003 s | 28 MB |
| `small.xlsx`: module (3.76 MB) | 0.10-0.11 s | 79-80 MB |
| `formulas.xlsx`, three outputs: native executable (193.6 MB) | 2.11-2.26 s | 542 MB |
| `formulas.xlsx`: module (188.5 MB) | 3.17-3.25 s | 861 MB |

The rows of this table were measured before the changes to the engine (cell index, constants, rounding) and with the `bincode` payload: today's native executable of `formulas.xlsx` answers in 1.35-1.41 s and 474 MB (see the model use below), and the module was not measured again.

For a small workbook the cost of the module is the start of Node and the compilation of the module; for a large one,
decoding its program, about 1.5 times the time and memory of the native executable.

### Sizes

| File | Bytes | gzip |
|---|---:|---:|
| `sheetrs` (all the tools, `dist`) | 16,105,512 | |
| `sheetc` (compiler, engine and the base of an executable) | 8,791,704 | 3,570,291 |
| `sheetrs_wasm_bg.wasm` (package of every tool; carries the two bases) | 16,982,478 | 5,457,150 |
| `sheetrs-wasm-0.4.1.tgz` | 5,635,391 | |
| `sheetrs-rt.wasm` (base module, engine) | 3,684,056 | 1,239,594 |
| `sheetrs-rt-template.wasm` (base, engine with readers and writers) | 4,708,523 | 1,621,735 |
| Python wheel, linux x86_64 (`manylinux_2_38`) | 6,961,523 | |
| `_sheetrs.abi3.so` inside the wheel | 16,285,856 | |

The wheel carries every tool in one extension, `abi3` for Python 3.13 and later, one wheel per platform.

## Against LibreOffice

The same flow, on the same input, in the same session, in LibreOffice 26.8.2 headless driven through UNO
(`scripts/measure_libreoffice_flow.py`, rounds alternated), for every flow that LibreOffice has an equivalent of.
`formulas.xlsx` unless said; LibreOffice stores with the filter `Calc MS Excel 2007 XML`. The SheetRS figures of the
read flows are those of the tables above (same hour, same machine); those of the edit and calculate flows were taken by
the contrast script in the same rounds as LibreOffice's.

### Load, edit, calculate, read

| Flow | SheetRS | LibreOffice |
|---|---|---|
| Load, calculate every formula, write the workbook | `sheetc calculate -o`: 10.8-11.1 s, 1598 MB | load, `calculateAll()`, store: 37.5-38.0 s, 2193 MB |
| Delete rows `Calc!3:10`, calculate, write | `sheetrs edit --delete-rows`: 23.3 s, 1646 MB *; Python `Book.delete_rows(...).save(out)`: 24.83-25.46 s (the process 26.87-27.56 s), 1655 MB | load, `Rows.removeByIndex`, `calculateAll()`, store: 47.7-48.9 s, 3286-3288 MB |
| The same plus formulas to values of `Summary` | `sheetrs edit --delete-rows … --formula-to-value Summary`: 41.0 s, 1890 MB * (31.5-32.6 s, 1639 MB with the conversion that calculates what it converts, `scripts/bench.py`) | the same steps with `.uno:ConvertFormulaToValue`: 47.7-49.5 s, 3293-3295 MB |
| Read the whole workbook (XLSX) | `sheetrs stats`: 1.73-1.77 s, 14 MB; Python `Book(path)`: 1.97-2.05 s, 17 MB | load: 10.2-10.6 s, 1348-1350 MB |
| Find a text in every sheet (XLSX; 14 930 cells in both) | `sheetrs query --find K00`: 2.32-2.38 s, 17-18 MB; Python `find`: 2.34-2.39 s after the open, 24 MB | load and `findAll`: 11.8-12.6 s, 1348-1349 MB (the search alone 2.0-2.1 s) |
| Read the whole workbook (ODS) | `sheetrs stats`: 2.90-2.95 s, 14 MB | load: 28.3-29.2 s, 1187 MB |

† Not measured again in the session of the rows above.

Phases of LibreOffice (range of the rounds), XLSX: load 10.2-11.6 s, deleting the rows 9.2-10.5 s, `calculateAll()`
6.6-7.5 s, conversion under 0.5 s, store 15.4-18.7 s. With the same work in the same order, SheetRS takes 0.19
of the time of LibreOffice and 0.53 of its memory in the calculation, 0.40 and 0.43 deleting rows, 0.66 and 0.44 when
the conversion goes with the deletion, and about 0.16 of the time reading. All the rows of this table, SheetRS and
LibreOffice, are from the same three alternated rounds, but the machine moves 10-20 % between sessions, so the ratios are
indicative.

### The model use: inputs, calculate, ranges

A model is used by giving inputs and reading output ranges, one time or many. SheetRS: `sheetc calculate --request`
(compiles in memory on every call), a compiled executable (`--request`, or `--serve` that answers one request per
line and keeps the model), and Python `Book(path).calculate(inputs, ranges).value_ranges()`. LibreOffice through UNO: open the
document once, set the input cell, calculate (`calculateAll()` on the first call, `calculate()` after, which
recalculates what the input moved), and read the output cells with `getDataArray()`; 100 calls with the document kept
open, the input changed by 0.1 % each time. `data_declared.xlsx` declares input `scale` (`Summary!B4`) and outputs
`total` and `scaled`; `formulas.xlsx` declares input `rate` (`Data!I1`) and outputs `total`, `month3` and `keys`.

`data_declared.xlsx` (600 005 formulas):

| Flow | One call from a cold start | Peak memory | 100 calls |
|---|---|---|---|
| `sheetc calculate --request` | 4.04-4.43 s | 684-687 MB | 100 separate runs: nothing is kept between calls |
| `sheetc build` of the executable (once) | 4.07-4.51 s | 683-703 MB | no LibreOffice equivalent |
| Compiled executable, `--request` | 0.003-0.004 s | 18 MB | 100 separate runs |
| Compiled executable, `--serve` | | 18 MB | 0.004-0.005 s in all (0.043 ms per call, the start of the process included) |
| Python `Book(path).calculate(...).value_ranges()` | 4.41-4.54 s | 747 MB | 100 separate runs: each call opens the workbook again |
| LibreOffice through UNO | open 10.78-12.15 s, first call 0.28-0.31 s | 361 MB | 1.0-1.2 s for the 100 calls (7.4-9.5 ms per call with `calculate()`) |

`formulas.xlsx` (2.4 M formulas):

| Flow | One call from a cold start | Peak memory | 100 calls |
|---|---|---|---|
| `sheetc calculate --request` | 6.27-6.84 s | 1152 MB | 100 separate runs: nothing is kept between calls |
| `sheetc build` of the executable (once) | 6.80-6.93 s | 1152 MB | no LibreOffice equivalent |
| Compiled executable, `--request` | 1.347-1.408 s | 474-475 MB | 100 separate runs |
| Compiled executable, `--serve` | | 474-475 MB | 40.2-43.6 s in all (0.40 s per call): every request calculates the formulas the input reaches |
| Python `Book(path).calculate(...).value_ranges()` | 9.61-10.82 s | 1267 MB | 100 separate runs: each call opens the workbook again |
| LibreOffice through UNO | open 11.16-14.91 s, first call 7.04-8.78 s | 1346-1348 MB | 182.9-190.3 s for the 100 calls (1.77-1.83 s per call with `calculate()`) |

Python and `sheetc calculate --request` keep no state, so 100 calls are 100 runs of the single call (not run). A
compiled executable answers a request in milliseconds when the workbook declares its inputs and outputs (the program
is the cone of the outputs), and keeps the model between requests with `--serve`: the 100 calls take 0.4 % of what
LibreOffice takes on `data_declared.xlsx` and 22 % on `formulas.xlsx`, whose input moves a column of 300 000 rows and
a running total.

### Notes and rows without an equivalent

- **No LibreOffice equivalent**: `sheetrs lint` (LibreOffice has no rules over formulas), the counters of
  `sheetrs stats` (the rows above compare their reading cost with loading the same file), `sheetc build` of an
  executable and the WebAssembly module of a compiled workbook (LibreOffice has no program to compile), and the calls of
  the WebAssembly package (LibreOffice is not a library for a page).
- The conversion flow does not compute the same values: in one `sheetrs edit` command the flags apply to the input
  file in a fixed order (see the README of `sheetrs-edit`), so the conversion reads the values before the deletion
  (`Summary!A1` 44400419159.41), while LibreOffice deletes first (`Summary!A1` 44399800557.29). The work is the same set
  of steps, so the times and memory compare, the values do not.
- Without `calculateAll()` LibreOffice converts the zeros that the file stores as results (the conversion takes
  35.6-36.5 s, 3272-4161 MB): it is faster, and wrong for this file.
- The search in an ODS is not the same flow: LibreOffice also looks at the results of formulas (29 760 cells in the
  ODS of this file), and `sheetrs query --find --include-formulas` takes 22.46-23.01 s on it (see the
  [`sheetrs query`](#sheetrs-query) table).
- Memory of LibreOffice is the `VmHWM` of `soffice.bin`; a deletion keeps the undo data, which is why it takes
  1.1 GB more than the calculation alone.

## Known limits

Measured at this state; none is being worked on here.

- **A deletion, replacement or name removal always pays for reading and compiling the workbook** when the plan leaves a
  formula to calculate, and `--calculate-all` calculates the whole workbook (`--delete-names base_rate` takes 20 s of an XLSX
  and 27 s of an ODS with it, 7 and 13 s by default, although no formula reads the name). Operations that could be applied in a
  batch before one recalculation are not yet ([#102](https://github.com/cosmoscalibur/sheetrs/issues/102)).
- **Memory of the editor with the engine**: 1.6-2.2 GB for 2.4 M formulas, against 19 MB of an editor that only
  streams (the conversions of an ODS). Intermediate parts of a plan of several passes are stored uncompressed
  ([#86](https://github.com/cosmoscalibur/sheetrs/issues/86)).
- **Deleting columns of an ODS** takes 24-25 s on `formulas.ods`, 1.5 times the XLSX and about as much as deleting rows
  of the same file (the 136-139 s that this row held before the empty run that closes each row was kept up to the last column
  were fixed: [#107](https://github.com/cosmoscalibur/sheetrs/issues/107)).
- **A search that also looks at the results of formulas** (`query --find … --include-formulas`) keeps every formula
  cell to infer the format of its result: 621-646 MB and 18-23 s on 2.4 M formulas, against 17 MB and 2.3 s of a search
  by value alone. Before the change that reads each sheet whole after eight windows (instead of the file once per
  referenced range) and finds the cells of a range by position, it did not finish in five minutes.
- **The XLSX reader** is not profiled in detail ([#83](https://github.com/cosmoscalibur/sheetrs/issues/83)).
- **`sheetc calculate -o`** writes the results truncated to 14 digits while the planner writes the full number
  ([#103](https://github.com/cosmoscalibur/sheetrs/issues/103)).
- **WebAssembly** has the memory of the host: a module of a large workbook needs 1.5 times the memory of the native
  executable, and the depth of the chains of references that are read at run time is that of the call stack of the host
  ([#95](https://github.com/cosmoscalibur/sheetrs/issues/95)).
- **Not measured**: the LibreOffice contrast of the edit flows on an ODS and of 100 repeated calls of the Python
  model use (it keeps no state: 100 runs of the single call); other platforms (macOS and Windows, [#99](https://github.com/cosmoscalibur/sheetrs/issues/99)), instruction
  counts (no hardware counters in the virtual machine) and the split of the time inside the passes of `sheetrs edit`.

## Pending: benchmark against third-party libraries

**Pending, not run.** Compare the reading (and the writing where the library has it) of SheetRS with the libraries
that [docs/references.md](references.md) lists: `calamine` (Rust), `fastexcel` (Python, on calamine), `openpyxl`
(Python) and `xlwings` (Python, which needs Excel). The same generated workbooks (`formulas.xlsx`, `data.xlsx`),
the same harness (`scripts/bench.py`: wall time and peak memory, min-max of 3 runs, one process per run), the same
reads (the whole workbook, one range of a large sheet, the cells that match a text) and, for the writers, the same
edits (a deletion of rows, a conversion of formulas to values). Add the results as a section of this document.
