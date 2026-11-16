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

Everything below comes from one run of `scripts/bench.py` on one machine, nothing else running.

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

With `scripts/benchmark.toml`: every default-active rule, thresholds fixed so that every run does the same rule work.

| Workbook | Wall time | Peak memory |
|---|---|---|
| `formulas.xlsx` | 16.36-16.80 s | 658-661 MB |
| `formulas.ods` | 21.52-21.69 s | 693-701 MB |

Where the time goes (XLSX): reading the workbook 15 %, rule execution 79 %, writing the report 5 %; the rules
that cost most are `INT401-403` (16 %) and `CALC202` (14 %): the table is in
[PROFILING.md](../sheetrs-linter/docs/PROFILING.md). The peak of the run is the floor of the workbook read
(329 MB) plus what the rules keep: the dependency graph of `CALC202` (about 110 MB) and the 1.2 M violations of
`CALC201` that this configuration produces.

### `sheetrs stats`

| Workbook | Wall time | Peak memory |
|---|---|---|
| `formulas.xlsx` | 1.73-1.77 s | 14 MB |
| `formulas.ods` | 2.90-2.95 s | 14 MB |

The counts are taken while the reader goes: the workbook is never built, and the formulas of an ODS are
not converted to Excel's notation (the counts do not read them). The peak is the floor of the harness; what
is left of the time is the archive and the XML: in the XLSX about 75 % is inflating, the CRC and tokenizing
(quick-xml), and the profile of the ODS has the same shape.

### `sheetrs query`

| Query | `formulas.xlsx` | `formulas.ods` |
|---|---|---|
| `--list-sheets` | 0.003-0.004 s, 17 MB | 1.36-1.40 s, 19 MB |
| `--range Calc!A1:H1000` | 0.011-0.020 s, 17 MB | 0.49 s, 19 MB |
| `--sheet Summary` | 0.003-0.004 s, 17 MB | 1.34-1.40 s, 19 MB |
| `--find K00` (values, every sheet; 14 930 cells) | 2.32-2.38 s, 17-18 MB | 3.61-3.82 s, 19 MB |
| `--find VLOOKUP --in formula` (every formula) | 6.17-6.29 s, 103 MB | 6.49-6.79 s, 108 MB |
| `--find K00 --include-formulas` (values and the results of 2.4 M formulas) | 17.74-18.51 s, 621 MB | 22.46-23.01 s, 646 MB |

The readers skip the sheets that are not asked for and stop after the last row a window reaches; an ODS has to read
its content part to know where the sheets are.

### `sheetrs edit`

The calculation mode decides what the engine does in a plan that leaves results stale (rows, columns, sheets or names
deleted, text replaced) or that converts formulas to values (see the README of `sheetrs-edit`, *The cache is the default*).
By default, as `LibreOffice` does when it opens a file and when a cell changes, the engine calculates the formulas the plan
reaches and what depends on them and every other formula keeps the result the file caches; with `--calculate-all`
(`calculateAll()` of `LibreOffice`) it calculates every formula, and the workbook is marked to be recalculated only if the
engine cannot (see the README of `sheetrs-edit`). The next two sections measure both on the referent size of every table
here, **300 000 rows** of `Calc` (`scripts/bench.py fixtures DIR`, `--rows 300000`: 2.4 M formulas, 3.6 M cells), in the
same build and the same session, minimum and maximum of 3 runs; the second one is the cache against `--calculate-all` for the
plans that leave results stale.

`formulas.xlsx` stores `0` as every cached result and says so (`fullCalcOnLoad`): a run on it has no results to keep and
calculates as `--calculate-all` does, so the workbooks of the comparisons are `formulas_cached.xlsx` (the same workbook
with the results LibreOffice computed) and `formulas.ods` (which stores them).

The table below was measured before the calculation modes, when every plan ended with the engine calculating every formula:
it is the `--calculate-all` scenario of today (`scripts/bench.py run DIR edit`; a row that deletes calculates the 2.4 M
formulas). The default is in the sections that follow.

| Scenario (flags) | `formulas.xlsx` | `formulas.ods` |
|---|---|---|
| `--delete-sheets Lookup` | 20.58-21.22 s, 1799 MB | 33.35-35.20 s, 1882 MB |
| `--delete-names base_rate` | 17.05-18.40 s, 1580 MB | 29.02-29.79 s, 1586 MB |
| `--formula-to-value Calc` | 15.81-16.53 s, 1566 MB | 22.93-23.62 s, 1600 MB |
| `--formula-to-value Calc!B:D Calc!1:1000` | 10.91-11.94 s, 1049 MB | 18.17-18.67 s, 1059 MB |
| delete sheet and name, convert `Calc` (one command) | 17.80-18.71 s, 1566 MB | 30.70-31.29 s, 1600-1608 MB |
| `--delete-rows Data!10:20` | 24.76-25.45 s, 1600 MB | 36.95-37.48 s, 1604 MB |
| `--delete-rows Data!10:20 --formula-to-value Calc` | 19.11-19.94 s, 1562 MB | 33.01-33.83 s, 1600-1608 MB |
| `--delete-rows Calc!3:10 --formula-to-value Summary` | 31.47-32.57 s, 1639 MB | 45.21-45.95 s, 1645 MB |
| `--delete-columns Data!B:D` | 22.17-23.10 s, 2152 MB | 30.83-31.75 s, 2166-2167 MB |
| rows and columns together | 23.07-24.03 s, 2152 MB | 31.50-32.34 s, 2167 MB |
| sheet, name, rows and columns (one command) | 23.10-23.86 s, 2150 MB | 31.86-33.28 s, 2164-2165 MB |

**A conversion of formulas to values has two scenarios**, both measured here (`scripts/bench.py run DIR edit --workbooks
formulas_cached.xlsx,formulas.ods --scenarios ...`). By default it converts the results the file caches for the formulas the
plan does not reach and what the engine calculates for the ones it reaches (a deletion or a replacement pushed with it): where
nothing is reached nothing is calculated, and an XLSX or an ODS that caches its results is streamed (the short path). With
`--calculate-all` the engine reads and compiles the workbook and calculates what it converts (the long path, the one that
does not depend on what the file caches).

| Scenario (flags) | Default (the cache) | `--calculate-all` |
|---|---|---|
| XLSX `formulas_cached.xlsx`, `--formula-to-value Calc` | 5.86-5.95 s, 18 MB | 18.86-19.32 s, 1590 MB |
| XLSX `formulas_cached.xlsx`, `--formula-to-value Calc!B:D Calc!1:1000` | 5.59-5.70 s, 18 MB | 14.45-14.79 s, 1171 MB |
| XLSX `formulas_cached.xlsx`, delete sheet and name, convert `Calc` | 8.89-8.97 s, 230 MB | 22.16-22.41 s, 1590 MB |
| XLSX `formulas_cached.xlsx`, `--delete-rows Data!10:20 --formula-to-value Calc` | 10.64-10.86 s, 265 MB | 23.75-24.57 s, 1572-1579 MB |
| ODS `formulas.ods`, `--formula-to-value Calc` | 24.4 s (single run) | 23.4 s (single run) |
| ODS `formulas.ods`, `--formula-to-value Calc!B:D Calc!1:1000` | 26.2 s (single run) | 24.3 s (single run) |
| ODS `formulas.ods`, delete sheet and name, convert `Calc` | 33.5 s (single run) | 31.3 s (single run) |
| ODS `formulas.ods`, `--delete-rows Data!10:20 --formula-to-value Calc` | 38.0 s (single run) | 35.1 s (single run) |

The default is 2.2-3.2 times faster (XLSX) and 2.4-4.3 times (ODS) than `--calculate-all` and holds 18-266 MB where it holds
1.0-1.6 GB. The price is what the file caches (the program that wrote it may not have calculated it), which is what
`LibreOffice` converts too. An ODS has no short path: `LibreOffice` calculates on load the formulas that cache a plain number in a cell with no format, which is every formula of `formulas.ods`, so the plan
starts with that calculation (see *The startup of an ODS* below) and the ODS rows of this table are single runs of the build of the startup, default and `--calculate-all` alike. What needs the engine holds the formulas of the workbook in memory, about 1.6-2.2 GB for
2.4 M formulas (a conversion of columns and rows only the cone of the cells it names: 1.05 GB for `Calc!B:D Calc!1:1000`). The ODS
rows that delete cost about 10 s more than the XLSX ones because the ODS is read and compiled (6-9 s of reading) and the output of the
deleting pass is read again by the engine. `--delete-columns` on an ODS took 136-160 s before the empty run that closes each row was
kept up to the last column ([#107](https://github.com/cosmoscalibur/sheetrs/issues/107)): a reader that finds that run short of the
edge takes it for formatted cells. LibreOffice was not run on this input, so there is no contrast for these rows; the contrast of the
edit flows is in [Against LibreOffice](#against-libreoffice).

### The cache and `calculate_all`

The default of `sheetrs edit` (see the README of `sheetrs-edit`) calculates only the formulas the plan reaches (what `LibreOffice` recalculates
after the same edit, *What a deletion reaches* in `sheetrs-compiler/docs/FIDELITY.md`) and the ones that depend on them; every other formula is a
constant that holds the result the file caches. It is measured against `--calculate-all` on the workbooks that cache true results, with
`scripts/bench.py run DIR cache` (300 000 rows, 2.4 M formulas, `dist` build, min-max of 3 runs, the same build and session as the artifact above; the XLSX rows were measured with 3 runs, and the figures of this round that
are a single run say so: only significant differences are meaningful). Each row is a plan whose recalculation has the share of the 2.4 M formulas that
the last column says.

| Workbook and scenario | Default (the cache) | `--calculate-all` | Formulas calculated (cone) |
|---|---|---|---|
| XLSX `--delete-names base_rate` (nothing reads it) | 7.06-7.39 s, 1235 MB | 18.59-18.99 s, 1603 MB | 0 |
| XLSX `--delete-rows Data!299990:299995` | 21.02-21.53 s, 1256 MB | 26.71-28.36 s, 1606-1607 MB | 59 |
| XLSX `--delete-rows Data!10:20` (the running total reads on) | 22.32-23.12 s, 1275-1276 MB | 27.71-28.78 s, 1622 MB | 300 074 (12.5 %) |
| XLSX `--delete-sheets Lookup` (a sheet goes: `LibreOffice` calculates every formula) | 23.69-25.50 s, 1800 MB | 23.77-24.24 s, 1800 MB | all |
| XLSX `--delete-columns Data!B:D` (every formula reads them) | 25.20-25.37 s, 2150 MB | 23.62-24.27 s, 2152 MB | 37.5 % that read the columns and what depends on them |
| ODS `--delete-names base_rate` | 26.8 s (single run) | 25.5 s (single run) | all (every formula is a root of the startup) |
| ODS `--delete-rows Data!299990:299995` | 35.28-36.12 s, 1586 MB (earlier build) | not measured | all |
| ODS `--delete-rows Data!10:20` | 34.0 s (single run) | 32.3 s (single run) | all |
| ODS `--delete-sheets Lookup` | 27.7 s (single run) | 28.1 s (single run) | all |
| ODS `--delete-columns Data!B:D` | 28.2 s (single run) | 28.0 s (single run) | all |

**The startup of an ODS (single runs, `formulas.ods`, `dist` build, one session).** `LibreOffice` calculates on load every ODS formula that caches a number, a currency or a date in a cell
with no number format (and a cached error, an array formula, a result with no type: `sheetrs-compiler/docs/FIDELITY.md`, *What LibreOffice calculates when it opens a file*), and
`formulas.ods` (which `LibreOffice` wrote) has all 2.4 M formulas as roots. `sheetrs edit` on an ODS therefore starts with that calculation, in the mode of the run (the roots by default, every
formula with `--calculate-all`), as the first pass that runs the engine. Where the earlier default spent its time, phases of `--formula-to-value Calc`, one run each (timers around each step, `SHEETC_TRACE` for the compiler):

| Phase | Default before (scans, two passes) | Default now (startup) |
|---|---|---|
| Plan: scan of the bytes for formulas with no result | 1.2 s | none |
| Plan: scan of the cells for what `LibreOffice` calculates on load (stops at the first one: the first formula is at byte 210 MB of 711 MB, after the `Data` sheet of constants, so it reads 29 % of `content.xml`) | 8.5 s | none |
| Pass 1: read and compile, calculate every root, write every result | 23.8 s | 25.2 s (read, compile, calculate, write and convert, one pass) |
| Pass 2: read the file again and convert the cached values | 6.2 s | none |
| Total | 41.0 s | 25.2 s |

The scans existed because the planner had to decide, before it ran, whether the plan needs the engine and whether a conversion needs the calculation first; the reader already tells the sink, cell by cell, which
formulas the load calculates (`CellSink::recalculated_on_load`), and the compile that has to read the workbook anyway uses that: the decision is folded into the one read. The second pass was not needed by the conversion:
LibreOffice's conversion writes the value the cell holds and forces no calculation (`probe_convert.py`: the cached `140`, `from-excel` and `99` of an XLSX stay, an ODS shows what it calculated on load), so the engine
that calculates the roots writes them, and the cells the conversion names that it did not calculate are converted with the result the file caches, in the same pass. Before this change (`sheetrs edit` ran every ODS
as `--calculate-all`, the build of the commit before) and now, same session, `formulas.ods`:

| Scenario | Before (every ODS as `--calculate-all`) | Now (default) | Now (`--calculate-all`) |
|---|---|---|---|
| `--formula-to-value Calc` | 24.8 s | 24.4 s | 23.4 s |
| `--formula-to-value Calc!B:D Calc!1:1000` | 19.1 s | 26.2 s | 24.3 s |
| delete sheet and name, convert `Calc` | 31.4 s | 33.5 s | 31.3 s |
| `--delete-rows Data!10:20 --formula-to-value Calc` | 36.7 s | 38.0 s | 35.1 s |
| `--delete-names base_rate` | 25.8 s | 26.8 s | 25.5 s |
| `--delete-rows Data!10:20` | 32.7 s | 34.0 s | 32.3 s |
| `--delete-sheets Lookup` | 27.5 s | 27.7 s | 28.1 s |
| `--delete-columns Data!B:D` | 26.4 s | 28.2 s | 28.0 s |
| `--hide-columns Data!B` (nothing stale) | 3.2 s | 27.6 s | 26.3 s |

Nothing is slower than the earlier default (41.0 s for the first row). Against the build that ran every ODS as `--calculate-all` the deletions and conversions cost 0-4 % more (noise of the session is of that order), the conversion of
`Calc!B:D Calc!1:1000` 37 % more (it no longer prunes the workbook to the cells it names: the startup writes every result, 52 MB against 33 MB), and an edit that leaves nothing stale pays the startup, 27.6 s where the file was
written as it is in 3.2 s: that is the price of the results `LibreOffice` shows when it opens the file. The memory was not measured again. An XLSX is not affected by the startup: `formulas_cached.xlsx` `--formula-to-value Calc` and
`--hide-columns Data!B` keep the figures above (not run again: the XLSX path did not change).

**The names of an ODS (+3 s of a deletion of rows or columns).** The rule that a name is reached by a deletion needs the names as they were before the pass, and an ODS gave them by reading all its
cells. They are now read from the skeleton of `content.xml` (the tables and the blocks of names, searched in the inflated bytes: no XML event for a cell). Measured with the binaries of the commit before and
after, three runs each, `formulas.ods`: `--delete-rows Data!10:20` 49.17-50.47 s before, 47.18-47.91 s after; `--delete-columns Data!B:D` 44.66-45.91 s before, 42.12-42.22 s after; rows and columns together
46.36-46.93 s before, 41.92-43.54 s after (5-10 %, 2-5 s: the first bounded step, kept; the inflate of the part and the reading of the rest of the file are what is left). These absolute figures are
of that session, when every plan also looked in the file for what is calculated on load; the rows of the table above are of the final build.

What the cache saves is the calculation and the writing of the results of the formulas that are not reached, and the memory
the program holds for them: 18-22 % of the time and 22-23 % of the peak where a share of the formulas is reached (12.5 % or less),
about 2.8 times the time where none is (the engine is not asked and no result is written). What it does not save is what the plan costs whatever the
engine does: the pass that deletes (7-8 s on this workbook) and, in the recalculation, reading the workbook, parsing its 2.4 M
formulas and building their graph (about 7 s); the second takes the same time with a cone of 59 formulas as with the whole workbook.
Two costs of this round, measured against the binaries of the previous commit in the same session: the rows that delete rows or columns of an
ODS cost about 3 s more (28.8-30.4 s before, 32.2-33.8 s now: the rule that a name is reached when a deletion moves its target needs the names
as they were before the pass, and an ODS has to be read to give them; an XLSX keeps them in a small part, and its rows did not change:
23.2-24.4 s before, 22.0-24.2 s now), and `--delete-names` reads no name at all (13.6-13.9 s before, 13.0-13.8 s now; a first version of the rule read
them and cost 16.9-17.3 s). The walk that finds the cone is made however many formulas are stale: the shortcut that calculated everything where more
than a quarter were stale or had no result is gone, because it wrote the engine's result where `LibreOffice` keeps the one the file caches
(`--delete-columns Data!B:D`, where it applied, is 26.1-27.6 s against 24.5-25.5 s before).

The cost of finding the cone is the walk over the formulas: for a workbook of this size 0.6-1.3 s (the dependencies of
the 2.4 M formulas and the reverse of the graph, which the compiler builds again for what it calculates). It is made however many formulas
are stale: if the cone turns out to be the whole workbook the loss is that walk (1-1.3 s on this workbook, 4-5 %), and what is gained is that the
formulas that nothing reaches keep the result the file caches, as in `LibreOffice`.

**No scan for a qualified `#REF!`.** The compiler reads `Data!#REF!` as the plain error, so no pass inflates or rewrites a sheet for it (an earlier version scanned the sheets of an XLSX for the mark `!#REF!` to write it bare: +0.25 s, 3.6 %, on `--delete-names base_rate` of `formulas_cached.xlsx`, and +0.66 s, 9 %, on `--hide-rows`; it was not kept).

### `sheetc calculate` and `sheetc build`

`sheetc calculate` has the calculation mode of the run (see the README of `sheetrs-compiler`, *The calculation mode*): by default what the
inputs it is given reach is calculated and the rest of the workbook is the result it caches, and `--calculate-all` calculates the whole
workbook. A calculation is built and run in one go, so the cells that can change are the inputs the request gives, not every input the
workbook declares (a build is run later with any of them and keeps the cone of all). The workbook written is the answer, so it carries what was
calculated and, for the rest, what the template caches; a formula nothing reached is not written again. On `formulas_cached.xlsx` (2.4 M formulas;
the workbook declares the input `rate` and its outputs; `scripts/bench.py run DIR calculate`, 300 000 rows, min-max of 3 runs, the same build
(`dist`) and session):

| Flow | Default (the cache) | `--calculate-all` |
|---|---|---|
| `sheetc calculate -i formulas_cached.xlsx -o OUT` (the workbook written, nothing given) | 5.88-6.36 s, 1226 MB | 22.71-24.23 s, 1519-1520 MB |
| `sheetc calculate -i formulas_cached.xlsx --range Summary!A1` (JSON) | 5.80-5.92 s, 1226 MB | 12.74-12.94 s, 1479-1489 MB |
| `sheetc calculate -i formulas_cached.xlsx --request '{"inputs": {"rate": 0.07}}' -o OUT` (the input given: its cone, 1.5 M formulas, is calculated and written) | 19.03-19.38 s, 1229 MB | 26.95-31.35 s, 1571-1598 MB |

**Why the workbook as the answer cost 80 % more, and what fixed it.** Before this round the same command took 15.80-15.93 s by default
(`--range` as JSON: 8.17-8.38 s; `--calculate-all -o`: 23.50-23.94 s; the same session and build profile, measured with the binaries of the
previous commit): the build kept in the program the cone of every input the workbook **declares** (`rate`: 1.5 M of the 2.4 M formulas) because a
program is run later with any of them, although the request gave none, and the answer, being the whole workbook, then calculated and wrote
those 1.5 M results (rewriting the sheet that holds them, 288 MB of XML). The calculation knows what it is given: its roots are the inputs of the
request (`Spec::given_inputs`) and the cells with no result or that nothing can vouch for; with none given nothing is reached and the workbook is
the template as it is (the output equals the input, sheet by sheet), which is what the default of `LibreOffice` shows. The cost that is left,
5.8-6.4 s, is reading the workbook, parsing its 2.4 M formulas and building the graph, as the JSON answer has. With the input given the cone is
calculated and written, as before the change: 19.5-21.2 s with the previous commit, 19.1-20.0 s now. A request read from the standard input
(`--request -`) or from a file is not known when the program is built, so the cone of every declared input is built then.
`formulas.xlsx`, the other workbook of the benchmark, stores `0` as every cached result and says so with `fullCalcOnLoad`: a run on it,
`sheetc calculate` included, has no result to keep and is whole, as `--calculate-all` is; the comparisons use `formulas_cached.xlsx`.

**The artifact** (`scripts/bench.py run DIR modes`, 3 runs of the build and of each run, the session of the final numbers of this round). A program built by default
carries the cached result of the formulas it keeps: 113.73 MB (as many with `--calculate-all`), against 95.13 MB when it is built with
`--force-calculate-all` (16 % lighter; it calculates all of them in every run). The build takes 9.11-9.48 s by default, 16.79-17.68 s with `--calculate-all` and
9.34-9.66 s with `--force-calculate-all`. A run of the executable takes about 2.1-2.5 s in every combination
(the input changed or nothing given, default or `--calculate-all`, built with or without force: 495-542 MB),
because what it costs is reading the 90-108 MB program and starting the engine, not calculating what it kept: on this workbook
the cache of the run saves nothing measurable (the formulas are cheap and the engine calculates the cone as fast as it clones the
cached results), and it pays only in the build and in `calculate`, which skip the calculation of what no input reaches. A
program built by default and one built with `--calculate-all` give the same answers (the same md5 in every run above).

`sheetc build` and the executable it writes, one run per row group (`sheetrs-compiler/scripts/measure_perf.py`; the
stages come from `SHEETC_TRACE`; median of 3):

| Workbook | Stage | Wall s | Peak MB |
|---|---|--:|--:|
| `small.xlsx` (0.03 MB) | build total (executable 9.02 MB) | 0.02 | 17 |
| | execute a request | 0.004 | 14 |
| `data.xlsx` (49.19 MB, no settings: inputs deduced) | build total (executable 114.78 MB) | 7.59 | 1540 |
| | read workbook / lower formulas / dependency graph / fold / assemble / write executable | 3.48 / 0.23 / 2.37 / 0.08 / 0.03 / 0.49 | |
| | engine start-up (`--meta`) / execute a request | 0.620 / 2.745 | 453 / 757 |
| `data_declared.xlsx` (inputs and outputs declared) | build total (executable 9.00 MB) | 4.52 | 704 |
| | read workbook / lower formulas / dependency graph / fold / assemble / write executable | 3.40 / 0.16 / 0.21 / 0.67 / 0.00 / 0.01 | |
| | engine start-up / execute a request | 0.003 / 0.002 | 14 / 14 |
| `formulas.xlsx` (28.98 MB) | build total (executable 93.54 MB) | 6.85 | 1152 |
| | read workbook / lower formulas / dependency graph / fold / assemble / write executable | 2.29 / 0.83 / 0.76 / 1.13 / 1.06 / 0.61 | |
| | engine start-up / execute a request | 0.722 / 1.994 | 423 / 496 |

The stage rows are parts of the build total of their workbook (the peak memory is that of the process). The build keeps one
copy of the graph (built straight into its compressed rows, and freed once the order and the taint are known), finds a
graph node by the row of its cell in the sorted node list instead of in a hash table, and does not keep the references
that lie inside an output (with nothing declared, almost every one): against the build before this, in the same
session, `formulas.xlsx` took 8.33 s and 1598 MB (1355 MB before the operands of a binary operator went in one allocation), `data.xlsx` 9.08 s and 1942 MB and `data_declared.xlsx` 4.81 s and
760 MB. The executables are byte-identical. The rows of the other flows that compile in memory (`sheetc calculate`
above, the recalculation of `sheetrs edit`, Python `calculate`, and the model use below) share these stages and
should take less memory and time; they were not measured again, nor were the start-up and request rows of the
executables, which the change does not touch. A workbook that declares its inputs and outputs keeps only what the outputs need, so its executable is small and
answers in milliseconds; one that does not carries every cell.

`sheetc build` and `calculate` hand the model to `compile` by value, which frees the text of the formulas as soon as they are lowered (nothing after reads it, only whether a cell is a formula): `formulas.xlsx` peaks at 1152 MB instead of 1309 MB and `data_declared.xlsx` at 704 MB instead of 724 MB, with the same time and the same program (`--emit-ir` identical). `data.xlsx` (about 600 000 short formulas, almost all of the peak in the cells and the graph) did not move: 1540-1584 MB either way, and its row above is as it was. The `formulas.xlsx` and `data_declared.xlsx` figures were measured on the same machine in the same session as their baselines; the stage rows were not measured again.

The rows of this table were measured again with the payload in `postcard` (see
[the executable format](../sheetrs-compiler/docs/DESIGN.md#executable-format)): with `bincode` the executables of
`data.xlsx` and `formulas.xlsx` were 207.19 and 193.60 MB and started in 0.56 and 0.81 s with 542 MB (same machine,
same session). The native-executable and WebAssembly-module rows of the next section were measured with the `bincode`
payload and have not been repeated: the sizes and the decoding time there are higher than today's.

### Python

The `Book` of the wheel (`maturin build --profile dist`) on `formulas.xlsx`; every step runs in its own process, which
opens the `Book` first (the statistics are read at open), so the peak of a step is that of its process:

| Step | Wall time of the step | Wall time of the process | Peak memory of the process |
|---|---|---|---|
| `Book(path)` (reads the workbook once, for the statistics) | 1.97-2.05 s | 1.99-2.06 s | 17 MB |
| `range(["Calc!A1:H1000"])` | 0.01 s | 1.92-1.94 s | 21 MB |
| `find("K00")` | 2.34-2.39 s | 4.26-4.40 s | 24 MB |
| `lint()` ‡ | 26.72-27.67 s | 28.79-29.85 s | 1021-1023 MB |
| `delete_rows("Calc!3:10").save(out)` (with the recalculation) | 24.83-25.46 s | 26.87-27.56 s | 1655 MB |
| `calculate({"rate": 0.07}, ["Summary!A1", "Summary!A3", "Summary!A4"]).value_ranges()` | 10.83-11.34 s | 12.86-13.38 s | 1698-1699 MB |

‡ Measured before the linter used less time and memory (see [`sheetrs lint`](#sheetrs-lint), which is about half
of it now); not measured again, as the wheel is not built here.

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

Phases of LibreOffice (range of the rounds), XLSX: load 10.7-11.2 s, deleting the rows 8.2-12.0 s, `calculateAll()`
7.0-8.0 s, conversion under 0.5 s, store 17.8-19.7 s. With the same work in the same order, SheetRS takes 0.29
of the time of LibreOffice and 0.73 of its memory in the calculation, 0.48 and 0.50 deleting rows, 0.83-0.86 and 0.57 when
the conversion goes with the deletion, and about 0.2 of the time reading. The rows of SheetRS marked `*` are one run taken
after the rest, and the LibreOffice column was not run again: the two columns come from different sessions, and the
machine moves 10-20 % between sessions, so the ratios are indicative.

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
| `sheetc calculate --request` | 4.96-5.14 s | 759-760 MB | 100 separate runs: nothing is kept between calls |
| `sheetc build` of the executable (once) | 4.79-4.89 s | 759-760 MB | no LibreOffice equivalent |
| Compiled executable, `--request` | 0.002-0.003 s | 28 MB | 100 separate runs |
| Compiled executable, `--serve` | | 28 MB | 0.004-0.005 s in all (0.04-0.05 ms per call, the start of the process included) |
| Python `Book(path).calculate(...).value_ranges()` | 8.24-8.47 s | 805 MB | 100 separate runs: each call opens the workbook again |
| LibreOffice through UNO | open 9.99-10.24 s, first call 0.29-0.31 s | 362 MB | 1.0-1.1 s for the 100 calls (7.4-7.6 ms per call with `calculate()`) |

`formulas.xlsx` (2.4 M formulas):

| Flow | One call from a cold start | Peak memory | 100 calls |
|---|---|---|---|
| `sheetc calculate --request` | 9.38-9.60 s | 1598 MB | 100 separate runs: nothing is kept between calls |
| `sheetc build` of the executable (once) | 8.90-9.05 s | 1598 MB | no LibreOffice equivalent |
| Compiled executable, `--request` | 2.067-2.125 s | 542 MB | 100 separate runs |
| Compiled executable, `--serve` | | 542 MB | 114.3-116.7 s in all (1.14 s per call): every request calculates the formulas the input reaches |
| Python `Book(path).calculate(...).value_ranges()` | 12.92-13.41 s | 1699-1700 MB | 100 separate runs: each call opens the workbook again |
| LibreOffice through UNO | open 10.31-10.53 s, first call 7.13-7.45 s | 1349-1360 MB | 175.8-179.4 s for the 100 calls (1.70-1.74 s per call with `calculate()`) |

Python and `sheetc calculate --request` keep no state, so 100 calls are 100 runs of the single call (not run). A
compiled executable answers a request in milliseconds when the workbook declares its inputs and outputs (the program
is the cone of the outputs), and keeps the model between requests with `--serve`: the 100 calls take 0.4 % of what
LibreOffice takes on `data_declared.xlsx` and 65 % on `formulas.xlsx`, whose input moves a column of 300 000 rows and
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
  38.4-39.7 s, 3276-3278 MB): it is faster, and wrong for this file.
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
- **Deleting columns of an ODS** takes 136-139 s on `formulas.ods`, five times the XLSX and about 3.5 times deleting rows
  of the same file.
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
