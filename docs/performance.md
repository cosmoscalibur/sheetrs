# Performance

Time and memory of every tool of SheetRS at the state of the repository, on workbooks that any checkout can
generate, and the same flows in LibreOffice 26.8.2 as the reference. The document holds the current numbers only:
what a change did to them is in the history of the commit that made it. A change that moves a number updates
the table of its flow here, and `sheetrs-linter/docs/PROFILING.md` when it targets the cost of a rule or a phase.

- [Method](#method)
- [Fixtures](#fixtures)
- [Results](#results): [lint](#sheetrs-lint), [stats](#sheetrs-stats), [query](#sheetrs-query),
  [edit](#sheetrs-edit), [calculate and the compiler](#sheetc-calculate-and-sheetc-build),
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
python3 scripts/bench.py run DIR                       # every group; or `run DIR lint,edit` for some
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

The engine calculates what a deletion or a replacement leaves stale; the workbook is marked to be recalculated
only if the engine cannot. A deletion of rows, columns or sheets (with conversions) finds the formulas it leaves
stale first: none, no recalculation pass; only the error `#REF!`, written without the engine; otherwise the engine
(see the README of `sheetrs-edit`). `formulas.xlsx` stores `0` as every cached result and says so (`fullCalcOnLoad`),
so no result of it can be trusted: every plan on it calculates the whole workbook, and the conversions calculate
first. In `formulas.ods` the results are stored. A cell with `*` is one run (the rows were measured one at a time
and not repeated); the others are the minimum and the maximum of 3 runs.

| Scenario (flags) | `formulas.xlsx` | `formulas.ods` |
|---|---|---|
| `--delete-sheets Lookup` | 22.48-22.60 s, 1839-1840 MB | 49.05-49.98 s, 1893 MB † |
| `--delete-names base_rate` | 16.3 s, 1646 MB * | 25.3 s, 1656 MB * |
| `--formula-to-value Calc` | 24.25-25.65 s, 1646 MB † | 5.89-5.98 s, 14 MB ‡ |
| `--formula-to-value Calc!B:D Calc!1:1000` | 25.99-26.44 s, 1646 MB † | 8.29-8.57 s, 19 MB † |
| delete sheet + name, convert `Calc` (one command) | 23.4 s, 1646 MB * | 12.4 s, 297 MB * |
| `--delete-rows Data!10:20` | 25.39-25.57 s, 1646 MB | 51.59-52.54 s, 1745 MB † |
| `--delete-rows Data!10:20 --formula-to-value Calc` | 24.9 s, 1646 MB * | 30.7 s, 1656 MB * |
| `--delete-rows Calc!3:10` | 23.3 s, 1646 MB * | not measured |
| `--delete-rows Calc!3:10 --formula-to-value Summary` | 41.0 s, 1890 MB * | not measured |
| `--delete-columns Data!B:D` | 21.59-22.07 s, 2195-2196 MB | 135.98-139.21 s, 2218 MB † |
| rows and columns together | 23.93-24.25 s, 2196 MB | 136.55-137.26 s, 2214 MB † |
| sheet, name, rows and columns (one command) | 24.81-26.28 s, 2200 MB † | 134.74-138.87 s, 2218 MB † |

† Not measured again after the last change of the analysis (it reads the cached results of the cells the plan leaves
alone, and a defined name that points to cells no deleted row or column touches no longer sends the plan to the
whole-workbook path): a conversion alone does not use the analysis; the ODS rows are those of
[#107](https://github.com/cosmoscalibur/sheetrs/issues/107) and were left out of the last round.

‡ Measured after the writers stopped allocating per cell (3 runs, same session as the binary before the change: 6.96-7.02 s, 14 MB,
the same output byte for byte). The writers alone, without the analysis or the engine, on `formulas.xlsx` / `formulas.ods`
(`WorkbookModifications` through `modify_workbook`, 14 MB peak in every case): conversion of `Calc` 3.97 → 3.07 s / 5.37 → 4.24 s,
`--delete-rows Data!10:20` 7.98 → 6.28 s / 11.55 → 9.84 s, `--delete-columns Data!B:D` 7.03 → 5.54 s / 10.67 → 8.95 s. The 136-139 s
of the ODS column deletion are not the writer: end to end it took 143 s and 2416 MB, so the rest is the analysis and the engine.

A conversion alone streams (the ODS rows: 19 to 298 MB); what needs the engine (a deletion that leaves a formula
stale, or an XLSX whose results are not stored) holds the formulas of the workbook in memory, about 1.6-2.2 GB for
2.4 M formulas: the engine still compiles and evaluates the whole workbook when it runs, and only writes the formulas
that are stale.

`formulas_cached.xlsx`, where the cached results are the true ones, so the formulas a deletion leaves stale are found
(one run of each; `sheetrs edit --dry-run` says what the plan does):

| Scenario (flags) | Wall time | Peak memory | What the plan does |
|---|---|---|---|
| `--delete-sheets Lookup` * | 41.32 s | 1825 MB | the engine calculates 1 500 004 formulas that read the sheet (`VLOOKUP` is not in the list of functions that pass the error on) |
| `--delete-names base_rate` * | 19.73 s | 1656 MB | the whole workbook is calculated (a name is not analyzed) |
| `--delete-rows Data!10:20` * | 42.40 s | 1797 MB | the engine calculates 300 052 formulas and 22 are given `#REF!`; the list is longer than 20 000, so the whole program runs and only those cells are written |
| `--delete-rows Data!10:20 --formula-to-value Calc` * | 22.59 s | 1655 MB | the engine calculates 1 formula |
| `--delete-rows Calc!3:10 --formula-to-value Summary` * | 31.43 s | 1655 MB | 299 991 formulas are given `#REF!`, no engine |
| delete sheet + name, convert `Calc` (one command) * | 9.71 s | 297 MB | the whole workbook is calculated (the plan has a name), after the conversion has taken the formulas of `Calc` |

The analysis reads the workbook once more (compiling its formulas, and the cached values of the cells that are not
numbers): on this workbook the dry run of a deletion takes 8-13 s, about what the engine costs, so a plan that ends
with no engine or with a few formulas is not much faster than one that calculates them all: the saving shows where
the workbook is read once and the engine would have evaluated what nothing reads. LibreOffice was not run on this
input, so there is no contrast for these rows; the contrast of the edit flows is in [Against LibreOffice](#against-libreoffice).

On `loan.xlsx` (1 462 formulas): `--delete-rows 'Loan!2'` gives `#REF!` to 1 083 formulas without the engine (the
chain from the payment, and the 360-row schedule that reads the rate, whose other operands are known not to be errors
from the cached results) and leaves 2 to the engine: `B8`, which reads by `INDIRECT`, and `B372`, a `SUM` over the
range that holds the formulas given the error; the rest (377) keep their result. `--delete-rows 'Loan!2'
--formula-to-value Loan` is one pass and no recalculation, with the values the workbook had.

### `sheetc calculate` and `sheetc build`

`sheetc calculate -i formulas.xlsx`: every formula calculated again, the workbook written (`-o OUT`) 10.39-10.82 s,
1598 MB; one declared output as JSON (`--range Summary!A1`) 10.52-10.54 s, 1598 MB.

`sheetc build` and the executable it writes, one run per row group (`sheetrs-compiler/scripts/measure_perf.py`; the
stages come from `SHEETC_TRACE`; median of 3):

| Workbook | Stage | Wall s | Peak MB |
|---|---|--:|--:|
| `small.xlsx` (0.03 MB) | build total (executable 9.02 MB) | 0.02 | 17 |
| | execute a request | 0.004 | 14 |
| `data.xlsx` (49.19 MB, no settings: inputs deduced) | build total (executable 114.78 MB) | 7.59 | 1540 |
| | read workbook / lower formulas / dependency graph / fold / assemble / write executable | 3.48 / 0.23 / 2.37 / 0.08 / 0.03 / 0.49 | |
| | engine start-up (`--meta`) / execute a request | 0.620 / 2.745 | 453 / 757 |
| `data_declared.xlsx` (inputs and outputs declared) | build total (executable 9.00 MB) | 4.52 | 724 |
| | read workbook / lower formulas / dependency graph / fold / assemble / write executable | 3.40 / 0.16 / 0.21 / 0.67 / 0.00 / 0.01 | |
| | engine start-up / execute a request | 0.003 / 0.002 | 14 / 14 |
| `formulas.xlsx` (28.98 MB) | build total (executable 93.54 MB) | 6.85 | 1309 |
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
| `sheetrs-wasm-0.4.0.tgz` | 5,635,391 | |
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
| The same plus formulas to values of `Summary` | `sheetrs edit --delete-rows … --formula-to-value Summary`: 41.0 s, 1890 MB * | the same steps with `.uno:ConvertFormulaToValue`: 47.7-49.5 s, 3293-3295 MB |
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

- **What the planner cannot tell is calculated whole.** A replacement, a deleted name, an input that says its
  results are stale (the `formulas.xlsx` of this document), a workbook with cycles, formulas the compiler keeps as the
  values the file cached, or a defined name that points into what is deleted or is not a plain reference: the engine
  calculates the whole workbook
  (`--delete-names base_rate` takes 16-25 s here, as much as the recalculation, although no formula reads the name).
  Operations that could be applied in a batch before one recalculation are not yet
  ([#102](https://github.com/cosmoscalibur/sheetrs/issues/102)).
- **Finding the stale formulas reads the workbook once more.** On `formulas.ods` the dry run of a deletion, which only
  does that, takes about 17 s; a plan that ends with the engine calculating most formulas pays it for nothing, and the
  rows of the ODS deletions above include it.
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
