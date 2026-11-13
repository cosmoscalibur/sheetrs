# sheetrs-compiler (`sheetc`)

Compile an XLSX or ODS spreadsheet into a standalone executable (or a library
call) that takes JSON in and returns JSON out.

```text
loan.xlsx ──┐                                ┌─ ./loan   (a single file: the runtime and the program)
            ├─▶ IR ──▶ fold · prune · order ─┼─ program.json  (the IR, portable)
loan.ods  ──┘  (format independent)          └─ Program::calculate() (library)
```

* **One IR, one engine.** Both formats are lowered to the same intermediate
  representation; there is no second execution engine to keep in sync.
* **Inputs and outputs are declared in the workbook**, in a settings sheet (or are not declared: then the
  outputs are the formula cells, as ranges, and the constants the outputs start from are the inputs). Cells that do not depend on a
  declared input are evaluated at build time and disappear; cells that do not feed an output are
  dropped. What remains is the minimal computation.
* **Hybrid compile/run-time.** Static references are resolved and ordered at
  build time (topological order, cycles isolated). `INDIRECT`, `OFFSET` and
  similar functions resolve their targets lazily at run time, on the same
  engine.
* **Faithful to LibreOffice.** The reference of the compiler and of the runtime (and of the
  modifications of `sheetrs edit`) is **LibreOffice 26.8 on Linux**. The engine behaves as it does, for an
  XLSX and an ODS alike (booleans as numbers, `0^0`, the 1899-12-30 null date, number
  formatting…); the cases where that differs from Excel are listed and tested in
  [docs/FIDELITY.md](docs/FIDELITY.md).
* **Execution is configurable where the file is silent:** the document decides
  the date system, iterative calculation and search settings; the build
  configuration (`sheetc.toml` and flags) decides the clock, the seed and how much of the workbook
  stays queryable, and the locale where the workbook does not give it (see *Locale*): any
  tag of the table, fixed at build time.

Architecture and decisions: [docs/DESIGN.md](docs/DESIGN.md). Differences between
Excel and LibreOffice: [docs/FIDELITY.md](docs/FIDELITY.md). Which functions and
operators exist and which are checked against LibreOffice:
[docs/COVERAGE.md](docs/COVERAGE.md).

## Quick start

Requires Rust ≥ 1.99 (`rust-toolchain.toml` pins it).

```bash
cargo build --profile dist -p sheetrs-compiler -p sheetrs-engine   # target/dist/{sheetc, sheetrs-rt}

# The workbook declares its inputs and outputs in a `settings` sheet
# (see below); compiling needs nothing else.
sheetc build -i tests/assets/loan.xlsx -o loan
# cells: 1500 total, 1090 reachable, 1500 kept | formulas: 1086 total ...

./loan --request '{"inputs": {"principal": 250000, "rate": 0.045, "years": 15}}'
# {"valueRanges":[{"name":"payment","range":"Loan!B4","values":[[1912.4832220336577]]},{"name":"interest",...},...]}

./loan --meta                  # document properties, settings sheet rows, exported names
./loan --info                  # technical description: shapes, cycles, diagnostics, build
./loan --version               # runtime version and the compiler that built it
./loan --serve                 # one request per line in, one answer per line out
```

`--info` also says how the executable was built (`build`): the compiler and its version, the
platform (`target`), the options that shape the program (`outputs`, `locale`, `fold`,
`features`), the workbook (`source`: file name, `xlsx` or `ods`, and the SHA-256 of the file) and
when, in UTC (`built_at`, `2026-10-30T12:00:00Z`). `SOURCE_DATE_EPOCH` (seconds since 1970) sets that
time, so a build can be repeated exactly. `sheetc --version` prints the version of the compiler.

The same works from `.ods`, and without producing a file:

```bash
sheetc calculate -i tests/assets/loan.ods --request '{"inputs": {"principal": 1000}}'
sheetc evaluate '=TEXT(1234.5,"#,##0.00")' --locale es          # "1.234,50"
sheetc inspect -i model.xlsx      # sizes, unsupported functions, cycles, run-time refs (never stops)
```

### Commands

| Command | Purpose |
| ------- | ------- |
| `build FILE [-o OUT]` | Compile to an executable. `--emit-ir PATH` also writes the IR as JSON. `--target wasm` writes a WebAssembly module instead ([below](#webassembly-target)). |
| `calculate FILE\|PROGRAM.json` | Compile in memory (or load an IR file) and calculate once. |
| `evaluate "=FORMULA"` | Evaluate a single formula. |
| `inspect FILE` | Report what the compiler sees. |
| `verify FILE` | Recompute the formulas against the values the file caches. |

### The settings sheet

The workbook parameterises itself. A sheet called `settings` (or the name given
with `--settings` / `settings = ...`) has these headers in its first row, in any
order:

| key | value | description | meta_type |
| --- | ----- | ----------- | --------- |
| `principal` | `Loan!B1` | Loan amount | `input` |
| `rate` | `Loan!B2` | Annual rate | `input` |
| `payment` | `Loan!B4` | Monthly payment | `output` |
| `version` | `1.0` | Model version | `info` |

* `meta_type` `input` or `output`: `key` is the name of the field in the JSON
  (an alias, like a named range but defined in this sheet) and `value` the cell
  or range it stands for; a row with no `key` is only the range, which then names itself
  (`Loan!B1`). `value` starts with the sheet (`Loan!B1`,
  `'My Sheet'!B2:C9`); a bare sheet name stands for its used range and a
  workbook's defined name is accepted too. An address without a sheet is an
  error: what it points at would depend on which sheet comes first.
* Any other `meta_type` is free metadata: `key` is the field name and `value` its
  literal value. A row with a `key` and no `meta_type` is metadata of the category `extra`.
* `description` documents the key.
* The executable prints with `--meta`, in this order: the properties the document
  declares (`meta_type` `document`: `title`, `subject`, `description`, `keywords`,
  `creator`, `last_modified_by`, `created`, `modified`, `generator` and the ones the
  file adds, named alike in an XLSX and an ODS), every row of the settings sheet (for
  inputs and outputs the value shown is the canonical range), the inputs and outputs that were
  **not** declared (so that what the program exposes is always listed: `input` and `output`
  rows, the value is the range) and the defined names that are aliases of what is exposed
  (`meta_type` `name`, the value is the definition).
* The same declarations can come from a **JSON file**, for a workbook that has no settings sheet
  (or one that must not change): `settings_file = "model.json"` in `sheetc.toml` (a path relative
  to the file) or `--settings-file`. It replaces the sheet; naming both is an error, and the
  checks below are the same. `inputs` and `outputs` are arrays of ranges, or of objects that give
  a range an alias; the two forms mix:

  ```json
  {
    "inputs":  ["Data!A1", { "name": "years", "range": "Data!A2", "description": "term" }],
    "outputs": [{ "name": "result", "range": "Data!B1:B3" }],
    "meta":    [ { "key": "version", "value": "1.2", "description": "model", "meta_type": "info" } ]
  }
  ```

  A range alone is named by its canonical text (`Data!A1`). `"inputs": []` declares that there are
  no inputs. An existing workbook compiles and calculates unmodified, with its inputs changed.

### What is an input and what is an output

Each of the two is **declared or not**, independently:

| | Declared | Not declared |
| - | -------- | ------------ |
| **Outputs** | exactly those ranges; the program keeps only what feeds them (pruned) | **deduced**: the cells that are calculated (the formulas, not the constants) as the rectangles that hold exactly them, named by their range (`Data!D2:F800`); with too many separate rectangles, the area that holds the formulas of each sheet |
| **Inputs** | exactly those ranges; what does not read them is evaluated at build time (folded) | **deduced**: after the workbook is pruned to what the outputs read, the constants a path of the dependency graph starts from (read by some formula, or inside a defined name or table) are the inputs, as the fewest rectangles that cover them (named by their range, `Data!A1:B800`); an isolated constant starts no path, so it is not an input (it can be declared as an output); nothing that reads an input is folded. A workbook with `INDIRECT`/`OFFSET` can read any cell, so there every constant counts. The analysis is one pass over the dependency graph (what it costs is in [docs/performance.md](../docs/performance.md#sheetc-calculate-and-sheetc-build)) |

Defined names and tables of the workbook are **aliases** in both modes: a name that lies inside the
outputs reads that output, and one that lies inside the inputs gives those inputs, so a request
may use `Total` or `Data!B2:C9` alike. A request's `inputs` key is an input, an alias or a range
inside the inputs (every cell of it an input; anything else is an error naming the cell).
`--meta` always lists the inputs and outputs that were exposed, declared or not, and the names that
are aliases. With nothing declared a workbook needs no adaptation to be compiled with those
defaults.

* There is no other way to declare inputs and outputs: no command-line list.
  A key names one thing: it cannot repeat in the sheet (not as the same type nor
  as another: an input, an output or a field), and it cannot coincide with a
  sheet, a defined name or a table of the workbook. All of that fails the
  build. The sheet is never part of the computation's outputs or queries.
* Rows are validated at build time (unknown range, duplicate, a row that is not an input or an
  output and has no `key`).

### The workbook as the answer: `--embed-template`, `-o` and `--in-place`

By default a program answers JSON: values in, values out. The other kind of answer is the
**workbook itself**, which suits whoever needs the file and not the numbers: a date comes out as a
date, a percentage as a percentage, in the formats the cells already have, and `sheetrs query` reads
the same values from it if they are needed.

```bash
sheetc build -i model.xlsx --embed-template -o model                            # the workbook is carried
./model --output answer.xlsx --request @request.json                            # {"inputs": {"Data!A1": 5}}
./model --request @request.json                                                 # the same calculation, as JSON
./model --output answer.xlsx --json --request @request.json                     # the workbook and the JSON, one calculation
sheetc calculate -i model.xlsx -o answer.xlsx --request @request.json           # in memory
sheetc calculate -i model.xlsx -o recalculated.xlsx                          # nothing given: every formula again
sheetc calculate -i model.xlsx --in-place                                    # replaces model.xlsx (atomically)
```

* The workbook travels inside the executable (it is made from `sheetc` itself, which knows how to
  write workbooks) and there is no other template: the cells are written by position, so a workbook
  of another layout could not be checked. An executable built without `--embed-template` is the usual JSON one (default stub `sheetrs-rt`), and does not write workbooks.
* The executable answers **the workbook, JSON or both**, decided when it runs (supporting JSON adds a
  negligible amount of code to it and costs no time when it is not used, so it is not a build option):
  * `--output OUT` writes a copy of the workbook, in its own format, with the results of the
    outputs that are formulas and the inputs given (`--request`) in their
    cells, and prints nothing; the request holds only `inputs`.
  * Without `--output` the answer is JSON, as in any executable (`{"valueRanges": …}`: the
    declared outputs, or the `--range` or `ranges` asked for; `--pretty`, `--serve`).
  * `--output OUT --json` gives both from **one calculation**: the workbook in `OUT` and the JSON
    on the standard output (`--range` and `--pretty` apply to it; `--serve`, which answers many
    requests, does not write a workbook).
  * The metadata (`--meta`, `--info`) is answered as in any executable, and says which answers the
    executable has. Errors are written in the workbook as the format has them (`#CIRC!` is
    `Err:522` in an ODS; an error LibreOffice gives with its own code, `Err:502`, is written so in
    an ODS and as `#VALUE!` in an XLSX).
* **An embedded executable is any executable with the workbook inside**: inputs and outputs are
  declared or deduced exactly as in the usual one (see above). With nothing declared the outputs are
  the formula cells (`--meta` lists the rectangles that hold them; with more than 4 096 separate
  rectangles, the used area of the formulas of each sheet) and the inputs the constants a path
  starts from; a request's `inputs` key is such a range or a defined name inside them
  (`{"inputs": {"Data!B2:C3": [[1, 2], [3, 4]], "Rate": 0.05}`). What the settings (sheet or file)
  declare narrows them as in any executable: the JSON gives the declared outputs (and `--range` is
  checked against them), only the declared inputs are accepted (what does not depend on them is
  calculated once, at build time), and `--meta` lists them.
* **The workbook answer writes what the program computed**: the inputs given in their cells and the
  outputs that are formulas. The formulas outside the program keep the result the template stored,
  so with outputs declared the answer refreshes only those; to refresh every formula, declare no
  outputs.
* **Recalculating a workbook** is `sheetc calculate -i FILE -o OUT` with nothing declared and
  nothing given: every formula is calculated again. `-o` never
  overwrites the input; `--in-place` replaces it through a temporary file and a rename, so a failure
  leaves the original. The executable has `-o` but no `--in-place`.
* The cells of an array formula other than its first keep what the file cached.

### WebAssembly target

`sheetc build -i FILE --target wasm -o book.wasm` writes a **WebAssembly module** of the workbook instead
of an executable: the program of the workbook added to a runtime module that holds the engine. The
module needs no JavaScript glue and no imports, so a browser or Node (the hosts the tests run it in)
instantiates it and calls its exports (`sheetrs_calculate`, `sheetrs_write`, `sheetrs_meta`…, described in
[sheetrs-wasm](../sheetrs-wasm/README.md#the-module-of-a-workbook)). It takes the same options as a
native build, including `--embed-template` (the module then also writes the
workbook with the answer).

The runtime is a *base module* that `sheetc` completes, the way a native executable is the
`sheetrs-rt` runtime plus a payload. There are two: `sheetrs-rt.wasm` (the engine, for a module that
answers JSON) and `sheetrs-rt-template.wasm` (also the readers and writers of XLSX and ODS, for a
module that writes a workbook). Build them with `scripts/build-wasm-rt.sh`, which puts them next to
`sheetc`; the release archives carry them. `sheetc` looks for them there, or where `--runtime PATH`
says, and **never falls back** to another file: a missing base is an error that names it.

```bash
scripts/build-wasm-rt.sh                                  # target/dist/sheetrs-rt{,-template}.wasm
sheetc build -i tests/assets/loan.xlsx --target wasm -o loan.wasm
node -e '…instantiate loan.wasm, call sheetrs_calculate…'
```

A module cannot read its own bytes, so the base has a slot (a marker, an address and a length) that
the compiler fills in: the payload goes in as one more data segment at the end of the memory of the
base, which is enlarged to hold it (`sheetrs_compiler::wasm_target`). The sizes of the bases and of
compiled workbooks, and the time of the module against the native executable, are measured in
[docs/performance.md](../docs/performance.md#webassembly), the only place where they are kept. The limit of the depth of chains that read one another at
run time is that of the call stack of the host, as for the tools in the WebAssembly package.

### What the compiler checks

All of these fail the build, reporting every offender at once:

* An input holding a **formula** (`input `rate` (Loan!B2) holds formulas in B2`):
  an input must be a constant.
* An output range with **no value and no formula** in it.
* The settings-sheet rules above (duplicates, unknown ranges, collisions).
* **Macros** (VBA, Basic, scripts): they are never run, so what depends on them
  would differ. This is *strict* mode, the default; `--no-strict` (or
  `strict = false`) compiles anyway and reports a warning, and `--strict` undoes
  a `strict = false` of the file.

A declared input is optional at run time: the value the workbook holds in that
cell is its default (a blank cell defaults to blank), and a request overrides it
only if it names it (`null` sets a blank).

### Checking a real workbook: `sheetc verify`

A workbook saved by Excel or LibreOffice holds, beside each formula, the value
the application computed. `sheetc verify -i FILE [--sheet NAME]...` recomputes the
formulas of the named sheets (all by default) with the inputs the file holds,
that is, its defaults, and compares with the cached values:

```text
sheet        formulas  match  differ  volatile  unparsed
Outputs          1204   1204        0         0         0
differences by kind: number: 2, text: 1
  Outputs!C14  [number] expected 1200000 got 1199999.94
      =ROUND(B14*0.19,0)
```

Numbers match within `--tolerance` (relative, default 1e-9) or `--abs-tolerance`
(default 1e-6, the noise of cancelling sums); text, booleans and error kinds must
be identical. `volatile` formulas (`NOW`, `TODAY`, `RAND`) are not compared, and
`unparsed` ones (syntax the compiler keeps as the cached value, like external
workbook references) cannot differ by construction. The exit code is 1 when
anything differs, so it can run in CI. Build options (`--locale`, `--config`…)
apply: a model written for `es-CO` needs `--locale es-CO`. Keep real workbooks
outside the repository.

`sheetc inspect -i FILE` counts the functions a workbook uses, most used first, with
the cumulative share and whether the engine has each one: the way to see which
functions matter for a given model.

### The request and the answer

The executable (and `sheetc calculate`) takes one JSON object, all of it optional, and
reads ranges the way the batch read of a spreadsheet API does (`ranges`,
`valueRanges`, `majorDimension` as in Google Sheets):

```json
{ "inputs": { "principal": 250000, "rate": 0.045 },
  "ranges": ["payment", "Total", "Loan!B5", "Loan!A11:B12", "Loan!C:C"],
  "majorDimension": "ROWS" }
```

* `inputs`: values for the inputs, by name, alias or range. Only inputs can be
  set, at every level. Left out, an input keeps the value stored in the file;
  unknown names are rejected.
* **Without `ranges`**, the answer is the outputs declared when the program was
  built, one entry each, named: `{"valueRanges": [{"name": "payment", "range": "Loan!B4",
  "values": [[1073.64]]}, {"name": "interest", "range": "Loan!B6", "values": [[186511.57]]}]}`.
* **With `ranges`**, the answer is exactly those ranges, in request order,
  repeats allowed, and not the declared outputs:
  `{"valueRanges": [{"name": "payment", "range": "Loan!B4", "values": [[1073.64]]}, ...]}`.
  `values` is always an array of rows (`"COLUMNS"`: of columns), even for one
  cell, rectangular, with `null` for blank cells (Google trims trailing blanks).
  `range` is the canonical text; `name` appears when it was asked by a name (and for every
  declared output).
  Open-ended ranges (`Loan!C:C`, `Loan!3:5`) end where the sheet's data ends.
* Each entry of `ranges` is a declared output's name, a defined name (levels
  `named` and `any`) or a range with its sheet (`'My Sheet'!B2:C9`; a bare
  address means the first sheet). Values are the engine's: number, string,
  boolean, `null`, or `{"error": "#DIV/0!"}` (also accepted on input); dates are
  serial numbers of the file's date system. Number formats are never applied.
* `--range RANGE` (repeatable) adds to `ranges` from the command line.
* `--serve` keeps the process running: it reads one request per line on standard
  input and writes one answer per line, until the input ends. Every line is
  evaluated from scratch (nothing carries over between lines); what it saves is
  starting the process for each request. A bad request gets
  `{"error_message": "..."}` and the loop goes on; blank lines are skipped.

### Which ranges can be asked for

What may be asked is decided when the program is built, by what is declared. Without `ranges` the
answer is the outputs as value ranges (the declared ones, or the formula cells as rectangles). With
`ranges`, which kind of answer it is decides what a range outside the outputs does:

| Declared outputs (settings sheet or `settings_file`) | `ranges` outside them | Where |
| --- | --- | --- |
| yes, **compiled program** (JSON answer: `sheetc calculate` without `-o`, an executable run without `--output`, `BookCompile::calculate` with `Mode::Values`, a WebAssembly module from `compile`) | **an error** that lists the declared outputs (`sheetrs_engine::Error::OutsideOutputs`, `declared: Some(["payment (Loan!B4)", ...])`): `ranges` is a selection among the declared outputs (an output's name, a defined name that lies inside them, or a range inside one) | the program has pruned the workbook to what the outputs read, so it cannot answer more |
| yes, **embedded mode** (the workbook is its own template: `sheetc calculate -o OUT --json`, `--in-place --json`, an embedded executable with `--output OUT --json`, `BookCompile::calculate` with `Mode::Embedded`, and the `calculate` of `sheetrs edit`, Python and WebAssembly, which are all this) | **allowed**: it is answered from the workbook as it was written (the inputs given are in it), which the engine calculates again with that range as its output, so a formula the outputs did not include has the value its formula gives and a constant the one it holds; the answer keeps the order of the request | the workbook is there to read. A range that is no range or name of the workbook, or whose cells are all empty, is still an error |
| no | the outputs are the formula cells (as the rectangles that hold them): in the compiled program `ranges` must lie inside them and a constant is an error (with no declared outputs there is no list to give); in the embedded mode the rest of the workbook is read as above | |

The fallback of the embedded mode is `sheetrs_compiler::memory::fill_workbook` (the one function that
`sheetc`, `BookCompile::calculate` and `sheetrs-edit` call), made from the options of the program (locale,
features) and the clock and seed of the run. It needs the compiler, so the WebAssembly base module that
writes workbooks (`sheetrs-rt-template.wasm`), which has the engine and not the compiler, answers as a compiled
program does: a module built with `--embed-template` refuses a range outside the declared outputs.

### Logs

The commands of `sheetc` write what they report besides their answer on the standard error, as
diagnostics (a level, a stable code, a message and its context fields; see `sheetrs-protocol`):
warnings and errors as text by default (`warning: …`), `-q` for the errors only, `-v` for what the
command did (`wrote OUT`, the summary of a build) and `-vv` for the detail, and `--log-format json` for a
JSON record per line (`{"level", "code", "message", "context"}`). The executables a build makes
answer in JSON and have none of these flags.

### Execution options

**The document decides** date system (`date1904` / `table:null-date`: one of the three of LibreOffice Calc,
1899-12-30, 1900-01-01 or 1904-01-01; any other null date is an error that names it),
iterative calculation (`calcPr` / `table:iteration`), search settings (regular
expressions, wildcards, whole cell, case sensitivity: ODS). There is no flag or
configuration key for them in `build`, `calculate` or compiled executables, and a flag
given anyway is an error: the engine is faithful to what the file says. What the file
does not say is LibreOffice's behavior. (`evaluate`, which has no file, accepts
`--date-system 1899|1900|1904` and `--iterate[=N[,EPS]]`; so does the library for models
built in memory.)

**The build configuration decides what a file cannot say.** Every option is a
flag of `build`/`calculate` and a key of `sheetc.toml`; the command line wins over the
file, and the file over the defaults. The file is `--config PATH`, else
`sheetc.toml` in the workbook's directory (only there: a build does not depend on
where it is started). Unknown keys are errors.

```toml
settings = "settings"   # --settings
# settings_file = "model.json"   # --settings-file: the same declarations as JSON
strict   = true         # --strict / --no-strict: refuse (or only warn about) macros
now      = 45292.5      # --now SERIAL: fix NOW()/TODAY() (default: the clock, at each run)
seed     = 42           # --seed N: fix RAND (default: seeded from the clock)
enable_features = ["LONGTEXT"]   # --enable-features L: compatibility features (default: none)

[locale]                # the user's regional settings; --locale TAG sets `tag`
tag = "es-CO"           # default: the locale of the workbook (see Locale); en-US, en-GB, es-ES,
                        # es-CO, es-MX and de-DE need nothing else
# A region the table lacks (es-BO) states these three; a language it lacks (sw) also
# months (12), days (7, Sunday first), true_text, false_text, currency:
# decimal = "."
# group = ","
# date_order = "dmy"    # mdy | dmy | ymd
# Optional everywhere, to change what the locale has:
two_digit_year_start = 1930
# date_patterns = ["D/M/Y", "D/M"], months_abbr, days_abbr, am, pm, currency_after, currency_parens
```

**Compatibility features** (`enable_features`, `--enable-features`) change how a formula is read
when the program is built. Every feature is off unless the list names it (a name that is not known
is an error): each one differs from LibreOffice 26.8, so the user decides. `LONGTEXT` joins the pieces of
`_xlfn._LONGTEXT("…","…")` into one text literal: LibreOffice 26.2 wrote a literal of more than 255
characters that way, and LibreOffice 26.8 reads the function as `#NAME?`. Only constant text pieces
are joined; any other argument leaves the function unknown, as without the feature.

**Locale** is fixed when the program is built: folding already used it, so a run
cannot change it (`--locale` on an executable is an error). The locale is a property of the workbook
(`sheetrs_io::Book::with_locale`, `Book(..., locale=...)` in Python, `new Book(bytes, locale)` in
WebAssembly, `--locale` here), and a build takes the first of these that is set: the `--locale` flag or the
`tag` of `[locale]` (or the `locale` option of a call in the bindings), the locale given to the workbook,
the locale of the system, for an XLSX and an ODS alike (`LC_ALL`, `LC_CTYPE`, then `LANG`, as LibreOffice
reads them on Linux, at the time of the build; `en-US` when there is none), so a program built from the
same workbook with the same configuration is the same on every machine only when the configuration
states the locale. A system locale that the table lacks is `en-US` with a diagnostic, as for LibreOffice; a
locale asked for that it lacks is an error. `sheetc evaluate` has no workbook: `en` unless `--locale` says another. Six locales are
supported: `en-US`, `en-GB`, `es-ES`, `es-CO`, `es-MX` and `de-DE` (`en`, `es` and `de` mean the first of
each), the ones the engine reproduces row by row against LibreOffice (`de-DE` but for the keywords of a
format code, see FIDELITY.md); their data is generated from LibreOffice
(`scripts/gen_locale_table.py`). `[locale]` is the user's own regional setting, like the one a user
defines in LibreOffice: it changes any field of a supported locale, and defines a region or language
the table lacks from what it states (the user's responsibility). A locale that is neither in the table
nor stated stops the build. Dates written as text, `TEXT`, `DOLLAR`,
number text and text ordering follow the locale (see
[docs/FIDELITY.md](docs/FIDELITY.md#locale-support) for what is checked against LibreOffice and
what differs). The language an ODS stores in its default cell style is informational: LibreOffice calculates with
the regional setting of the machine (see *Locale* in FIDELITY.md for the two components of its locale).

The `[locale]` keys are the application-wide regional settings that Excel and
LibreOffice let a user change outside any document (separators, two-digit-year
window, list and currency conventions). Other application-wide settings can be
added to the file as they are verified; per-document ones stay the document's.

## Function support

The goal is the whole catalogue of LibreOffice; [docs/COVERAGE.md](docs/COVERAGE.md), generated by
`scripts/function_coverage.py` from the engine's registry and the catalogue LibreOffice reports, gives the
count and lists what is missing. Among them are aggregates and conditional aggregates (`SUM`…`SUMIFS`, `COUNTIFS`, `MAXIFS`,
`SUMPRODUCT`), math and rounding (spreadsheet-exact `ROUND`), logic and special
forms (`IF`, `IFS`, `IFERROR`, `SWITCH`, `CHOOSE`, lazy), text (`TEXT` with
number/date format codes, `TEXTJOIN`, `SEARCH` with wildcards…), lookup
(`VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT`),
dates (`DATE`, `EDATE`, `EOMONTH`, `DATEDIF`, `WEEKDAY`…) and finance (`PMT`,
`FV`, `PV`, `NPV`, `IRR`). Functions are in OpenDocument's namespaces (`core::SUM`, `microsoft::IFS`); the names each format gives them are in [docs/FUNCTION_NAMES.md](docs/FUNCTION_NAMES.md). A function of LibreOffice's catalogue that is not implemented (`docs/COVERAGE.md` lists them) stops
the build, naming it; a name LibreOffice does not know either is `#NAME?`, as there. `inspect` lists what
a workbook needs without stopping. See [docs/FIDELITY.md](docs/FIDELITY.md#what-stops-a-build) for
what else stops a build or warns.

Correctness is checked against LibreOffice: `tests/semantics/oracle.rs` recomputes the
formulas from `tests/assets/semantics.{xlsx,ods}` and compares them with the
values LibreOffice cached (see `scripts/gen_semantics_fixture.py`).

## Limitations

See [docs/DESIGN.md](docs/DESIGN.md#known-limitations) for the full list. The
ones that matter first: no external-workbook references (cached values are
kept), no dynamic-array spilling (array results take their top-left value),
no `LET`/`LAMBDA`, and run-time `INDIRECT`/`OFFSET` keep the whole workbook
reachable.

## Roadmap

1. **Calculation engine** (this release): IR, parser, graph ordering, folding,
   pruning, iteration, locale, JSON I/O, executables.
2. **Data validation**: read validations from the files and compile them into
   input checks (list, range, whole/decimal, custom formula) that run before
   evaluation and report which input broke which rule.
3. **Spreadsheet as source**: embed programs in Rust crates
   (`include_bytes!` + typed accessors generated from the declared bindings),
   native code generation for the static part and an HTTP wrapper around `--serve`. (The
   WebAssembly target and the Python bindings exist already: see above and
   [sheetrs-py](../sheetrs-py/README.md).)

Details of each step are in the design document.

## Related work

Other spreadsheet calculation engines written in Rust, consulted for ideas and for doubts about a
function (LibreOffice stays the reference here):

* [IronCalc](https://github.com/ironcalc/IronCalc) — a Rust engine and library (with Python
  bindings) that reads and writes XLSX and documents Excel compatibility of its formulas.
* [GridCraft](https://github.com/storytold/gridcraft) — a Rust spreadsheet application whose
  calculation layer is split into crates (formula parser, functions, model, XLSX I/O) and whose
  CLI can check its recalculation against the values cached in a file; it uses XLSX as its native format.

`sheetc` differs in what it produces (a standalone executable with declared inputs and outputs)
and in its reference (LibreOffice 26.8 for both XLSX and ODS).

## Development

```bash
cargo test -p sheetrs-compiler -p sheetrs-engine           # unit, integration, LibreOffice oracle, CLI
cargo clippy --all-targets -- -D warnings
cargo fmt --check
cargo run --release --example stress        # scale check
```

`sheetrs-compiler` lives in the SheetRS workspace and uses its readers
(`sheetrs-io`), the formula AST (`sheetrs-formula`) and the engine
(`sheetrs-engine`) as path dependencies. Cargo features:
`load` (default) enables the file frontend (`sheetrs-io` and `sheetrs-answer`, which writes the answer into the
workbook) and the `sheetc` binary. The compiler does not depend on `sheetrs-edit`: the editor depends on the compiler. The engine
is its own crate, `sheetrs-engine`, whose binary `sheetrs-rt` is the runtime stub embedded into
compiled executables).
