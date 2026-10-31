# sheetrs-compiler (`sheetc`)

Compile an XLSX or ODS spreadsheet into a standalone executable (or a library
call) that takes JSON in and returns JSON out.

```text
loan.xlsx ──┐                                ┌─ ./loan   (single file, ~4 MB, ~2 ms/run)
            ├─▶ IR ──▶ fold · prune · order ─┼─ program.json  (the IR, portable)
loan.ods  ──┘  (format independent)          └─ Program::execute() (library)
```

* **One IR, one engine.** Both formats are lowered to the same intermediate
  representation; there is no second execution engine to keep in sync.
* **Inputs and outputs are declared in the workbook**, in a settings sheet. Cells that do not depend on an input are
  evaluated at build time and disappear; cells that do not feed an output are
  dropped. What remains is the minimal computation.
* **Hybrid compile/run-time.** Static references are resolved and ordered at
  build time (topological order, cycles isolated). `INDIRECT`, `OFFSET` and
  similar functions resolve their targets lazily at run time, on the same
  engine.
* **Faithful to LibreOffice.** The reference of the compiler and of the runtime (and of the
  modifications of `sheetedit`) is **LibreOffice 26.8 on Linux**. The engine behaves as it does, for an
  XLSX and an ODS alike (booleans as numbers, `0^0`, the 1899-12-30 null date, number
  formatting…); the cases where that differs from Excel are listed and tested in
  [docs/FIDELITY.md](docs/FIDELITY.md).
* **Execution is configurable where the file is silent:** the document decides
  the date system, iterative calculation and search settings; the build
  configuration (`sheetc.toml` and flags) decides the locale (any BCP-47 tag,
  fixed at build time), the clock, the seed and how much of the workbook stays
  queryable.

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
sheetc build tests/assets/loan.xlsx -o loan
# cells: 1493 total, 1090 reachable, 1493 kept | formulas: 1086 total ...

echo '{"inputs": {"principal": 250000, "rate": 0.045, "years": 15}}' | ./loan
# {"outputs":{"payment":1912.4832220336577,"interest":94246.9799660584,"balance":...}}

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
sheetc run tests/assets/loan.ods <<< '{"inputs": {"principal": 1000}}'
sheetc calc '=TEXT(1234.5,"#,##0.00")' --locale es          # "1.234,50"
sheetc inspect model.xlsx      # sizes, unsupported functions, cycles, run-time refs (never stops)
```

### Commands

| Command | Purpose |
| ------- | ------- |
| `build FILE [-o OUT]` | Compile to an executable. `--emit-ir PATH` also writes the IR as JSON. |
| `run FILE\|PROGRAM.json` | Compile in memory (or load an IR file) and run once. |
| `calc "=FORMULA"` | Evaluate a single formula. |
| `inspect FILE` | Report what the compiler sees. |

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
  or range it stands for. `value` starts with the sheet (`Loan!B1`,
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
  inputs and outputs the value shown is the canonical range) and, only when the
  build has `outputs = "named"` or `"any"`, the defined names it can read
  (`meta_type` `name`, the value is the definition).
* There is no other way to declare inputs and outputs: no command-line list.
  A key names one thing: it cannot repeat in the sheet (not as the same type nor
  as another: an input, an output or a field), and it cannot coincide with a
  sheet, a defined name or a table of the workbook. All of that fails the
  build. The sheet is never part of the computation's outputs or queries.
* Rows are validated at build time (unknown range, duplicate, missing
  `key`) and a sheet with no `output` row is an error unless
  `outputs = "any"`. Without a settings sheet, every formula range is an output
  (keyed by its range text) and nothing is an input.

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
the application computed. `sheetc verify FILE [--sheet NAME]...` recomputes the
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
apply: a model written for `es-CO` needs `--locale es-CO`. Keep real workbooks in
`tests/production/`, which git ignores.

`sheetc inspect FILE` counts the functions a workbook uses, most used first, with
the cumulative share and whether the engine has each one: the way to see which
functions matter for a given model.

### The request and the answer

The executable (and `sheetc run`) takes one JSON object, all of it optional, and
reads ranges the way the batch read of a spreadsheet API does (`ranges`,
`valueRanges`, `majorDimension` as in Google Sheets):

```json
{ "inputs": { "principal": 250000, "rate": 0.045 },
  "ranges": ["payment", "Total", "Loan!B5", "Loan!A11:B12", "Loan!C:C"],
  "majorDimension": "ROWS" }
```

* `inputs`: values for the declared inputs, by name. Only declared inputs can be
  set, at every level. Left out, an input keeps the value stored in the file;
  unknown names are rejected.
* **Without `ranges`**, the answer is the outputs declared when the program was
  built: `{"outputs": {"payment": 1073.64, "interest": 186511.57}}`, a scalar for
  a single cell and an array of rows for a range.
* **With `ranges`**, the answer is exactly those ranges, in request order,
  repeats allowed, and not the declared outputs:
  `{"valueRanges": [{"name": "payment", "range": "Loan!B4", "values": [[1073.64]]}, ...]}`.
  `values` is always an array of rows (`"COLUMNS"`: of columns), even for one
  cell, rectangular, with `null` for blank cells (Google trims trailing blanks).
  `range` is the canonical text; `name` appears when it was asked by a name.
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

What may be asked is decided when the program is built (`--outputs`):

| Level | `ranges` may name | Without `ranges` | Kept in the program |
| ----- | ----------------- | ---------------- | ------------------- |
| `declared` (default) | the declared outputs, or ranges inside them | the declared outputs | the minimum: only what feeds them |
| `named` | also the workbook's defined names (and tables), by name | the declared outputs | also what feeds the named ranges |
| `any` | any range, like the range API of a spreadsheet application (not the settings sheet), and the names | an error: `ranges` is required | every cell; formulas that depend on no input are still folded |

A range the level does not allow is an error naming the level.

### Execution options

**The document decides** date system (`date1904` / `table:null-date`),
iterative calculation (`calcPr` / `table:iteration`), search settings (regular
expressions, wildcards, whole cell, case sensitivity: ODS). There is no flag or
configuration key for them in `build`, `run` or compiled executables, and a flag
given anyway is an error: the engine is faithful to what the file says. What the file
does not say is LibreOffice's behavior. (`calc`, which has no file, accepts
`--date-system 1899|1900|1904|YYYY-MM-DD` and `--iterate[=N[,EPS]]`; so does the library for models
built in memory.)

**The build configuration decides what a file cannot say.** Every option is a
flag of `build`/`run` and a key of `sheetc.toml`; the command line wins over the
file, and the file over the defaults. The file is `--config PATH`, else
`sheetc.toml` in the workbook's directory (only there: a build does not depend on
where it is started). Unknown keys are errors.

```toml
settings = "settings"   # --settings
outputs  = "declared"   # --outputs: declared | named | any
fold     = true         # --no-fold
strict   = true         # --strict / --no-strict: refuse (or only warn about) macros
now      = 45292.5      # --now SERIAL: fix NOW()/TODAY() (default: the clock, at each run)
seed     = 42           # --seed N: fix RAND (default: seeded from the clock)
enable_features = ["LONGTEXT"]   # --enable-features L: compatibility features (default: none)

[locale]                # the user's regional settings; --locale TAG sets `tag`
tag = "es-CO"           # default en; en-US, en-GB, es-ES, es-CO, es-MX need nothing else
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
cannot change it (`--locale` on an executable is an error). Five locales are supported:
`en-US`, `en-GB`, `es-ES`, `es-CO` and `es-MX` (`en` and `es` mean the first of each), the ones the
engine reproduces row by row against LibreOffice; their data is generated from LibreOffice
(`scripts/gen_locale_table.py`). `[locale]` is the user's own regional setting, like the one a user
defines in LibreOffice: it changes any field of a supported locale, and defines a region or language
the table lacks from what it states (the user's responsibility). A locale that is neither in the table
nor stated stops the build. Dates written as text, `TEXT`, `DOLLAR`,
number text and text ordering follow the locale (see
[docs/FIDELITY.md](docs/FIDELITY.md#locale-support) for what is checked against LibreOffice and
what differs). The document's *language* is informational and does not change evaluation
(LibreOffice evaluates an `es-ES` document with the system locale).

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
   native code generation for the static part, WASM runtime, language bindings,
   and an HTTP wrapper around `--serve`.

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
(`sheetrs-core`), the formula AST (`sheetrs-formula`) and the engine
(`sheetrs-engine`) as path dependencies. Cargo features:
`load` (default) enables the file frontend and the `sheetc` binary. The engine
is its own crate, `sheetrs-engine`, whose binary `sheetrs-rt` is the runtime stub embedded into
compiled executables).
