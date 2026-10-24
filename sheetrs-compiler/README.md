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
* **Faithful to the source application.** Behavior that differs between Excel
  and LibreOffice (booleans as numbers, `0^0`, 1900 leap-year quirk, number
  formatting…) follows the file's dialect; every case is listed and tested in
  [docs/FIDELITY.md](docs/FIDELITY.md).
* **Execution is configurable where the file is silent:** the document decides
  the date system, iterative calculation, search settings and dialect; the build
  configuration (`sheetc.toml` and flags) decides the locale (any BCP-47 tag,
  fixed at build time), the clock, the seed and how much of the workbook stays
  queryable.

Architecture and decisions: [docs/DESIGN.md](docs/DESIGN.md). Differences between
Excel and LibreOffice: [docs/FIDELITY.md](docs/FIDELITY.md).

## Quick start

Requires Rust ≥ 1.99 (`rust-toolchain.toml` pins it).

```bash
cargo build --profile dist -p sheetrs-compiler -p sheetrs-engine   # target/dist/{sheetc, sheetrs-rt}

# The workbook declares its inputs and outputs in a `settings` sheet
# (see below); compiling needs nothing else.
sheetc build tests/assets/loan.xlsx -o loan
# cells: 1493 total, 1090 reachable, 1493 kept | formulas: 1086 total ...

echo '{"Loan!B1": 250000, "Loan!B2": 0.045, "Loan!B3": 15}' | ./loan
# {"Loan!B4":1912.4832220336577,"Loan!B6":94246.9799660584,"Loan!B8":...}

./loan --meta                  # the settings sheet: inputs, outputs, custom metadata
./loan --info                  # technical description: shapes, cycles, diagnostics
./loan --serve                 # one JSON object per line in, one per line out
```

The same works from `.ods`, and without producing a file:

```bash
sheetc run tests/assets/loan.ods <<< '{"Loan!B1": 1000}'
sheetc calc '=TEXT(1234.5,"#,##0.00")' --locale es          # "1.234,50"
sheetc inspect model.xlsx      # sizes, unsupported functions, cycles, run-time refs
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
| `Loan!B1` | | Principal | `input` |
| `Loan!B4` | | Monthly payment | `output` |
| `version` | `1.0` | Model version | `info` |

* `meta_type` `input` or `output`: `key` is a cell or range (`Sheet1!A1`,
  `'My Sheet'!B2:C9`, an unqualified address on the first sheet, or a defined
  name). That is all the engine knows: JSON keys are the **canonical range text**
  (`Sheet1!B2:C9`, sheet quoted when needed). `value` and `description` are
  documentation.
* Any other `meta_type` is free metadata (`key`, `value`, `description`).
* The executable prints every row with `--meta`, inputs and outputs included.
* There is no other way to declare inputs and outputs: no command-line list, no
  names. The sheet is never part of the computation's outputs or queries.
* Rows are validated at build time (unknown range, duplicate, missing
  `meta_type`) and a sheet with no `output` row is an error unless
  `outputs = "any"`. Without a settings sheet, every formula range is an output
  and nothing is an input.

### How much stays queryable (`--outputs`)

| Level | Caller may read | Kept in the program |
| ----- | --------------- | ------------------- |
| `declared` (default) | the declared outputs | the minimum: only what feeds them |
| `named` | the declared outputs and the workbook's defined names | also what feeds the named ranges |
| `any` | any range, like the range API of a spreadsheet application (not the settings sheet) | every cell; formulas that depend on no input are still folded |

Ranges beyond the declared outputs are requested when running:
`./loan --query 'Loan!B5' --query 'Loan!A11:B12'`, or `"$query": [...]` in the
JSON object (also per line with `--serve`). A range the level does not allow is
an error naming the level. Inputs stay strict at every level: only the declared
ones can be set.

JSON mapping: a single cell is a scalar, a range is an array of rows (a flat
array is accepted for one-row/one-column ranges). Numbers, strings, booleans and
`null` (blank) map to themselves; an error is `{"error": "#DIV/0!"}` (also
accepted on input). Dates are spreadsheet serial numbers of the file's date
system. An input left out of the JSON keeps the value stored in the file;
unknown keys are rejected.

### Execution options

**The document decides** date system (`date1904` / `table:null-date`),
iterative calculation (`calcPr` / `table:iteration`), search settings (regular
expressions, wildcards, whole cell, case sensitivity: ODS) and dialect (the
file extension). There is no flag or configuration key for them in `build`,
`run` or compiled executables, and a flag given anyway is an error: the engine
is faithful to what the file says. (`calc`, which has no file, accepts
`--date-system 1899|1900|1904`, `--dialect` and `--iterate[=N[,EPS]]`; so does
the library for models built in memory.)

**The build configuration decides what a file cannot say.** Every option is a
flag of `build`/`run` and a key of `sheetc.toml`; the command line wins over the
file, and the file over the defaults. The file is `--config PATH`, else
`sheetc.toml` in the workbook's directory (only there: a build does not depend on
where it is started). Unknown keys are errors.

```toml
settings = "settings"   # --settings
outputs  = "declared"   # --outputs: declared | named | any
fold     = true         # --no-fold
now      = 45292.5      # --now SERIAL: fix NOW()/TODAY() (default: the clock, at each run)
seed     = 42           # --seed N: fix RAND (default: seeded from the clock)

[locale]                # the user's regional settings; --locale TAG sets `tag`
tag = "es-CO"           # any BCP-47 tag; default en
decimal = ","           # each of these overrides what the tag gives
group = "."
true_text = "VERDADERO"
false_text = "FALSO"
currency = "$"
two_digit_year_start = 1930
date_order = "dmy"      # mdy | dmy | ymd
```

**Locale** is fixed when the program is built: folding already used it, so a run
cannot change it (`--locale` on an executable is an error). Any locale tag is
accepted. Text ordering and equality use the ICU collator for the tag. For `en`
and `es` the separators, names and date order are the hand-checked ones verified
against Excel and LibreOffice. For any other tag they come from CLDR through
ICU4X (separators, month and day names, short-date order) and the build lists
what is **unverified**: boolean text and currency keep the English defaults and
native digits are not parsed, unless the `[locale]` table overrides them. The
document's *language* is informational and does not change evaluation
(LibreOffice evaluates an `es-ES` document with the system locale).

The `[locale]` keys are the application-wide regional settings that Excel and
LibreOffice let a user change outside any document (separators, two-digit-year
window, list and currency conventions). Other application-wide settings can be
added to the file as they are verified; per-document ones stay the document's.

## Function support

About 130 functions, grouped in namespaces (`core::`, `ods::`, `xlsx::`):
aggregates and conditional aggregates (`SUM`…`SUMIFS`, `COUNTIFS`, `MAXIFS`,
`SUMPRODUCT`), math and rounding (spreadsheet-exact `ROUND`), logic and special
forms (`IF`, `IFS`, `IFERROR`, `SWITCH`, `CHOOSE`, lazy), text (`TEXT` with
number/date format codes, `TEXTJOIN`, `SEARCH` with wildcards…), lookup
(`VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT`),
dates (`DATE`, `EDATE`, `EOMONTH`, `DATEDIF`, `WEEKDAY`…) and finance (`PMT`,
`FV`, `PV`, `NPV`, `IRR`). Dialect-only functions live in their namespace
(`ods::EASTERSUNDAY`). Unsupported functions evaluate to `#NAME?` and are listed
by `build`/`inspect`.

Correctness is checked against LibreOffice: `tests/oracle.rs` recomputes ~200
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

## Development

```bash
cargo test -p sheetrs-compiler -p sheetrs-engine           # unit, integration, LibreOffice oracle, CLI
cargo clippy --all-targets -- -D warnings
cargo fmt --check
cargo run --release --example stress        # scale check
```

`sheetrs-compiler` lives in the SheetRS workspace and uses its readers
(`sheetrs`), the formula AST (`sheetrs-formula`) and the engine
(`sheetrs-engine`) as path dependencies. Cargo features:
`load` (default) enables the file frontend and the `sheetc` binary. The engine
is its own crate, `sheetrs-engine`, whose binary `sheetrs-rt` is the runtime stub embedded into
compiled executables).
