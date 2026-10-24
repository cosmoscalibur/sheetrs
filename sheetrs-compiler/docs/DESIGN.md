# Design

Goal: turn a spreadsheet file into something that can be executed — an
executable today, a library, an application or a service later — with the
calculation engine first and data validation next.

The previous contents of this repository (a `pest` grammar targeting a
Python transpiler) were set aside: the strategy below needs a different shape.

## Pipeline

```text
          format-specific                     format-independent
 ┌──────────────────────────┐   ┌───────────────────────────────────────────────┐
 │ file ──sheetrs──▶ Model  │──▶│ parse ▶ resolve refs ▶ graph ▶ fold ▶ prune ▶ │──▶ Program
 │ (cells, formulas, names) │   │ (IR)     (OFFSET/INDIRECT)  (SCC)   (engine)    │    (IR + order)
 └──────────────────────────┘   └───────────────────────────────────────────────┘
                                                                   │
                           ┌───────────────────────┬───────────────┴───────────┐
                           ▼                       ▼                           ▼
                  executable (bundle)       program.json              Program::execute
```

The code is split in three crates: `sheetrs-formula` parses and prints formulas;
`sheetrs-engine` holds `value`, `addr`, `ir`, `options`, `dates`, `coerce`, `format`,
`matching`, `functions`, `engine`, `run`, `bundle` and the run-time `cli`;
`sheetrs-compiler` holds `loader`, `model`, `formula` (lowering), `compile` and
`graph`.

| Module | Role |
| ------ | ---- |
| `loader` (feature `load`) | `sheetrs` workbook → `Model`. The only code that knows XLSX vs ODS: dialect, defined-name syntax, error codes. |
| `model` | Format-independent workbook: cells, formula text, names, date system. |
| `formula` | Lowers the `sheetrs-formula` AST (lossless; parsing lives there) to IR: resolves sheets and names, picks function namespaces, folds literal ranges. |
| `ir` | `Expr`, `Program`, namespaced `FuncId`. Serializable. |
| `compile` | Cone of influence, SCC order, constant folding, slot assignment. |
| `graph` | CSR adjacency + iterative Tarjan SCC. |
| `engine` | Demand-driven evaluator; iterative calculation; array context. |
| `functions` | Registry (`core`/`xlsx`/`ods`) and ~130 functions. |
| `options`, `coerce`, `format`, `dates` | Date system, locale, iteration; conversions; `TEXT`; serial dates. |
| `run`, `bundle`, `cli` | JSON mapping, executable payload, command line. |

## One IR, one engine

The requirement is explicit: no double execution engine. Everything after the
loader sees only `Model`/`Program`. Cell values (`Value`), references
(`CellRef`/`Rect`), operators and function calls are the same whatever the
source. ODS-isms (`;` separators, `[.A1]` references, `of:=`,
`COM.MICROSOFT.` prefixes, `Err:522`) are resolved by `sheetrs` and the loader
or by the function-name resolution described next.

## Function namespaces

A function is identified by `(namespace, NAME)`, serialized as `ns::NAME`
(`core::SUM`, `ods::EASTERSUNDAY`):

* `core` — semantics shared by both formats (the vast majority).
* `xlsx`, `ods` — functions that exist in one format only, or behave
  differently there. Registering `ods::TEXT` next to `core::TEXT` is all it
  takes to give ODS its own format-code behavior.

The parser resolves unqualified names with the **dialect** of the source file:
dialect namespace first, then `core`. Format-specific prefixes select a
namespace explicitly: `_xlfn.`/`_xlws.` (XLSX extension functions) and
`COM.MICROSOFT.`/`ORG.OPENOFFICE.`/`ORG.LIBREOFFICE.` (ODS). The IR stores the
resolved `ns::NAME`, so a program built from an ODS file runs identically
anywhere, and a function missing from a namespace is a `#NAME?` plus a
diagnostic rather than silent fallback to the wrong semantics.

Operators and shared coercions cannot be namespaced, so they consult the
**dialect** (`EngineOptions.dialect`, derived from the file extension, forced
with `--dialect`): `0^0`, odd roots of negative numbers, booleans as numbers,
number-to-text format, direct text arguments to `SUM`. The full list, with the
evidence behind each case, is in [FIDELITY.md](FIDELITY.md).

## Hybrid compile / run time

Most references are static, and for them the compiler does the work once:
dependency order, dead-cell elimination, constant folding. Some are not
(`INDIRECT(A1)`, `OFFSET(x, n, 0)` with an input-dependent `n`). Those cannot
be ordered at build time, so the engine is **demand driven**:

* cells are evaluated in the precomputed order, so ordinary references find
  their operands ready;
* a reference to a cell that is not ready (only possible through dynamic
  references) evaluates it on the spot, memoizes it, and treats a cell found
  *in progress* as a circular reference;
* the compiler resolves whatever it can first: `INDIRECT("Sheet2!B3")`,
  `OFFSET(A1, 2, 0)` and `A1:B2` built from literal operands become plain
  references;
* if a dynamic reference survives in the cone of an output, the whole workbook
  becomes reachable (any cell may be read). That is conservative and correct;
  narrowing it from the shape of the argument (`"D"&n` ⇒ column D) is on the
  roadmap.

So the “JIT” is not machine code: it is lazy resolution on the same
interpreter, which keeps one engine and still removes every cost that static
analysis can remove. Native code generation for the static part is a later
step that can reuse the same IR.

## Inputs, outputs and what is optimized

The workbook's settings sheet (`input` and `output` rows) declares the cells or
ranges exchanged as JSON; `--outputs` decides how far a caller's choice of
outputs reaches (below). The compiler:

1. Starts from the outputs and walks dependencies backwards, **cutting at
   inputs** (an input cell's own formula is ignored). Everything else is dead.
2. Orders the cone with Tarjan SCC (dependencies first, cycles as groups).
3. Marks a cell *live* if it depends — transitively — on an input, a volatile
   function (`RAND`, `NOW`, `TODAY`) or a dynamic reference. All other cells
   are **folded**: evaluated at build time by the same engine and replaced by
   constants.
4. Keeps only live formulas plus the constants they (or the outputs) read.

The compiled loan example drops from 1086 formulas to 1 for `payment` (an
intermediate computation that depends on inputs stays; the 360-row schedule
nobody asked for is gone).

### Query levels

`declared` starts the walk from the declared outputs only. `named` adds the
workbook's defined names as further roots, `any` adds every formula of every
sheet (except the settings sheet) and keeps the workbook's constants. Folding
works identically at every level, so `any` still drops the formulas that depend
on no input, and the program grows to the size of the data (`stress`: 50 k folded
formulas keep 1 cell when declared, 100 001 when any range is queryable). A
runtime query is checked against the level (`Queryable` in the program): rectangles
inside a declared output, a named range, or any range outside the excluded
sheets.

### Who configures what

The date system, iteration, search settings and dialect are the document's and
are never overridable. What a file cannot say comes from the build configuration
(`sheetc.toml`, overridden by flags): the locale (fixed at build time, so
folding may use it; any BCP-47 tag, see FIDELITY.md), the clock, the seed,
`--no-fold`, the settings sheet's name and the query level. A run can only fix
`NOW()` and the seed. Because the locale cannot change after the build, the
compiler no longer needs to track which formulas depend on it: that analysis
existed only for the former `--runtime-locale` and was removed. Who decides what
is in [FIDELITY.md](FIDELITY.md#who-decides-what).

## Scale

Measured on a synthetic 49 MB XLSX (600 000 rows × 11 columns, 6.6 M cells, 600 000
formulas), release build, `cargo run --release -p sheetrs-compiler --example timing`:

| phase | time | peak memory |
| ----- | ---: | ----------: |
| read the file (`sheetrs`) | 12 s | 1.0 GB |
| convert to the model (consuming the workbook) | 5.5 s | 1.3 GB |
| compile, output depends on one input (everything else folded) | 4.8 s | 1.3 GB |
| compile, 2.4 M input cells and 300 000 live formulas | ~6 s | 1.3 GB |
| run (JSON in, JSON out) | 0.04 ms – 0.84 s | — |

Choices that make this possible: constants are never graph nodes (they stay in
the model and are read in place, also while folding); the graph holds only
formulas, inputs and shared range nodes; maps keyed by cell address use a fast
non-cryptographic hasher; small ranges are not cached while large ones are; the
loader consumes the workbook instead of copying it. An earlier version, which
copied every cell into the graph and into a temporary program, needed 69 s and
3.4 GB for the same file. Most of what remains is the reader itself (about 150
bytes per cell); a streaming reader that skips cells outside the cone is the next
step if this has to go lower.

### Cost of ICU collation

Text comparison goes through an ICU4X collator (~110 ns per comparison); ASCII
strings, nearly all cell text, are compared for equality without it. On the 49 MB
workbook, 600 000 text equalities cost ~0.26 s in total and 600 000 ordering
comparisons ~0.24 s (`--example bench_text`), a small part of a run that also
builds the arrays. Build time: the ICU and `fancy-regex` dependencies add ~20 s
to a clean release build, once (incremental builds are unaffected; changing the
compiler recompiles in ~10 s); the runtime stub grows by ~3 MB (3.9 MB stripped).
Collation affects only text *ordering and equality*, never display: the engine
computes cell values, and cell number formats are not read at all
(`TEXT`/`FIXED`, whose result is a string value, are the only users of format
codes).

## Dependency ordering: in-house, no graph library

Requirements: evaluation order, cycle detection (for iterative calculation),
huge node counts, one pass for both. A measured comparison (random dependency
graph, ~3 edges per node, build + order, release build; one-off benchmark, not
kept in the tree) gave:

| nodes | CSR + Tarjan (here) | petgraph `tarjan_scc` | petgraph `toposort` | ultragraph |
| ----: | ------------------: | --------------------: | ------------------: | ---------: |
| 100 k | 3.3 ms  | 10.9 ms | 14.0 ms | 2.4 s |
| 1 M   | 73 ms   | 177 ms  | 178 ms  | 31.8 s |

petgraph is a fine library and `tarjan_scc` returns the same condensation
order, but nothing needs it: the in-house version is ~100 lines, iterative (no
stack limit on 100k-long chains), 2.4× faster here, and returns components
already flattened in schedule order. It is therefore the only implementation and
petgraph is **not** a dependency, not even for tests: the unit test checks the
SCC partition and the dependencies-first order against plain BFS reachability on
random graphs, which needs no library. ultragraph (a causal-hypergraph library)
builds in quadratic time here and was rejected.

Range dependencies are modeled with shared *range nodes* (one per distinct
rectangle) so `SUM(A$1:A$10000)` used by 50 000 cells costs one node, not 500 M
edges.

## Reusing `sheetrs`

Yes, reuse it. Container parsing (zip, shared strings, shared formulas, ODS
table repetition, covered cells, defined names, error cells) is where
format-specific bugs live, and `sheetrs` already carries the fixes and the XLSX/ODS
test files. Re-implementing it would add a second reader to keep faithful for no
gain in the engine.

Reuse is conditional on the reader being faithful, so the contract is: **the
reader returns what the file stores; the engine owns every semantic decision.**
Auditing against LibreOffice (see FIDELITY.md) found one violation, fixed in the
`sheetrs` reader (same workspace): ODS dates were numbered from 1899-12-31, a convention of
the reader's own that is neither LibreOffice's nor Excel's (one lower from
1900-03-01 on), and `office:time-value` was returned as display text. Dates are
now days since the document's `table:null-date`, times a fraction of a day. No
calendar adjustment happens in the reader; the date system (`1899`, `1900`,
`1904`) is applied by the engine. The reader also exposes the calculation settings a document stores
(`Workbook::settings`: iteration, search mode, case sensitivity, two-digit-year
window, language) **as stored, with no defaults applied**, so the engine can tell
"not stated" from "stated as the default". Remaining reader-side items
(FIDELITY.md) are an upstream wish list: typed values, formula-string results
that look numeric, `#VALUE!` placeholders, defined-name scope.

## Inputs and outputs are ranges

The engine's interface is two lists of cells/ranges. No names, no schema:
`Program.inputs`/`outputs` hold rectangles and JSON uses their canonical text as
keys. Names, types, units and documentation belong to a later layer that maps
friendly names to ranges.

## Parser: hand-written, not `pest` (in `sheetrs-formula`)

The formula language is small and regular, and the hot path is "parse
hundreds of thousands of short formulas". The parser is a single pass over the
bytes (no token vector, no parse tree) producing IR directly, with Pratt
precedence (Excel's: unary minus binds tighter than `^`, `%` postfix, `:`
tightest). Specializations a generic grammar would not give: references and
`A:A`/`1:3` ranges recognized by address shape, sheet-qualified references with
quoting, defined names expanded in place, function prefixes mapped to
namespaces, arity checked against the registry at parse time, unknown syntax
(external references, structured references) degrading to the cached value with
a diagnostic.

## Execution parameters

`EngineOptions`: dialect, date system (1900 with the 1900-02-29 quirk, 1904, or
1899 = real calendar, the ODS default),
iteration (`max_iterations`, `max_change`; Excel defaults 100 / 0.001), locale
(decimal/group separators, boolean names, month and day names; `en` and `es`
built in; text order from ICU4X collators) plus fixed `now` and `rand` seed for determinism. The
source file's own setting (`date1904`) is the default. Locale affects text
conversion at run time (`TEXT`, `&`, `VALUE`, `FIXED`); function *names* are not
localized because files always store the canonical ones.

## Executable format

`sheetc build` copies a runtime stub — `sheetrs-rt` (~4 MB with the ICU collation data, engine only) or
the full binary — and appends the program as JSON plus a 16-byte trailer
(`XLSLANG1` + length). At start the binary looks for the trailer in its own
file; if present it runs the program, otherwise it is the compiler. No
toolchain is needed to build, the artifact is a single file, startup is ~2 ms.
Programs are plain data, so the same bytes can be embedded in a library or
loaded by `sheetc run program.json`.

## Testing

* Unit tests per module (parser precedence, dates, formats, rounding, graph).
* The graph is checked against BFS reachability on random graphs.
* **Oracle**: `scripts/gen_semantics_fixture.py` writes ~200 formulas over
  fixed data, has LibreOffice recalculate and save XLSX and ODS, and
  `tests/oracle.rs` requires the engine to reproduce those cached values from
  the formulas, per format. `tests/probe.rs` does the same for ~120 probes of
  Excel/LibreOffice differences in the `ods` dialect and pins the Excel side. Documented gaps (LibreOffice vs Excel semantics,
  reader limitations) are listed in the test with the reason.
* `tests/pipeline.rs`: pruning, folding, cycles, iteration, dynamic references,
  options, JSON mapping.
* `tests/cli.rs`: compiles `loan.xlsx` and `loan.ods` into executables and
  runs them, including `--serve` and error exit codes.
* `examples/stress.rs`: 100 k-formula chain compiles in ~0.12 s and runs in
  ~17 ms; 20 k `VLOOKUP`s over a 10 k-row table take ~1 s.

## Known limitations

* **External references** (`[1]Sheet!A1`) and structured references keep the
  value cached in the file. Defined names lose their sheet scope in XLSX when
  read through `sheetrs`.
* **Arrays**: operators over ranges use legacy implicit intersection outside
  array context (`SUMPRODUCT`, `INDEX`, `MATCH`, `XLOOKUP` and CSE formulas
  evaluate element-wise, scalar functions are lifted there). There is no
  spilling: a formula returning an array exposes its top-left value.
  `LET`, `LAMBDA`, `FILTER`, `SORT`, … are not implemented.
* **Anchored accumulations** such as `SUM($A$1:A5000)` down a column are
  quadratic in graph size and run time (5 000 rows: 1 s). Fix planned:
  segment-tree dependency nodes and prefix-sum caching.
* **Run-time `INDIRECT`/`OFFSET`** make the whole workbook reachable.
* **Dialect differences not yet modeled** are listed in FIDELITY.md (regex mode
  of ODS documents, Unicode case mapping, locale-bound text dates, …).
* `NOW`/`TODAY` use the system clock (pass `--now` on WASM).

## Roadmap

### Data validation

The IR already separates inputs from the rest of the program. Validation rules
(list, whole/decimal bounds, date ranges, text length, custom formula) become
predicates (`Expr`) attached to input bindings, evaluated before the run:

1. `sheetrs` exposes data validations (XLSX `<dataValidations>`, ODS
   `table:content-validations`) — upstream change.
2. `compile` lowers each rule to an IR predicate over the input cell and, for
   list rules, a constant set or a range reference (itself part of the
   program, so lists can come from the workbook).
3. `Program::execute` returns `{ "errors": [{ "input": "rate", "cell": "B2",
   "rule": "decimal between 0 and 1", "message": … }] }` instead of running on
   invalid input; the executable exits with a distinct code.

### Spreadsheet as the source of a library, application or service

1. **Library**: `Program::from_bytes(include_bytes!(…))` and typed accessors
   generated from the declared bindings (`build.rs` helper first, proc-macro
   later), so `loan.xlsx` is a Rust module with `LoanInput`/`LoanOutput`.
2. **Service**: `--serve` already gives a line protocol; wrap it in an HTTP
   server (one program instance per worker, programs are immutable and `Sync`).
3. **WASM**: the engine has no OS dependency besides the clock; ship
   `sheetrs-engine` as a WASM module for browsers and edge runtimes.
4. **Native code**: emit Rust (or Cranelift) for the folded, ordered static
   part and keep the interpreter for dynamic references — same IR, same
   semantics, AOT speed.
5. **Bindings**: Python/JS wrappers over `Program::execute`.

### Engine

* Segment-tree dependency nodes; hash index cache for exact `VLOOKUP`/`MATCH`.
* Narrow dynamic reachability from `INDIRECT`/`OFFSET` argument shapes.
* Incremental recalculation for `--serve` (re-evaluate only the cone of the
  changed inputs).
* More functions, driven by the unsupported list `inspect` prints for real
  workbooks.
