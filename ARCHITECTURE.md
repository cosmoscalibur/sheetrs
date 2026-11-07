# Architecture Documentation

The **SheetRS Suite** is organized as a Cargo workspace containing a shared core library and three specialized binary crates. This architecture promotes code reuse, consistency, and testability.

## Reference Behavior

The reference behavior of the compiler, the runtime and the modifications of `sheetedit` is **LibreOffice 26.8 on Linux**. What an operation gives where the formats or the applications differ is what LibreOffice gives for it (`sheetrs-compiler/docs/FIDELITY.md`).

## Workspace Structure

```
sheetrs/
├── sheetrs-io/       # Input of the suite: workbook model + XLSX/ODS readers (was sheetrs-core)
├── sheetrs-formula/    # Lossless formula AST (parse, print, rewrite references)
├── sheetrs-protocol/   # The contract: typed requests/responses of run, query, find, replace
├── sheetrs-linter/     # Rules, configuration, violations (lib) + `sheetlint`
├── sheetrs-stats/      # Workbook metrics (lib) + `sheetstats`
├── sheetrs-query/      # Reading and finding cells by range, sheet or workbook (lib) + `sheetquery`
├── sheetrs-edit/       # Workbook modification (lib) + `sheetedit`
├── sheetrs-engine/     # IR, function library, run-time evaluation, `sheetrs-rt`
├── sheetrs-compiler/   # Loader, dependency analysis, folding, `sheetc`
├── sheetrs-wasm/       # Bindings for the browser (wasm-bindgen)
├── sheetrs-py/         # Bindings for Python (PyO3, maturin)
└── Cargo.toml          # Workspace configuration
```

## Layering

The suite follows a hexagonal (ports and adapters) layout. A **tool crate** is an application
core: it owns the decisions about a workbook (the linter's rules, the statistics, the edit
operations, the compilation) and exposes them as a Rust library. Around it:

- **Driven side** (what a tool needs): the workbook model and the readers of `sheetrs-io`. The
  model (`Workbook`, `Cell`, `Grid`, `CellSink`, `WorkbookReader`) is the port; the XLSX and ODS
  parsers are its adapters. A tool never sees a file format.
- **Driving side** (who calls a tool): the executables (`sheetlint`, `sheetstats`, `sheetedit`,
  `sheetc`) and the bindings (`sheetrs-wasm`, `sheetrs-py`). Each is an adapter that translates
  arguments and renders results; none holds a decision. What several adapters need to present
  (violations grouped by sheet and cell, flat records) is part of the tool's library
  (`sheetrs_linter::report`), not copied into each adapter.

Dependencies point inward: adapters → tool libraries → `sheetrs-io` / `sheetrs-formula` /
`sheetrs-protocol`. Tool crates do not depend on each other, with one deliberate exception:
`sheetrs-edit` depends on `sheetrs-query`, because a replacement is a search followed by a
write, and the two must agree on which cells a search finds (its matcher and the text of a cell
that a search sees live in `sheetrs-query`; the writing stays in `sheetrs-edit`).

### What belongs in `sheetrs-protocol`

`sheetrs-protocol` is the **contract** of the suite as typed JSON, with no behaviour: it depends
on `sheetrs-formula` (the grammar of references), `serde` and `serde_json`, so a client (a web
page, a Python program, another service) can speak the protocol without any tool.

| Module | What it says |
| ------ | ------------ |
| `run` | The request a compiled program answers (`inputs`, `ranges`, `majorDimension`) and its response (`valueRanges` or `outputs`). `sheetrs-engine` parses its requests with it |
| `query` | `Query` (scopes: workbook, a sheet, a range; properties: `value`, `formula`, `number_format`, `display`; an optional `Find`; a locale), `CellRecord`, `QueryResponse` |
| `find` | `Find` (text, mode literal/regex/wildcards, match case, entire cells, search in value or formula, `include_formulas`) and `Replace` (a `Find` and the text that takes its place) |
| `reference` | A range as a caller writes it (`'My Sheet'!B2:C9`) |

The array of a cell is not a property: a query does not say whether a formula is an array
formula, and a selective read does not know it.

### What belongs in `sheetrs-io`

`sheetrs-io` (named `sheetrs-core` before 0.4.0) is the **input side** of the suite and a **shared kernel**, not "the public library": the public surface of the suite is
every library crate and every binding. A module is in the kernel when at least two tools need it
and it has no meaning outside the workbook model. The test, applied to the code that was in the
crate:

| Code | Used by | Decision |
| ---- | ------- | -------- |
| `reader` (model, `Grid`, `CellSink`, XLSX/ODS parsers), `ExcelError` | linter, stats, compiler, edit | kernel |
| `reader::parser_utils`, the XLSX shared-formula and part-name helpers the writer reuses | readers and edit | kernel |
| `rules`, `config`, `violation`, `report`, `Linter` | linter only (rules exist only to be run by it) | `sheetrs-linter` |
| `writer` (XLSX/ODS stream editing, `ValueScope`) | edit only | `sheetrs-edit` |

Moving code out of the kernel is how its boundary stays honest: a feature of one tool (a new rule,
a new edit operation) never changes `sheetrs-io`, and a second consumer of a tool's code is
served by that tool's library, not by growing the kernel. Reading is shared because four crates
read workbooks; writing is not, because one does.

## 1. sheetrs-io

The input of the suite and its shared kernel: the format-agnostic workbook model and the streaming readers that fill it.

- **`reader`**: Abstracts over file formats (XLSX, ODS). `Source` (a file or bytes, with its
  `Format`) is the input every tool takes: a reference that reads nothing until a tool does, and
  not a model of the workbook.
  - Uses `zip` and matching XML parsers (`quick-xml` or custom) for efficient data reading.
  - Custom stream-based parsing logic for both formats ensures minimal memory footprint.
  - `WorkbookReader` is the common interface of the formats. Its one required method for cells,
    `read_sheets_into(sink)`, hands each cell to a `CellSink` as it is parsed; every consumer of
    the readers gets its cells that way, so a change in how a format is read is made once.
  - `Grid<T>` (`reader/grid.rs`) is the one store of cells: a sorted list of columns and items per
    sheet with the offset where each row starts. The linter's `Sheet.cells` is a `Grid<Cell>`
    (filled by the default `GridSink`, with no intermediate map) and the compiler's model keeps a
    `Grid<ModelCell>` per sheet in its own `CellStore`. Reading in row order appends; a walk is
    in row order. A `Cell` holds the value, formula, array flag and number format, and not its own
    position: the position is the cell's place in the `Grid`, which iteration and `CellSink::cell`
    hand over with it. The cell is 56 bytes (the hash map's entry was 80 plus its table).
  - `Selection` (sheets, `Window`s of rows and columns, and the `Properties` wanted of each cell) is
    what a selective read takes: `read_selection_into(selection, sink)` does not parse a sheet that
    is not selected, stops the reading of a sheet after the last row a window reaches (an ODS ends
    with the last table selected), skips the cells outside the windows and does not build the
    properties that were not asked for. A cell reaches the sink as in a full read, so the two agree
    on what a cell is (the tests compare them over both formats). XLSX rows are expected in
    ascending order, as Excel requires.
  - `SheetState` (`Visible`, `Hidden`, `VeryHidden`) is what a sheet's `state` says; an ODS has the
    first two.
  - `ExcelError` is the error value a cell can hold (`#DIV/0!`, …).

## 1b. sheetrs-formula

A dependency-free crate with the formula **AST**: it keeps what analysis and
rewriting need (absolute `$` markers, the function name as written, the text of
numbers, explicit parentheses, omitted arguments) and nothing engine-specific.

- `parse`: single-pass Pratt parser (Excel syntax as the readers normalize it).
- `print`: canonical spelling; `parse(print(parse(s))) == parse(s)`.
- `parse_spanned`: keeps the source and the span of every reference, so
  whitespace and spelling are recoverable and `Formula::rewrite` applies an edit
  by replacing only the references that change.
- `edit`: what a formula becomes when rows or columns are inserted or deleted
  or a sheet is renamed or deleted (`#REF!` for what disappears, ranges that
  shrink and grow).
- `r1c1`: relative form of a formula as seen from its own cell.

Lint rules still use string heuristics (`rules/helpers.rs`); migrating them to
the AST is incremental and not part of this crate's introduction.

## 1c. sheetrs-engine and sheetrs-compiler

Two crates, split along the line that matters for reuse and build time:

- **`sheetrs-engine`** executes a compiled `Program` (the format-independent IR)
  with JSON in and JSON out: values, the function registry, dates, text conversion, collation, search patterns and the
  demand-driven evaluator. It knows nothing about XLSX or ODS and does not
  depend on the readers; its binary `sheetrs-rt` is the stub a compiled
  executable is made from.
- **`sheetrs-compiler`** turns a workbook into a `Program`: `loader` (the only code
  that knows the file format, on top of `sheetrs-io`) → `Model` → lowering of the
  `sheetrs-formula` AST to the IR → cone of influence, SCC order, constant
  folding. Its binary is `sheetc`.

The date system, iteration, search settings and reference syntax come from the document;
see [sheetrs-compiler/docs/DESIGN.md](sheetrs-compiler/docs/DESIGN.md) and
[sheetrs-compiler/docs/FIDELITY.md](sheetrs-compiler/docs/FIDELITY.md).

## 1d. Reading cells: selections and the query tool

The readers can read **what is asked and nothing else** (`reader::Selection`): a window of rows
and columns per sheet (or one for every sheet), and the properties wanted (`value`, `formula`,
`number_format`). A sheet that is not selected is not parsed, the reading of a sheet stops after the
last row a window reaches (an ODS ends with the last table selected), cells outside the windows
are skipped, and a property not asked for is not built. `sheetrs-query` is the tool on top of
it, with the same answer for an XLSX and an ODS wherever the two files say the same thing.

Two texts of a cell are different things and `sheetrs-query::text` keeps them apart: what a person
**sees** (`display`: the value written with its number format, in a locale) and what a **search**
sees (the text of the input line of the cell: `1234.5` for a cell that shows `1,234.50`, `25.6%`,
the date `03/15/2023`, and for a formula the notation of LibreOffice, `=SUM(A1,B1)`). Number
formats are read for every family of data style of an ODS (`reader::ods_styles`) and written as the
format code an XLSX holds for the same cell. A LibreOffice error that Excel does not have
(`Err:502`, `Err:522`) keeps its code in an ODS (`ExcelError::Libre`) and is `#VALUE!` in an XLSX,
as LibreOffice itself writes it.

## 2. Tools

Each tool is a crate `sheetrs-<tool>`: a library with the tool's logic and, for the three
original tools, the executable (`src/main.rs`) that adapts it to the command line.

- **`sheetlint`** (crate `sheetrs-linter`):
  - `rules`: each rule is a standalone struct implementing `WalkerRule`; the `WorkbookWalker`
    runs them in a single pass; they are registered in `rules::registry`, which is also the
    catalog of rules (the rules exist only inside the linter).
  - Categories: `ERR` (Errors), `CALC` (Calculations), `REF` (References), `INT` (Interruptions), `CPX` (Complexity), `VUL` (Vulnerability), `DATA` (Data Issues), `EXT` (External), `HID` (Hidden), `FILE` (Files), `VBA` (Macros).
  - `config`: TOML configuration. Hierarchical loading: Default -> Global Config -> Sheet
    Overrides; rules enabled or disabled by ID or category.
  - `violation`: `Violation`, `RuleId`, `Severity`, `ViolationScope`.
  - `report`: `group_by_scope` (book → sheet → cell, one entry per sheet) and `ViolationRecord`
    (flat form for the bindings).
  - `Linter` runs the rules over a `Workbook` it is given; it reads nothing.
    A `LinterConfig` is valid by construction (parsing validates: unknown rules, categories or
    parameters are an error), so no adapter validates it again.
  - `main.rs` and `formatter.rs`: arguments, colored text and streamed JSON.

- **`sheetstats`** (crate `sheetrs-stats`):
  - Focuses on aggregating metrics.
  - Uses the same `Workbook` model but collects different data points (counts, distributions).
  - `stats.rs`: `collect(&Source)` computes workbook and per-sheet metrics in one streaming pass
    and is what the bindings call; it is a function because there is nothing to configure.
    `main.rs` handles arguments and output.

- **`sheetedit`** (crate `sheetrs-edit`):
  - The writer stream-edits the archive with `zip` and `quick-xml` (removing sheets, converting
    formulas to values, deleting rows and columns) without fully rewriting the file, applying the
    same changes to each format.
  - `ValueScope` selects the cells to convert (a sheet, columns or rows) independently of the
    file format.
  - `Editor` (module `editor`, the API of the crate): made for a `Source`, it collects changes
    in order, each addressing the workbook as the previous ones left it (what running them one
    at a time gives), and `save` runs them. Rows and columns are typed (`RowRange`,
    `ColumnRange`, with `FromStr`/`Display`), so a deletion with the wrong kind of range cannot
    be built. Marking the output to be recalculated when opened is the editor's decision alone
    (`marks_recalculation`: it does when a change leaves stale cached results, and there is no
    option to ask for it or avoid it), applied once in the last pass.
  - `plan` (private): the planner behind `Editor`, pure logic with no I/O. `passes()` plans the
    sequence into the fewest passes over the file that give the same workbook: deletions of one
    sheet and axis share a pass, with the later positions translated through the earlier
    deletions; deletions of different sheets and on rows vs columns, name and sheet removals and
    the recalculation mark commute; a conversion after another change, and two deletions that
    touch (which the writer would take for one band), start a new pass. The `sheetedit` flags
    are all written against the input, in no order: applied as conversions, rows, columns,
    names, sheets they are one pass. A pass is applied by the writer in the order defined once
    in `StructureEdits::edits`. A test generates sequences and compares the plan with the
    operations run one at a time, in both formats. A reference that spans sheets would tie
    sheets together, and is refused (issue #75).
  - `writer` (private): `xlsx_writer.rs` and `ods_writer.rs` are the mechanism of a pass; neither
    the writer nor the plan is public, so the editor is the one way in.
  - The writer streams: each pass reads and writes every worksheet event by event, and the plan
    is a few bands and names, so memory does not depend on the size of the workbook.
  - `main.rs` handles arguments.

## 3. Bindings

The bindings are adapters: they convert arguments and results and hold no analysis logic. What
they present comes from the tool libraries (`sheetrs_linter::report`, `sheetrs_stats`,
`sheetrs_edit::Editor`), so every surface reports the same findings. The API of a binding is
organized by tool, as the crates are: the rules are a part of the linter, so they are offered
next to `lint`, not as a separate service.

- **`sheetrs-wasm`** (`cdylib`, wasm-bindgen): bytes in, structured values (`lint_workbook_records`,
  `lint_rules`, `workbook_stats`) or the human text report out. Packaged for npm by
  `scripts/package-wasm.sh` (`wasm-bindgen --target web`, template in `sheetrs-wasm/npm`).
- **`sheetrs-py`** (`cdylib`, PyO3 with `abi3`): the extension `sheetrs._sheetrs`, presented by
  the `python/sheetrs` package as the crates are laid out (type stubs): `sheetrs` (core: `open`,
  `Source`, `SheetrsError`), `sheetrs.linter` (`Linter`, `Violation`), `sheetrs.stats`
  (`collect`), `sheetrs.edit` (`Editor`). `Source` is the Python face of
  `sheetrs_io::reader::Source`: a reference to a workbook that reads nothing, with no behavior.
  The tool is the subject and takes it, as in Rust. There is no eager mode: a workbook is
  streamed, never loaded. The functions release the GIL while they work. It is built by maturin,
  which enables the `extension-module` feature; the crate has no Rust test harness (it would have
  to link libpython), its tests are `sheetrs-py/tests` (pytest). Neither binding crate is
  published to crates.io.

## Design Principles

1. **Format Agnostic**: Rules and logic should work on the abstract `Cell` and `Workbook` models, not specific file implementation details, whenever possible.
2. **Fail-Fast Configuration**: Invalid configuration (unknown rules, bad types) causes immediate startup failure to prevent silent misconfiguration.
3. **Hierarchical Violations**: Errors are reported with context (File -> Sheet -> Cell) to make debugging easier.
4. **Performance First**:
   - Use streaming readers where possible.
   - Avoid loading unused data (e.g., loading values when only checking formulas).
