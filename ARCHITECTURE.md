# Architecture Documentation

The **SheetRS Suite** is organized as a Cargo workspace containing a shared core library and three specialized binary crates. This architecture promotes code reuse, consistency, and testability.

## Reference Behavior

The reference behavior of the compiler, the runtime and the modifications of `sheetedit` is **LibreOffice 26.8 on Linux**. What an operation gives where the formats or the applications differ is what LibreOffice gives for it (`sheetrs-compiler/docs/FIDELITY.md`).

## Workspace Structure

```
sheetrs/
├── sheetrs-core/       # Shared kernel: workbook model + XLSX/ODS readers
├── sheetrs-formula/    # Lossless formula AST (parse, print, rewrite references)
├── sheetrs-linter/     # Rules, configuration, violations (lib) + `sheetlint`
├── sheetrs-stats/      # Workbook metrics (lib) + `sheetstats`
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

- **Driven side** (what a tool needs): the workbook model and the readers of `sheetrs-core`. The
  model (`Workbook`, `Cell`, `Grid`, `CellSink`, `WorkbookReader`) is the port; the XLSX and ODS
  parsers are its adapters. A tool never sees a file format.
- **Driving side** (who calls a tool): the executables (`sheetlint`, `sheetstats`, `sheetedit`,
  `sheetc`) and the bindings (`sheetrs-wasm`, `sheetrs-py`). Each is an adapter that translates
  arguments and renders results; none holds a decision. What several adapters need to present
  (violations grouped by sheet and cell, flat records) is part of the tool's library
  (`sheetrs_linter::report`), not copied into each adapter.

Dependencies point inward: adapters → tool libraries → `sheetrs-core` / `sheetrs-formula`. Tool
crates do not depend on each other.

### What belongs in `sheetrs-core`

`sheetrs-core` is a **shared kernel**, not "the public library": the public surface of the suite is
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
a new edit operation) never changes `sheetrs-core`, and a second consumer of a tool's code is
served by that tool's library, not by growing the kernel. Reading is shared because four crates
read workbooks; writing is not, because one does.

## 1. sheetrs-core

The shared kernel: the format-agnostic workbook model and the streaming readers that fill it.

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
  that knows the file format, on top of `sheetrs-core`) → `Model` → lowering of the
  `sheetrs-formula` AST to the IR → cone of influence, SCC order, constant
  folding. Its binary is `sheetc`.

The date system, iteration, search settings and reference syntax come from the document;
see [sheetrs-compiler/docs/DESIGN.md](sheetrs-compiler/docs/DESIGN.md) and
[sheetrs-compiler/docs/FIDELITY.md](sheetrs-compiler/docs/FIDELITY.md).

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
  - `stats.rs` (`WorkbookStats::from_path`/`from_reader`) computes workbook and per-sheet
    metrics and is what the bindings call; `main.rs` handles arguments and output.

- **`sheetedit`** (crate `sheetrs-edit`):
  - `writer`: `xlsx_writer.rs` and `ods_writer.rs` apply the same `WorkbookModifications` to each
    format, stream-editing the archive with `zip` and `quick-xml` (removing sheets, converting
    formulas to values, deleting rows and columns) without fully rewriting the file.
  - `ValueScope` selects the cells to convert (a sheet, columns or rows) independently of the
    file format.
  - `EditPlan` and `Operation` (module `plan`, exported at the crate root): a **sequence** of
    operations on one workbook, each addressing the workbook as the previous ones left it (what
    running them one at a time gives). Rows and columns are typed (`RowRange`, `ColumnRange`, with
    `FromStr`/`Display`), so a deletion with the wrong kind of range cannot be built; the plan is
    a collection (`Extend`, `FromIterator`, `IntoIterator`). Marking the output to be recalculated
    when opened is the plan's decision (`Recalculation`: `Auto` marks it when an operation leaves
    stale cached results, `Always`, `Never`), applied once in the last pass, not an operation of
    the sequence. `passes()` plans the sequence into the fewest passes over
    the file that give the same workbook: deletions of one sheet and axis share a pass, with the
    later positions translated through the earlier deletions; deletions of different sheets and
    on rows vs columns, name and sheet removals and the recalculation mark commute; a conversion
    after another change, and two deletions that touch (which the writer would take for one
    band), start a new pass. `against_input()` is the declarative case (a command line: all
    written against the input, in no order), always one pass. A pass is applied by the writer in
    the order defined once in `StructureEdits::edits`. A test generates sequences and compares
    the plan with the operations run one at a time, in both formats. A reference that spans
    sheets would tie sheets together, and is refused (issue #75).
  - The writer streams: each pass reads and writes every worksheet event by event, and the plan
    is a few bands and names, so memory does not depend on the size of the workbook.
  - `main.rs` handles arguments.

## 3. Bindings

The bindings are adapters: they convert arguments and results and hold no analysis logic. What
they present comes from the tool libraries (`sheetrs_linter::report`, `sheetrs_stats`,
`sheetrs_edit::writer`), so every surface reports the same findings. The API of a binding is
organized by tool, as the crates are: the rules are a part of the linter, so they are offered
next to `lint`, not as a separate service.

- **`sheetrs-wasm`** (`cdylib`, wasm-bindgen): bytes in, structured values (`lint_workbook_records`,
  `lint_rules`, `workbook_stats`) or the human text report out. Packaged for npm by
  `scripts/package-wasm.sh` (`wasm-bindgen --target web`, template in `sheetrs-wasm/npm`).
- **`sheetrs-py`** (`cdylib`, PyO3 with `abi3`): the extension `sheetrs._sheetrs`, re-exported flat
  by the `python/sheetrs` package (type stubs). `open` makes a `Source`, the Python face of
  `sheetrs_core::reader::Source`: a reference to a workbook that reads nothing, with no `save`.
  The tools take it: `Linter.lint(source)` (with `Violation`, its result type), `Source.stats()`,
  and the operations on it, which make an `EditPlan` over a `sheetrs_edit::EditPlan` and are the
  only objects with `save`. There is no eager mode: a workbook is streamed, never loaded. The
  functions release the GIL while they work. It is built by maturin, which enables the
  `extension-module` feature; the crate has no Rust test harness (it would have to link
  libpython), its tests are `sheetrs-py/tests` (pytest). Neither binding crate is published to
  crates.io.

## Design Principles

1. **Format Agnostic**: Rules and logic should work on the abstract `Cell` and `Workbook` models, not specific file implementation details, whenever possible.
2. **Fail-Fast Configuration**: Invalid configuration (unknown rules, bad types) causes immediate startup failure to prevent silent misconfiguration.
3. **Hierarchical Violations**: Errors are reported with context (File -> Sheet -> Cell) to make debugging easier.
4. **Performance First**:
   - Use streaming readers where possible.
   - Avoid loading unused data (e.g., loading values when only checking formulas).
