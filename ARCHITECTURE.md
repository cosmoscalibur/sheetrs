# Architecture Documentation

The **SheetRS Suite** is organized as a Cargo workspace containing a shared core library and three specialized binary crates. This architecture promotes code reuse, consistency, and testability.

## Reference Behavior

The reference behavior of the compiler, the runtime and the modifications of `sheetedit` is **LibreOffice 26.8 on Linux**. What an operation gives where the formats or the applications differ is what LibreOffice gives for it (`sheetrs-compiler/docs/FIDELITY.md`).

## Workspace Structure

```
sheetrs/
├── sheetrs-core/       # Core library (readers, writers, rules, config)
├── sheetrs-formula/    # Lossless formula AST (parse, print, rewrite references)
├── sheetrs-engine/     # IR, function library, run-time evaluation, `sheetrs-rt`
├── sheetrs-compiler/   # Loader, dependency analysis, folding, `sheetc`
├── sheetrs-linter/     # Linter CLI
├── sheetrs-stats/      # Statistics CLI
├── sheetrs-edit/       # Workbook editing CLI (`sheetedit`)
└── Cargo.toml          # Workspace configuration
```

## 1. sheetrs-core

The heart of the suite. It encapsulates all domain logic, file parsing, rule enforcement, and configuration management.

### Key Modules

- **`reader`**: Abstracts over file formats (XLSX, ODS).
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

- **`rules`**: Implements the linting logic.
  - Each rule is a standalone struct implementing the `WalkerRule` trait.
  - The `WorkbookWalker` executes all rules in a single pass over the workbook.
  - Rules are registered in a central `Registry`.
  - Categories: `ERR` (Errors), `CALC` (Calculations), `REF` (References), `INT` (Interruptions), `CPX` (Complexity), `VUL` (Vulnerability), `DATA` (Data Issues), `EXT` (External), `HID` (Hidden), `FILE` (Files), `VBA` (Macros).

- **`config`**: Handles TOML configuration.
  - Hierarchical loading: Default -> Global Config -> Sheet Overrides.
  - Supports enabling/disabling rules by ID or Category.

- **`writer`**: Handles file modification.
  - `xlsx_writer.rs` and `ods_writer.rs` apply the same `WorkbookModifications` to each format.
  - Uses `zip` and `quick-xml` to stream-edit the archive structure (e.g., removing sheets, converting formulas to values) without fully rewriting the file, ensuring speed and preserving metadata.
  - `ValueScope` selects the cells to convert (a sheet, columns or rows) independently of the file format.

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

## 2. CLI Tools

The tool crates are thin wrappers around `sheetrs-core`. Each is named `sheetrs-<tool>` and builds an
executable with the short name below.

- **`sheetlint`** (crate `sheetrs-linter`):
  - Responsible for loading configuration (`sheetlint.toml`).
  - Iterates over the workbook using `sheetrs-core` readers.
  - Formats output (Text/JSON).

- **`sheetstats`** (crate `sheetrs-stats`):
  - Focuses on aggregating metrics.
  - Uses the same `Workbook` model but collects different data points (counts, distributions).
  - `stats.rs` computes workbook and per-sheet metrics; `main.rs` handles arguments and output.

- **`sheetedit`** (crate `sheetrs-edit`):
  - Handles command-line arguments for operations.
  - Delegates complex logic (like ZIP manipulation) to `sheetrs_core::writer`.

## Design Principles

1. **Format Agnostic**: Rules and logic should work on the abstract `Cell` and `Workbook` models, not specific file implementation details, whenever possible.
2. **Fail-Fast Configuration**: Invalid configuration (unknown rules, bad types) causes immediate startup failure to prevent silent misconfiguration.
3. **Hierarchical Violations**: Errors are reported with context (File -> Sheet -> Cell) to make debugging easier.
4. **Performance First**:
   - Use streaming readers where possible.
   - Avoid loading unused data (e.g., loading values when only checking formulas).
