# Architecture Documentation

The **SheetRS Suite** is organized as a Cargo workspace containing a shared core library and three specialized binary crates. This architecture promotes code reuse, consistency, and testability.

## Workspace Structure

```
sheetrs/
├── sheetrs/    # Shared library (logic, rules, readers)
├── sheetrs-formula/    # Lossless formula AST (parse, print, rewrite references)
├── xlslang/            # Compiler + calculation engine (IR, functions, CLI)
├── sheetlint/          # Linter CLI
├── sheetstats/         # Statistics CLI
├── sheetcli/           # Manipulation CLI
└── Cargo.toml          # Workspace configuration
```

## 1. sheetrs

The heart of the suite. It encapsulates all domain logic, file parsing, rule enforcement, and configuration management.

### Key Modules

- **`reader`**: Abstracts over file formats (XLSX, ODS).
  - Uses `zip` and matching XML parsers (`quick-xml` or custom) for efficient data reading.
  - Custom stream-based parsing logic for both formats ensures minimal memory footprint.
  - `Workbook` trait defines the common interface for all formats.

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
- `edit`: what a formula becomes when rows or columns are inserted or deleted
  or a sheet is renamed or deleted (`#REF!` for what disappears, ranges that
  shrink and grow).
- `r1c1`: relative form of a formula as seen from its own cell.

Lint rules still use string heuristics (`rules/helpers.rs`); migrating them to
the AST is incremental and not part of this crate's introduction.

## 1c. xlslang

Compiles a workbook into a program with declared inputs and outputs and runs it
on a demand-driven engine. Layers: `loader` (the only code that knows the file
format, on top of `sheetrs`) → `Model` → lowering of the `sheetrs-formula` AST
to a format-independent IR → compiler (cone of influence, SCC order, constant
folding) → `Program` → engine. The date system, iteration, search settings and
dialect come from the document; see [xlslang/docs/DESIGN.md](xlslang/docs/DESIGN.md)
and [xlslang/docs/FIDELITY.md](xlslang/docs/FIDELITY.md).

## 2. CLI Tools

The binary crates are thin wrappers around `sheetrs`.

- **`sheetlint`**:
  - Responsible for loading configuration (`sheetlint.toml`).
  - Iterates over the workbook using `sheetrs` readers.
  - Formats output (Text/JSON).

- **`sheetstats`**:
  - Focuses on aggregating metrics.
  - Uses the same `Workbook` model but collects different data points (counts, distributions).
  - `stats.rs` computes workbook and per-sheet metrics; `main.rs` handles arguments and output.

- **`sheetcli`**:
  - Handles command-line arguments for operations.
  - Delegates complex logic (like ZIP manipulation) to `sheetrs::writer`.

## Design Principles

1. **Format Agnostic**: Rules and logic should work on the abstract `Cell` and `Workbook` models, not specific file implementation details, whenever possible.
2. **Fail-Fast Configuration**: Invalid configuration (unknown rules, bad types) causes immediate startup failure to prevent silent misconfiguration.
3. **Hierarchical Violations**: Errors are reported with context (File -> Sheet -> Cell) to make debugging easier.
4. **Performance First**:
   - Use streaming readers where possible.
   - Avoid loading unused data (e.g., loading values when only checking formulas).
