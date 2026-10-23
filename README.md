# SheetRS Suite

The **SheetRS Suite** is a high-performance toolkit for processing, linting,
and manipulating spreadsheets (XLSX and ODS). Written in Rust, it is designed
for speed, safety, and ease of integration into CI/CD pipelines.

The suite consists of three specialized CLI tools, a compiler, and two
libraries:

- **sheetlint**: Advanced spreadsheet linter with hierarchical rule enforcement.
- **sheetstats**: Detailed statistics and dependency analysis for workbooks.
- **sheetcli**: General-purpose spreadsheet operations (convert, modify,
  repair).
- **xlslang**: Compiler from a spreadsheet to a standalone executable (JSON in,
  JSON out) with a format-independent IR and a calculation engine faithful to
  Excel and LibreOffice. See [xlslang/README.md](xlslang/README.md).
- **sheetrs-formula**: Lossless formula AST (parser, printer, reference
  rewriting for row/column/sheet edits, R1C1 form), shared by the engine and,
  later, by the linter rules.

## Prerequisites

- [Rust](https://www.rust-lang.org/tools/install) ≥ 1.99 (edition 2024)

## Installation

### From Source

```bash
# Clone the repository
git clone https://github.com/cosmoscalibur/sheetrs.git
cd sheetrs

# Install all tools
cargo install --path sheetlint
cargo install --path sheetstats
cargo install --path sheetcli
```

## Development

```bash
# Build the workspace
cargo build

# Run all tests
cargo test

# Lint
cargo clippy --all-targets -- -D warnings

# Format check
cargo fmt --check
```

CI checks are defined in `.github/workflows/ci.yml`.

No environment variables are required. Configuration is via `sheetlint.toml`.

## Project Structure

```text
sheetrs/
├── sheetrs/        # Core library (parsers, rules, config, writer)
├── sheetrs-formula/# Formula AST: parse, print, rewrite references
├── xlslang/        # Spreadsheet → executable compiler and engine
├── sheetlint/      # Linter CLI
├── sheetstats/     # Statistics CLI
├── sheetcli/       # Manipulation CLI
├── sheetrs-wasm/   # WASM bindings for browser use
├── docs/           # Documentation index + WASM demo
├── tests/          # Shared test assets (minimal_test.xlsx/.ods)
├── scripts/        # Benchmark scripts
└── Cargo.toml      # Workspace configuration
```

## Tools Overview

### 1. sheetlint

A comprehensive linter for detecting errors, security issues, performance
bottlenecks, and style violations.

**Features:**

- **Hierarchical Reporting**: Violations grouped by file → sheet → cell.
- **Configurable**: TOML-based configuration with global and per-sheet
  overrides.
- **Formats**: Support for JSON and text output.

**Usage:**

```bash
# Lint a file
sheetlint workbook.xlsx

# Lint with custom config
sheetlint workbook.xlsx --config sheetlint.toml

# CI/CD mode (JSON output) (:warning: experimental/unstable)
sheetlint workbook.xlsx --format json > report.json
```

**Example Rules:**

- `ERR102`: Excel error cells (#DIV/0!, etc.)
- `EXT802`: External workbook references
- `CPX503`: Excessive conditional formatting
- `DATA703`: Inconsistent date formats
- `CPX501`: Excessive sheet counts

See the [rule reference](sheetlint/README.md#rule-reference) for the full,
authoritative catalog.

### 2. sheetstats

Provides deep insights into workbook structure, complexity, and dependencies.

**Features:**

- **Workbook and sheet stats**: Sheets, named ranges, file size, and per-sheet
  used cells, formulas, errors, empty cells and XML size.
- **Dependencies**: Builds a graph of inter-sheet dependencies (upcoming).

**Usage:**

```bash
# Get general stats
sheetstats workbook.xlsx
```

### 3. sheetcli

A swiss-army knife for spreadsheet manipulation.

**Features:**

- **Modification**: Remove sheets, delete named ranges, convert formulas to
  their values by sheet, column or row, in XLSX and ODS files.

**Usage:**

```bash
# Remove sensitive sheets and save to new file
sheetcli input.xlsx --remove-sheets "Secrets" "Admin" --output cleaned.xlsx

# Remove named ranges
sheetcli input.xlsx --remove-ranges "OldRange" --output cleaned.xlsx

# Freeze the inputs sheet and columns B:D of the model as values
sheetcli master.xlsx --convert-to-values "Inputs" "Model!B:D" --output run.xlsx
```

## Roadmap

- **Error Propagation Tracing**: Future versions may trace only the root cause
  error cell rather than reporting all affected cells.
- **Python Bindings**: PyO3 bindings for direct integration with Python data
  workflows.
- **Performance Review**: Continuous optimization for large workbooks (>1M
  cells).

## Documentation

- [Architecture](ARCHITECTURE.md) — technical breakdown of the library and CLI
  tools.
- [Documentation Index](docs/README.md) — coding patterns, profiling, analysis,
  and WASM demo.
- [Contributing](CONTRIBUTING.md) — development setup, testing, and PR process.

## License

MIT
