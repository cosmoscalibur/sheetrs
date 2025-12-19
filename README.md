# SheetCraft Suite

The **SheetCraft Suite** is a high-performance toolkit for processing, linting,
and manipulating spreadsheets (XLSX and ODS). Written in Rust, it is designed
for speed, safety, and ease of integration into CI/CD pipelines.

The suite consists of three specialized CLI tools:

- **sheetlint**: Advanced spreadsheet linter with hierarchical rule enforcement.
- **sheetstats**: Detailed statistics and dependency analysis for workbooks.
- **sheetcli**: General-purpose spreadsheet operations (convert, modify,
  repair).

## Installation

### From Source

Requires [Rust](https://www.rust-lang.org/tools/install) (latest stable).

```bash
# Clone the repository
git clone https://github.com/cosmoscalibur/sheetcraft.git
cd sheetcraft

# Install all tools
cargo install --path sheetlint
cargo install --path sheetstats
cargo install --path sheetcli
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

**Key Rules:**

- `ERR001`: Error cells (#DIV/0!, etc.)
- `SEC001`: External links
- `PERF006`: Excessive conditional formatting
- `UX002`: Inconsistent date formats
- `SM001`: Excessive sheet counts

### 2. sheetstats

Provides deep insights into workbook structure, complexity, and dependencies.

**Features:**

- **General Stats**: Counts of sheets, cells, formulas, values.
- **Dependencies**: Builds a graph of inter-sheet dependencies (upcoming).

**Usage:**

```bash
# Get general stats
sheetstats workbook.xlsx
```

### 3. sheetcli

A swiss-army knife for spreadsheet manipulation.

**Features:**

- **Modification**: Remove sheets, delete named ranges.

**Usage:**

```bash
# Remove sensitive sheets and save to new file
sheetcli input.xlsx --remove-sheets "Secrets" "Admin" --output cleaned.xlsx

# Remove named ranges
sheetcli input.xlsx --remove-ranges "OldRange" --output cleaned.xlsx
```

## Roadmap

- **ODS Support**: Partial support implemented. Full parity with XLSX is in
  progress.
- **Python Bindings**: PyO3 bindings for direct integration with Python data
  workflows.
- **Windows & WASM**: Windows support is likely functional but untested. WASM
  target for browser-based linting is planned.
- **Performance Review**: Continuous optimization for large workbooks (>1M
  cells).

## Known Issues

- **ERR001**: Currently reports all affected cells. Future versions may trace
  only the root cause error.

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for a detailed technical breakdown of the
`sheetcraft-core` library and the CLI tool implementations.

## License

MIT
