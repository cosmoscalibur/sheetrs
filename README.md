# SheetRS Suite

The **SheetRS Suite** is a high-performance toolkit for processing, linting,
and manipulating spreadsheets (XLSX and ODS). Written in Rust, it is designed
for speed, safety, and ease of integration into CI/CD pipelines.

The suite consists of three specialized CLI tools, a compiler, and two
libraries, available as executables, Rust crates, a WebAssembly module and a
Python package (see [Distribution](#distribution)):

- **sheetlint**: Advanced spreadsheet linter with hierarchical rule enforcement.
- **sheetstats**: Detailed statistics and dependency analysis for workbooks.
- **sheetedit**: General-purpose spreadsheet operations (convert, modify,
  repair).
- **sheetc** (`sheetrs-compiler`): Compiler from a spreadsheet to a standalone
  executable (JSON in, JSON out) with a format-independent IR and a calculation
  engine (`sheetrs-engine`) faithful to Excel and LibreOffice. See
  [sheetrs-compiler/README.md](sheetrs-compiler/README.md).
- **sheetrs-formula**: Lossless formula AST (parser, printer, reference
  rewriting for row/column/sheet edits, R1C1 form), shared by the engine and,
  later, by the linter rules.

## Prerequisites

- [Rust](https://www.rust-lang.org/tools/install) ≥ 1.99 (edition 2024)

## Distribution

| Surface | Package | What it gives |
| ------- | ------- | ------------- |
| Executables | GitHub release assets (`sheetrs-vX.Y.Z-<target>`) | `sheetlint`, `sheetstats`, `sheetedit`, `sheetc`, `sheetrs-rt` |
| Rust libraries | `sheetrs-linter` (rules, configuration, violations), `sheetrs-stats`, `sheetrs-edit`, `sheetrs-core` (workbook model and readers, shared by all), `sheetrs-formula`, `sheetrs-engine`, `sheetrs-compiler` | The same code the executables run |
| WebAssembly | `sheetrs-wasm` (npm tarball in the release) | Lint and statistics in the browser, structured and text output |
| Python | `sheetrs` (wheels and sdist in the release; [sheetrs-py](sheetrs-py/README.md)) | `open` → `Source` (`Linter.lint`, `stats`, editing → `EditPlan.save`) |

### Python

```bash
pip install sheetrs          # or the wheel from the GitHub release
```

```python
import sheetrs

source = sheetrs.open("book.xlsx")          # a reference; nothing is read
for v in sheetrs.Linter(config_path="sheetlint.toml").lint(source):
    print(v.rule_id, v.sheet, v.cell, v.message)
print(source.stats()["total_formulas"])
# the operations on a source make an EditPlan, which has save()
source.remove_sheets("Secrets").delete_rows("Data!3:10").save("clean.xlsx")
```

Wheels are `abi3` (Python ≥ 3.10). Details in [sheetrs-py/README.md](sheetrs-py/README.md).

### WebAssembly

```bash
scripts/package-wasm.sh      # builds target/npm/sheetrs-wasm-X.Y.Z.tgz
```

The package is an ES module (`--target web`); see
[sheetrs-wasm/README.md](sheetrs-wasm/README.md) for the API.

## Installation

### Executables

Download the archive for your platform from the
[releases](https://github.com/cosmoscalibur/sheetrs/releases) page (each has a
`.sha256` file).

### From Source

```bash
# Clone the repository
git clone https://github.com/cosmoscalibur/sheetrs.git
cd sheetrs

# Install all tools
cargo install --path sheetrs-linter
cargo install --path sheetrs-stats
cargo install --path sheetrs-edit
cargo install --path sheetrs-compiler   # sheetc
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

# Python bindings (needs Python ≥ 3.10)
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop -m sheetrs-py/Cargo.toml && pytest sheetrs-py
```

CI checks are defined in `.github/workflows/ci.yml`.

No environment variables are required. Configuration is via `sheetlint.toml`.

## Project Structure

```text
sheetrs/
├── sheetrs-core/   # Shared kernel: workbook model and XLSX/ODS readers
├── sheetrs-formula/# Formula AST: parse, print, rewrite references
├── sheetrs-engine/ # IR, functions, run-time evaluation (JSON in/out)
├── sheetrs-compiler/ # Workbook → program compiler (`sheetc`)
├── sheetrs-linter/ # Rules, config, violations + `sheetlint`
├── sheetrs-stats/  # Workbook metrics + `sheetstats`
├── sheetrs-edit/   # Workbook modification + `sheetedit`
├── sheetrs-wasm/   # WASM bindings for browser use
├── sheetrs-py/     # Python bindings (PyO3, maturin)
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

See the [rule reference](sheetrs-linter/README.md#rule-reference) for the full,
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

### 3. sheetedit

A swiss-army knife for spreadsheet manipulation.

**Features:**

- **Modification**: Remove sheets, delete named ranges, convert formulas to
  their values by sheet, column or row, in XLSX and ODS files.

**Usage:**

```bash
# Remove sensitive sheets and save to new file
sheetedit input.xlsx --remove-sheets "Secrets" "Admin" --output cleaned.xlsx

# Remove named ranges
sheetedit input.xlsx --remove-ranges "OldRange" --output cleaned.xlsx

# Freeze the inputs sheet and columns B:D of the model as values
sheetedit master.xlsx --convert-to-values "Inputs" "Model!B:D" --output run.xlsx
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
