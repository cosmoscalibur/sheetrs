# SheetRS Suite

The **SheetRS Suite** is a high-performance toolkit for processing, linting,
and manipulating spreadsheets (XLSX and ODS). Written in Rust, it is designed
for speed, safety, and ease of integration into CI/CD pipelines.

The suite has two executables: `sheetrs`, with the subcommands `query`, `edit`, `lint` and `stats`,
and `sheetc`, the compiler (`build`, `calculate`, `inspect`, `verify`). Behind them are the libraries,
available as Rust crates, a WebAssembly module and a Python package (see
[Distribution](#distribution)):

- **sheetrs lint**: Advanced spreadsheet linter with hierarchical rule enforcement.
- **sheetrs stats**: Detailed statistics and dependency analysis for workbooks.
- **sheetrs query**: Read the cells of a workbook by range, sheet or workbook, with the properties
  asked for (value, formula, number format, display), find the cells that match a search, and list
  its sheets (`--list-sheets`) and defined names (`--list-names`).
- **sheetrs edit**: General-purpose spreadsheet operations (convert, modify,
  repair).
- **sheetc** (`sheetrs-compiler`): Compiler from a spreadsheet to a standalone
  executable (JSON in, JSON out) with a format-independent IR and a calculation
  engine (`sheetrs-engine`) faithful to Excel and LibreOffice. See
  [sheetrs-compiler/README.md](sheetrs-compiler/README.md).
- **sheetrs-formula**: Lossless formula AST (parser, printer, reference
  rewriting for row/column/sheet edits), shared by the engine and,
  later, by the linter rules.

## Prerequisites

- [Rust](https://www.rust-lang.org/tools/install) ≥ 1.99 (edition 2024)

## Distribution

| Surface | Package | What it gives |
| ------- | ------- | ------------- |
| Executables | GitHub release assets (`sheetrs-vX.Y.Z-<target>`) | `sheetrs` (`query`, `edit`, `lint`, `stats`), `sheetc`, `sheetrs-rt` |
| Rust libraries | `sheetrs-linter` (rules, configuration, violations), `sheetrs-stats`, `sheetrs-query` (reads and finds cells), `sheetrs-edit`, `sheetrs-protocol` (the typed JSON contract), `sheetrs-io` (workbook model and readers, shared by all; the writers behind the feature `write`), `sheetrs-formula`, `sheetrs-engine`, `sheetrs-compiler`, `sheetrs-answer` (the answer of a program written into its workbook) | The same code the executables run |
| WebAssembly | `sheetrs-wasm` (npm tarball in the release) | Every tool in the browser and in Node (lint with SARIF, statistics, query, edit, calculate), and the compiler of a workbook into a WebAssembly module of its own |
| Python | `sheetrs` (wheels and sdist in the release; [sheetrs-py](sheetrs-py/README.md)) | `Book`: queries (`range`, `find`, `lint`, `stats`) and operations that accumulate until `save` (`delete_rows`, `replace`, `formula_to_value`, `calculate`…; `value_ranges()` gives the answer of a `calculate` in memory) |

### Python

```bash
uv add "sheetrs @ git+https://github.com/cosmoscalibur/sheetrs#subdirectory=sheetrs-py"      # in a uv project
pip install "git+https://github.com/cosmoscalibur/sheetrs#subdirectory=sheetrs-py"                 # or with pip
```

Other ways to install are in [Installation](#installation).

```python
import sheetrs

book = sheetrs.Book("book.xlsx")                    # a reference; nothing is read
r = book.range(["Data!A1:C10", "Summary!B2"])       # several ranges, one pass
for v in book.lint("sheetrs.toml"):
    print(v.rule_id, v.sheet, v.cell, v.message)
book.delete_sheets("Secrets").delete_rows("Data!3:10").calculate()   # accumulate
book.save("clean.xlsx")                             # the cut: plan and run
```

Wheels are `abi3` (Python 3.13 to 3.15, as [SPEC 0](https://scientific-python.org/specs/spec-0000/) asks). Details in [sheetrs-py/README.md](sheetrs-py/README.md).

### WebAssembly

```bash
scripts/package-wasm.sh          # builds target/npm/sheetrs-wasm-X.Y.Z.tgz
scripts/package-wasm.sh --demo   # also puts the package in docs/pkg, for the demo page (docs/index.html)
scripts/build-wasm-rt.sh         # the base modules of `sheetc build --target wasm`, next to sheetc
```

The package is an ES module (`--target web`); see
[sheetrs-wasm/README.md](sheetrs-wasm/README.md) for the API.

## Installation

The packages are not on crates.io or PyPI yet ([#108](https://github.com/cosmoscalibur/sheetrs/issues/108)): install
them from GitHub. Linux is the supported platform (x86_64 and aarch64); macOS and Windows are
[#99](https://github.com/cosmoscalibur/sheetrs/issues/99).

### Executables

Download the archive for your platform from the
[releases](https://github.com/cosmoscalibur/sheetrs/releases) page (each has a
`.sha256` file): it holds `sheetrs`, `sheetc` and `sheetrs-rt`.

### Rust: executables and libraries from GitHub

Needs [Rust](https://www.rust-lang.org/tools/install) ≥ 1.99 (`rust-toolchain.toml` picks it).

```bash
cargo install --locked --git https://github.com/cosmoscalibur/sheetrs sheetrs-cli        # sheetrs: query, edit, lint, stats
cargo install --locked --git https://github.com/cosmoscalibur/sheetrs sheetrs-compiler   # sheetc (and sheetrs-rt)

# a tag, a branch or a revision instead of the default branch
cargo install --locked --git https://github.com/cosmoscalibur/sheetrs --tag vX.Y.Z sheetrs-cli
```

As a library, name the crate that has what you need (`sheetrs-io` reads and writes workbooks; the tools are
extension traits of its `Book`):

```toml
[dependencies]
sheetrs-io = { git = "https://github.com/cosmoscalibur/sheetrs", features = ["write"] }
sheetrs-query = { git = "https://github.com/cosmoscalibur/sheetrs" }
```

### Python: from GitHub

Needs Python ≥ 3.13 and, because it builds the extension, Rust ≥ 1.99 (`maturin` is fetched by the build).

```bash
uv add "sheetrs @ git+https://github.com/cosmoscalibur/sheetrs#subdirectory=sheetrs-py"      # in a uv project
pip install "git+https://github.com/cosmoscalibur/sheetrs#subdirectory=sheetrs-py"                 # or with pip
uv run --with "sheetrs @ git+https://github.com/cosmoscalibur/sheetrs#subdirectory=sheetrs-py" python script.py   # a one-off script with uv
```

Without a compiler, use the wheel of a [release](https://github.com/cosmoscalibur/sheetrs/releases)
(`abi3`, Linux x86_64 and aarch64):

```bash
uv add https://github.com/cosmoscalibur/sheetrs/releases/download/vX.Y.Z/<wheel file>.whl   # in a uv project
pip install https://github.com/cosmoscalibur/sheetrs/releases/download/vX.Y.Z/<wheel file>.whl
```

### WebAssembly: from GitHub

The npm package is a tarball of the release:

```bash
npm install https://github.com/cosmoscalibur/sheetrs/releases/download/vX.Y.Z/sheetrs-wasm-X.Y.Z.tgz
```

To build it, see [WebAssembly](#webassembly) above (`scripts/package-wasm.sh`; it needs `wasm-bindgen-cli` at the
version `Cargo.lock` pins).

### From a checkout

```bash
git clone https://github.com/cosmoscalibur/sheetrs.git
cd sheetrs
cargo install --path sheetrs-cli        # sheetrs: query, edit, lint, stats
cargo install --path sheetrs-compiler   # sheetc
```

## Development

```bash
# Build the workspace
cargo build

# Run all tests (the Python bindings are tested with pytest, below)
cargo test --workspace --exclude sheetrs-py

# Lint
cargo clippy --workspace --exclude sheetrs-py --all-targets -- -D warnings

# Format check
cargo fmt --check

# Python bindings (needs Python ≥ 3.13)
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop -m sheetrs-py/Cargo.toml && pytest sheetrs-py
```

CI checks are defined in `.github/workflows/ci.yml`.

No environment variables are required. Configuration is via `sheetrs.toml`.

## Project Structure

```text
sheetrs/
├── sheetrs-io/   # Input: workbook model and XLSX/ODS readers (was sheetrs-core); feature `write`: the writers
├── sheetrs-formula/# Formula AST: parse, print, rewrite references
├── sheetrs-engine/ # IR, functions, run-time evaluation (JSON in/out)
├── sheetrs-compiler/ # Workbook → program compiler (`sheetc`)
├── sheetrs-answer/ # The answer of a program written into its workbook (`fill_template*`)
├── sheetrs-cli/    # The `sheetrs` executable: `query`, `edit`, `lint`, `stats`
├── sheetrs-linter/ # Rules, config, violations (+ `lint` subcommand, feature `cli`)
├── sheetrs-stats/  # Workbook metrics (+ `stats` subcommand)
├── sheetrs-query/  # Reading and finding cells (+ `query` subcommand)
├── sheetrs-protocol/ # The typed JSON contract of the tools
├── sheetrs-edit/   # Workbook modification (+ `edit` subcommand)
├── sheetrs-wasm/   # WASM bindings for browser use: every tool + the compiler
├── sheetrs-rt-wasm/ # The base module of the WebAssembly target of `sheetc`
├── sheetrs-py/     # Python bindings (PyO3, maturin)
├── docs/           # Documentation index + WASM demo
├── tests/          # Shared test assets (lint_book.xlsx/.ods and their generator)
├── scripts/        # Benchmark scripts
└── Cargo.toml      # Workspace configuration
```

## Tools Overview

Every `sheetrs` subcommand and every `sheetc` command writes what it reports besides its answer on
the standard error as diagnostics: warnings and errors as text by default, `-q` for the errors only,
`-v` for what the command did and `-vv` for the detail, and `--log-format json` for a JSON record per
line (`level`, `code`, `message`, `context`).

### 1. sheetrs lint

A comprehensive linter for detecting errors, security issues, performance
bottlenecks, and style violations.

**Features:**

- **Hierarchical Reporting**: Violations grouped by file → sheet → cell.
- **Configurable**: TOML-based configuration with global and per-sheet
  overrides.
- **Formats**: `--format human` (default), `json` (the document for an interface of your own) and
  `sarif` (SARIF 2.1.0, for code scanning, editors and CI), from the CLI, the Rust library, the
  WebAssembly package and Python.

**Usage:**

```bash
# Lint a file
sheetrs lint -i workbook.xlsx

# Lint with custom config
sheetrs lint -i workbook.xlsx --config sheetrs.toml

# JSON for an interface of your own
sheetrs lint -i workbook.xlsx --format json > report.json

# CI/CD and code scanning (SARIF 2.1.0)
sheetrs lint -i workbook.xlsx --format sarif > report.sarif
```

**Example Rules:**

- `ERR102`: Excel error cells (#DIV/0!, etc.)
- `EXT802`: External workbook references
- `CPX503`: Excessive conditional formatting
- `DATA703`: Inconsistent date formats
- `CPX501`: Excessive sheet counts

See the [rule reference](sheetrs-linter/README.md#rule-reference) for the full,
authoritative catalog.

### 2. sheetrs stats

Provides deep insights into workbook structure, complexity, and dependencies.

**Features:**

- **Workbook and sheet stats**: Sheets, named ranges, file size, and per-sheet
  used cells, formulas, errors, empty cells and XML size.
- **Dependencies**: Builds a graph of inter-sheet dependencies (upcoming).

**Usage:**

```bash
# Get general stats
sheetrs stats -i workbook.xlsx
```

### 3. sheetrs edit

A swiss-army knife for spreadsheet manipulation.

**Features:**

- **Modification**: Delete sheets, names, rows and columns; formula to value (convert formulas to
  their values) by sheet, column or row, in XLSX and ODS files.
- **Find & Replace**: `--replace TEXT --with TEXT` over the workbook, a sheet or a
  range, as LibreOffice's Find & Replace does it (literal, regular expression or
  wildcards; match case; entire cells; in values or in formulas).

**Usage:**

```bash
# Delete sensitive sheets and save to new file
sheetrs edit -i input.xlsx --delete-sheets "Secrets" "Admin" --output cleaned.xlsx

# Delete defined names
sheetrs edit -i input.xlsx --delete-names "OldRange" --output cleaned.xlsx

# Freeze the inputs sheet and columns B:D of the model as values
sheetrs edit -i master.xlsx --formula-to-value "Inputs" "Model!B:D" --output run.xlsx

# Replace text in a sheet; the replacement is entered as if it were typed (12 makes a number)
sheetrs edit -i book.ods --replace "2023" --with "2024" --scope Data --in-place

# Replace inside formulas, written in LibreOffice's notation (=SUM($'My Sheet'.A1,B1))
sheetrs edit -i book.xlsx --replace "B1" --with "C1" --in formula -o out.xlsx
```

## Roadmap

- **Error Propagation Tracing**: Future versions may trace only the root cause
  error cell rather than reporting all affected cells.
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
