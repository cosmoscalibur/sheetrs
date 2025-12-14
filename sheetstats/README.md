# sheetstats

**sheetstats** is a CLI tool for analyzing the complexity and structure of spreadsheets. It provides metrics that help identify "heavy" workbooks, complex dependencies, and potential optimization targets.

## Installation

```bash
cargo install --path .
```

## Usage

```bash
sheetstats <FILE> [OPTIONS]
```

### Options

- `--format <FORMAT>`: Output format (text/json).

## Metrics Reported

- **General**: Total sheets, named ranges and size.
- **Content**: Formulas, Cells and non-empty cells by sheets.
