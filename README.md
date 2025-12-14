# xlslint

Fast Excel/ODS linter with hierarchical violation reporting, written in Rust.

## Features

- **Fast**: Optimized for performance, ideal for CI/CD pipelines
- **Extensible**: Easy to add new linter rules
- **Hierarchical Reporting**: Violations organized by book → sheet → cell
- **Configurable**: TOML-based configuration with global and sheet-specific settings
- **Multiple Formats**: XLSX (Excel) and ODS (LibreOffice) support

## Installation

```bash
cargo install --path xlslint-cli
```

## Usage

### Basic Usage

```bash
# Lint a file with default rules
xlslint myfile.xlsx

# Lint with custom configuration
xlslint myfile.xlsx --config xlslint.toml

# JSON output for CI/CD
xlslint myfile.xlsx --format json

# Show only errors
xlslint myfile.xlsx --errors-only
```

### Configuration

Create a `xlslint.toml` file to customize linter behavior:

```toml
[global]
# Disable specific rules
disabled_rules = ["UX001"]

[rules.PERF003]
threshold_rows = 30
threshold_columns = 30

[sheets."Summary"]
# Disable rules for specific sheets
disabled_rules = ["PERF002"]
```

## Default Active Rules

The following rules are active by default:

### Unresolved Errors
- **ERR001**: Error cell values (#DIV/0!, #REF!, etc.)
- **ERR002**: Broken named ranges

### Security and Privacy
- **SEC001**: External links in formulas or URLs
- **SEC004**: Macros and VBA code detection

### Formatting and Usability
- **UX001**: Inconsistent number formatting

### Performance
- **PERF001**: Unused named ranges
- **PERF002**: Unused sheets
- **PERF003**: Large used range (>20 rows/cols beyond data)
- **PERF006**: Excessive conditional formatting

## Available Rules

All rules can be enabled/disabled via configuration. The following additional rules are available:

### Security and Privacy
- **SEC002**: Hidden sheets
- **SEC003**: Hidden columns or rows

### Formatting and Usability
- **UX004**: Blank rows/columns in used ranges

### Performance
- **PERF004**: Volatile functions (NOW, RAND, INDIRECT, etc.)
- **PERF005**: Duplicate formulas

### Structural and Maintainability
- **SM001**: Excessive sheet counts (default: >50)
- **SM002**: Duplicate sheet names
- **SM003**: Long formulas (default: >255 chars)
- **SM004**: Long text cells (default: >255 chars)
- **SM005**: Merged cells
- **SM006**: Non-descriptive sheet names

### Formula Best Practices
- **FORM001**: Whole column/row references (A:A, 1:1)
- **FORM002**: Empty string tests (suggest ISBLANK)
- **FORM003**: Deep formula nesting (default: >5 levels)
- **FORM004**: Deeply nested IF statements (default: >5 levels)
- **FORM005**: Circular references

## Development

### Building

```bash
cargo build --release
```

### Testing

```bash
cargo test
```

### Adding New Rules

See [ARCHITECTURE.md](ARCHITECTURE.md) for details on adding new linter rules.

## License

MIT OR Apache-2.0
