# sheetcli

**sheetcli** is a command-line utility for performing operations on spreadsheet files. It handles conversions, cleaning, and structural modifications.

## Installation

```bash
cargo install --path .
```

## Usage

### 1. Modification

Perform destructive operations. **Note**: destructive operations generally require the `--output` flag to prevent accidental overwrites, or explicit confirmation if supported.

```bash
# Remove specific sheets
sheetcli <FILE> --remove-sheets "Sheet1" "Sheet2" --output <OUT_FILE>

# Remove named ranges
sheetcli <FILE> --remove-ranges "MyRange" "OldData" --output <OUT_FILE>

# Combined operations
sheetcli <FILE> --remove-sheets "Temp" --remove-ranges "TempRange" -o cleaned.xlsx
```

## Supported Formats

- **Input**: XLSX (Full support), ODS (Partial support for conversion).
- **Output (Modification)**: XLSX.
