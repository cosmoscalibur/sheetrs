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

#### Convert formulas to values

`--convert-to-values` replaces formulas by their cached values, e.g. to turn a
master workbook into a production run model where argument cells hold fixed
values and need no computation. Each scope selects a whole sheet, columns or
rows:

```bash
# Whole sheet, columns B to D, and rows 3 to 10
sheetcli master.xlsx --convert-to-values "Inputs" "Model!B:D" "Model!3:10" -o run.xlsx

# Single column or row; quote sheet names containing '!' as in formulas
sheetcli master.xlsx --convert-to-values "Model!F" "'R!D'!2" -o run.xlsx
```

- Values are the results the spreadsheet application stored when it last
  saved the file; the file is not recalculated. A formula without a stored
  value keeps its formula.
- Formulas that return text become text cells. Error results become error
  constants.
- A shared formula group whose first cell is converted gives each remaining
  cell outside the scope its own formula, so they keep calculating.
- Array formulas and data tables are converted only when one scope covers
  their whole range.
- The calculation chain (`xl/calcChain.xml`) is removed; Excel rebuilds it.
- An unknown sheet name is an error.

### 2. Inspection

Non-destructive operations to inspect file internals.

```bash
# List all named ranges and check for errors (#REF!)
sheetcli <FILE> --list-ranges
```

### 3. Dry Run

Preview what operations would be performed without writing to disk.

```bash
sheetcli <FILE> --remove-sheets "Old" --dry-run
```

## Supported Formats

- **Input**: XLSX (Full support), ODS (Read-only for inspection/linting purposes in this tool).
- **Output (Modification)**: XLSX.
