# sheetcli

**sheetcli** is a command-line utility for performing operations on spreadsheet files. It handles conversions, cleaning, and structural modifications.

## Installation

```bash
cargo install --path .
```

## Usage

### 1. Modification

Perform destructive operations. They write a new file given by `--output`,
which must not be the input file.

```bash
# Remove specific sheets
sheetcli <FILE> --remove-sheets "Sheet1" "Sheet2" --output <OUT_FILE>

# Remove named ranges
sheetcli <FILE> --remove-ranges "MyRange" "OldData" --output <OUT_FILE>

# Combined operations
sheetcli <FILE> --remove-sheets "Temp" --remove-ranges "TempRange" -o cleaned.xlsx

# ODS files are modified the same way
sheetcli <FILE>.ods --remove-sheets "Temp" -o cleaned.ods
```

- Removing a sheet also removes the named ranges local to it. A workbook keeps
  at least one sheet.
- The operations of one command are applied together, reading each sheet
  once.

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
  constants in XLSX. ODF has no error constants, so in ODS they keep their
  text (`#DIV/0!`), which LibreOffice shows as text and sheetrs still reads as
  an error.
- Array formulas and data tables are converted only when one scope covers
  their whole range.
- A dynamic array formula whose result Excel keeps as a rich value, such as a
  spill error, caches `#VALUE!` in its place; it keeps its formula (#62).
- An unknown sheet name is an error.

XLSX only:

- A shared formula group whose first cell is converted gives each remaining
  cell outside the scope its own formula, so they keep calculating.
- The calculation chain (`xl/calcChain.xml`) is removed when formulas are
  converted or sheets removed; Excel rebuilds it.
- `--remove-ranges` keeps tables, which sheetrs lists as named ranges (#63).

ODS only:

- Repeated rows and cells (`table:number-rows-repeated`,
  `table:number-columns-repeated`) are split where a scope starts or ends.
- `--remove-ranges` also removes database ranges, which sheetrs lists as named
  ranges.

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

- **Input**: XLSX and ODS.
- **Output (Modification)**: the input format. Every modification supports
  XLSX and ODS (LibreOffice).
