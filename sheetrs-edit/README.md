# sheetedit

**sheetedit** is a command-line utility for performing operations on spreadsheet files. It handles conversions, cleaning, and structural modifications.

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
sheetedit <FILE> --remove-sheets "Sheet1" "Sheet2" --output <OUT_FILE>

# Remove named ranges
sheetedit <FILE> --remove-ranges "MyRange" "OldData" --output <OUT_FILE>

# Combined operations
sheetedit <FILE> --remove-sheets "Temp" --remove-ranges "TempRange" -o cleaned.xlsx

# ODS files are modified the same way
sheetedit <FILE>.ods --remove-sheets "Temp" -o cleaned.ods
```

The reference behavior of every operation is **LibreOffice 26.8 on Linux**: what a formula
becomes after an operation is what LibreOffice gives when the same operation is made in it
(`docs/FIDELITY.md` of `sheetrs-compiler` lists where Excel differs). The check is made on XLSX
and ODS files against LibreOffice: `sheetrs-core/scripts/compare_structure_with_libreoffice.py` makes
each operation in LibreOffice through UNO and with `sheetedit` on the same file, opens both results
in LibreOffice and compares formulas, values, notes, page breaks and shapes. Excel is not part of
it.

- Removing a sheet also removes the named ranges local to it. A workbook keeps
  at least one sheet.
- The references to a removed sheet become `#REF!` (`SUM(Data!A1:A3)` is `SUM(#REF!)`), and a
  name that pointed at it keeps a definition that is `#REF!`.
- A removed named range stays in the formulas that use it, as in LibreOffice, where it is the
  unknown name `#NAME?`; another name whose definition used it has `#NAME?` there.
- The operations of one command are applied together, reading each sheet
  once.

#### Delete rows and columns

```bash
# Rows 10 to 20 and column B of one sheet; `Sheet!7` and `Sheet!C` delete one
sheetedit book.xlsx --delete-rows "Model!10:20" --delete-cols "Model!B" -o shorter.xlsx
```

Rows and columns are those of the input file, whatever the order and the other operations of
the command: `--delete-rows "Model!3:4" "Model!7"` deletes rows 3, 4 and 7 of the file, and a
`--convert-to-values` of the same command reads the file's coordinates too.

The formulas of **every sheet** follow, as LibreOffice does it:

| A reference to… | Becomes |
|-----------------|---------|
| cells after the deletion | the cells' new place (`A6` with rows 4-5 deleted is `A4`) |
| a deleted cell or a range with no cell left | `Sheet!#REF!` (`#REF!` in an ODS); a removed sheet leaves a bare `#REF!` (see the last item of this list) |
| a range that loses cells | the range of what is left (`A4:A8` is `A4:A6`) |
| whole rows or columns | whole rows or columns (`A:A` stays when rows go), `#REF!` when all go |
| a 3D reference (`Sheet1:Sheet3!A1`) | follows a deletion in a sheet of its span only when the deleted rows or columns touch what it names; one before it does not move it. A removed sheet at an end of the span passes the end to the next sheet; `#REF!` when no sheet of the span stays |

Also adjusted: the cells and their rows or columns (widths, hidden, merged areas), shared and
array formulas, defined names and print ranges, conditional formats, data validations,
hyperlinks, filters (the range shrinks around the deleted columns and their criteria go), the conditional formats, validations and sparklines of the extensions, notes (XLSX comments and threaded comments, with their shapes), manual page
breaks, the anchors of shapes and the series of the charts of an XLSX, and the database ranges and
print ranges of an ODS. A formula whose text changes loses its cached result, which
is no longer valid; one that does not change keeps it.

- `--recalc-on-open` makes the application that opens the file compute every formula (XLSX:
  `fullCalcOnLoad` in `workbook.xml`, which Excel and LibreOffice clear when they save; ODS has no
  such marker, so the cached results of the formulas are dropped). Without it the cached results
  that stay are the ones the file had, and a formula that depends on the deleted cells without
  changing its text (`SUM(A:A)`) keeps a result that is out of date.
- A deletion stops with an error when it cannot be made right. In both formats: it covers part of
  an array formula, removes the first cell of a merged area that stays, or removes the cell a
  shape is anchored to together with the rows (or columns) before it, because LibreOffice then
  places the shape by its size and the height of the rows, which the file does not hold. In an
  XLSX: a sheet that changes has tables, header or footer images or data tables; the workbook has pivot tables. In an ODS: the document has pivot tables, filters, consolidations or
  scenarios.
- Shapes, images and charts (the drawings of an XLSX, the shapes in the cells of an ODS) follow
  their anchor cell as LibreOffice moves them. A shape whose cell is deleted stays: it is
  anchored to the cell with the index of the one it left, which is where LibreOffice keeps it
  when the rows or columns that pass have the same height or width. A shape that resizes with its
  cells (`editAs="twoCell"`, the default, in an XLSX; `table:end-cell-address` in an ODS) has its
  end corner follow its end cell the same way; any other keeps its size in rows and columns.
  Manual page breaks go with their rows.
- A bare `#REF!` in an XLSX formula is read by LibreOffice as an unknown name (`#NAME?`), also
  in the files it saves itself; a deleted cell is written `Sheet!#REF!`, which it reads as `#REF!`.
  The references to a removed sheet have no sheet to write: they are the bare `#REF!`.

#### Convert formulas to values

`--convert-to-values` replaces formulas by their cached values, e.g. to turn a
master workbook into a production run model where argument cells hold fixed
values and need no computation. Each scope selects a whole sheet, columns or
rows:

```bash
# Whole sheet, columns B to D, and rows 3 to 10
sheetedit master.xlsx --convert-to-values "Inputs" "Model!B:D" "Model!3:10" -o run.xlsx

# Single column or row; quote sheet names containing '!' as in formulas
sheetedit master.xlsx --convert-to-values "Model!F" "'R!D'!2" -o run.xlsx
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
  `table:number-columns-repeated`) are split where a scope starts or ends, and shortened where
  a deletion cuts them.
- `--remove-ranges` also removes database ranges, which sheetrs lists as named
  ranges.

### 2. Inspection

Non-destructive operations to inspect file internals.

```bash
# List all named ranges and check for errors (#REF!)
sheetedit <FILE> --list-ranges
```

### 3. Dry Run

Preview what operations would be performed without writing to disk.

```bash
sheetedit <FILE> --remove-sheets "Old" --dry-run
```

## Supported Formats

- **Input**: XLSX and ODS.
- **Output (Modification)**: the input format. Every modification supports
  XLSX and ODS (LibreOffice).
