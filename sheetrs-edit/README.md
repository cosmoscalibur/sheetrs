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
and ODS files against LibreOffice: `sheetrs-edit/scripts/compare_structure_with_libreoffice.py` makes
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
  once. They are not run in the order they are written: they are planned (below).

#### How the operations of one command are planned

The flags of one command are all written against the **input** file, in no order: `--delete-rows
"Model!3:4" "Model!7"` deletes rows 3, 4 and 7 of the file, and a `--convert-to-values` of the same
command reads the file's rows too. The command is applied in a single pass over the file, in an
order that does not depend on how the flags were written:

1. formulas converted to values, while every cached result is still there (a conversion on a sheet
   that is removed goes with it);
2. in each sheet, rows deleted **from the last to the first**, overlapping or touching ranges
   merged, so a deletion shifts only what is below it and the next one is still at the position the
   input gives it;
3. in each sheet, columns deleted the same way;
4. named ranges removed;
5. sheets removed (the deletions of a removed sheet go with it);
6. the mark to recalculate when the file is opened.

The order has two kinds of dependency and no more: a conversion must read values that a deletion
or a removal would lose, and the deletions of one sheet shift one another. Operations on
different sheets do not depend on each other, because a reference moves only with the rows and
columns of the sheet it names, so the order between sheets is free. That holds because references
across sheets are not supported (below). `--dry-run` prints the passes.

**As a library** (`sheetrs_edit::Editor`), an editor is made for a `Source` and collects changes in
order, and each one addresses the workbook as the previous ones left it, as when they are run one
at a time: after `delete_rows` of row 3, row 3 is what was row 4. `save(output)` plans the sequence
into the fewest passes over the file that give what running them one at a time gives; changes that
cannot share a pass (a conversion after another change, which must read what that change dropped;
two deletions that touch, which the writer would see as one band) start a new one, and
`explain()` says what `save` will do. The flags above are the same changes applied in a fixed order
(conversions, rows, columns, names, sheets), which is always one pass whatever order the flags
came in. The planner and the writer are not part of the API: the editor is. Tests run generated
sequences both ways, in XLSX and ODS, and compare the workbooks.

```rust
let source = Source::from_file("in.xlsx")?;
Editor::new(&source)
    .delete_rows(["Data!3:10".parse()?])
    .remove_sheets(["Scratch"])
    .save("out.xlsx")?;
```

**Recalculation is decided by the editor, with no option for it.** Deleting rows or columns, or
removing sheets or names, leaves the cached results of formulas that read what went (their text
may not change, so their value is stale). An editor with one of them marks the output to be
recalculated when it is opened, as the last thing it does, after every operation (an ODS applies
the mark by dropping every cached result, which in the middle of a sequence would change what a
later conversion finds). A plan that only converts formulas to values changes no other result and
does not mark. The mark cannot be asked for or avoided, in the command line (`--recalc-on-open`
was removed) or in the library: whoever does not know what a stale cached result is would get a
wrong file by choosing.

Memory does not grow with the workbook: a pass reads and writes each worksheet as a stream, one
cell (XLSX) or row (ODS) at a time, and the plan is a handful of bands and names (measured at
13 MB from 200,000 to 1,600,000 rows). Only the parts that are small by nature are read whole (the
workbook part, its relationships and, for a sheet whose cells move, its notes and drawings). A
plan of several passes keeps the file between them in a temporary folder.

#### References across sheets are not supported

A 3D reference (`Sheet1:Sheet3!A1`, `$S1.A1:$S3.B2` in an ODS) spans every sheet between its two
ends, which ties the sheets together: what a deletion in one of them does to it, and in which order
several deletions in different sheets act, are not implemented (issue #75). A command that deletes
rows or columns or removes sheets in a workbook with a formula that holds one (in a cell, a name, a validation or a chart)
**stops with an error that names the issue**, and writes nothing, instead of leaving the reference
wrong. Commands that move nothing (`--convert-to-values`, `--remove-ranges`) are not affected.

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

Also adjusted: the cells and their rows or columns (widths, hidden, merged areas), shared and
array formulas, defined names and print ranges, conditional formats, data validations,
hyperlinks, filters (the range shrinks around the deleted columns and their criteria go), the conditional formats, validations and sparklines of the extensions, notes (XLSX comments and threaded comments, with their shapes), manual page
breaks, the anchors of shapes and the series of the charts of an XLSX, and the database ranges and
print ranges of an ODS. A formula whose text changes loses its cached result, which
is no longer valid; one that does not change keeps it.

- The output is marked to be recalculated when it is opened, by the plan and without a flag, when
  rows, columns, sheets or names were removed (XLSX: `fullCalcOnLoad` in `workbook.xml`, which
  Excel and LibreOffice clear when they save; ODS has no such marker, so the cached results of the
  formulas are dropped). Without the mark, a formula that depends on the deleted cells without
  changing its text (`SUM(A:A)`) would keep a result that is out of date.
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
