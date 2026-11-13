# sheetrs edit

**sheetrs edit** is a command-line utility for performing operations on spreadsheet files. It handles conversions, cleaning, and structural modifications.

## Installation

`sheetrs edit` is a subcommand of the `sheetrs` executable (crate `sheetrs-cli`), which also holds `query`, `edit`, `lint` and `stats`:

```bash
cargo install --path sheetrs-cli   # from the root of the workspace
```

## Usage

### 1. Modification

Perform destructive operations. They write a new file given by `--output`, which must not be
the input file, or replace the input with `--in-place` (written to a temporary file and renamed,
so a failure leaves the input as it was).

```bash
# Delete specific sheets
sheetrs edit -i <FILE> --delete-sheets "Sheet1" "Sheet2" --output <OUT_FILE>

# Delete defined names
sheetrs edit -i <FILE> --delete-names "MyRange" "OldData" --output <OUT_FILE>

# Combined operations
sheetrs edit -i <FILE> --delete-sheets "Temp" --delete-names "TempRange" -o cleaned.xlsx

# ODS files are modified the same way
sheetrs edit -i <FILE>.ods --delete-sheets "Temp" -o cleaned.ods
```

### 2. Find & Replace

```bash
sheetrs edit -i <FILE> --replace TEXT --with TEXT [--scope SHEET | --scope 'Sheet!A1:C9']... \
    [--regex | --wildcards] [--match-case] [--entire-cells] [--formatted] [--filtered] [--in value|formula|comment] \
    (--output <OUT_FILE> | --in-place)

# A sweep: repeat the pair. Each --replace takes the --with in the same position; every pair is an
# operation, in order, that sees what the one before it left, and the options are shared
sheetrs edit -i <FILE> --replace 2024 --with 2025 --replace Q1 --with T1 -o out.xlsx
```

The counts must match: `--replace` twice and `--with` once is an error that says so.

It does what the Find & Replace of LibreOffice does (observed in **26.8**, the reference version; the
check is `tests/replace.rs`, which compares the result with the workbook LibreOffice leaves after the
same `replaceAll`, in XLSX and ODS; `tests/assets/replace/make_fixtures.py` makes them):

- Case is ignored unless `--match-case`. `--entire-cells` asks the whole text of the cell, and the
  **first** match must be the whole text (`a|ab` does not match `ab`; `a*b` does not match `abab`).
- `--wildcards`: `*` (as little as it can), `?`, and `\` to escape (not `~`, which is `COUNTIF`'s).
- `--regex`: ICU syntax, the subset that was compared with LibreOffice (groups `$1`, `&` for the
  whole match, `\t`, `\&`, `\$`, `\\` in the replacement); a pattern outside it is refused, and the anchors `^` and `$` are outside it (`sheetrs query --find` reads them: what a replacement does with them, an empty match or a match at the end of the text, is not reproduced).
- The text searched in a cell is the one of its **input line**: `1234.5` for a cell that shows
  `1,234.50`, `25.6%`, the date `03/15/2023` (the ISO date when the format of the cell is an ISO
  one). Every hit of the cell is replaced.
- A replacement leaves out the rows an autofilter has hidden, as LibreOffice does with
  `SearchFiltered` off; `--filtered` reaches them too. Rows hidden by hand are always replaced. The
  rows are those the reader reports as filtered (an ODS marks them, an XLSX counts the rows hidden
  inside the range of the autofilter of the sheet; see `sheetrs-query`). Compared with `replaceAll` of
  LibreOffice in `tests/replace.rs` (`input_filtered.*`, cases `filtered_left_out` and `filtered_included`).
- With `--formatted` the text searched is the one the cell **shows** (`1,234.50`, `15 Mar 2023`,
  `$10.00`); the text after the replacement is entered as typed, so `1,234.50` → `9,999.50` stays a
  number and `15 Mar 2023` → `20 Apr 2023` becomes text, as in LibreOffice. The text of the format
  around the number is accepted when typing (`x9` in a cell of format `"x"0` is `9`). The result of a
  formula is searched with the format its cell has in the file (the format it would take from its
  result is not worked out here).
- The replacement is **entered as if it were typed**: `12` makes a number, `50%` a percentage,
  `03/15/2024` a date, `TRUE` a boolean, `=1+1` a formula, `'12` a text, nothing an empty cell; a
  cell with the text format keeps text. A cell in `General` takes the number format LibreOffice
  gives what was typed, and keeps its style in every other respect (font, fill, borders); a cell
  with a format of its own keeps it, except that a boolean always takes the boolean format. The
  formats are those of the `en-US` locale (other locales: issue #84), checked against LibreOffice
  26.8 in both formats (`tests/replace.rs`, cases `typed_general` and `typed_formats`):

  | Typed in a `General` cell | Format given (`value type` in an ODS) |
  |---|---|
  | `50%`, `-50%`, `1.5%` | `0.00%` (`percentage`) |
  | `03/15/2024`, `1/2`, `Mar 5`, `March 15, 2024`, `15-Mar-24` | `MM/DD/YY` (`date`) |
  | `2024-03-15` | `YYYY-MM-DD` (`date`) |
  | `03/15/2024 10:20`, with `PM` or seconds | `MM/DD/YY HH:MM AM/PM` (`date`) |
  | `2024-03-15 10:20:30` / `2024-03-15T10:20` | `YYYY-MM-DD HH:MM:SS` / `YYYY-MM-DD"T"HH:MM:SS` (`date`) |
  | `12:30`, `10:20 PM`, `1:2` | `HH:MM:SS AM/PM` (`time`) |
  | `25:00`, `48:30` / `12:30:45.5` | `[HH]:MM:SS` / `[HH]:MM:SS.00` (`time`) |
  | `TRUE`, `false` | `BOOLEAN` (`boolean`) |
  | `1E3`, `1e-3`, `12.5E-2` | `0.00E+00` (`float`) |
  | `$12.50`, `($5)`, `5$` | `[$$-409]#,##0.00;[RED]-[$$-409]#,##0.00` (`currency`, USD) |
  | `1 1/2`, `-1 1/2` | `# ?/?` (`float`) |
  | `12`, `1,234.5`, `(5)`, `.5`, `00123` | stays `General` (`float`) |
  | `12.50 USD`, `15 Mar 2024`, `2024/03/15`, `€5` | stays `General`, text |
- `--in formula` searches the formula in LibreOffice's notation (`=SUM($'My Sheet'.A1,B1)`, string
  literals included) and writes the result in the notation of the file (Excel's in an XLSX,
  OpenFormula's in an ODS). A shared formula group whose first cell is replaced gets a formula of
  its own in each cell. A reference to another workbook is searched as LibreOffice shows it,
  `'file:///folder/test.xlsx'#$Sheet1.B3`, with the file resolved against the folder of the
  document (`tests/replace.rs`, `a_formula_with_an_external_reference_...`), and it is written back
  as the file writes it (`[1]Sheet1!B3` in an XLSX, the address of the file in an ODS).
- `--in comment` replaces in the text of the notes (XLSX legacy comments and ODS annotations), as the
  `replaceAll` of LibreOffice does: the author stays, an XLSX note is saved as one run (Arial 10) and
  an ODS one as plain paragraphs. Verified against notes LibreOffice made, in both formats
  (`tests/replace.rs`, `tests/notes.rs`).
- Rewritten parts of the archive are deflated at level 4 (the size of level 6, a third less time).
- A search by value reaches formula cells, which become constants, as in LibreOffice.
- What the replacement leaves stale is calculated again by the engine as the last pass; the output is marked to be recalculated when opened only if the engine cannot calculate the workbook (see *Recalculation* below).

A text typed as a formula is compiled as LibreOffice compiles it (observed in 26.8,
`tests/assets/replace/make_fixtures.py`), and the cell keeps what it understood:

| Typed | The cell holds | Error |
| ----- | -------------- | ----- |
| `SUM(1,2` | `SUM(1,2)` (what is left open at the end is closed) | none |
| `foo(1)`, `A1+B`, `$Nope.A1` | the same in lowercase: `foo(1)`, `A1+b`, `$nope.a1` | `#NAME?` |
| `SUM(1,2))`, `A1)` | the text as typed | `Err:508` |
| `A1 B1`, `1,2` | the text as typed | `Err:509` |
| `1+*2` | the text as typed | `Err:510` |
| `1+(`, `((` | `1+()`, `(())` | `Err:511` |
| `A1+`, `A1:`, `SUM(A1:A3)*` | the text as typed | `Err:520` |
| `#REF!+1` | the text as typed | `#REF!` |

An XLSX writes the errors as Excel's (`Err:508`–`Err:511` are `#VALUE!`, `Err:520` is `#N/A`), an ODS
as `Err:5xx` (`#NAME?` and `#REF!` as such). The probed classes are the ones above; any other
malformed text is given the error of the first rule it breaks and is not checked against LibreOffice.

**Not reproduced yet.** These are differences from LibreOffice that nobody approved, so they are
bugs waiting for a fix or a decision (see the pending decisions of the project): the cells of an
array formula are left as they were, as LibreOffice's `replaceAll` leaves them (and the command says so).

The reference behavior of every operation is **LibreOffice 26.8 on Linux**: what a formula
becomes after an operation is what LibreOffice gives when the same operation is made in it
(`docs/FIDELITY.md` of `sheetrs-compiler` lists where Excel differs). The check is made on XLSX
and ODS files against LibreOffice: `sheetrs-edit/scripts/compare_structure_with_libreoffice.py` makes
each operation in LibreOffice through UNO and with `sheetrs edit` on the same file (a band as one
band, two rows given to one command as the band they are, and one row after the other as two saved
steps, which is what two `removeByIndex` are), opens both results in LibreOffice and compares
formulas, values, notes, page breaks and shapes. A manual page break before the first row is not
compared: LibreOffice keeps it in memory after a deletion (`T!1:2` leaves the break of row 3 at row 1)
and reads none from a file, which has no way to say it. Excel is not part of it.

- Deleting a sheet also deletes the named ranges local to it. A workbook keeps
  at least one sheet.
- The references to a deleted sheet become `#REF!` (`SUM(Data!A1:A3)` is `SUM(#REF!)`), and a
  name that pointed at it keeps a definition that is `#REF!`.
- A deleted name stays in the formulas that use it, as in LibreOffice, where it is the
  unknown name `#NAME?`; another name whose definition used it has `#NAME?` there.
- The operations of one command are applied together, reading each sheet
  once. They are not run in the order they are written, in any layer: they are planned (below).

#### How the operations are planned

**The operations are not a sequence, in any layer** (the command line, the Rust library, Python and
WebAssembly). Each one is written against the workbook **as it is** (its rows, its columns and its
sheets), and the planner decides the order they run in, so that nobody has to think about how the
workbook looks after each step: `--delete-rows "Model!3:4" "Model!7"` deletes rows 3, 4 and 7 of the
file, `delete_rows("Model!3")` and `delete_rows("Model!4")` pushed in any order and split in any
number of calls delete the same rows 3 and 4, and a `--formula-to-value` of the same command
reads the file's rows too. Pushing the same operations in another order gives the same file, byte for
byte. The rule is one canonical order:

1. the cells written (the template of an answer) and the replacements, each decided against the
   workbook as the ones before it leave it; the operations of one kind keep the order they were
   given in (a sweep of replacements is a sequence of texts: each one sees what the previous one
   left);
2. the calculation: the inputs are written and every formula calculated (one step; several keep
   their order);
3. in one pass over the file: formulas converted to values, while every cached result is still there
   (a conversion on a sheet that is removed goes with it); in each sheet, rows deleted **from the
   last to the first**, overlapping or touching ranges merged, so a deletion shifts only what is
   below it; in each sheet, columns deleted the same way; names deleted; sheets deleted (the
   deletions of a deleted sheet go with it);
4. when the deletions left cached results stale, the formulas that are stale (see *Recalculation*
   below) are given their new result as the last pass: `#REF!` written directly where it follows from
   a deleted reference, the engine for the others; only if the engine cannot is the workbook marked
   to be recalculated when it is opened, and the run says why.

**The planner takes as many passes as the operations need, and no plan fails because of how its
operations were grouped.** It does not promise one pass over the file and give up otherwise: it
promises the cheapest path, and when that is several passes (a replacement, a calculation, the
conversion and the deletions, the recalculation of what they left stale) it runs them, writing the
workbook between passes in a temporary file (or in memory, for `write_to`). What a conversion reads
is a decision of the semantics and not of the grouping: every operation describes the workbook as it
is, so a conversion pushed with a deletion converts the values the workbook has, whichever was pushed
first; the values after a deletion are what a second `write` of the saved workbook converts (its
results are stale and are calculated again first). Two deletions that touch are one band, which the
writer does for any band (the shapes in it are placed by the sizes of the rows and columns: see
below), so they never need to be split.

A conversion reads the results the workbook caches. When what runs before it (a replacement) left
them stale, or the file itself says it is stale, they are calculated again first, with the engine of
`LibreOffice`'s semantics, in a pass of its own. A conversion **does not** see a deletion pushed with
it: the deletion runs after, as in `--formula-to-value F --delete-rows T!4`. To convert what a
deletion leaves (what `LibreOffice` shows after deleting and then converting) make two cuts: save the
deletion and convert in the next command or the next `write`; the saved workbook is the input of the
second step, which finds its results stale and calculates them again. There is no sequential mode.

Operations on different sheets do not depend on each other, because a reference moves only with the
rows and columns of the sheet it names. That holds because references across sheets are not
supported (below). `--dry-run` prints the passes. `--locale TAG` (and `--config FILE`, whose `locale` in `[global]` is its default; `sheetrs.toml` of the current directory when there is one) is the locale of the calculation that an edit leaves to do (`TEXT`, the conversion of a number to text): by default the one of the system (`LC_ALL`, `LC_CTYPE` or `LANG`; `en-US` when there is none), for an XLSX and an ODS alike. The typed replacement of `--replace` still reads its text in `en-US` (issue #84). `--now SERIAL` and `--seed N` fix `NOW()`/`TODAY()` and `RAND()` in that calculation (the `now` and `seed` of `sheetc`, and of a `calculate` step of the library); without them the engine reads the clock, as a spreadsheet does.

**As a library** (`sheetrs_edit::BookEdit`, a trait for `sheetrs_io::Book`), the changes are collected
in the `Book` (`Book::from_file(..)?`, `Book::from_bytes(..)?`; its operations are `book.operations()`)
and each one is written against the workbook as it is, as above. `write(path)` (to a path),
`write_in_place()` (the original, replaced atomically) and `write_to(impl io::Write)` (memory) run
them in the fewest passes over the file, and `explain()` says what the write will do (`explain(&operations)`
says it for a list of operations, with no workbook). The planner (`sheetrs_edit::plan`, private) is a set
of functions that take the operations of the `Book`; the writers are `sheetrs_io::write`. Tests
compare the result of generated sets of operations with the writer given all of them at once, in XLSX
and ODS, and push the same operations in every order and compare the bytes.

```rust
use sheetrs_edit::BookEdit;
use sheetrs_io::Book;

let mut book = Book::from_file("in.xlsx")?;
book.delete_rows(["Data!3:10".parse()?, "Data!15".parse()?])   // the array of its targets
    .delete_sheets(["Scratch", "Tmp"])
    .replace([], [pair("2024", "2025"), pair("Q1", "T1")])   // a sweep, one operation per pair, in order; `pair` builds a `Replace`
    .write("out.xlsx")?;
```

Every operation takes the array of its targets (an `IntoIterator`), because the tool is for sweeps.
`replace(scopes, replaces)` takes the scopes shared by the sequence of `Replace` specs, and each
spec is an operation as if `replace` were called once per spec: it sees what the one before it left,
so the order of the specs matters. The options of a spec are its own (a replacement that needs
other options is another element); one replacement is `[spec]`.

**Recalculation is the planner's, with no option for it, and the mark is never decided beforehand.**
A plan that deletes rows, columns, sheets or names, or replaces text, ends with a pass in which the engine
(the semantics of LibreOffice) calculates every formula again. The result a file cached is not a goal
(see [approved differences](../docs/approved-differences.md)): what a cell holds is the value its formula
gives now, whatever the workbook cached, so no analysis decides which formulas are stale and the plan is
the same for every workbook.

The workbook is written with fresh results and no mark. Only when the engine cannot calculate it (a
function it does not have, a reference it does not read) is the workbook marked to be recalculated
when it is opened (`fullCalcOnLoad` in an XLSX; an ODS has no marker, so the cached results of its
formulas are dropped) and the run says why (a `Diagnostic` with the code `RECALC_NOT_SUPPORTED` in `Report::diagnostics`,
written by the command line to the standard error): the spreadsheet program that opens it does the recalculation then. Nothing
assigns the mark before the engine has run, and nothing takes it off except the engine having
calculated a workbook that came with it. A run that only converts formulas to values changes no other
result and has nothing to calculate. The mark cannot be asked for or avoided, in the command line or in the library. What the engine costs is in
[`docs/performance.md`](../docs/performance.md#sheetrs-edit). `--dry-run` says the passes and the calculation that ends them.

**`calculate` is an operation, and `sheetrs edit` has no flag for it.** In the library,
`BookEdit::calculate(Calculation { inputs, ranges, major, answer, now, seed })` writes the inputs and calculates every
formula of the workbook as it is, before the conversions and the deletions (*Data ▸ Calculate* in
LibreOffice). The engine is the one of `sheetrs-compiler` (the semantics of LibreOffice), which this
crate depends on: nothing is handed in, and the same one recalculates the results that a conversion of
formulas to values reads after a change that leaves them stale.

- **The engine can calculate the workbook:** it is written with fresh results and **without** the mark
  to recalculate it when it is opened (unless a deletion that comes after it leaves results stale).
- **It cannot** (a function it does not have, a reference it does not read): the run goes on, the workbook is written as the other operations leave it and **marked** to be
  recalculated when it is opened (the same mark as above: `fullCalcOnLoad` in an XLSX, the cached
  results dropped in an ODS; LibreOffice recalculates both when it opens them, checked on 26.8), and the
  `Report` the write returns holds a `Diagnostic` (level `warning`, code `RECALC_NOT_SUPPORTED`, context
  fields `operation` and `reason`): the recalculation that is needed will be done by the spreadsheet
  program that opens the file later. The command line writes it to the standard error (see
  *Diagnostics* in `sheetrs-protocol`).
- **The value ranges** (`ranges`: ranges, in the cells of the workbook as it is; the workbook's own outputs
  when omitted) are kept in the `Report` when the `Calculation` asks for the `answer`: `Report::value_ranges` is
  the answer of the protocol (`sheetrs_protocol::ValueRanges`, `{"valueRanges": [...]}`) and `Report::write_json(out)` writes it as JSON, the way `write` writes
  the workbook; to read them without writing the workbook anywhere, `write_to(std::io::sink())`. With
  the answer asked for, a workbook the engine cannot calculate is an error (`Error::Calculation`),
  since there is no answer to give; `write_json` of a report that kept none is `Error::NoCalculation`.

**The clock and the seed.** `Calculation::now` (the serial number of `NOW()`, from which `TODAY()` follows) and
`Calculation::seed` (`RAND()`) are the `now` and `seed` of `sheetc`, and reach the engine as the options of the
build and of the run it already has; `None` is the system clock. They are the clock of the whole plan, not only
of the step: the recalculation that the planner adds after what leaves results stale (a deletion, a
replacement) takes the `now` and `seed` of the last `calculate` step that sets them, so a plan that fixes the
clock gives the same workbook whichever pass calculates. Where a plan has no `calculate` step the recalculation
takes the ones given to the book (`Book::with_volatile`, `sheetrs_io::Volatile`; `sheetrs edit --now --seed`
set it), and a step's own value wins over the book's. The seed is the engine's: an even seed and the odd one
after it give the same sequence.

`sheetrs edit` has no `--calculate` flag: calculating is a step the planner inserts where a conversion needs fresh results (`BookEdit::calculate` is the explicit step in the library). To calculate a file with inputs and ranges, `sheetc calculate -i FILE -o OUT` (with `--request` to give them).

A pass reads and writes each worksheet as a stream, one
cell (XLSX) or row (ODS) at a time, and the plan is a handful of bands and names, so a plan with no recalculation
(a conversion of an ODS, whose results are stored) keeps the memory flat; a plan that leaves formulas stale that only the engine can calculate ends with
the engine running, which holds the formulas of the workbook (the times and the memory per scenario are in
[`docs/performance.md`](../docs/performance.md#sheetrs-edit)). Only the parts that are small by nature are read whole (the
workbook part, its relationships and, for a sheet whose cells move, its notes and drawings). A
plan of several passes keeps the file between them in a temporary folder.

#### References across sheets are not supported

A 3D reference (`Sheet1:Sheet3!A1`, `$S1.A1:$S3.B2` in an ODS) spans every sheet between its two
ends, which ties the sheets together: what a deletion in one of them does to it, and in which order
several deletions in different sheets act, are not implemented (issue #75). A command that deletes
rows or columns or removes sheets in a workbook with a formula that holds one (in a cell, a name, a validation or a chart)
**stops with an error that names the issue**, and writes nothing, instead of leaving the reference
wrong. Commands that move nothing (`--formula-to-value`, `--delete-names`) are not affected.

#### Delete rows and columns

```bash
# Rows 10 to 20 and column B of one sheet; `Sheet!7` and `Sheet!C` delete one
sheetrs edit -i book.xlsx --delete-rows "Model!10:20" --delete-columns "Model!B" -o shorter.xlsx
```

Rows and columns are those of the input file, whatever the order and the other operations of
the command: `--delete-rows "Model!3:4" "Model!7"` deletes rows 3, 4 and 7 of the file, and a
`--formula-to-value` of the same command reads the file's coordinates too.

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

- When rows, columns, sheets or names were removed, the engine calculates the formulas again as the
  last pass, so a formula that depends on the deleted cells without changing its text (`SUM(A:A)`)
  has the result it has now. If the engine cannot calculate the workbook, the output is marked to be
  recalculated when it is opened (XLSX: `fullCalcOnLoad` in `workbook.xml`, which Excel and LibreOffice
  clear when they save; ODS has no such marker, so the cached results of the formulas are dropped) and
  the run warns.
- A deletion stops with an error when it cannot be made right. In both formats: it covers part of
  an array formula, or removes the first cell of a merged area that stays (LibreOffice does it; not
  reproduced yet). In an
  XLSX: a sheet that changes has tables, header or footer images or data tables; the workbook has pivot tables. In an ODS: the document has pivot tables, filters, consolidations or
  scenarios.
- Shapes, images and charts (the drawings of an XLSX, the shapes in the cells of an ODS) are placed
  as LibreOffice places them (`sheetrs_io::write::shapes`, from `ScDrawLayer`, checked against 26.8
  on every band of rows and of columns up to five wide of the fixtures:
  `tests/assets/oracle/shapes_delete.json`, made by `scripts/make_shapes_oracle.py`). A shape that
  lies wholly inside the deleted band is deleted with it (its outline counts: one that starts where the
  band starts stays). The cell it is anchored to keeps its index when it is deleted and moves up
  otherwise; a rectangle, picture or chart whose two ends cross has them put in order. Its size, when
  it resizes with its cells, comes from its end cell, or is kept (scaled by the size of its cell) when
  both ends are in one cell; a shape that does not resize keeps its size. All of it by the sizes of
  the rows and columns, which are in the file: the `ht` of each row and the `width` of each column of
  an XLSX (the width of a column is in characters, so a column that differs from the default is an
  approximation by the width of a digit of 7.4 pixels), the styles of the rows and columns of an ODS.
- A bare `#REF!` in an XLSX formula is read by LibreOffice as an unknown name (`#NAME?`), also
  in the files it saves itself; a deleted cell is written `Sheet!#REF!`, which it reads as `#REF!`.
  The references to a removed sheet have no sheet to write: they are the bare `#REF!`.

#### Convert formulas to values

`--formula-to-value` replaces formulas by their values, e.g. to turn a
master workbook into a production run model where argument cells hold fixed
values and need no computation. Each scope selects a whole sheet, columns or
rows:

```bash
# Whole sheet, columns B to D, and rows 3 to 10
sheetrs edit -i master.xlsx --formula-to-value "Inputs" "Model!B:D" "Model!3:10" -o run.xlsx

# Single column or row; quote sheet names containing '!' as in formulas
sheetrs edit -i master.xlsx --formula-to-value "Model!F" "'R!D'!2" -o run.xlsx
```

- Values are the results the spreadsheet application stored when it last saved the file, as long as
  they are fresh. When they are not (an XLSX marked to recalculate when it is opened, an ODS with no
  stored result in its formulas, or a change earlier in the same run that leaves results stale: a
  deletion, a replacement), the results are calculated again with the engine before they are
  converted, as an application that calculates as it goes shows them, so a run such as
  `--replace A --with B --formula-to-value S` in one invocation converts what the formulas give
  after the replacement. When the engine cannot calculate the workbook (a function or syntax it does
  not have), the conversion is an error and nothing is written. A formula without a stored value
  keeps its formula.
- Formulas that return text become text cells. Error results become empty
  texts, as in LibreOffice's *Data ▸ Calculate ▸ Formula to Value* (checked on
  `=1/0`, `=NA()`, `=SQRT(-1)` and an error the file caches), in XLSX and in
  ODS. A bare `#REF!` that LibreOffice reads as an unknown name is an error
  result too, so it is an empty text as well.
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
- `--delete-names` keeps tables, which sheetrs lists as named ranges (#63).

ODS only:

- Repeated rows and cells (`table:number-rows-repeated`,
  `table:number-columns-repeated`) are split where a scope starts or ends, and shortened where
  a deletion cuts them.
- `--delete-names` also removes database ranges, which sheetrs lists as named
  ranges.

### 3. Dry Run

Preview what operations would be performed (the passes `explain()` gives in the library) without writing to disk.

```bash
sheetrs edit -i <FILE> --delete-sheets "Old" --dry-run                      # the passes; no --output or --in-place needed, nothing is written
sheetrs edit -i <FILE> --delete-sheets "Old" --output out.xlsx --dry-run   # also says where the result would go; still writes nothing
```

## Supported Formats

- **Input**: XLSX and ODS.
- **Output (Modification)**: the input format. Every modification supports
  XLSX and ODS (LibreOffice).

## How it is tested

Each operation is judged against `LibreOffice` 26.8 in each format on its own:

- `tests/book.rs` runs the operations of `tests/assets/book/cases.json` (delete rows and columns,
  delete sheets and names, formula to value, replace, write cells, save) on the integrated
  workbook `tests/assets/book/book.{xlsx,ods}` (merged cells, hidden rows, columns and sheet, array
  formulas, a conditional format, a validation, a hyperlink, a note, a chart, local and global
  names, an external reference, rich text) and compares the result with
  `expected_<case>.*`, the workbook `LibreOffice` leaves after doing the same operation: sheets,
  names, structure and every cell. `scripts/gen_book_fixtures.py` makes all of it (with the Python of
  LibreOffice and `scripts/libreoffice-server.sh` listening).
- `tests/replace.rs` does the same for Find & Replace on `tests/assets/replace/`.
- The tests of rows, columns, sheets and names deleted start from `tests/assets/structure/`
  (`scripts/gen_structure_fixtures.py`).


XLSX, deleting the cell a chart is anchored to.

## Hide and show

`--hide-rows`, `--show-rows`, `--hide-columns`, `--show-columns`, `--hide-sheets` and `--show-sheets`
(`hide_rows`, `show_rows`, … in the Rust, Python and WebAssembly `Book`) change the visibility of
rows, columns and sheets, and keep them in the file (`hidden` of `<row>`, `<col>` and `<sheet>` in an
XLSX; `table:visibility` and the `table:display` of the table style in an ODS).

- The rows and columns are those of the workbook as it is: the changes run in a pass of their own
  before any deletion, so `--delete-rows T!3 --hide-rows T!5` hides the row that is the fifth of the
  input. A later change of the same row wins; on the command line `--show-*` is applied after
  `--hide-*`.
- A row hidden here is hidden by hand: a search or replacement with `filtered` off still reaches it.
  Showing a row shows it whatever hid it (by hand or by a filter).
- A hidden row changes what `SUBTOTAL` and `AGGREGATE` add up, so the results are calculated again
  after it; hiding columns or sheets changes no result.
- One sheet must stay visible (the run fails otherwise). Hiding a very hidden XLSX sheet leaves it very
  hidden; showing makes it visible. If the sheet that opens first is hidden, the workbook opens on the
  first visible one.
- A row or column that the file does not declare is added (a hidden `<row>`, a `<col>`).

Checked by opening the results in LibreOffice 26.8 (rows, columns and sheet states read back).
