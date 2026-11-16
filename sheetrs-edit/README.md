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
- An empty `--replace ''` matches nothing, in every mode (literal, regex, wildcards; with or without
  `--entire-cells`; with an empty or a non-empty `--with`), as `replaceAll` does: the run is answered
  at once, with no error (not even for a sheet in `--scope` that is not there), no cell is read or
  replaced (`Replaced in 0 cells`) and the report is empty (probed in XLSX and ODS,
  `tests/replace_anchors.rs`).
- `--regex`: ICU syntax, the subset that was compared with LibreOffice (groups `$1`, `&` for the
  whole match, `\t`, `\&`, `\$`, `\\` in the replacement); a pattern outside it is refused. The anchors `^` and `$` anchor the text of the cell, not its lines (`^` is its start; `$` its end or the place before the line terminator that ends it), and a pattern may match the empty text (`x*`, `^`, `$`, `a*$`). A replacement does what LibreOffice does: it is made in the text itself and the search goes on after it, in what is left (`^a` replaced by nothing empties `aa`); a match of no characters is not replaced, except at the very end of the text (`abc` with `$` replaced by `X` gives `abcX`; with `^` or `x*` it does not change) and no search begins at the end of the text (`a` with `a*` replaced by `X` gives `X`; `a` and a line break with `$` is not changed, because the one empty match, before the line break, is skipped). LibreOffice counts in UTF-16 units, and so does the replacement after a skipped empty match (a cell that ends in an emoji is not given a `$` match by `x*`). Approved difference (`docs/approved-differences.md`): when a pattern that skipped an empty match before a character outside the Basic Multilingual Plane can match half of it (`^|.`), LibreOffice writes a text that is not valid (it brought a LibreOffice server down in the probe); the replacement goes on after the whole character, because that text cannot be stored. Compared with `Sheet.replaceAll` in XLSX and ODS: `tests/replace_anchors.rs` (oracle `tests/assets/replace_anchors/oracle.json`, made by `make_anchors_oracle.py`).
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
2. the calculation: the inputs are written and the formulas they reach calculated, or every formula with `calculate_all` (one
   step; several keep their order);
3. in one pass over the file: formulas converted to values (the cells they name are calculated by the
   engine and written as values in that same pass; a conversion on a sheet that is removed goes with
   it); in each sheet, rows deleted **from the
   last to the first**, overlapping or touching ranges merged, so a deletion shifts only what is
   below it; in each sheet, columns deleted the same way; names deleted; sheets deleted (the
   deletions of a deleted sheet go with it);
4. when the deletions left cached results stale, the formulas that are stale (see *Recalculation*
   below) are given their new result as the last pass by the engine (a reference that went is `#REF!`
   in the functions that read it, see *Delete rows and columns*); only if the engine cannot is the workbook marked
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

A conversion converts, by default, **what LibreOffice converts**: the results the file caches for the cells nothing reaches
(*Data ▸ Calculate ▸ Formula to Value* reads the value a cell holds), and what the engine calculates for the cells the plan
reaches (a replacement pushed with it, a deletion, volatile formulas; and, in an ODS, the formulas `LibreOffice` calculates when it opens the file:
*The startup of an ODS*, below; *The cache is the default*, below). With
`--calculate-all` (`calculate_all=True` in Python, `calculateAll = true` in WebAssembly) it converts **what the formulas
calculate**: the engine (the semantics of `LibreOffice`) calculates the cells it names and what they read in an XLSX, and every formula in an ODS (its startup),
and the pass that writes the file writes them as values (so a conversion costs
the engine's reading of the workbook, not a calculation of every formula followed by a pass of its own). **A conversion is one pass in every mode**: where the
engine runs for it (always for an ODS; for an XLSX when its results are stale or the run calculates every formula) it calculates in the pass that writes the
conversion, and the cells of the scopes that it did not calculate are converted with the result the file caches, which the same pass reads. What runs before
it (a replacement) is seen. A cell that depends on something the engine cannot calculate (another workbook, a macro, an extension:
*What the engine cannot calculate*, below) keeps the result the file caches, as `LibreOffice` shows it, and the report says so; one that the run calculates and
that only `LibreOffice` can give a result for (a macro, a function it has and the engine lacks) is an error that names the cause. An XLSX whose results are right
and that keeps the cache converts the cached results and needs no engine. Two scenarios, both documented and measured (`docs/performance.md`, *sheetrs edit*;
2.4 M formulas, 300 000 rows).
A conversion **does not** see a deletion pushed with
it: the deletion runs after, as in `--formula-to-value F --delete-rows T!4`. To convert what a
deletion leaves (what `LibreOffice` shows after deleting and then converting) make two cuts: save the
deletion and convert in the next command or the next `write`; the saved workbook is the input of the
second step. There is no sequential mode.

Operations on different sheets do not depend on each other, because a reference moves only with the
rows and columns of the sheet it names. That holds because references across sheets are not
supported (below). `--dry-run` prints the passes. `--locale TAG` (and `--config FILE`, whose `locale`, `now` and `seed` in `[global]` are the defaults of `--locale`, `--now` and `--seed`; `sheetrs.toml` of the current directory when there is one) is the locale of the calculation that an edit leaves to do (`TEXT`, the conversion of a number to text): by default the one of the system (`LC_ALL`, `LC_CTYPE` or `LANG`; `en-US` when there is none), for an XLSX and an ODS alike. The typed replacement of `--replace` still reads its text in `en-US` (issue #84). `--now SERIAL` and `--seed N` fix `NOW()`/`TODAY()` and `RAND()` in that calculation (the `now` and `seed` of `sheetc`, and of a `calculate` step of the library; their defaults are the `now` and `seed` of `[global]` in the configuration file); without any the engine reads the clock, as a spreadsheet does. `--external-from-cache`, `--macro-from-cache` and `--extension-from-cache` (defaults: the keys of the same name in `[global]` of the configuration file; off) are explained in *What the engine cannot calculate*.

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

**Recalculation is the planner's, and the mark is never decided beforehand.**
A plan that deletes rows, columns, sheets or names, or replaces text, ends with a pass in which the engine
(the semantics of LibreOffice) calculates the formulas the plan reaches and what depends on them and leaves the others
as the file caches them (the default, as LibreOffice does), or, with `--calculate-all`, calculates every formula
again (*The cache is the default*, below). The plan is the same for every workbook; what the engine is asked for in its
last pass depends on the mode.

The workbook is written with fresh results and no mark. Only when the engine cannot calculate it (a
syntax it does not read; **not** for a macro or a function only `LibreOffice` has in a cell the run calculates, which is an error: *What the engine cannot
calculate*) is the workbook marked to be recalculated
when it is opened (`fullCalcOnLoad` in an XLSX, which `sheetc` reads too: a marked workbook is a run with `calculate_all` in both; an ODS has no marker, so the cached results of its
formulas are dropped) and the run says why (a `Diagnostic` with the code `RECALC_NOT_SUPPORTED` in `Report::diagnostics`,
written by the command line to the standard error): the spreadsheet program that opens it does the recalculation then. Nothing
assigns the mark before the engine has run, and nothing takes it off except the engine having
calculated a workbook that came with it. A run that only converts formulas to values changes no other
result and has nothing to calculate. The mark cannot be asked for or avoided, in the command line or in the library. What the engine costs is in
[`docs/performance.md`](../docs/performance.md#sheetrs-edit). `--dry-run` says the passes and the calculation that ends them.

**What the engine cannot calculate.** Three things a workbook can depend on that the engine does not reproduce, each
detected explicitly while the formulas are read (the cause is a typed `sheetrs_compiler::Unsupported`). There is no option for them: the engine
does what `LibreOffice` does, which was observed (the table is in [`FIDELITY.md`](../sheetrs-compiler/docs/FIDELITY.md), *What depends on what the engine cannot
calculate*, with `tests/assets/from_cache/probe.py`):

| Kind | Detected as | A cell no edit reaches | A cell the run calculates | Warning in the `Report` |
| ---- | ----------- | ---------------------- | ------------------------- | ----------------------- |
| external | a reference to another workbook (`[1]Sheet1!A1` in an XLSX; an ODS is read the same way) or a defined name that holds one | the result the file caches (the stored value of the link), also with `calculate_all` | the same (a formula that mixes it with local cells that an edit reaches keeps it too) | `EXTERNAL_FROM_CACHE` |
| macro | a workbook that carries macros (`vbaProject.bin`, macrosheets, a `Basic/` or `Scripts/` library, event macros) and a function that no spreadsheet program has, which a macro may define; the macros are never run | the result the file caches | an error that names the function (`LibreOffice` gives `#MACRO?` or the macro's value, by the security level) | `MACRO_FROM_CACHE` |
| extension | a function the engine does not implement: an add-in that is not installed (`_xll.NAME`, `COM.SUN.STAR.SHEET.ADDIN.*`) or a function of `LibreOffice` it lacks (`WEBSERVICE`, `FILTERXML`, `DDE`) | the result the file caches | an add-in: `#NAME?`, as `LibreOffice`; `WEBSERVICE`, `FILTERXML`, `DDE`: an error that names the function (their result is the one only `LibreOffice` gives) | `EXTENSION_FROM_CACHE` |

A cell that no edit reaches keeps the result the file caches, as it is, and what reads it is calculated from it; the `Report` carries one warning per kind (`Diagnostic`,
level warning) with the number of formulas, the first cell and what they depend on (where the engine runs: an XLSX conversion that keeps the cache and has
nothing stale converts the file's results and runs no engine, so it says nothing). The result the file caches is the one of the program that wrote it: an external
workbook may have changed, a macro may give another value, an extension may compute what the cache does not say, and `LibreOffice` keeps it all the same until
something dirties the cell. A workbook that carries macros is not refused (`sheetc` does refuse it unless `--no-strict`, a policy of a build). A name no
spreadsheet program has, in a workbook without macros, is `#NAME?`, as in `LibreOffice`.

**The startup of an ODS.** `LibreOffice` runs the rule of the load every time it opens an ODS (the formulas that cache a number in a cell with no format, a cached error,
no result, a few functions), so `sheetrs edit` on an ODS **starts** with that calculation, in the mode the run was asked: the roots of the load by default, every formula with
`--calculate-all`. The roots are the only difference between the modes at the start. The startup is the first pass that runs the engine, and a pass of its own, the first,
when nothing else in the plan needs the engine (hiding columns, hiding a sheet); `--dry-run` says it first (`startup of an ODS: ...`). The roots are found in the one read that
compiles the workbook: no scan of the file comes before. After the startup each operation applies its own rule (a deletion reaches a cone in the default mode and every formula with
`--calculate-all`; a conversion writes what `LibreOffice` has calculated; hiding or showing rows reaches the cells that read them). An ODS has no mode of its own: the one of the run
decides the startup and the operations. An XLSX has no startup: it keeps the functions of the rule only. OBSERVED (`oracle.json`, every fixture, both formats): an ODS edit
that leaves nothing stale (`book.ods`, hide `Data!B`) shows the numbers `LibreOffice` calculates on load (`rate*2` 14, `ROW()` 8) and keeps a cached text (`from-excel`).

**The cache is the default, `--calculate-all` calculates everything.** What `LibreOffice` does when it opens a file and
when a cell changes is to keep the results the file caches and to calculate what the change reaches: that is what
`sheetrs edit` does by default, and `--calculate-all` (`Book::with_calculate_all(true)`, `calculate_all=True` in Python,
`calculateAll = true` in WebAssembly) emulates `calculateAll()`: every formula is calculated. By default the recalculation
calculates only the formulas the plan reaches and the ones that depend on them; every other formula is read as a constant
that holds the result the file caches, and keeps it, byte for byte (what LibreOffice saves of a workbook it opened and
edited; the probes and the oracle that compares them, `tests/assets/write_to/make_oracle.py`: `reread_default`). A formula is
reached when:

- the writer left it with no cached result (the formula read a deleted cell or a range that shrank, so the
  plan rewrote it; a formula the file never gave a result, as in an ODS written by a program that does not store them);
- `LibreOffice` calculates it when it opens the file, whatever it caches (in both formats: `HYPERLINK`, every type of `CELL`, `INFO`,
  `FORMULATEXT`, the volatile functions, `INDIRECT` and `OFFSET`, and what reads them; in an ODS also a formula that caches a number, a currency
  or a date in a cell with no number format, a cached error, an array formula, or a result with no type, and not what reads those: *What
  LibreOffice calculates when it opens a file*, in `FIDELITY.md`), and a run that keeps the cache does the same, even when no edit reaches
  anything: an ODS starts with them (*The startup of an ODS*, below), an XLSX that holds one of those functions ends with the recalculation that load does; a
  cell that depends on another workbook, a macro or an extension the engine cannot calculate keeps the result the file caches;
- it reads a cell a replacement wrote, or reads a range with a row that was hidden or shown and is a `SUBTOTAL`, or sits in a group of
  adjacent formulas of one shape (`LibreOffice` groups them; hiding calculates the members that read the row);
- rows or columns were deleted and it calls `ROW` or `COLUMN` and its own cell moved, or the cell or range it is given moved
  or went (`ROWS`, `COLUMNS`, `SHEET`, `ISFORMULA`... keep their result when cells only move), or it reads a range that
  goes to the last row (or column) of a sheet that lost rows (or columns), or it reads through a defined name whose target the
  deletion moved, shrank or removed (a band that begins at or before the last row, or column, of the target; a name made of
  other names is reached when one of them loses cells, not when it only moves);
- a sheet was deleted: every formula is reached (`LibreOffice` calculates all of them);
- a defined name was deleted and it reads through a name (any name; in an XLSX also one that holds a bare `#REF!`).

What `LibreOffice` keeps is kept too: a formula that only reads cells that moved, a formula that already held an error `#REF!`
(`COUNT(#REF!)` caching 777 stays 777 when rows go), a formula of another sheet. These rules are what the probes of
`tests/assets/write_to/probe_dirty.py` show for both formats (the same in an XLSX and in an ODS that caches text) and the oracle
`oracle.json` (fixtures `dirty`, `classes` and `indirect` and the older ones) holds against `LibreOffice` re-read; `write_to_libreoffice.rs` runs
the default mode now against it for every fixture in both formats; they are in
[`sheetrs-compiler/docs/FIDELITY.md`](../sheetrs-compiler/docs/FIDELITY.md) (*What LibreOffice calculates when it opens a file*, *What a deletion reaches*).

What depends on those is found with the dependency graph the compiler already builds (through names, ranges and whole
columns), and is calculated too: for the cells reached the result is that of `--calculate-all`. The run calculates **every**
formula, as `--calculate-all` does, where the file says its results are stale (an XLSX marked to be
recalculated when it is opened), where a replacement and a deletion meet (the cells of the
replacement are not where they were). A workbook that reads cells at run time (`INDIRECT`, `OFFSET` with an argument that is not known) is not
calculated whole: those formulas are calculated, as `LibreOffice` does, and the cells they may read hold what the file caches. However many formulas are stale, the others keep their result (an earlier version calculated
everything where more than a quarter were; the walk that finds the cone costs 1-1.3 s on 2.4 M formulas, see `docs/performance.md`). The result a file caches is not a goal of
the project's fidelity (see [approved differences](../docs/approved-differences.md)) but the default is LibreOffice's, and
what `--calculate-all` gives is the choice of the caller who wants every formula to hold what the engine gives (the issue that
asked for the cache, [#126](https://github.com/cosmoscalibur/sheetrs/issues/126), shows cells where Excel and the engine
differ, and a `TEXT` calculated with another locale than the one the workbook was written with). What the cache costs and
saves is in [`docs/performance.md`](../docs/performance.md): the engine, the writing of the
results and the memory of the calculation shrink with what is not reached; the reading of the workbook and the passes of the plan do not.

**`calculate` is an operation, and `sheetrs edit` has no flag for it.** In the library,
`BookEdit::calculate(Calculation { inputs, ranges, major, answer, now, seed })` writes the inputs and calculates the
formulas they reach (every formula with `calculate_all`, as `Book::with_calculate_all(true)` asks) before the conversions and
the deletions (*Data ▸ Calculate* in LibreOffice): the inputs it is given are the declared inputs of the build of the engine's
program, so what they do not reach keeps the result the file caches. A calculation has the mode of the run, as `sheetc
calculate` does; `force_calculate_all` does not apply to it. The engine is the one of `sheetrs-compiler` (the semantics of LibreOffice), which this
crate depends on: nothing is handed in, and the same one calculates the cells that a conversion of
formulas to values converts and recalculates what the deletions leave stale.

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
  when omitted; see *Which ranges* below) are kept in the `Report` when the `Calculation` asks for the `answer`: `Report::value_ranges` is
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
set it, with `now` and `seed` of `[global]` in `sheetrs.toml` as their defaults), and a step's own value wins
over the book's. The book's `now` and `seed` also reach the engine in `sheetrs_compiler::memory`
(`Options::with_book`), where the options of a call win over them.

**Which ranges.** `calculate` is the embedded mode of the engine (the workbook is its own template: inputs,
`ranges` or the `answer` need no compiled program). With no output declared (a settings sheet named `settings`
with `output` rows; the library takes no `settings_file`) the outputs are the formula cells, and with some
declared they are those; `ranges` is a selection among them when it lies inside, and a range **outside** them is
not refused as it is in a compiled program (`sheetc calculate` without `-o`, which lists the declared outputs in its
error): it is read, in memory and with no recalculation, from the workbook just written, with the machinery of
`query` (`sheetrs_answer` does it for every path that writes a workbook). A constant has the value it holds, the
`inputs` are in it, and a formula the outputs left out is calculated too (the workbook is the answer, so the build
declares only the inputs and the whole workbook comes out calculated, under the rules of the cache); the answer keeps
the order of the request. A text that is no range or name of the workbook is an
error (`Error::Calculation`). Details and the table of the cases: [`sheetrs-compiler`](../sheetrs-compiler/README.md#which-ranges-can-be-asked-for).

`sheetrs edit` has no `--calculate` flag: calculating is a step the planner inserts at the end of what leaves results stale and in a conversion (`BookEdit::calculate` is the explicit step in the library). To calculate a file with inputs and ranges, `sheetc calculate -i FILE -o OUT` (with `--request` to give them).

A pass reads and writes each worksheet as a stream, one
cell (XLSX) or row (ODS) at a time, and the plan is a handful of bands and names, so a plan with no recalculation
(a default conversion of an ODS, whose results are stored) keeps the memory flat; a plan that leaves formulas stale, or converts with `--calculate-all`, that only the engine can calculate ends with
the engine running, which holds the formulas of the workbook it needs (the times and the memory per scenario are in
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
| a deleted cell or a range with no cell left | `#REF!` (`[#REF!]` in an ODS), as the files of LibreOffice and Excel spell it; a removed sheet leaves the same `#REF!` (see the last items of this list) |
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
  `tests/assets/oracle/shapes_delete.json`, made by `scripts/make_shapes_oracle.py` in the file LibreOffice saves after the deletion and opens again: the reference of the project; three bands of an ODS give another height there than the memory of LibreOffice, a difference proposed in `docs/approved-differences.md`). A shape that
  lies wholly inside the deleted band is deleted with it (its outline counts: one that starts where the
  band starts stays). The cell it is anchored to keeps its index when it is deleted and moves up
  otherwise; a rectangle, picture or chart whose two ends cross has them put in order. Its size, when
  it resizes with its cells, comes from its end cell, or is kept (scaled by the size of its cell) when
  both ends are in one cell; a shape that does not resize keeps its size. All of it by the sizes of
  the rows and columns, which are in the file: the `ht` of each row and the `width` of each column of
  an XLSX (the width of a column is in characters, so a column that differs from the default is an
  approximation by the width of a digit of 7.4 pixels), the styles of the rows and columns of an ODS.
- **The reference is the file saved and opened again.** What LibreOffice shows in memory right after the
  same deletion is not what it shows when someone opens the file it saves (a reference that went is `#REF!`
  there and `COUNT` of it is `#REF!`; opened again it is 0), so the result of `sheetrs edit` is the one of the
  re-read, in both formats (`CONTRIBUTING.md`, *Parity and fidelity*; the oracle is
  `tests/assets/write_to/oracle.json`, `reread`).
- A deleted cell, range or sheet is written as the bare `#REF!` in an XLSX and `[#REF!]` in an ODS, the text the
  files of LibreOffice and Excel hold for both: LibreOffice reads the bare `#REF!` in an XLSX as an unknown name
  (`#NAME?`), also in the files it saves itself, and `[#REF!]` in an ODS as the error value `#REF!`. Written so, the file
  shows the same values as the file LibreOffice saves after the same edit (`COUNT` 0, `COUNTA` 1,
  `COUNT(ref;5)` 1, `IFERROR(ref;7)` 7, `SUM`, `MAX`... `#REF!` in an ODS and `#NAME?` in an XLSX), which
  `Sheet!#REF!`, the spelling of earlier versions for deleted cells, does not give (LibreOffice reads it as a reference with no
  cell: `COUNT` is `#REF!` and `IFERROR` does not catch it). The engine caches what LibreOffice shows when it opens
  the file saved after the same edit, for deleted cells and for a **deleted sheet** alike (`SUM`, `AVERAGE`, `MAX`,
  `COUNTIF`... are `#NAME?` in an XLSX there and here); a qualified `#REF!` of the input (`Data!#REF!`) is read as the plain error, and its text is kept in a formula the run does not touch.
- **The spelling of a reference that went is the one of the files, not the one LibreOffice shows in
  memory** (`=$Data.A#REF!*2` there; `[#REF!]*2` in the ODS it saves, `#REF!*2` in the XLSX it saves, and the same
  in the files of `sheetrs edit`): the memory of LibreOffice is not the reference (`CONTRIBUTING.md`, *Parity and
  fidelity*), so this is not a difference.
- **A `#REF!` in the text is read as the error it is.** The text does not tell a reference that a deletion
  removed from an error written in the formula, and LibreOffice does not either when it opens the file, so the
  recalculation reads both the same way and with the semantics of the format (see above).

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

- By default the values are the results the file caches for the formulas the plan does not reach, as
  `LibreOffice` converts them, and what the engine calculates for the formulas it reaches (the cells that a
  deleted row, a replacement or a hidden row changes, what has no stored value, a volatile formula, and in an ODS
  what `LibreOffice` calculates when it opens the file): the engine calculates them in the pass that writes the
  conversion, and `--replace A --with B --formula-to-value S` in one invocation converts what
  the formulas give after the replacement. With `--calculate-all` the values are what the formulas **calculate**: the
  engine calculates the cells to convert (and what they read) in the pass that writes them (every formula in an ODS). A cell
  that depends on what the engine cannot calculate (below) keeps the result the file caches; one the run calculates and that only `LibreOffice` can give a result for is
  an error that names the cause; nothing is written then. A formula without a stored value that
  the engine does not calculate (it cannot) keeps its formula.
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
- `tests/calculation_mode.rs` runs the same edits by default (the cache) and with `calculate_all` on `tests/assets/cached/`
  (`make_fixtures.py`: formulas that cache a result the engine would not give, so that a formula that kept its
  cache is told from one calculated again): the cells reached are those of a full calculation, the others are the
  file's (the XML of the cell, byte for byte, in an XLSX), and the cases that are calculated whole are equal to a run
  without the option; in each format.


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
