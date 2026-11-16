# sheetrs-io

The input and output of the SheetRS suite. `Book` is the one object: a workbook as a file or as bytes,
whose format (XLSX, XLSM or ODS) is detected from the content. It reads the workbook model, streams its
cells, and, with the feature `write`, keeps the operations pending on it and writes the result; the
other crates add their verbs to it (`BookQuery`, `BookEdit`, `BookCompile`, `BookLint`, `BookStats`).

```rust,no_run
use sheetrs_io::Book;

let book = Book::from_file("book.xlsx")?;
println!("{} {:?}", book.format().extension(), book.date_system()?);
# Ok::<(), sheetrs_io::Error>(())
```

The locale is a property of the `Book`: `Book::with_locale` gives it (a `locale::LocaleTag`, which parses
from `"es-CO"`), and every tool that takes the book follows it (`TEXT` and the conversion of a number to
text when a workbook is calculated, the `display` of a query). Without it the book has the locale its
system gives (`Book::default_locale`, `locale::system`: `LC_ALL`, `LC_CTYPE` or `LANG`, as `LibreOffice`
reads them on Linux), for an XLSX and an ODS alike; `en-US` is what there is when the system has none and in
WebAssembly. With the feature `config`, `locale::from_config`
reads the `locale` of `[global]` in a `sheetrs.toml`.

The API is experimental: names and shapes may change.

## What the writers change of a result when they write it

The writers (`sheetrs_io::write`, the feature `write`: `modify_workbook`, `CellWrites`) rewrite the XLSX or the ODS as a stream. They write a
result in three ways: **a conversion of formulas to values** (`ValueScope`, `sheetrs edit --formula-to-value`: the cell loses its formula and holds
the result), **the result next to its formula** (`CellWrites::result`: what a calculation gives, in a cell that keeps its formula) and **a constant**
(`CellWrites::constant`: a number, a text, a boolean, an empty cell). This section lists what each kind of result becomes in the file, against what
`LibreOffice` 26.8 writes for the same thing. The reference is the file `LibreOffice` saves and opens again (`CONTRIBUTING.md`, *Parity and
fidelity*): the **status** is what `LibreOffice` shows when it opens the file of the writers, against what it shows of its own file; the columns
*here* and *LibreOffice* are the text of the cell in the file.

How it was probed (not assumed): `sheetrs-edit/tests/assets/writer_results/make_oracle.py` makes, with `LibreOffice`, `kinds.{xlsx,ods}` (and
`kinds_1904.*`, in the 1904 date system) with one formula for each case of the *Case* column in the sheet `K` and one constant in the sheet `C`;
`LibreOffice` converts the formulas (*Data > Calculate > Formula to Value*), saves and opens the file again; `sheetrs edit` and
`BookEdit::write_cells` do the same on the same file and `LibreOffice` opens theirs. `oracle.json` holds both views and the XML of every cell
(`raw_*`), and `sheetrs-edit/tests/writer_results.rs` checks them (`differences_*` is the list of the cells where the two views differ; the
test pins it). To run the probe again: `scripts/libreoffice-server.sh`, `cargo build -p sheetrs-cli`, then
`/opt/libreoffice26.8/program/python sheetrs-edit/tests/assets/writer_results/make_oracle.py`.

Every row that **differs** is in `docs/approved-differences.md` as proposed (the maintainer decides). A row marked *same* is the same for
`LibreOffice` when it opens the file; where only the text of the cell differs, the row says so.

| Kind of result (case of the probe) | Format | Here | LibreOffice | Status |
| ---------------------------------- | ------ | ---- | ----------- | ------ |
| **Error**, every kind: `#DIV/0!`, `#N/A`, `#NAME?`, `#REF!`, `#VALUE!`, `#NUM!`, `#NULL!`, `Err:502`, `Err:522`, an error through a reference, an error in a cell with a format; converted to a value (default and `calculate_all`) | XLSX | an empty inline string (`t="inlineStr"`, `<is><t xml:space="preserve"/></is>`); a blank cell for `LibreOffice` | an empty shared string (`t="s"`); a blank cell | same: the error is not kept, `LibreOffice` leaves an empty text too (`.uno:ConvertFormulaToValue`, re-read); test `a_conversion_leaves_an_empty_text_where_the_formula_gave_an_error` |
| the same | ODS | a string cell with an empty paragraph (`office:value-type="string"`, `<text:p/>`) | the same | same, same test |
| Error **next to its formula** (a calculation, the formula stays) | XLSX | `t="e"` and the error (`#VALUE!` for the `Err:5xx` an XLSX cannot hold) | the same | same (`differences_recalc` holds no error case) |
| the same | ODS | `calcext:value-type="error"` and the code of `LibreOffice` (`Err:502`, `Err:522`) | the same | same |
| Integer, decimal of up to 15 significant digits, negative, zero (`integer`, `negative`, `zero`, `IFERROR gives a number`, `reference to an empty cell`) | XLSX | `<v>5</v>`, no `t` (a number) | `t="n"` | same |
| the same | ODS | `office:value-type="float"`, `office:value="5"` | the same | same |
| Number of **more than 15 significant digits** (`fraction` 1/3, `float sum 0.1+0.2`, `2^53+1`, `date and time` 45366.56302083333, and the constants `123456789012345678` and `2^53+1 as a double`) | XLSX | all the digits that read back to the same double: `0.3333333333333333`, `0.30000000000000004`, `9007199254740992` | 15 significant digits: `0.333333333333333`, `0.3`, `9007199254740990`, `1.23456789012346E+017` | **differs**: the value differs after the 15th digit and `LibreOffice` shows the same text; here the file holds the value the engine calculated (or the constant given), `LibreOffice` rounds when it saves. In the default mode an XLSX conversion converts what the file caches (15 digits already), so only a run that calculates (`calculate_all`, a recalculation, a constant) shows it; an ODS is always calculated. Test `numbers_keep_all_their_digits_where_libreoffice_keeps_fifteen` |
| the same | ODS | the same, in plain decimal (`9007199254740992`) | the same 15 digits, in scientific form from 1E+15 (`9.00719925474099E+015`) | **differs**, same row |
| **Very large and very small** numbers (`1E+300`, `1E-300`, `1E-15`, `5E-324`, `1.7976931348623157E+308`) | both | the plain decimal expansion of the double (Rust's `Display` has no exponent: up to 330 characters, `0.000000000000001` for `1E-15`) | `1E+300`, `1E-015` | same value (re-read is the same number); the text of the cell is longer. Not a change of the result |
| **Negative zero** (`negative zero`) | XLSX, ODS | a result is `0` (the engine gives none), the constant `-0.0` is `-0` | `0` in an XLSX, `-0` in an ODS | same (shown as `0`) |
| **NaN** (constant) | XLSX | `NaN` | `NaN` | same |
| the same | ODS | `NaN` | `0` | same re-read (`LibreOffice` reads `NaN` as 0), the file differs |
| **+inf, -inf** (constants) | XLSX | `inf`, `-inf` | `INF`, `-INF` | **differs**: `LibreOffice` reads `inf` as 0 and `INF` as infinity |
| the same | ODS | `inf`, `-inf` | `0` | same re-read (0), the file differs |
| **Text**: `abc`, a leading or trailing space, only spaces, two inner spaces, a newline, a tab, unicode, a leading apostrophe, `#DIV/0!` as text, `<`, `&` and `'`, 300 characters | XLSX | an inline string (`t="inlineStr"`, `xml:space="preserve"`, the characters as they are: a line break or a tab is the character itself) | a shared string (`t="s"`) with the same characters (a line break is `&#10;`, a tab `&#9;`) | same (the text is the same; inline and shared strings are two forms of one cell) |
| the same | ODS | `office:value-type="string"` with the text in `<text:p>` (`<text:s/>` for spaces, `<text:tab/>`, a paragraph for each line) and the attribute `office:string-value` | the same, without `office:string-value` | same |
| **Text a cell would read as something else** (constants: a leading `=`, `1`, `TRUE`) | XLSX, ODS | a text cell, the number format of the cell stays (General) | a text cell whose format is `@` | **differs** in the format of the cell (the text is the same; the cell is still text for `LibreOffice` after the file is opened, and the format `@` is not set). Converted formulas keep the format they had, and the same |
| **Empty text**, converted (`=""`), and an empty text next to a formula | XLSX | an empty inline string; a blank cell for `LibreOffice` | an empty shared string; a blank cell | same |
| the same | ODS | a string cell with an empty paragraph | the same | same |
| **Empty text as a constant** (`Constant::Text("")`) | XLSX | an empty inline string; a blank cell for `LibreOffice` | no cell | same |
| the same | ODS | a string cell with an empty paragraph (`Text("")`) | an empty cell | **differs**: `LibreOffice` shows a text cell here and an empty one for its own file |
| **Boolean** (`TRUE()`, `FALSE()`, `1=1`; constants) | XLSX | `t="b"` and `1` or `0` | `t="b"` (a result); a constant boolean of its API is a number with the format `TRUE`/`FALSE` and the formula `TRUE()` | same (shows `TRUE`) |
| the same | ODS | `office:value-type="boolean"`, `office:boolean-value` | the same | same |
| **Date, time, date and time, a date before the epoch, `9999-12-31`, a duration over 24 hours**, in the usual and the 1904 date systems | XLSX | the serial number, the style of the cell (and so its format) stays | the same | same (the date system of the workbook is not touched) |
| the same | ODS | `office:value-type="float"` with the serial, and the style of the cell: `LibreOffice` shows a date, a time, a duration by the format | typed values: `date` (`office:date-value`), `time` (`office:time-value`), the same style | same for `LibreOffice` (it reads the format of the style); the type of the cell differs, so a program that reads `office:value-type` sees a number |
| **Percentage, currency, scientific, fraction, thousands separator** | XLSX | the number, the style stays | the same | same |
| the same | ODS | `float` and the style (the percentage is `0.25` and shows `25.00%`; a currency is `1234.5`) | `percentage`, `currency` with `office:currency` | same for `LibreOffice` (the format of the style decides what it shows); the type of the cell differs |
| **Array formula** (`{1;2}`) | XLSX, ODS | the first cell is the value `1` (a scope that covers the whole array converts it) | the same | same |
| Any of the above in a workbook with the **1904 date system** | XLSX, ODS | as above | as above | same: `kinds_1904.*` give the same rows |

What was **not** probed (nothing here says it is the same): rich text in a cell, hyperlinks and notes of the cells that are converted, a data table, a
shared formula group converted in part (`writer_basic.rs` and `ods_writer.rs` have tests of their own, without `LibreOffice`), dates typed as
`office:date-value` in the input of a conversion with `calculate_all`, and locales other than `en-US`.
