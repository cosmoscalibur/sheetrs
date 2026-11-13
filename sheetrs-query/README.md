# sheetrs-query

Reads the cells of an XLSX or ODS workbook **by range, sheet or workbook**, with the properties
asked for, and finds the cells that match a search. Library (`sheetrs_query::BookQuery::query`) and the
`sheetrs query` subcommand (feature `cli`; the `sheetrs` executable is the crate `sheetrs-cli`). The request and the response are the typed JSON of
[`sheetrs-protocol`](../sheetrs-protocol/README.md).

Only what is asked is read: the readers skip the sheets that are not in scope, stop after the last
row a window reaches and do not build the properties that are not asked for (the cost of ten rows of a sheet
of 4 million cells is in [docs/performance.md](../docs/performance.md#sheetrs-query)). Empty cells are not returned.

## Command line

```bash
sheetrs query -i book.xlsx --range 'Data!A1:C10' -p value,formula          # a range
sheetrs query -i book.ods --sheet Data -p display                          # a sheet, as it is seen
sheetrs query -i book.xlsx --workbook -p value,number_format --jsonl       # the whole workbook, one cell per line
sheetrs query -i book.xlsx --find 'total' --entire-cells                   # the cells whose value is "total"
sheetrs query -i book.xlsx --find '=SUM(' --in formula                     # the cells whose formula holds it
sheetrs query -i book.xlsx --query '{"scope": {"range": "Data!A1:B2"}, "properties": ["display"]}'
sheetrs query -i book.xlsx --list-sheets                                   # the names of the sheets, in order
sheetrs query -i book.xlsx --list-names                                    # the defined names and what they refer to
```

| Option | |
| ------ | - |
| `--range R`, `--sheet S`, `--workbook` | Where to read (repeatable); the whole workbook when none is given. A range without a sheet is on the first one |
| `-p, --property` | `value` (default), `formula`, `number_format`, `display`, `comment` |
| `--find TEXT` | Keep the cells where the text is found, as LibreOffice's Find & Replace finds it |
| `--regex`, `--wildcards` | The text is a regular expression (the subset of ICU that is reproduced) or has wildcards (`*`, `?`, `\` to escape) |
| `--match-case`, `--entire-cells` | As in Find & Replace |
| `--in value\|formula\|comment` | What the search looks at (default `value`); `comment` is the text of the notes |
| `--include-formulas` | A search by value also looks at the results of formulas (off by default: it reaches what was typed, not what was computed) |
| `--formatted` | Search the text each cell **shows** (`1,234.50`, `15 Mar 2023`, `$10.00`) instead of its input line (`SearchFormatted` of LibreOffice) |
| `--filtered` | Also search the rows an autofilter has hidden (`SearchFiltered` of LibreOffice); off by default, as there |
| `--locale TAG` | Locale of `display` and of the dates a search sees (default `en-US`) |
| `--dates` | Give the numbers that are dates or times as dates (see below); without it, the serial the file stores |
| `--query JSON\|@file` | The whole query as JSON; replaces the options above |
| `--list-sheets`, `--list-names` | The names of the sheets (a JSON list, in order) or the defined names with what they refer to (a JSON object); nothing else is read |
| `--jsonl`, `--pretty` | One JSON object per cell per line; indented JSON |

Each cell is `{"sheet", "cell", "row", "col"}` plus the properties asked for: `value` (number,
text, boolean, or an error, which is the object `{"error": "#DIV/0!"}` so that it is not taken for a text that starts with `#`; `Err:522` for an ODS error Excel does not have),
`formula` (without the `=`, in the syntax of the file: Excel's in an XLSX, `OpenFormula` in an ODS) (a search in formulas sees a reference to another workbook as LibreOffice shows it, `'file:///folder/test.xlsx'#$Sheet1.B3`, the file resolved against the folder of the document), `number_format` (the format code,
`0.00%`, `yyyy-mm-dd`), `display` and `comment` (the text of the note of the cell: XLSX legacy comments, with
rich runs joined, and ODS `office:annotation`; the paragraphs are joined with `\n`; a note on an empty cell is a cell of the
answer too; an XLSX threaded comment is the text of the legacy note that Excel writes for it, which holds the
first message after `Comment:`; LibreOffice ignores the threaded part, and so does this reader, so a reply that only
that part holds is not found: `tests/libreoffice_find.rs`, fixture `threaded.xlsx`). Every cell of an array formula
reports its formula as `{=…}` (the interior cells too, as LibreOffice shows them), so a search in formulas finds all of them.
Whether a formula is an array formula is not a property of its own.

## Dates as dates

By default every number is the serial the file stores (`45366.4375`), in the date system of the file. With
`--dates` (`"dates": true` in a query, `dates=True` in Python, `dates: true` in JavaScript, `dates: true` in `RangeOptions`/`FindOptions`) a number whose
**format** is a date, a time or both (the recognition `display` uses) is given as a typed value, ISO 8601
text with its type, worked out with the date system of the file (one of the three of LibreOffice Calc: 1899-12-30, 1900-01-01
or 1904-01-01; an ODS with another null date is an error that names it):

```bash
sheetrs query -i book.xlsx --range 'Data!A1:C1' --dates --jsonl
# {"sheet":"Data","cell":"A1","row":1,"col":1,"value":{"type":"date","value":"2024-03-15"}}
# {"sheet":"Data","cell":"B1","row":1,"col":2,"value":{"type":"datetime","value":"2024-03-15T10:30:00"}}
# {"sheet":"Data","cell":"C1","row":1,"col":3,"value":{"type":"time","value":"10:30:00"}}
```

| Format of the cell | `value` |
| ------------------ | ------- |
| a date (`yyyy-mm-dd`, `dd/mm/yyyy`, `d-mmm-yy`…) | `{"type": "date", "value": "2024-03-15"}` |
| a date and a time | `{"type": "datetime", "value": "2024-03-15T10:30:00"}` |
| a time of day (`hh:mm:ss`, `h:mm AM/PM`, `mm:ss`) | `{"type": "time", "value": "10:30:00"}`: the day is dropped, as the format drops it |
| an elapsed time (`[h]:mm:ss`), a percentage, `General`, text… | unchanged |

The rules are the ones LibreOffice shows (the tests compare them with it, in the three date systems):

- the serial 0 of the usual system is 1899-12-30, also for an XLSX (it has no 1900 leap-year day, as
  LibreOffice reads it); a date before 1582-10-15 is a Julian-calendar date, as LibreOffice shows it;
- whole seconds are cut, not rounded (`0.999999` of a day is `23:59:59`, never `24:00:00`); when the format
  writes fractions of a second (`ss.000`) the text has them (up to three digits), rounded without reaching
  the next second;
- a serial whose year is outside 1 to 9999 stays a number (ISO 8601 writes four digits and a Python
  `datetime.date` stops there; LibreOffice goes on to the year 32767);
- a date format shows the day of the whole part of the serial: `45366.999999` is `2024-03-15`;
- a formula takes the format LibreOffice gives its result (see below), so `=A1+1` over a date is a date.

## What a person sees and what a search sees

`display` is the value written with its number format in a locale. A **search** looks at another
text, the one LibreOffice puts in the input line of the cell (observed in 26.8, `en-US`):

- a number as it was typed: `1234.5` for a cell that shows `1,234.50`, `25.6%` for a percentage;
- a date as the input line writes it: `03/15/2023`, or the ISO date when the format of the cell is
  an ISO one (`2023-03-15`);
- a formula in the notation of LibreOffice, not Excel's: `=SUM(A1,B1)`, `=$'My Sheet'.A1`;
  - the name of a function LibreOffice does not have is written in lowercase from an XLSX
    (`=casa()`, `=_xlfn.casa()`);
  - a bare `#REF!` (the error written where an operand goes, as Excel does when the cell a formula
    read is deleted) is the name `#ref!` that LibreOffice does not know when it comes from an XLSX, and
    the cell shows `#NAME?` whatever error the file caches; an ODS holds it as `#ref!`, which
    LibreOffice reads as the error itself (`=#REF!+5`, `#REF!`). A `#REF!` that is part of a reference
    (`Sheet!#REF!`) is the error the file caches. Tested against `errors.xlsx` and `errors.ods`;
- an error is never found, and the result of a formula only with `--include-formulas`.

With `formatted` (`--formatted`) the text searched is the `display` of the cell, as LibreOffice does
with `SearchFormatted`: `1,234.50` and `15 Mar 2023` are found, `1234.5` and `03/15/2023` are not.
It applies to the values (and to the results of formulas with `include_formulas`) and not to a search
in formulas or notes. The searches of both options are compared with `findAll` of LibreOffice in
`tests/libreoffice_find.rs`.

**Filtered rows (`SearchFiltered`).** LibreOffice skips by default the rows that an autofilter hides
and searches them when `SearchFiltered` is true (`filtered`, `--filtered`); rows hidden by hand are
always searched. The reader reports the filtered rows of the sheets searched
(`CellSink::filtered_rows`) and the search leaves them out, in a range scope too:

- an **ODS** marks them (`table:visibility="filter"`), as LibreOffice writes them;
- an **XLSX** has no mark of its own: a row counts as filtered when it is hidden **inside the range
  of the autofilter of the sheet** (`<autoFilter ref>` of the worksheet), the rule LibreOffice
  applies when it opens the file. A row hidden by hand outside that range is searched, and so is
  every row of a filter kept in a table (`xl/tables`), which LibreOffice does not read back as a
  filter either (its ODS does keep it). The header row of a filter is filtered too when it does not
  satisfy the criterion, as LibreOffice shows it.

The searches, in both formats and with and without the option, are compared with `findAll` of
LibreOffice in `tests/libreoffice_find.rs` (fixtures `tests/assets/filtered*.{xlsx,ods}`, made by LibreOffice with
`tests/assets/make_filtered_fixtures.py`; the oracle by `make_find_oracle.py`). A search that does not look for filtered rows
reads an XLSX sheet to its end, because the range of the autofilter comes after the rows.

A regular expression outside the subset of ICU that was compared with LibreOffice
(`sheetrs_engine::regex_subset`) is refused instead of giving another result.

The anchors `^` and `$` are in that subset for a search, as LibreOffice reads them in a cell: `^` is
the start of the text and `$` its end (or the place before the line terminator that ends it), never
the start or the end of a line, so `^(Code|Resumen Contadia)$` finds the cells with exactly one of
the two texts, in the middle of a pattern (`a^b`, `a$b`) they match nothing, and they work inside
groups and alternations (`(^a|b$)`). `.` still leaves out the line breaks. A pattern that can match
the empty text (`^`, `$`, `^$`, `a*$`) is refused. `sheetrs edit --replace` refuses the anchors:
what LibreOffice does with them when it replaces (the empty match, the match at the end of the text)
is not reproduced. The searches with anchors are compared with `findAll` of LibreOffice in
`tests/libreoffice_find.rs` (workbooks `tests/assets/anchors.*`, made by
`tests/assets/make_anchors_fixtures.py`).

## The format of a formula cell

LibreOffice gives a formula cell whose style is `General` the format its result has: `=A1` over a
date shows a date, `=A1+1` the default date format (`MM/DD/YY`), `=A4*2` over a currency a
currency. `number_format` and `display` follow it (`src/inherit.rs`; the rules are in its
documentation and the probes are the fixture `sheetrs-io/tests/assets/inherit.{xlsx,ods}`, made by
`sheetrs-io/tests/assets/oracle/make_inherit.py`, checked against LibreOffice cell by cell; text,
`@` and several arguments have their own, `inherit_text.*`, made by `make_inherit_text.py`).

- **Which cells.** Only those whose own format is `General`, and only when LibreOffice interprets
  them as it loads the file: in an **ODS**, every formula cell without a style; in an **XLSX**, those
  that call a volatile function (`RAND`, `RANDBETWEEN`, `RANDARRAY`, `NOW`, `TODAY`, `INDIRECT`,
  `OFFSET`, `CELL`) or depend on one. The other cells of an XLSX keep what the file says.
- **Rules.** A cell, `INDIRECT`, `OFFSET`, `INDEX`, `CHOOSE`, `IF` and `IFERROR` give the format of the
  cell they point at; `+` and `-` combine the classes of the operands (a date plus a number is a
  date, a time plus a time a duration, a currency wins); `*` and `/` keep a currency, otherwise
  `General`; `x%` is `0.00%`; an aggregate takes the format of the first of its arguments that holds a
  number (the cells without one, text and empty, do not count; a number written in the formula gives
  `General`), `SUM` that of the last cell of that argument, the others that of the first; the text
  format (`@`) is never taken from a cell, the result is `General`; comparisons, `AND`, `OR`, `NOT`, `IS…` give `BOOLEAN`; a
  scientific number becomes `0.00E+00` in an aggregate. In an ODS a time is shown
  `HH:MM:SS AM/PM` and a percentage `0.00%`, whichever format they come from.
- **Choosing a cell.** `IF`, `CHOOSE` and `INDIRECT` take the format of the branch or the cell their
  condition, index or text names, worked out from constants, cells (their cached value), defined names
  (a cell, a range or a constant) and a list of functions (`ISNUMBER`, `ISTEXT`, `ISBLANK`, `LEN`,
  `ABS`, `INT`, `ROUND`, `MOD`, `ROW`, `COLUMN`, `CONCATENATE`, `UPPER`, `LOWER`, `LEFT`, `RIGHT`,
  `ADDRESS`, `SUM`, `COUNT`, `MAX`, `MIN`, `AND`, `OR`, `NOT`, `IF`); the probes are
  `sheetrs-io/tests/assets/inherit_more.*`.
- **Limits.** A condition, index or text that depends on anything else (another function, a volatile
  one, a name that is a formula) gives `General`; a name defined in more than one scope is not
  resolved; `SUM` over a whole column or row too; only the default formats of `en-US` are written.

## Library

```rust
use sheetrs_io::Book;
use sheetrs_query::{BookQuery, Query};

let book = Book::from_file("book.xlsx")?;
let query: Query = serde_json::from_str(r#"{"scope": {"sheet": "Data"}, "properties": ["value", "display"]}"#)?;
for cell in book.query(&query)?.cells {
    println!("{}!{} = {:?}", cell.sheet, cell.cell, cell.display);
}
```

`sheetrs_query::{Matcher, text::TextContext}` are shared with
[`sheetrs-edit`](../sheetrs-edit/README.md), which replaces what this crate finds.
