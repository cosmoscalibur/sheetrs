# sheetrs-query

Reads the cells of an XLSX or ODS workbook **by range, sheet or workbook**, with the properties
asked for, and finds the cells that match a search. Library (`sheetrs_query::run`) and executable
(`sheetquery`). The request and the response are the typed JSON of
[`sheetrs-protocol`](../sheetrs-protocol/README.md).

Only what is asked is read: the readers skip the sheets that are not in scope, stop after the last
row a window reaches and do not build the properties that are not asked for (ten rows of a sheet of
4 million cells take under 10 ms and 13 MB, in either format). Empty cells are not returned. The
answer is the same for an XLSX and an ODS wherever the two files say the same thing.

## Command line

```bash
sheetquery book.xlsx --range 'Data!A1:C10' -p value,formula          # a range
sheetquery book.ods --sheet Data -p display                          # a sheet, as it is seen
sheetquery book.xlsx --workbook -p value,number_format --jsonl       # the whole workbook, one cell per line
sheetquery book.xlsx --find 'total' --entire-cells                   # the cells whose value is "total"
sheetquery book.xlsx --find '=SUM(' --in formula                     # the cells whose formula holds it
sheetquery book.xlsx --query '{"scope": {"range": "Data!A1:B2"}, "properties": ["display"]}'
```

| Option | |
| ------ | - |
| `--range R`, `--sheet S`, `--workbook` | Where to read (repeatable); the whole workbook when none is given. A range without a sheet is on the first one |
| `-p, --property` | `value` (default), `formula`, `number_format`, `display` |
| `--find TEXT` | Keep the cells where the text is found, as LibreOffice's Find & Replace finds it |
| `--regex`, `--wildcards` | The text is a regular expression (the subset of ICU that is reproduced) or has wildcards (`*`, `?`, `\` to escape) |
| `--match-case`, `--entire-cells` | As in Find & Replace |
| `--in value\|formula` | What the search looks at (default `value`) |
| `--include-formulas` | A search by value also looks at the results of formulas (off by default: it reaches what was typed, not what was computed) |
| `--locale TAG` | Locale of `display` and of the dates a search sees (default `en-US`) |
| `--query JSON\|@file` | The whole query as JSON; replaces the options above |
| `--jsonl`, `--pretty` | One JSON object per cell per line; indented JSON |

Each cell is `{"sheet", "cell", "row", "col"}` plus the properties asked for: `value` (number,
text, boolean, or an error written `#DIV/0!`; `Err:522` for an ODS error Excel does not have),
`formula` (without the `=`, Excel's syntax in both formats), `number_format` (the format code,
`0.00%`, `yyyy-mm-dd`) and `display`. Whether a formula is an array formula is not a property.

## What a person sees and what a search sees

`display` is the value written with its number format in a locale. A **search** looks at another
text, the one LibreOffice puts in the input line of the cell (observed in 24.2.7, `en-US`):

- a number as it was typed: `1234.5` for a cell that shows `1,234.50`, `25.6%` for a percentage;
- a date as the input line writes it: `03/15/2023`, or the ISO date when the format of the cell is
  an ISO one (`2023-03-15`);
- a formula in the notation of LibreOffice, not Excel's: `=SUM(A1,B1)`, `=$'My Sheet'.A1`;
- an error is never found, and the result of a formula only with `--include-formulas`.

A regular expression outside the subset of ICU that was compared with LibreOffice
(`sheetrs_engine::regex_subset`) is refused instead of giving another result.

## Library

```rust
use sheetrs_io::reader::Source;
use sheetrs_query::{run, Query};

let source = Source::from_file("book.xlsx")?;
let query: Query = serde_json::from_str(r#"{"scope": {"sheet": "Data"}, "properties": ["value", "display"]}"#)?;
for cell in run(&source, &query)?.cells {
    println!("{}!{} = {:?}", cell.sheet, cell.cell, cell.display);
}
```

`sheetrs_query::{Matcher, text::TextContext}` are shared with
[`sheetrs-edit`](../sheetrs-edit/README.md), which replaces what this crate finds.
