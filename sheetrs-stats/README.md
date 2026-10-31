# sheetstats

**sheetstats** is a CLI tool for analyzing the complexity and structure of spreadsheets. It provides metrics that help identify "heavy" workbooks, complex dependencies, and potential optimization targets.

## Installation

```bash
cargo install --path .
```

## Usage

```bash
sheetstats <FILE> [OPTIONS]
```

### Options

- `--format <FORMAT>`: Output format (human/json).

## Metrics Reported

Workbook totals:

- **Sheets** and **named ranges**.
- **File size** on disk.
- Sum of every per-sheet count below (used cells, formulas, errors, empty cells, comments, charts,
  tables).

Per sheet, in workbook order (human output also shows each sheet's share of
the workbook total):

| Metric | Definition |
| ------ | ---------- |
| Used cells | Cells holding a value or a formula. Cells kept only for their style are not counted. |
| Formulas | Used cells with a formula. |
| Errors | Used cells whose value, or cached formula result, is an error (`#DIV/0!`, `#REF!`, ...). |
| End of sheet | Bottom right corner of the used area (`GV5439`): the last row and the last column that hold a used cell. Absent in a sheet without used cells. |
| Comments | Notes of the cells: the comments of an XLSX (the threaded ones when the sheet has no others) and the annotations of an ODS. |
| Charts | Charts the sheet shows: the chart parts its drawing relates to in an XLSX, the embedded objects of an ODS. |
| Tables | Table objects of an XLSX; named database ranges of an ODS (the one of an autofilter is not a table). |
| Empty cells | Cells without value or formula inside the used area, which spans from A1 to the last row and column holding a used cell. |
| XML size | Uncompressed size of the XML describing the sheet: the worksheet part in XLSX, the `<table:table>` element of `content.xml` in ODS. |

XML size depends on each format's markup, so it compares sheets within a
workbook, not across formats.

### JSON output

```json
{
  "total_sheets": 1,
  "total_named_ranges": 0,
  "total_file_size": 8812,
  "total_used_cells": 12,
  "total_formulas": 3,
  "total_errors": 1,
  "total_empty_cells": 4,
  "total_comments": 1,
  "total_charts": 0,
  "total_tables": 0,
  "total_xml_size": 5120,
  "sheets": [
    {
      "name": "Data",
      "used_cells": 12,
      "formulas": 3,
      "errors": 1,
      "empty_cells": 4,
      "end_of_sheet": "D4",
      "comments": 1,
      "charts": 0,
      "tables": 0,
      "xml_size": 3584
    }
  ]
}
```

The cells are counted as the reader reads them and are not kept, so the memory of a run does not
grow with the size of the workbook.
