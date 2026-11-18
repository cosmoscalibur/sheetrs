# Specifications

Behaviour of `LibreOffice` 26.8 that sheetrs reproduces and that no public specification
states, grouped as the crates are. Each page lists its observations with the test that
shows them, and its open defects: differences between sheetrs and the oracle that are known
and not fixed. Every such difference is a defect.

| Area | Crate | Covers |
| --- | --- | --- |
| [`model/`](model/) | `sheetrs-core` | What a cell holds, dates, storage limits: whatever does not depend on a file format |
| [`io/`](io/) | readers and writers | Behaviour common to formats goes in `io/`, behaviour of one format in `io/<format>/` |
| [`io/xlsx/`](io/xlsx/) | `sheetrs-xlsx` | Values, workbook, cell placement, formulas and writing of XLSX |
| [`ops/`](ops/) | `sheetrs-testkit` and the planner | Operations applied to a workbook |

A page belongs to the lowest area whose crate owns the behaviour: a limit of the store is a
`model/` page even when a format triggers it. A behaviour that holds for several formats moves
up to the area they share when the second format shows it.
