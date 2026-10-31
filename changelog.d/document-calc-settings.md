### Added

- `Workbook::settings` (`DocumentSettings`): the calculation settings a document stores, as stored and with no defaults applied (`None` = not stated): iterative calculation (XLSX `calcPr`, ODS `table:iteration`), case sensitivity, whole-cell matching, regular expressions and wildcards in search criteria, the two-digit-year window (`table:null-year`), precision as shown, and the document language (ODS default cell style). `WorkbookReader` gains `read_settings`.
