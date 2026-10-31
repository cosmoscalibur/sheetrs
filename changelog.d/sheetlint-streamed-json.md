### Changed

- `sheetlint --format json` writes the document violation by violation instead of building a `serde_json::Value` tree of all of them and then its text (about 4 KB per violation). The output is byte-identical (a test compares it with the tree's document; the three benchmark workbooks match by MD5). Delgado workbook: 38.5 s and 3.4 GB peak → 14.5 s and 527 MB; the 1.2 M-formula workbook: 33 s and 1.8 GB → 23 s and 0.9 GB.
