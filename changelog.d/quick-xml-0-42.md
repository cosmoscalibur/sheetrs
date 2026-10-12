### Added

- Add `parser_utils::attr_value`, which returns an XML attribute value with entities resolved and XML 1.0 whitespace normalization applied.

### Changed

- Update `quick-xml` to 0.42 and `regex` to 1.13, and raise the minimum versions of `serde` (1.0.229), `toml` (1.1.8), `anyhow` (1.0.104), `clap` (4.6.7), `reqwest` (0.13.5), `chrono` (0.4.45) and `serde_json` (1.0.151). XML names, attributes and text are now compared and read as `&str`. `regex` 1.13.1 fixes incorrect match offsets in some patterns. quick-xml 0.42 validates UTF-8 on every event, the main reason reading the ODS production workbook takes 7.8% more user-space instructions.
- Speed up reading: shared XLSX formulas read their reference parts by group index, and ODS formula normalization copies the formula fewer times. A full `sheetlint` run on the production workbooks takes 2.0% fewer user-space instructions on XLSX and 1.0% more on ODS than before this update.
- Attribute values follow XML 1.0 attribute-value normalization: literal tabs and line breaks become spaces, while character references such as `&#10;` keep the line break. The test workbooks parse identically.
- Invalid UTF-8 in CDATA sections is reported as an error instead of being replaced, as text and attributes already were.

### Fixed

- Fix `sheetcli --remove-sheets` and `--remove-ranges` ignoring sheets and named ranges whose names contain XML entities (e.g. `R&D`).
