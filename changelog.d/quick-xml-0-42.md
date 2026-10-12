### Added

- Add `parser_utils::attr_value`, which returns an XML attribute value with entities resolved and XML 1.0 whitespace normalization applied.

### Changed

- Update `quick-xml` to 0.42 and `regex` to 1.13, and raise the minimum versions of `serde` (1.0.229), `toml` (1.1.8), `anyhow` (1.0.104), `clap` (4.6.7), `reqwest` (0.13.5), `chrono` (0.4.45) and `serde_json` (1.0.151). XML names, attributes and text are now compared and read as `&str`. `regex` 1.13.1 fixes incorrect match offsets in some patterns.
- Attribute values follow XML 1.0 attribute-value normalization: literal tabs and line breaks become spaces, while character references such as `&#10;` keep the line break. The test workbooks parse identically.
- Invalid UTF-8 in CDATA sections is reported as an error instead of being replaced, as text and attributes already were.

### Fixed

- Fix `sheetcli --remove-sheets` and `--remove-ranges` ignoring sheets and named ranges whose names contain XML entities (e.g. `R&D`).
