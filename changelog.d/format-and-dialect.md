### Changed

- There is no dialect: the engine behaves as LibreOffice 26.8 for an XLSX and an ODS alike. The `dialect` key of `sheetc.toml`, `--dialect`, Excel's 1900 leap-year calendar and every Excel-semantics branch of the engine are gone; the Excel answers where LibreOffice differs are documented in `FIDELITY.md`, not computed. What varies per file is the document's settings and the namespace of the functions (`core`, `microsoft`, `openoffice`, `libreoffice`, `legacy`, registered as `ns::NAME`).
- The date system is the document's null date, any date: an ODS's `table:null-date` is read in full (it used to be only 1899-12-30 or 1904) and numbers its date cells and every date function; `--date-system` accepts `1899`, `1900`, `1904` or `YYYY-MM-DD`. Before the null date, serials are negative in every system.

### Added

- `sheetc` reads `SyntaxStringRef` from an ODS `settings.xml` (`DocumentSettings::string_ref_syntax` in `sheetrs`): `INDIRECT` and `ADDRESS` use Calc notation (`Sheet.A1`) in a document created in LibreOffice and Excel notation (`Sheet!A1`) in an XLSX or an XLSX saved as ODS.
- A bare `CEILING`/`FLOOR` in an ODS is `core::CEILING` (the OpenDocument function: signs must agree); `COM.MICROSOFT.CEILING` and an XLSX `CEILING` are `microsoft::CEILING` (Excel's). `FLOOR(x,0)` is `Err:502` in LibreOffice; both Excel functions need two arguments.
- `tests/assets/native.ods` (a spreadsheet created in LibreOffice 26.8) and `tests/assets/null_date_{1900,1904,1950}.ods` join the oracle tests; `scripts/compare_formats.py`, `scripts/function_namespaces.py` (`docs/FUNCTION_NAMES.md` and the generated `functions/namespaces.rs`), `scripts/gen_native_fixture.py` and `scripts/gen_null_date_fixture.py` document that LibreOffice runs both formats with one engine and what each format changes.
