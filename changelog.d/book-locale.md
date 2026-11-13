### Added

- The locale is a property of the workbook (issues #106 and #84): `Book::with_locale` in Rust, `--locale` in `sheetrs edit` (and `sheetrs query`, which had it), `Book(..., locale=...)` in Python, `new Book(bytes, locale)` in WebAssembly. It reaches `calculate`, the recalculation that follows an edit, `TEXT` and the conversion of a number to text, and the `display` and the dates of a query: `TEXT(350*49799,"[$$-en-US]#.###.##0")` is `$17.429.650` and `"x "&(0.5543*100)&"%"` is `x 55,43%` with `es-CO`, as in LibreOffice. The `locale` option of a single call (`evaluate`, `inspect`, `compile`, a query) still wins.
- Without a locale given, the workbook has the one of its format: an ODS says it (the language and country of its default cell style), an XLSX takes the locale of the system (`LC_ALL`, `LC_CTYPE`, then `LANG`, as LibreOffice reads them on Linux; `en-US` when there is none and in WebAssembly). `sheetrs query` and `sheetrs edit` read the `locale` of `[global]` in `sheetrs.toml` (`--config`, or the file of the current directory) between the two; the linter refuses a value that is not a locale tag.
- The locale `de-DE` (separators, names of months and days, text dates), with its fixtures. The keywords of its format codes (`TT.MM.JJJJ`, `[ROT]`) are read in English; the difference is a proposal in `docs/approved-differences.md`.
- `sheetrs_io::locale`: `LocaleTag`, `system`, `from_environment` and, with the feature `config`, `from_config`.

### Changed

- A build, a calculation or an edit of a workbook that states no locale no longer assumes `en`: it takes the one of the format (see above). A workbook calculated with a locale other than `en-US` in the environment of the process gives the text of that locale.
- `[$-407]` is read (German); the examples of a language the engine has no names for are now `[$-40C]` (French).
