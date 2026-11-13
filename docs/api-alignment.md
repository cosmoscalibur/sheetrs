# API alignment across the four layers

SheetRS is used from four layers: the **command line** (`sheetrs`, `sheetc`), **Rust** (the crates), **Python**
([`sheetrs`](../sheetrs-py/README.md)) and **WebAssembly** ([`sheetrs-wasm`](../sheetrs-wasm/README.md)). They
expose the same features with the same names and the same results; this document holds the rules that keep them
aligned and the table that maps each operation to its form in every layer.

> **The API is experimental and unstable in every layer.** There is no contract: names, signatures and JSON shapes
> change without notice until a version says otherwise.

## Rules

1. **Behavior is LibreOffice's.** The reference is LibreOffice 26.8.x for every result; a difference exists only if it
   is approved and recorded in [approved-differences.md](approved-differences.md). Every feature works on XLSX and ODS.
2. **Names come from the interface language of LibreOffice** (delete, find, replace, calculate, range), the same in all
   layers (`snake_case` in Rust and Python, `camelCase` in WebAssembly, `--kebab-case` flags).
3. **One entry point: `Book`.** In Rust `sheetrs_io::Book`; in Python `sheetrs.Book`; in WebAssembly `Book`. The
   command line is the same thing with one command per tool: `sheetrs query|edit|lint|stats` and
   `sheetc build|calculate|inspect|verify`. Each tool adds its operations to `Book` (in Rust, as an extension trait
   per crate: `BookQuery`, `BookEdit`, `BookCompile`, `BookLint`, `BookStats`).
4. **The origin is bytes or a path; the format is detected from the content** (the ODS `mimetype`, the XLSX
   `[Content_Types].xml`, the macro content type of an `xlsm`). No layer has a `format` argument and the extension is
   never consulted.
5. **Operations are written against the workbook as it is** (original coordinates) and they accumulate on the `Book`
   until a cut (`save`, `write`, `--output`). The order in which they are given does not matter: the planner decides
   the order and the passes over the file, writing intermediate workbooks when it needs them. To see the workbook
   after an operation, cut there (one command per operation, or a `save` in between).
6. **Every operation takes a list of targets** (`"Data!3:10"`, `"Data!B:D"`, a sheet name, a defined name).
7. **`calculate(inputs, ranges, major)`** is a request (`inputs`, `ranges`, `majorDimension`), and its answer is
   `valueRanges` (`[{range, name, values}]`, as the values of the Google Sheets API). The engine recalculates what an
   operation leaves stale; the workbook is marked to be recalculated on opening only when the engine cannot.
8. **Reading is lazy, statistics are not.** Opening a `Book` in Python and WebAssembly reads the workbook once for its
   statistics: `format`, `date_system`, `has_macros`, `sheet_names`, `sheets[name]` (`name`, `used_range`, `state`,
   counts), `defined_names` and the eleven `total_*` counters are properties.
9. **Typed results with explicit conversions:** `LintReport`, `Stats`, `ValueRanges`, `Inspection`, `Report` and the
   blocks of a read are objects; `.to_dict()` / `.to_json()` (`toJSON()` in WebAssembly) and, for the lint report,
   `.to_sarif()` give the document. Nothing is converted implicitly.
10. **Bytes and streams in memory in every layer:** `Book` reads from bytes, a path or a stream, and `save` writes to a
    path, a stream or returns bytes; the standard input and output of the command line carry the request and the
    answer (`--request -`).
11. **`dates` is a flag:** by default dates are serial numbers; with `dates` (`--dates`, `dates=True`, `dates: true`)
    they come as `{"type": "date" | "time" | "datetime", "value": "<ISO 8601>"}`. A workbook has one of three date
    systems (1899-12-30, 1900-01-01, 1904-01-01); another is an error.
12. **Structured logs:** every diagnostic has a level, a stable code, a message and context fields. The command line
    writes them with `--log-format json` (`-q`, `-v` choose the levels); Python sends them to `logging` (with `extra`)
    and to `warnings`; Rust returns `Diagnostic`s; WebAssembly calls `onDiagnostic`.

## What a layer does not expose, and why

| Not in | What | Why |
| ------ | ---- | --- |
| Python | `compile` (`sheetc build`) | a compiled workbook is an executable or a module: its consumers are the command line and WebAssembly; Python gets the same answers from `calculate` |
| Python, WebAssembly | `verify` (`sheetc verify`) | it compares the cached results of a file with a recalculation: a check of the project, run from the command line |
| WebAssembly | files by path | the host gives bytes (`Uint8Array`) and receives bytes |
| Rust, command line | the shortcut properties of rule 8 | a statistics read is one call (`BookStats::stats()`, `sheetrs stats`); `Book::sheet_names()` and `Book::defined_names()` read on demand |

## Operations

`Book` in Rust is `sheetrs_io::Book` with the traits of each crate imported (`use sheetrs_edit::BookEdit;`...).

| Operation | Command line | Rust | Python | WebAssembly |
| --------- | ------------ | ---- | ------ | ----------- |
| Open | `-i FILE` | `Book::from_file(path)`, `Book::from_bytes(data)`, `Book::from_reader(r)` | `Book(path \| bytes \| stream)` | `new Book(bytes)` |
| Cells of ranges | `sheetrs query --range R`, `--sheet S`, `--workbook`, `-p PROP` | `BookQuery::range(ranges, &RangeOptions)` | `book.range(ranges, properties, major=, dates=)` | `book.range(ranges, options?)` |
| Find | `sheetrs query --find TEXT` (`--regex`, `--wildcards`, `--match-case`, `--entire-cells`, `--in`, `--formatted`, `--filtered`, `--include-formulas`) | `BookQuery::find(&Find, &FindOptions)` | `book.find(text, regex=, wildcards=, match_case=, entire_cells=, search_in=, formatted=, filtered=, scope=...)` | `book.find(text, options?)` |
| The whole query | `sheetrs query --query JSON` | `BookQuery::query(&Query)` | `book.range`, `book.find` | `book.query(query?)` |
| Sheets and names | `sheetrs query --list-sheets`, `--list-names` | `Book::sheet_names()`, `Book::defined_names()` (read on demand) | `book.stats().sheets.names()`, `book.stats().defined_names` | `book.stats().toJSON().sheets`, `book.stats().definedNames` |
| Statistics | `sheetrs stats [--format json]` | `BookStats::stats()` | `book.stats()` (`Stats`: `total_*`, `sheets`, `has_macros`, `defined_names`) | `book.stats()` (`Stats`: `toJSON()`, `hasMacros`, `definedNames`) |
| Lint | `sheetrs lint [-c CONFIG] [-f human\|json\|sarif] [-e]` | `BookLint::lint(&Linter)` | `book.lint(config_path)` | `book.lint(configToml?)` |
| Delete sheets | `sheetrs edit --delete-sheets S...` | `BookEdit::delete_sheets(names)` | `book.delete_sheets(names)` | `book.deleteSheets(names)` |
| Delete defined names | `--delete-names N...` | `delete_names(names)` | `book.delete_names(names)` | `book.deleteNames(names)` |
| Formulas to values | `--formula-to-value SCOPE...` | `formula_to_value(scopes)` | `book.formula_to_value(scopes)` | `book.formulaToValue(scopes)` |
| Delete rows | `--delete-rows ROWS...` | `delete_rows(rows)` | `book.delete_rows(rows)` | `book.deleteRows(ranges)` |
| Hide / show rows | `--hide-rows ROWS...` / `--show-rows ROWS...` | `hide_rows(rows)` / `show_rows(rows)` | `book.hide_rows(rows)` / `book.show_rows(rows)` | `book.hideRows(ranges)` / `book.showRows(ranges)` |
| Hide / show columns | `--hide-columns COLUMNS...` / `--show-columns COLUMNS...` | `hide_columns(columns)` / `show_columns(columns)` | `book.hide_columns(columns)` / `book.show_columns(columns)` | `book.hideColumns(ranges)` / `book.showColumns(ranges)` |
| Hide / show sheets | `--hide-sheets SHEET...` / `--show-sheets SHEET...` | `hide_sheets(names)` / `show_sheets(names)` | `book.hide_sheets(names)` / `book.show_sheets(names)` | `book.hideSheets(names)` / `book.showSheets(names)` |
| Delete columns | `--delete-columns COLUMNS...` | `delete_columns(columns)` | `book.delete_columns(columns)` | `book.deleteColumns(ranges)` |
| Find and replace | `--replace TEXT --with TEXT` (repeatable), `--regex`, `--wildcards`, `--match-case`, `--entire-cells`, `--formatted`, `--filtered`, `--in`, `--scope` | `replace(...)` | `book.replace(...)` | `book.replace(replaces, options?)` |
| Calculate with a request | `sheetc calculate -i FILE --request JSON --range R [-o OUT]` | `BookEdit::calculate(Calculation { inputs, ranges, major, .. })`, `BookCompile::calculate(...)` | `book.calculate(inputs, ranges, major=)` | `book.calculate({inputs, ranges, majorDimension})` |
| The answer in memory | `sheetc calculate` without `-o` (the JSON on the standard output) | `value_ranges()` | `book.value_ranges()` | `book.valueRanges()` |
| Plan | `--dry-run` | `explain()` | `book.explain()` | `book.explain()`, `book.pending` |
| The cut | `-o OUT`, `--in-place` | `write(path)`, `write_to(writer)`, `write_in_place()` | `book.save(dest=None)`, `book.to_bytes()` | `book.save()` (`Uint8Array`) |
| Inspect a program | `sheetc inspect` | `BookCompile::inspect(&Options)` | `book.inspect()` | `book.inspect(options?)` |
| Compile | `sheetc build [--target native\|wasm]` | `BookCompile::compile`, `build_wasm` | not exposed | `book.compile(options?)` |
| Verify | `sheetc verify --reference cache\|libreoffice` | `BookCompile::verify` | not exposed | not exposed |
| One formula | `sheetc evaluate` | `sheetrs_compiler::memory::evaluate` | `sheetrs.evaluate(formula)` | `evaluateFormula(formula)` |
| List the lint rules | not exposed (the rules are in the [rule reference](../sheetrs-linter/README.md)) | `sheetrs_linter::rules::registry::rule_metadata(&config)` | `sheetrs.lint_rules(config_path)` | `lintRules()` |

The wording of the flags and the options of a call is in `sheetrs <tool> --help`, in the
[Python stub](../sheetrs-py/python/sheetrs/_sheetrs.pyi) and in the
[WebAssembly README](../sheetrs-wasm/README.md); the rules of the planner are in the
[README of `sheetrs-edit`](../sheetrs-edit/README.md).
