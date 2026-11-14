# sheetrs (Python)

Python bindings ([PyO3](https://pyo3.rs)) of the SheetRS suite: query, search, edit, calculate and
lint XLSX and ODS workbooks with the same engine as the `sheetrs query`, `sheetrs edit`, `sheetc`,
`sheetrs lint` and `sheetrs stats` executables. One wheel per platform carries everything (`abi3`,
Python 3.13 to 3.15); there are no optional parts and no dependency besides Python.

```python
import io
import sheetrs

book = sheetrs.Book("book.xlsx")          # a path, bytes or a binary stream (io.BytesIO)
book = sheetrs.Book("book.xlsx", locale="es-CO")   # the regional settings of the workbook (below)
book = sheetrs.Book("book.xlsx", now=45366.75, seed=7)   # NOW() and RAND() fixed in every calculation (below)

# Opening reads nothing. The statistics are a tool, like the others: `stats()` reads the workbook.
stats = book.stats()
stats.sheets.names(), stats.sheets["Data"].used_range, stats.total_formulas
stats.defined_names, stats.has_macros, stats.to_dict()

# Queries read what they need. Several ranges, of any sheets, are read in one pass.
r = book.range(["Data!A1:C10", "Summary!B2"], properties=("value", "formula"))
r["Data!A1:C10"].values                   # a list of rows; None where a cell is empty
r["Data!A1:C10"].formulas
book.range("Data!A1:C10", dates=True)   # date cells as datetime.date / time / datetime
book.find("2024", search_in="values", entire_cells=True)   # a list of Match

# Operations accumulate on the same Book. Nothing runs until save(), the cut.
# Every operation takes the array of its targets: a sweep is one call, not one per target.
(book.delete_rows(["Data!3:10", "Data!15"])           # simultaneous within the call
     .delete_sheets(["Tmp", "Scratch"])
     .formula_to_value(["Report!B:D", "Totals"])
     .replace({"2024": "2025", "Q1": "T1"})      # pairs: each is an operation, in the order given
     .calculate())                         # the engine recalculates the edited workbook
print(book.explain())                      # the passes over the file
clean = book.save("clean.xlsx")            # a Book on the written file
clean.report.to_dict()                     # {"replaced": 3, "written": 0, "skipped": [], "diagnostics": []}
clean.save(io.BytesIO())                   # or to a stream; save() with no destination replaces the original

# calculate() is a step of the plan too. To get its value ranges: in memory, or as JSON when saving.
model = sheetrs.Book("model.xlsx").calculate(inputs={"Data!B2": 1500}, ranges=["Summary!B2:B5"])
model.value_ranges()["Summary!B2:B5"].values    # a ValueRanges of Blocks; .to_dict() gives {"Summary!B2:B5": [[...], ...]}
model.save("answer.xlsx", json="answer.json")   # the workbook calculated, and its value ranges as JSON
print(book.lint("sheetrs.toml").to_sarif())     # the lint report; to_json(), to_text(), to_dict()
book.inspect().unsupported_functions       # what sheetc inspect says
sheetrs.evaluate("=SUM(1,2,3)")            # 6
```

## The `Book`

A `Book` is a workbook: a path, an `os.PathLike`, the bytes of a file (`bytes`, `bytearray`) or a binary
stream that reads and seeks (`io.BytesIO`, a file opened with `"rb"`, read from its start). The format
(XLSX, XLSM or ODS) is the one of the content, not of the extension, and is not an argument. It reads
nothing when it is made and keeps nothing of the workbook: it is not a navigable model of it, there are no row or cell
objects, and what a query does not ask for is neither read nor kept.

| | Members | What they do |
| --- | --- | --- |
| **Format** | `format` | A property: `"xlsx"`, `"xlsm"` or `"ods"`, from the content. |
| **Settings** | `locale`, `now`, `seed` | Read only properties of what was given to `Book(...)`: the locale, the serial number of `NOW()` and the seed of `RAND()`; `None` when not given. |
| **Operations** | `delete_sheets`, `delete_names`, `formula_to_value`, `delete_rows`, `delete_columns`, `hide_rows`, `show_rows`, `hide_columns`, `show_columns`, `hide_sheets`, `show_sheets`, `replace`, `calculate(inputs=None, ranges=None, major="rows", now=None, seed=None)` | Accumulate on the same `Book`, which each one returns, so the calls chain. Each takes the **array of its targets** (one `str`, or any iterable of `str`; an empty one does nothing; anything else in it is a `TypeError`). |
| **The cut** | `save(dest=None, *, json=None)`, `to_bytes()`, `value_ranges()`, `explain()` | `save` plans all the operations in the fewest passes over the file and runs them, and writes the workbook: `save(dest)` a new one to a path or a binary stream, `save()` over the original (the file, or the stream the `Book` was made from); `json=` a path or stream for the value ranges of `calculate`. `to_bytes()` runs them in memory and returns the workbook as `bytes`. `value_ranges()` runs them in memory and gives the value ranges. `explain` says what it will do. |
| **Queries** | `range`, `find`, `lint`, `inspect`, `stats` | Read what they need; `stats()` reads the workbook in a single streaming pass and gives a `Stats` (the eleven `total_*`, `sheets[name]` with `name`, `used_range` (the rectangle of the used cells, `"D10:GV5439"`, not anchored at `A1`; `None` for a sheet with no cells), `state` and the counts, `has_macros`, `defined_names`). They raise `PendingOperationsError` with operations pending. |
| **Free functions** | `evaluate(formula, locale=..., now=..., seed=...)`, `lint_rules()` | The value of one formula (`en` unless `locale` says another; `now` and `seed` fix `NOW()` and `RAND()`); the rules of the linter. |

Names are those of the LibreOffice Calc interface (*Delete Sheet*, *Manage Names ▸ Delete*,
*Formula to Value*, *Calculate*, *Find & Replace*).

Every result is typed and writes itself explicitly, with no implicit `dict`: `to_dict()` gives the
content as Python dicts and lists, `to_json()` as JSON text (and, for the lint report, `to_sarif()` and
`to_text()`). The same content, in two forms, as `to_dict` and `to_json` of a table are: `Stats`,
`Sheet`, `ValueRanges`, `Report`, `Diagnostic`, `Inspection` and `LintReport`.

### Operations accumulate, `save` is the cut

- Each operation addresses the workbook **as it is**, in the coordinates of the original, whatever the
  other operations are: `delete_rows("T!3")` and `delete_rows("T!4")` are the rows 3 and 4 of the
  workbook, in one call or in two. A scope is `"Sheet"`, `"Sheet!B:D"` (columns) or `"Sheet!3:10"` (rows).
- Every operation takes the array of its targets, because the tool is for sweeps: `delete_sheets`,
  `delete_names`, `formula_to_value`, `delete_rows` and `delete_columns` take a `str` or an
  iterable of `str` (a list, a tuple, a generator); `replace(scope=...)`, `calculate(ranges=...)`
  and `range(ranges)` too. `replace` takes pairs for a sweep: `replace({"a": "b", "b": "c"})` or
  `replace([("a", "b"), ("b", "c")])` with no `replacement`; every pair is an operation of the plan,
  in order, that sees what the pair before it left, exactly as if `replace` were called once per
  pair (so the order matters: here `a` ends as `c`); the options (`regex`, `match_case`,
  `entire_cells`, `search_in`, `formatted`, `filtered`) and the `scope` are shared. `find` takes one
  text: its answer is one ordered list of matches, and a regular expression (`regex=True`,
  `"a|b"`) covers a sweep of searches.
- `save` is where the plan is made, so it sees all the operations: it runs them in the fewest
  streaming passes over the file. The operations are **not a sequence**: each is written against the
  workbook as it is (rows, columns and sheets of the original), and the order they were called in changes
  nothing, because the planner decides the order they run in (replacements, then the calculation, then
  the conversions and the deletions in one pass; the conversion reads the values the workbook has, before
  the deletions). To work step by step, `save` between steps and open what was saved. Each pass reads and
  writes the file once, so memory does not grow with the workbook.
- `save(dest)` writes to a path (never the file being edited) or to a binary stream (`io.BytesIO()`) and
  returns a `Book` on what it wrote; this `Book` keeps its operations. `save()` replaces the original
  and returns the same `Book`, which then has none and describes the workbook as it is now: a file is
  replaced atomically, and the stream the `Book` was made from is rewritten from its start and cut
  where the workbook ends. A `Book` opened from bytes has no original to replace.
- `to_bytes()` is the cut to memory: the workbook the operations make, as `bytes` in the format of the
  source, with no file and no temporary file (the editor writes to a buffer). They are the bytes
  `save(path)` writes for the same plan, `Book(book.to_bytes())` reads them back, and
  with no operations it returns the workbook as it is. Like `save(path)` it keeps the operations of
  the `Book`; `report` and the diagnostics say what the run did.
- `calculate(inputs=None, ranges=None, major="rows", now=None, seed=None)` is an operation like the others, not a query: the engine,
  which reproduces LibreOffice, writes `inputs` (`{"Data!B2": 1500}`; a number, text, boolean or a
  list of rows for a range) and calculates every formula of the workbook as it is (before the conversions and
  the deletions of the same `save`, which are written against it too). `save` writes the workbook with the results and
  without the mark to recalculate it on open (as for any `save` after a deletion or a replacement: the
  engine calculates what they leave stale; only a workbook it cannot calculate is marked, with a
  `RecalculationWarning`), so no other program has to open it for its results to be
  up to date. `ranges` (a range or a list of them) are the ranges the answer holds, listed as `major` says
  (`"rows"` or `"columns"`, the `majorDimension` of the request); without them, the outputs the
  workbook declares or, with none, its formula cells. `now` (the serial number that `NOW()` and `TODAY()`
  give) and `seed` (the seed of `RAND()`) fix the volatile functions, as `--now` and `--seed` of `sheetc` do,
  and win over the `now` and `seed` of the `Book` (below); without either the engine reads the clock. They also fix the recalculation that the same `save` runs after
  an edit that leaves results stale (a deletion, a replacement).
- **The value ranges** are read where the user decides: `value_ranges()` gives a `ValueRanges` in memory (a mapping
  from the name of each one to a `Block`; `value_ranges().to_dict()` is `{name: list[list]}`; it runs the
  plan with the workbook written nowhere, so it is allowed with operations pending, and raises
  `sheetrs.Error` without a `calculate` step), and `save(dest, json="answer.json")` writes them as JSON
  (the answer of `sheetc calculate`, `value_ranges().to_json()`) besides the workbook, to a path or a stream. The `calculate` of a `Book` is
  the embedded mode (the workbook is its own template): `ranges` may be inside the outputs the workbook declares (a
  selection among them) or outside them, and a range outside is read, in memory and with no recalculation, from the
  workbook just written with the machinery of `range` (the inputs given are in it, a constant has the value it holds
  and a formula no output includes has the result the template stored: declare it as an output to have it
  calculated), in the order of the request; with no output declared the outputs are the formula cells and the rest is
  read the same way. A range that is no range or name of the workbook, or an input that is not one, raises
  `sheetrs.Error`. The workbook can be read from memory at any time: `sheetrs.Book(book.calculate(...).to_bytes()).range("Data!A1:C9")`.
  (The compiled program, which writes no workbook, refuses a range outside the declared outputs and lists
  them: see [`sheetrs-compiler`](../sheetrs-compiler/README.md#which-ranges-can-be-asked-for).)
- **`use_cache=True`** (a keyword of `save`, `to_bytes` and `explain`; `False`, calculating every formula,
  is the default): the recalculation after an edit calculates only the formulas the operations reach
  (the ones that read a deleted cell or a range that shrank, a replaced cell, a `SUBTOTAL` when rows are hidden
  or shown, a function of position or a name when rows, columns, sheets or names are deleted) and the ones that depend on
  them; every other formula keeps, as the file has it, the result the application that wrote it cached. It takes
  less time and memory the fewer formulas the edit reaches and is the same as without the option where it
  reaches most of them (see `sheetrs-edit/README.md`, *Using the cache*). It also decides what
  `formula_to_value` converts: what the formulas calculate (the engine calculates the cells of the scopes while the
  file is written), or, with `use_cache=True`, the results the file caches for the formulas the edit does not reach.
- **Recalculation on open**, with no option: a workbook whose results the changes leave stale (rows,
  columns, sheets or names deleted) and that no `calculate` calculated again is saved marked so that
  the application that opens it recalculates every formula. When the engine cannot calculate the
  workbook (a function it does not have, a reference it does not read) a `calculate` step does not
  stop the save: the workbook is written as the other operations leave it, marked, and
  `sheetrs.RecalculationWarning` (a `UserWarning`) says that the recalculation that is needed will be
  done by the spreadsheet program that opens it later (see *Diagnostics*). With `value_ranges()` or `json=`
  there is no answer to give, and it is an error.
- **3D references** (`Sheet1:Sheet3!A1`) are not supported by the editor: a deletion in a workbook
  with a formula that holds one raises `sheetrs.Error`, naming issue #75.

### Queries wait for the operations

A query on a `Book` that has operations nothing has run raises `sheetrs.PendingOperationsError`: the
answer would be about a workbook that does not exist yet. `save()` first and query the `Book` it
returns, or run the query before adding the operations if it does not depend on them. `format`,
`explain()` and `repr()` work with operations pending.

### What a query gives back

- **`range(ranges, properties=("value",), *, major="rows", dates=False)`** returns
  `{range: Block}` in the order asked. `ranges` is a range, a defined name that is a range, or a list
  of them. `properties` are `"value"`, `"formula"`, `"number_format"`, `"display"` and `"comment"`: only
  what is asked is read, and a property that was not asked raises `AttributeError`. A `Block`
  iterates and indexes its rows, has a `len`, and `to_dicts()` takes the first row as the header. A
  bounded range is as big as the range, with `None` for an empty cell; a range open at one end (`A:C`,
  `3:10`) is cut at its last cell. `major="columns"` gives lists of columns.
- **`find(text, *, regex, wildcards, match_case, entire_cells, search_in, include_formulas, formatted, filtered, scope, properties, by, dates)`**
  searches as Find & Replace of LibreOffice does and returns a list of `Match` (`sheet`, `cell`,
  `row`, `col` and the properties asked). A regular expression (`regex=True`) may have the anchors `^`
  and `$`, which anchor the text of the cell (`"^(Code|Total)$"`), and a pattern that can match the
  empty text (`x*`, `^`, `$`, `a*$`); `find` and `replace` read them alike, as LibreOffice does: an
  empty match is a hit only at the very end of the text (`"abc"` with `$` replaced by `X` is
  `"abcX"`, `find("$", regex=True)` finds every cell that does not end in a line break, and `"^"`
  none). The empty `text` matches nothing in every mode: `find("")` is `[]` and `replace("", ...)`
  changes no cell, at once and with no error.
- **`value_ranges()`** gives the value ranges of the `calculate` step as a `ValueRanges` (the `ranges` asked
  for, by the text they were asked with, or the outputs of the workbook by name); a formula with a
  function the engine does not have raises `sheetrs.UnavailableFunctionError`, and any other workbook
  the engine cannot calculate `sheetrs.Error`.
- **`lint(config_path=None)`** takes the **path** of a `sheetrs.toml` (the default rules without it; an
  unknown rule, category or parameter is a `ValueError`) and returns a `LintReport`: a sequence of
  `Violation` that writes itself as `sheetrs lint` does: `to_text()`, `to_json()` (the document for an
  interface of your own), `to_sarif()` (SARIF 2.1.0), all as text, and `to_dict()`. Each takes a `file`
  that names the workbook in the document.
- **`stats()`** is a `Stats` that reads the workbook each time it is called; `stats().to_dict()` is
  `sheetrs stats --format json`.
- **`inspect()`** is `sheetc inspect`: an `Inspection` with the sheets, the date system, the cells and
  formulas, the circular groups, the functions the workbook uses (`FunctionUse` with `name`, `calls` and
  `implemented`) and the ones the engine lacks. It compiles the workbook in memory and calculates nothing.

### The locale of a workbook

`Book(source, locale="es-CO")` gives the workbook its regional settings (a tag; `ValueError` when it is
not one): the separators and names that give a text to a number and read a number from a text. They reach
`calculate` and the recalculation that follows an edit (`TEXT("#.##0,00")`, a number joined to a text with
`&`: `55,43%` and not `55.43%`), `range` and `find` (`display`, the dates a search reads) and the
`inspect`; the `Book` that `save` returns keeps it, and `book.locale` says it (`None` when it was not
given). Without it the workbook, an XLSX or an ODS, has the locale of the system (`LC_ALL`, `LC_CTYPE` or
`LANG`, as LibreOffice reads them on Linux, when it is needed; `en-US` when there is none); the language an
ODS stores in its default cell style is not read, as in LibreOffice. The `locale` of a single call (`evaluate(formula,
locale=...)`) wins over it. The engine has `en-US`, `en-GB`, `es-ES`, `es-CO`, `es-MX` and `de-DE`; the
others are an error when the engine needs them (see the `sheetrs-compiler` README).

### The clock and the seed of a workbook

`Book(source, now=45366.75, seed=7)` fixes the volatile functions in every calculation of the workbook, as
`--now` and `--seed` of `sheetc` do: `now` is the serial number that `NOW()` and `TODAY()` give (`TODAY()` is its
integer part) and `seed` the seed of `RAND()` and its relatives (a non-negative integer; distinct seeds give
distinct streams). They reach `calculate`, the recalculation that follows an edit and `inspect`; the `Book` that
`save` returns keeps them, and `book.now` and `book.seed` say them (`None` when not given, and the engine reads
the clock). The `now` and `seed` of a single call (`calculate(now=..., seed=...)`, `evaluate(formula, now=...)`)
win over them, field by field. `RAND` of LibreOffice cannot be seeded, so `seed` is an extension of SheetRS with
no LibreOffice counterpart to match (`sheetrs-compiler/docs/FIDELITY.md`).

### Diagnostics

What a run reports besides its result is a `Diagnostic`: a `level` (`"error"`, `"warning"`, `"info"`,
`"debug"`), a stable `code` that a program can match on (`"RECALC_NOT_SUPPORTED"`: the engine could not
calculate the workbook, so it was written marked to be recalculated), the `message` for a person and the
`context` fields the message was made from (`operation`, `reason`). They are in `report.diagnostics`
after a `save` or a `to_bytes`, and each one also goes to the `sheetrs` logger of the `logging`
module (the `code` and the fields in `extra`: `record.code`, `record.reason`) and, a warning, through
`warnings.warn` (a `RecalculationWarning` for that code, a `UserWarning` for any other).

```python
import logging
logging.getLogger("sheetrs").addHandler(logging.StreamHandler())   # or logging.basicConfig()
```

### Types of a cell

| The cell is | Python |
| --- | --- |
| a number | `float`, or `int` when it is an exact integer |
| text | `str` |
| a boolean | `bool` |
| empty | `None` |
| an error | `sheetrs.CellError` (its `code` is `#DIV/0!`, `#N/A`, or `Err:502` for the ones only LibreOffice has) |
| a date, a time or both, with `dates=True` | `datetime.date`, `datetime.time`, `datetime.datetime` |

A date is a number (the serial the file stores) unless `dates=True`, in every layer of
SheetRS. The date system is the one of the file, one of the three of LibreOffice Calc (1899-12-30, the default and
the 1900 system of an XLSX; 1900-01-01, StarCalc 1.0; 1904-01-01); a workbook with another null date is
an `UnsupportedDateSystemError` when it is read. An
elapsed time (`[h]:mm:ss`) is not a time of day and stays a number. A `CellError` is not a `str`: a cell
with the text `#N/A` is a `str`.

### Errors

`sheetrs.Error` is the base of what the tools raise. `FormatError` (not an XLSX, XLSM or ODS workbook) and `InvalidRangeError` (a range, sheet or name that is not valid) are also
`ValueError`; `UnsupportedDateSystemError` (a date system that is none of the three of LibreOffice
Calc, named in the message) is a `FormatError`; `UnavailableFunctionError` and `PendingOperationsError`
are `Error`. A file that cannot
be read or written is an `OSError` (`FileNotFoundError`, …), and an argument that is not allowed is a
`ValueError` or `TypeError`.

## Install and develop

Not on PyPI yet ([#108](https://github.com/cosmoscalibur/sheetrs/issues/108)); from GitHub (needs Rust ≥ 1.99, `maturin` comes with the build):

```bash
uv add "sheetrs @ git+https://github.com/cosmoscalibur/sheetrs#subdirectory=sheetrs-py"   # in a uv project
pip install "git+https://github.com/cosmoscalibur/sheetrs#subdirectory=sheetrs-py"              # or with pip
```

or a wheel of a [release](https://github.com/cosmoscalibur/sheetrs/releases), with no compiler
(`uv add <wheel url>` or `pip install <wheel url>`). More in the [repository README](../README.md#installation).

```bash
python -m venv .venv && . .venv/bin/activate
pip install maturin pytest
maturin develop -m sheetrs-py/Cargo.toml    # builds and installs the extension in the venv
pytest sheetrs-py
```

The CI builds the wheels (manylinux x86_64 and aarch64) and the sdist with `maturin`, and runs the tests
on Python 3.13, 3.14 and 3.15 (see `.github/workflows/`). macOS and Windows wheels are out of scope for now
(Linux and WebAssembly are the priority): issue #99.

## Stability

The API is **experimental and unstable in every layer** (command line, Rust, Python, WebAssembly) and
may change without notice: a rename or a removal needs no deprecation period. The `.pyi` stubs must say
what the extension has.
