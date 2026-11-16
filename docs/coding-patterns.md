# Coding Patterns

Conventions and patterns for the SheetRS Suite. Read this before modifying
core modules.

## General Principles

### Separation of Concerns (SoC)

Rules must be **format-agnostic**. The rule receives a "clean" string, and the
parser handles the "dirt" of the file format.

- Format-specific parsing logic belongs in `<format>_parser.rs` (e.g.,
  `ods_parser.rs`, `xlsx_parser.rs`).
- Common general logic belongs in `parser_utils.rs`, `workbook.rs`, or a new
  file (requires approval).

### One used area

The bounds of a sheet's cells are found once, in `sheetrs-io` (`reader/used_area.rs`): the readers
put an `AreaSink` in front of whatever `CellSink` they are given, so the loaded model and every
streaming consumer get the same numbers, in an XLSX and in an ODS. `Sheet::used_area` is the
rectangle of the **used cells** (a value or a formula), 0-based, which is what `sheetrs stats` shows
(`used_range`); `Sheet::kept_area` also holds the cells kept only for their style and the hidden rows
and columns, which is what REF304 and REF305 measure (`Sheet::kept_extent`, rows `0..=last_row` and
columns `0..=last_col`). The `<dimension>` element is not read. A tool needing the area never
tallies its own.

### Format-Agnostic Design

Rules and logic operate on abstract `Cell` and `Workbook` models, not on
specific file format implementation details. This ensures that every rule works
identically on XLSX and ODS inputs.

### Fail-Fast Configuration

Invalid configuration (unknown rules, bad types) causes immediate startup
failure. Silent misconfiguration is not acceptable.

## Rule Patterns

### The Rule Trait

Each rule is a standalone struct implementing the `WalkerRule` trait. Rules are
registered in a central `Registry`.

Categories: `ERR` (Errors), `CALC` (Calculations), `REF` (References),
`INT` (Interruptions), `CPX` (Complexity), `VUL` (Vulnerability),
`DATA` (Data Issues), `EXT` (External), `HID` (Hidden), `FILE` (Files),
`VBA` (Macros).

### The Walker Pattern

All rules implement the `WalkerRule` trait, which provides lifecycle hooks
for single-pass cell-by-cell traversal:

- `on_workbook_start()` — called once before traversal begins.
- `on_sheet_start()` — called at the start of each sheet.
- `on_cell(sheet, at, cell, …)` — called for each cell during the walk, in row
  order, with its `(row, column)` as `at`: a `Cell` does not repeat its position
  (the sheet's `Grid` has it). Collect data into `LinterContext`.
- `on_sheet_end()` — called at the end of each sheet.
- `on_workbook_end()` — called after all cells are visited. Emit violations
  based on collected data.

Hooks return `()` and push each violation into the `out: &mut Vec<Violation>`
sink that the walker passes to every rule, so a hook that emits nothing leaves
`out` untouched. A unit test collects a hook's output with
`rules::collect_violations(|out| rule.on_cell(&sheet, at, &cell, &mut ctx, out))`.

Hooks take `&mut self`: rules keep their collected state in plain fields
(e.g. `HashMap`), without locks. The registry creates fresh instances for every
lint run (`registry::create_walker_rules`), so each run starts from empty state.
The trait requires `Send` so that a future per-sheet parallel walk can move
rule instances between threads and merge their state.

`Linter::lint_workbook` sorts violations by scope, then rule ID, with a stable
sort, so a rule that emits several violations with the same scope sets their
order. The walker visits cells in `HashMap` order, so such a rule sorts what it
collected (cells, grouping keys, defined names) before emitting, which keeps
the cost proportional to the findings rather than to the cells walked.

### Hierarchical Violations

Violations are reported with full context: File → Sheet → Cell. Use
`ViolationScope` to specify the granularity.

### Multi-ID Rules

When multiple rule IDs share the same collection logic (e.g., INT401–403 all
scan for formula interruptions), a single rule struct reports all of them:

- `InterruptionRule::new(kinds, config)` takes the interruption kinds to
  report and emits each violation under its own rule ID.
- `create_all_walker_rules()` lists one single-kind instance per rule ID, so
  rule metadata, configuration tokens and enablement stay per ID.
- `create_walker_rules()` merges the enabled INT IDs into one instance for the
  run, so cells are classified once however many of them are enabled.

### Rule Dependencies

Some rules emit the violations of other rules or collect data that other rules
read: CALC202 fills `cell_dependencies`, `referenced_sheets` and
`circular_cells`; ERR102 emits ERR103; CPX504 emits the CPX505/506/508/509
cascade. `registry::rule_providers(id)` lists the providers of each rule.

Disabling a rule only removes its own violations. `Linter::with_config` runs
the enabled rules plus their providers (`registry::run_rule_ids`, in execution
order) and `Linter::lint_workbook` drops the violations of rules that are not
enabled. A new rule that reads context data collected by another rule, or whose
violations another rule emits, must be added to `rule_providers`.

### Formula Normalization

The INT rules normalize formulas to R1C1-style relative offsets so that
structurally identical formulas in different cells produce the same pattern
string (e.g., `=A1+B1` in cell C1 and `=A2+B2` in cell C2 both become
`=R[0]C[-2]+R[0]C[-1]`).

Absolute references (`$A$1`) are normalized to fixed markers (`R0C0`) to
distinguish them from relative references. The regex skips false matches
on function names (e.g., `LOG10`) and sheet qualifiers (e.g., `SHEET1!`)
by checking surrounding context.

That is an internal pattern to compare formulas, not the R1C1 notation: SheetRS reads and writes
the A1 notation only (see `docs/approved-differences.md`).

### Argument Separator Normalization

ODS uses `;` as the function argument separator while XLSX uses `,`. The ODS
parser normalizes `;` → `,` during `normalize_ods_reference()` (via
`normalize_arg_separator()`), preserving semicolons inside string literals.
This ensures rules receive Excel-style formulas regardless of input format
(SoC: parser handles format dirt, rules stay format-agnostic).

### Gap Policy (Interruption Context Boundaries)

When scanning for interruptions, `find_interruptions` must decide how far
apart two matching formulas can be and still belong to the same "run". The
algorithm uses a **gap=1 policy**:

- **1 absent cell** between formula cells → bridged as an Empty interruption
  (likely an oversight: missing formula, formatting artifact).
- **2+ absent cells** → the run is broken. The two formula blocks are treated
  as **separate contexts** (headers, spacing, different data sections).

This is a conservative choice that avoids merging unrelated formula regions.
A spreadsheet column often contains multiple independent formula contexts
separated by blank rows. Merging them would produce false "Interrupted by
Other/Empty" violations.

**Performance:** The gap=1 policy makes the algorithm strictly O(n) in the
number of cells. No range iteration is needed — only a single position check
(`pos == run_end + 2`) per cell transition.

### Configuration Parameters

Rules may accept configuration parameters via `LinterConfig`:

- `calc203_allow_double_negative` (bool) — If `true`, `--` (double negative /
  coercion to number) is not flagged by CALC203.
- `locale` (a locale tag such as `"es-CO"`, in `[global]`; not set by default) — The locale of the
  workbook for the tools that read the same file (`sheetrs query`, `sheetrs edit`; the lint rules do
  not use it). A value that is not a tag is a configuration error. In `sheetc.toml` the key is
  `tag` of the `[locale]` table. The resolution order is in *The locale of a `Book`*.
- `now` (a number, in `[global]`; not set by default) and `seed` (a non-negative integer) — The serial
  number of `NOW()`/`TODAY()` and the seed of `RAND()` in the recalculation of `sheetrs edit`, under its
  `--now` and `--seed` flags (the lint rules do not use them). A value of another type is a configuration
  error. In `sheetc.toml` they are the keys `now` and `seed`.
- `enable_features` (array of names, default empty) — Compatibility features,
  all off unless listed; known: `LONGTEXT` (ERR102 accepts the `#NAME?` of
  `_xlfn._LONGTEXT(…)`). Unknown names make the configuration invalid.
- `int_min_formula_sequence` (int, default 3) — Minimum formula run length
  for INT401–403 to consider a sequence.
- `ref310_min_adjacent_data` (int ≥ 1, default 1) — Minimum number of contiguous
  adjacent lines that continue the block (constants after constants, the same
  formula pattern after formulas) beyond a range boundary to trigger REF310
  (Longer Ref Expected). PerfectXL uses 1; keep parity with it.

### Shared Formula Helpers

`rules/helpers.rs` provides formula-parsing utilities shared across rules:

- **`is_inside_string(formula, pos)`** — Check if a byte position is inside
  a `"..."` string literal. Used to skip false function matches.
- **`extract_args(formula, paren_pos)`** — Extract top-level function
  arguments by balanced parenthesis counting. Handles nested calls and
  string literals. Used by CALC204, CALC205.
- **`parse_formula_range(range_str)`** — Parse a range reference (e.g.,
  `A1:B10`, `$A$1`) into 0-based `(start_row, start_col, end_row, end_col)`.
  Strips `$` signs. Used by CALC205.

`reader/parser_utils.rs` provides the column-letter conversions used by both
parsers and rules:

- **`column_letters_to_index(reference)`** — Column letters to a 0-based
  index (`A` → 0, `XFD` → 16383), skipping non-letter characters.
- **`column_letters(index)`** — 0-based index to column letters as an
  allocation-free `impl Display` (built with `fmt::from_fn`). Write it with
  `write!`/`format!`; call `to_string()` only when an owned `String` is needed.
- **`read_text_node(reader)`** — Collect an element's text, resolving entity
  and character references.
- **`push_general_ref(text, reference)`** — Append what an `Event::GeneralRef`
  stands for. quick-xml reports `&quot;`, `&lt;`, `&#10;` apart from the
  surrounding `Event::Text`, so every loop that collects element text must
  handle `GeneralRef` too, and must read with `trim_text(false)`: trimming
  applies to each fragment between references and drops the spaces around
  them.

## Tests

- **A test is named for what it proves**, by itself, in a few words, and the name does not start with
  `test`: `the_empty_search_text_finds_nothing_in_every_mode`, not `test_empty_search`. No comparative
  filler (`too`, `also`), and no sentence that tells the setup instead of the behavior. A name that needs
  more than about 70 characters is testing more than one behavior or describing how: split the test or say
  the behavior shorter. (Some older tests of the linter still start with `test_`; a test that is added or
  touched follows this rule.)
- **Python** (`sheetrs-py/tests`) follows the same rule: `python_functions = ["[a-z]*"]` in `pyproject.toml`
  collects every public function of a test module, so a helper starts with an underscore (`_make_xlsx`) and
  is never collected, and a class or a constant starts with a capital. Python has no logic of its own: its
  tests prove the interface contract (arguments reach the engine, defaults, `sheetrs.Error` and the warnings
  that Rust's errors and reports become). The order of a plan, what a change reaches, the stale rules and the
  text of `explain` are asserted in Rust (`sheetrs-edit`), where the logic is.

## Performance Guidelines

- Use streaming readers — avoid loading data that is not needed.
- Compile static regexes once, in a module-level `LazyLock<Regex>`. Build
  derived search patterns (e.g. `"FUNC("` strings) the same way.
- For case-insensitive formula matching in `on_cell()`, use
  `LinterContext::formula_upper(sheet, at, cell)`, which converts each cell's formula once
  for all rules. Use the original formula to inspect string literal content.
- In regex replacements on hot paths, implement `regex::Replacer::replace_append`
  and write into the output buffer instead of returning a `String` per match.
- Prefer `include_bytes!` in tests over file I/O.
- Do not keep a copy of the workbook's cells in a rule. A rule reads `sheet.cells` (a `Grid`, in
  row order) in `on_sheet_end`/`on_workbook_end`, where it has the sheet; what it must keep
  between hooks is the facts it needs, not the cells. A set of every position cost 40 bytes a
  cell (REF309), and a classification of every cell with a copy of its formula pattern, three
  times over (INT401–403): +590 MB on a 6.6 M-cell workbook. Repeated text (formula patterns) is
  interned and kept as an id.
- The dependency graph (`LinterContext::cell_dependencies`) is a `CellDependencies`: the formula
  cells, sorted, with their references in one flat arena. A `HashMap<Cell, Vec<Cell>>` cost about
  100 bytes a formula and a hash of each cell for the search of loops (CALC202 alone on 2.4 M
  formulas: 782 → 440 MB and 11.3 → 4.5 s with this and one parse per formula). A rule that has to find equal texts among millions (VUL601) keeps the hash of
  each text and the position (16 bytes), sorts them and compares the texts of the equal hashes only,
  not a map of the texts. Read a formula once for what a rule needs from it (CALC202 gets references
  and links from one parse), and read a parameter of the configuration once per sheet, not per cell.
- `compiler::compile` takes the model as `&Model` (the caller keeps it) or as `Model`, which it owns and from which it frees the text of every formula once all are lowered (`ModelCell::release_formula`: `formula` stays `Some("")`, so `is_some()` still says the cell is a formula). Pass the model by value when nothing reads it afterwards (`sheetc build`/`calculate`, `BookCompile`): `formulas.xlsx` peaks at 1152 MB instead of 1309 MB. Code after the lowering must not read the text of a formula.
- Output is written as it is produced (`sheetrs lint --format json` and `--format sarif` serialize each violation):
  building a tree of the whole document costs about 4 KB a violation.
- Memory is measured as the peak of the run (`measure_perf.py` does it for `sheetc`; for the
  linter, `wait4`'s `ru_maxrss` or `/usr/bin/time -v`), with the human and the SARIF output, and a
  change must leave the output byte-identical (compare the MD5).

## CLI Conventions

### Exit Codes

- `0` — success (no violations, or operation completed)
- `1` — violations found (sheetrs lint) or operation failed
- `2` — configuration or argument error

### Output Streams

- `stdout` — primary output (violations, stats, results)
- `stderr` — what a run reports besides its result, as `sheetrs_protocol::Diagnostic` records written by
  the `Logger` (`--log-format human|json`, `-q`, `-v`): errors and warnings by default, what the run did
  with `-v`, the detail with `-vv`

### Output Formats

Support `--format human` (default) and machine-readable formats. The linter writes its own `json`
(flexible, for an interface of your own) and, when asked, [SARIF 2.1.0](https://docs.oasis-open.org/sarif/sarif/v2.1.0/sarif-v2.1.0.html)
(`--format sarif`) for the tools that read the standard. Every format is a library function
(`sheetrs_linter::output`) that the CLI and the bindings call, not CLI code.

## Library Conventions

### Names

The name of an action or of an interface (a subcommand, a flag, a public function or method, a
binding, a WASM export) comes from the language of the LibreOffice Calc interface; if LibreOffice has
none, from xlwings, then openpyxl, then similar tools outside spreadsheets (for the compiler and the
linter, tools of their area). The name is chosen from what SheetRS exposes, never the other way
round. One concept has one verb (`delete` for sheets, names, rows and columns). The Rust API stays
idiomatic Rust where that differs (`write` and not Save, `io::Write` sinks, constructors, errors).

### Rust idioms

- **Conversions are traits.** A text becomes a type with `FromStr` (`"T!A3".parse::<Reference>()`,
  `"1904".parse::<DateSystem>()`), a total conversion is `From` (`Scope::from("Data!A1:C9")`), a fallible
  one `TryFrom`; there is no `parse_*` or `from_*` function that a trait says. A constructor that opens
  something is `open` or `from_*` (`Book::from_file`, `Book::from_bytes`).
- **Getters have no `get_`** (`Workbook::sheet`, `Sheet::cell`); a bare name for what is read, `set_` only
  where a setter is meant.
- **A parameter that is a mode is an enum, not a `bool`** where its meaning is not clear at the call site
  (`loader::from_workbook(wb, Format::Ods)`); options of a search that mirror the protocol keep their
  field names.
- **Construction is by value, plan operations are by `&mut self`.** Configuration of a value that is then
  used (`TextContext::new(..).with_date_system(..)`) consumes `self`; the verbs that accumulate in a
  `Book` take `&mut self` and return it.
- **`#[non_exhaustive]`** on a public enum that the library produces and that callers match on, when it
  may grow (errors, `Format`, `Level`, `LogFormat`, `DateValue`, the output formats, the rule
  categories, `FromCalcError`, `ExcelError`, `EntryFormat`). Not on a protocol enum that callers construct
  and that every tool of the workspace must handle in full (`Property`, `Scope`, `Mode`, `SearchIn`,
  `Operation`, `Typed`), nor on the closed sets that a formatter or a rule maps one by one (`SheetState`,
  `NameScope`, `Severity`, `ViolationScope`, `MajorDimension`, `DateSystem`): adding a variant there is a
  change that every consumer has to implement, so the compiler should say where, and a wildcard arm
  would ignore the new one in silence.
- **The surface is what is documented.** A module that no other crate uses is private (`mod`), not
  `pub mod`; what must be reachable is re-exported at the crate root. The public types are `Send + Sync`
  (a test per crate: `tests/send_sync.rs`) and the entry points have a documented example
  (`no_run` where it needs a file).
- **The clock is an input.** The engine reads `NOW()`, `TODAY()` and the seed from the options of the
  run (`now`, `seed`) and, when they are absent, from `clock::unix_now`, a system clock that a host
  without one (WebAssembly) replaces once with `clock::set_wall_clock`; nothing else in the workspace reads
  the time. The names are the same in every layer: `now` and `seed` of `sheetc`, of a `Calculation`
  (Rust), of `Book.calculate(now=, seed=)` (Python) and of the options of `calculate(request, {now, seed})`
  (WebAssembly), and `--now`/`--seed` of `sheetrs edit` (`Book::with_volatile`) for the recalculation an
  edit leaves to do.
- **The clock and the seed are properties of the `Book`, like its locale.** Every tool that builds or
  calculates (`calculate`, the recalculation of an edit, `build`, `inspect`, `verify`) takes them from the
  book: `Book::with_volatile(Volatile { now, seed })` in Rust (`sheetrs_compiler::memory::Options::with_book`
  applies it under the options of a call), `Book(..., now=, seed=)` in Python, `new Book(bytes, locale,
  {now, seed})` in WebAssembly. The order, from the one that wins: the `now`/`seed` of a single call (a
  `calculate` step, the options of `inspect`, `compile` and `evaluate`), those of the book, the configuration
  file (`now` and `seed` in `[global]` of `sheetrs.toml`, `Volatile::configured`; `sheetc` has its own
  `sheetc.toml`) and the clock. A command line has its own flags that feed the book internally
  (`--now`/`--seed` of `sheetrs edit`) or the build configuration (`sheetc build`, `calculate`, `evaluate`,
  `verify`, and the executables it makes), and a new tool that evaluates volatile functions gets the same two
  flags and the same order. The seed is a SheetRS extension: `RAND` of LibreOffice cannot be seeded.

### The locale of a `Book`

The locale (the regional settings that give a text to a number and read a number from a text) is a
property of the workbook: `Book::with_locale` (a `sheetrs_io::locale::LocaleTag`, which parses with
`FromStr`) and `Book::locale`. Every tool that takes the `Book` follows it: `TEXT` and the conversion of
a number to text in `calculate` and in the recalculation that follows an edit (`sheetrs-edit` gives
every intermediate `Book` of a plan the one of its source, `Book::with_locale_of`), the `display` of
a query and the dates a search reads. The first of these that is set decides, and the same name,
`locale`, is the argument in every layer:

1. the `locale` of a single call, which wins over the rest: the `locale` of a `Query` (`--locale` of
   `sheetrs query`), of `evaluate`, of the options of `inspect` and `compile`, and `--locale` and the
   `[locale]` table of `sheetc` (the build configuration of a program is one thing: its flags win over
   its file);
2. the locale given to the `Book`: `Book::with_locale`, `--locale` of `sheetrs edit`,
   `Book(..., locale=...)` in Python, `new Book(bytes, locale)` in WebAssembly;
3. the configuration file: `locale` in `[global]` of `sheetrs.toml` (`--config`, else the one of the
   working directory) for `sheetrs query` and `sheetrs edit`; these command lines read it
   (`sheetrs_io::locale::configured`, feature `config`) and give it to the `Book` as the item above,
   so that a flag wins over it. (`sheetc` has its own build configuration, `[locale] tag` of
   `sheetc.toml`, which is item 1.) An invalid value fails when the
   configuration is read (the shape of the tag) or when it is first used (a locale the engine's table
   lacks), before any work;
4. the locale of the system (`Book::default_locale`, `locale::system`), for an XLSX and an ODS alike:
   `LC_ALL`, `LC_CTYPE`, then `LANG`, the ones `LibreOffice` reads on Linux (`C` and `POSIX` ask for none),
   read when it is needed, and `en-US` when there is none and in WebAssembly. The language an ODS stores in
   its default cell style is not read: `LibreOffice` does not calculate with it. A system locale that the
   engine's table lacks is `en-US` (it is not an error, as it is not one for `LibreOffice`), and a locale
   that was asked for is an error. **Tests** never depend on the system: `.cargo/config.toml` pins
   `LC_ALL` for `cargo test`, `conftest.py` does for pytest, and a test of the system's locale sets
   the variables in a child process.

`sheetrs_compiler::locale::resolve` is the one place that applies the order for the engine. The
adapters (`sheetrs-py`, `sheetrs-wasm`, the command lines) only convert the argument.

### The calculation mode: the cache is the default, `calculate_all` calculates everything

The default of every run is what `LibreOffice` does when it opens a file and when a cell changes: a
formula that has a result keeps it, and the formulas that a change reaches (and what depends on them)
are calculated. The opt-in is `calculate_all`, which emulates `calculateAll()` of the UNO API: every
formula is calculated. The option is a property of the run, off by default, and its name is
`calculate_all` in every layer, converted only to the convention of each language:
`Book::with_calculate_all` and `Book::calculate_all` (`sheetrs-io`; the planner of `sheetrs-edit` and the
compiler read it), `--calculate-all` of `sheetrs edit`, `sheetc calculate`, `sheetc build` and the
compiled executables, the keyword `calculate_all` of `save`, `to_bytes`, `explain` and `value_ranges` in
Python, and the property `calculateAll` of the `Book` in WebAssembly (`sheetrs_set_calculate_all` in the
module of a compiled workbook). `force_calculate_all` is the option of a build only (`sheetc build
--force-calculate-all`, the key `force_calculate_all` of the options of `compile`): the program stores
no cached result for the formulas it keeps. The option is never implied by another.

What each tool does with it:

- A **build** folds, as constants, the values the workbook stores (what `calculate_all` replaces by the
  results of calculating the whole workbook first) and keeps in the program (`Program::cache`) the cached
  result of every formula it keeps: a run whose inputs leave a formula and what it reads as they were takes
  it (`Engine::take_cached`), and only what a changed input reaches is calculated. A run with
  `calculate_all` (`Overrides::calculate_all`) ignores it. `sheetrs_compiler::Spec` has `calculate_all` and
  `force_calculate_all`.
- A **calculation** (`calculate`, in `sheetrs-edit`, `sheetrs-answer` and `sheetc calculate`) builds in the
  mode it is given (`force_calculate_all` does not apply) and runs the program with it. It is the calculation that
  tells its build what it is for: the workbook as the answer (`Spec::whole_workbook`) and the inputs it was given
  (`Spec::given_inputs`, the only cells that can change, so the only cone that is calculated). A build, embedded or
  not, keeps its own rules whatever calls it.
- An **edit** (`sheetrs edit`) hands the compiler what its plan changed (`sheetrs_compiler::Cache`, built
  by the planner from the operations it has run, `sheetrs-edit/src/plan.rs`): the formulas those changes
  reach, and the ones that depend on them, are calculated, and the rest keep their result; with
  `calculate_all` the whole workbook is calculated. The conversion of formulas to values converts the
  cached results of the cells nothing reaches and what the engine calculates for the others (all of it
  calculated first, with `calculate_all`).

### What `LibreOffice` calculates on load is a root of the cache, not a walk

A formula that `LibreOffice` calculates when it opens a file is calculated by the default mode too (`keep_cached`, `sheetrs-compiler`), and the
kind of root says whether what reads it is: the functions it calculates on load in both formats (`HYPERLINK`, `CELL`, `INFO`, `FORMULATEXT`, the
volatile ones, `INDIRECT`, `OFFSET`) are roots **with** the walk (`Flags::volatile`, `Flags::on_load`: what reads them is calculated), and the
ODS cells whose cached result it recalculates (`ModelCell::recalculated_on_load`, told by the reader through `CellSink::recalculated_on_load`
before the cell) are roots **without** it (a cell that caches a text, or a number in a cell with a format, keeps its result whatever the cell it
reads gives). A rule of this kind is written from a probe of `LibreOffice` (`sheetrs-edit/tests/assets/write_to/probe_*.py`), checked in the oracle
(`make_oracle.py`) and run now against it by `write_to_libreoffice.rs` for every fixture and both formats; a fixture whose cells give another
result each time (`NOW`, `RAND`) cannot be compared, so it is only probed.

### The payload schema and the ABI of a module

The program a compiled executable or WebAssembly module carries is serialized with `postcard`, which is positional: a field
inserted, removed, moved or retyped in `Program` or in what it holds changes what the same bytes mean, and an old payload is
read as garbage or fails far from the cause. So `sheetrs_engine::bundle::SCHEMA` is the number of that layout, written next to
the magic of the payload and checked by the runtime before it decodes (a refusal names both numbers); `sheetc` refuses a stub or
a base module whose schema or ABI is not its own. **A change of the order or of the type of a field of what is serialized bumps
`SCHEMA`** (the test `the_schema_number_follows_the_layout_of_the_program` pins the bytes of a program that sets every field and
fails until you bump it and write the bytes again). `bundle::WASM_ABI_VERSION` (`sheetrs_abi`) is another number: an integer
bumped only on an incompatible change of the exports of a module or of what they mean; an added export (`sheetrs_set_calculate_all`,
`sheetrs_payload_schema`) never bumps it, a host looks for the export. A module is only valid with the `sheetc` that built it.

### Versioning

The workspace uses [semantic versioning](https://semver.org/):

- **Major**: breaking changes to public API
- **Minor**: new features, new rules, backward-compatible changes
- **Patch**: bug fixes, performance improvements

These are the rules for when the API stops being experimental; while it is experimental (see API
stability) a minor or a patch version may also change it.

### Public API Surface

Every tool is a library plus an executable; the bindings are adapters
over those libraries. Where code lives follows who uses it (see `ARCHITECTURE.md`, *Layering*):

- `sheetrs-io` (shared kernel): `reader` (the `Workbook` model, `Grid`, `CellSink`, `ExcelError`,
  the XLSX/ODS parsers, and `Format`), `locale` (`LocaleTag`, the locale of the system and of a
  configuration file) and `Book` (`sheetrs_io::Book`: the one reference to a workbook that every tool takes, with its locale and the operations pending on it). Add to it only what two tools need and has no meaning outside the
  workbook model; code used by one tool belongs to that tool's crate.
- `sheetrs-linter`: `rules` (`WalkerRule`, the registry and the catalog), `config`, `violation`,
  `report` (`group_by_scope`, `ViolationRecord`), the output formats (`human`, `json`, `sarif`, and
  `output::write_report` over an `OutputFormat`), `Linter`.
- `sheetrs-stats`: `BookStats::stats` and its result `WorkbookStats`.
- `sheetrs-edit`: `BookEdit` (a trait for `Book`; every change is written against the workbook as it is and the order they are pushed in changes nothing: the planner, a set of functions that take the `Book`'s operations, runs them in a canonical order; `write(path)`, `write_in_place()` and `write_to(out)` plan them into the fewest passes over the file and `explain` says how; work step by step with a `write` between steps), `RowRange`, `ColumnRange`, `ValueScope`. `calculate(Calculation)` is an operation of the plan run by the engine of `sheetrs-compiler` (the editor depends on it; nothing is injected, and the same engine calculates again the results a conversion of formulas to values reads after a change that leaves them stale): the workbook is written with fresh results and without the mark to recalculate on open (also for what a deletion or a replacement leaves stale: the planner ends the run with the engine calculating it again), or, only where the engine cannot calculate it, marked and with a `Warning` in the `Report` the write returns, which also holds the value ranges (`Calculation::answer`, `Report::value_ranges`, `Report::write_json`). The planner is private: `BookEdit` is the one way in. The writers are `sheetrs_io::write` (feature `write`, with their own error `sheetrs_io::write::Error`): a crate that only reads never enables it, and a new writer capability goes there, not in `sheetrs-edit`. `sheetrs-answer`: `fill_template*` (the answer of a program written into its workbook).
- `sheetrs-engine` and `sheetrs-compiler`: each has its own `Error` (typed, `thiserror`). The compiler builds a `Model` from a `Book` (`loader::load`; `BookCompile` has `compile`, `calculate`, `inspect`, `verify`) and does not re-export the engine: what is the engine's is imported from `sheetrs_engine`.

**Interface convention (every crate and binding).** There is one interface, the workbook: `Book`
(`sheetrs_io::Book`; no `Source`, no `Editor`, no other type in front of it). It holds where the workbook is,
its format and the operations pending on it, and what it does itself is reading the workbook
(`workbook`, `workbook_into`, `analysis_into`, `selection_into`, `date_system`, `defined_names`). Each tool
adds its verbs in its own crate as an extension trait of `Book` (`BookQuery::query`, `BookEdit::delete_rows`
and `write`, `BookCompile::calculate`, `BookLint::lint`, `BookStats::stats`), which the caller imports with
`use`; the core cannot depend on the tools, so this is how a tool adds to the `Book` without the core
knowing it. An operation that is pending lives in the `Book` (`push`, `operations`), is written against
the workbook as it is, and is ordered by the planner of the tool, not by the caller. Python has one
object, `Book`, a lazy reference to a workbook whose operations accumulate until `save` (the tools are its
methods); there is no facade crate, since each tool has its own executable and use. A new tool follows the
rule instead of choosing a shape of its own.

Keep logic in the libraries and out of the executables and the bindings (`sheetrs-wasm`,
`sheetrs-py`): an adapter converts arguments and results and nothing else, and what two adapters
both present (a grouping, a flat record) goes in the library. The API of a binding is organized by
tool (the rules are part of the linter, so they sit next to `lint`). A change to a binding's API
updates its README (and the type stubs of `sheetrs-py`), and a new capability is exposed in both
bindings.

Path dependencies between publishable crates carry a `version`
(`sheetrs-io = { path = "../sheetrs-io", version = "0.4.1" }`) so that they can be published;
`cargo set-version --workspace <version>` (cargo-edit) updates them with the workspace version.

### API stability

The API is **experimental and unstable in every layer** (command line, Rust, Python, WebAssembly) and
may change without notice: a rename or a removal needs no deprecation period. The stubs must say what
the extension has.
