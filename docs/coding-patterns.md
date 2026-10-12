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
- `on_cell()` — called for each cell during the walk. Collect data into
  `LinterContext`.
- `on_sheet_end()` — called at the end of each sheet.
- `on_workbook_end()` — called after all cells are visited. Emit violations
  based on collected data.

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

### Formula Normalization

The INT rules normalize formulas to R1C1-style relative offsets so that
structurally identical formulas in different cells produce the same pattern
string (e.g., `=A1+B1` in cell C1 and `=A2+B2` in cell C2 both become
`=R[0]C[-2]+R[0]C[-1]`).

Absolute references (`$A$1`) are normalized to fixed markers (`R0C0`) to
distinguish them from relative references. The regex skips false matches
on function names (e.g., `LOG10`) and sheet qualifiers (e.g., `SHEET1!`)
by checking surrounding context.

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
- `int_min_formula_sequence` (int, default 3) — Minimum formula run length
  for INT401–403 to consider a sequence.
- `ref310_min_adjacent_data` (int, default 1) — Minimum number of contiguous
  adjacent data cells beyond a range boundary to trigger REF310 (Longer Ref
  Expected). PerfectXL uses 1.

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
- **`column_index_to_letters(index)`** — 0-based index to column letters.

## Performance Guidelines

- Use streaming readers — avoid loading data that is not needed.
- Compile static regexes once, in a module-level `LazyLock<Regex>`. Build
  derived search patterns (e.g. `"FUNC("` strings) the same way.
- For case-insensitive formula matching in `on_cell()`, use
  `LinterContext::formula_upper(sheet, cell)`, which converts each cell's formula once
  for all rules. Use the original formula to inspect string literal content.
- In regex replacements on hot paths, implement `regex::Replacer::replace_append`
  and write into the output buffer instead of returning a `String` per match.
- Prefer `include_bytes!` in tests over file I/O.

## CLI Conventions

### Exit Codes

- `0` — success (no violations, or operation completed)
- `1` — violations found (sheetlint) or operation failed
- `2` — configuration or argument error

### Output Streams

- `stdout` — primary output (violations, stats, results)
- `stderr` — progress messages, warnings, errors

### Output Formats

Support `--format text` (default, human-readable) and `--format json`
(machine-readable, CI integration).

## Library Conventions

### Versioning

The workspace uses [semantic versioning](https://semver.org/):

- **Major**: breaking changes to public API
- **Minor**: new features, new rules, backward-compatible changes
- **Patch**: bug fixes, performance improvements

### Public API Surface

The `sheetrs` library crate exports:

- `reader` module — `Workbook` trait, format parsers
- `rules` module — rule registry and `WalkerRule` trait
- `config` module — TOML configuration loading
- `writer` module — workbook modification utilities
- `violation` module — violation types and formatting

### Backward Compatibility

Avoid breaking changes to the public API. When deprecating, add
`#[deprecated]` with a migration note and remove in the next major version.
