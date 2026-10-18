# Formula Representation Plan (XLSX/ODS)

Status: **proposal** — no code yet. Decisions marked *Open* need an answer
before the work that depends on them starts.

## 1. Scope and correction of the previous version

Three independent processes need a structured view of formulas:

1. **Semantic linting** — rules that depend on the structure of formulas and
   on the relations inside the workbook.
2. **Structural transformations** — edits that rewrite formulas, such as
   deleting or inserting rows/columns and renumbering every reference.
3. **Calculation engine** — the workbook converted into an executable that
   receives input cell values and computes the workbook's results, including
   data validation, filters and conditional formatting. It must be faster and
   use less memory than recalculating the original file in Excel or
   LibreOffice.

The previous version of this document started from a common AST and fitted
the three processes to it. It also misread the engine's "three optimization
levels" as an axis separate from the three access modes; they are the same
axis (§5.2). This version derives each process's representation from what
the process needs, then checks what, if anything, is worth sharing (§6).

## 2. Method

For each process:

- **Information needed** from each formula and from the workbook.
- **Operations** performed on it.
- **When the cost is paid** — per run, per edit, or once at compile time —
  and what dominates performance.
- **Minimal representation** meeting the above.

A component is shared only if every consumer needs it in the same form and
sharing adds no construction or traversal cost to any of them.

## 3. Process 1 — semantic linting

### 3.1 What current rules need

Classification of every formula rule in `sheetrs/src/rules/`:

| Need | Rules |
| ---- | ----- |
| Function-name tokens | VUL602, VUL604, VUL606, CPX504 (IF calls) |
| Token adjacency (operator/literal sequences) | CALC203, VUL603, CALC201 (number literals) |
| Parenthesis/argument depth and counts | CPX505, CPX506, CPX508, CPX509 (length) |
| Argument boundaries of one call | CALC204 (4th argument of lookups), CALC205 (SUM arguments) |
| Reference tokens, decoded | REF307, EXT801, CPX507, REF309, REF310, REF301 (names), REF306 (sheets) |
| Relative shape of the formula | INT401–403, VUL601 |
| Workbook dependency graph | CALC202, ERR102/ERR103 |

No current rule needs operator precedence or a tree: everything above is
computable from one token stream plus a stack for parentheses and arguments
(`helpers::extract_args` already does this by counting parentheses on the
raw string).

A future rule that needs to know **which operands an operator applies to**
(e.g. arithmetic on a text cell) needs precedence. That rule can build a
tree for the formulas it inspects, on demand, without storing one for every
formula.

### 3.2 Cost profile

The lint cost is paid on every run, over every formula (1.86 M on the
production file). Rules currently scan the same string repeatedly with
separate regexes; CALC202 alone is 20.1% of samples, INT401–403 11.9%
(`sheetlint/docs/PROFILING.md`). Throughput and the absence of
per-formula allocations are what matter.

### 3.3 Representation

- **Token stream** per formula, produced once per cell and consumed by all
  rules in `on_cell` (zero-allocation iterator, or a reusable token buffer
  held in `LinterContext` like `formula_upper` today).
- **Relative shape hash** computed from the tokens (replaces the INT regex
  normalization).
- **Lint dependency graph**: cell → cell edges and cell → range edges with
  a per-sheet interval index to answer "which ranges contain this cell".
  This fixes CALC202's documented gap (ranges reduced to their corners). It
  stores no values, types or schedule.

No AST is stored.

## 4. Process 2 — structural transformations

### 4.1 What it needs

For each reference in each formula of the workbook (every sheet, since other
sheets may point into the edited one): its exact byte span in the **raw
source text**, its kind (cell, area, whole rows, whole columns, name,
external), its absolute/relative flags and sheet qualifier, and whether it
sits inside a string literal. Function names matter only to warn about
`INDIRECT`, `OFFSET`, `ADDRESS`, whose targets are computed and cannot be
rewritten.

The same reference handling applies to the non-cell sites: defined names,
conditional formatting and data validation ranges and formulas, merged
cells, hidden ranges, tables, autofilter, print areas, XLSX shared-formula
`ref` ranges.

### 4.2 Operations

A local function of each reference and the edit (Excel semantics):

- reference inside the deleted band → `#REF!`;
- area partly deleted → shrinks; fully deleted → `#REF!`;
- reference after the band → shifts (absolute references shift too: `$`
  controls copying, not structural edits);
- whole-row/whole-column references outside the edited axis are unchanged.

An XLSX shared-formula group whose cells end up with different relative
offsets is split into individual formulas, as `--convert-to-values` already
does.

No dependency graph, evaluation or precedence is involved.

### 4.3 Cost profile and representation

Paid once per edit, inside the streaming writer, which already reads every
sheet's XML. The minimal representation is the **reference token stream of
the raw dialect text**; the edit is applied by splicing the changed spans,
so untouched formulas are written byte-for-byte and edited ones keep their
spacing and prefixes. Nothing is kept in memory beyond the formula being
written.

The reader's normalized `Cell.formula` cannot be used: it is already lossy
(ODS references rewritten, local sheet qualifiers stripped, `A1:A1`
collapsed, `;` turned into `,`).

## 5. Process 3 — calculation engine

### 5.1 Goal

A compiler that turns a workbook into an equivalent executable:

```text
workbook.xlsx/.ods ──compile──► executable ──(inputs: cell values)──► outputs
                                              ├─ computed values
                                              ├─ data validation verdicts
                                              ├─ filter state
                                              └─ applied conditional formats
```

It is faster and lighter than the original applications because it can
drop what they must keep: per-cell objects, dynamic dependency tracking,
generic variant values, UI state, and anything the outputs do not depend on.

### 5.2 Optimization levels = access levels

| Level | Readable after a run | Kept | Constants |
| ----- | -------------------- | ---- | --------- |
| **L1 — any cell** | every cell (constant or formula) and every CF/DV result | everything | stored; fillable cells are inputs |
| **L2 — computed only** | formula cells and CF/DV results | formulas, CF, DV, and the constants they use | used constants kept; unused ones dropped |
| **L3 — explicit output** | declared outputs only | backward slice from the outputs | non-input constants folded into code |

Inputs are the cells the executable accepts values for. *Open:* how they
are declared (explicit list; every constant cell; or cells carrying data
validation, which in form-like workbooks already mark the input cells). At
L3 the choice decides how much can be folded: a constant that may be an
input cannot be folded.

### 5.3 What it needs

- The **full grammar**: precedence, unary/postfix operators, union and
  intersection, array constants, implicit intersection, structured
  references, names (sheet-scoped and global), external references.
- **Resolved, typed** operands: function signatures, number/text/bool/error/
  array/range types, so values can be unboxed where types are provable.
- **Workbook calculation settings** stored in the file (§5.4).
- **Data validation** definitions, **autofilter** definitions and state, and
  **conditional formatting** rules — none of which the reader parses today
  (CF is only counted; DV, autofilter criteria and calculation settings other
  than the 1904 date system are not read).

### 5.4 File-stored calculation settings

The same formula gives different results depending on settings saved inside
the file. The engine reads them and compiles the matching semantics, so
the choice costs nothing at run time.

| Setting | XLSX | ODS (`table:calculation-settings`) | Example |
| ------- | ---- | ---------------------------------- | ------- |
| Criteria syntax | wildcards only | `table:use-regular-expressions`, `table:use-wildcards` | `A1="abc"`, `=COUNTIF(A1:A3;"a.*")`: regex → 1; wildcards → 0 (`.` is literal) |
| Whole-cell criteria | always whole cell | `table:search-criteria-must-apply-to-whole-cell` | `A1="abc"`, `=COUNTIF(A1:A3;"ab")`: whole cell → 0; partial → 1 |
| Precision as shown | `calcPr fullPrecision="0"` | `table:precision-as-shown` | `A1=1/3` shown as `0.33`, `=A1*3`: full precision → 1; as shown → 0.99 |
| Date origin | `workbookPr date1904` | `table:null-date` | `=DATE(2024;1;1)`: 45292 (1900) vs 43830 (1904) |
| Iteration | `calcPr iterate`, `iterateCount`, `iterateDelta` | `table:iteration` | `A1=B1+1`, `B1=A1/2`: iteration on → converges to 2; off → circular-reference error |

Differences between the applications that the file does not record (function
implementation details) are covered by the oracle: each file stores the
results of the application that last calculated it, and the engine is tested
against those cached values.

### 5.5 Data validation, filters, conditional formatting

- **Data validation** compiles to a predicate per validated range, run when an
  input is set. Error style decides the behaviour: `stop` rejects the value;
  `warning` and `information` accept it and report. List validations whose
  source depends on other cells (dependent dropdowns) are graph nodes.
- **Autofilter**: the filter's hidden rows feed `SUBTOTAL(1xx,…)` and
  `AGGREGATE`. Excel and LibreOffice do **not** re-apply a filter when values
  change; hidden rows stay as saved until the user re-applies it.
  Re-applying after each calculation makes the filter depend on results that
  may depend on the filter (through `SUBTOTAL`). *Open:* faithful mode
  (hidden rows as saved, an input), re-apply mode, or both.
- **Conditional formatting** compiles to predicates per rule and range,
  evaluated in priority order with `stopIfTrue`. Rule types that depend on
  the whole range (top-N, above average, duplicates, color scales, data
  bars) need aggregate nodes over the range. The output per cell is the
  resulting format.

### 5.6 Cost profile

Compilation is paid once per workbook version; execution is paid on every
input set and is the target metric. Compile time still matters at 1.86 M
formulas, and code size is bounded by compiling per **formula group** (one
relative shape over a contiguous region) rather than per cell.

### 5.7 Representation

```text
parser ─► shape AST (one per distinct relative shape, compile time only)
       ─► typed IR per formula group (resolved refs, function ids, types)
       ─► dataflow graph over groups, ranges, DV, CF, filter nodes
            (regions split where dependencies overlap partially;
             SCC blocks for iterative calculation)
       ─► slice by level (L1/L2/L3) ─► optimizations ─► back end
```

Optimizations at IR level: constant folding (L3), common subexpressions
across cells (same `SUM` range used by many cells), running aggregates
(`SUM($A$1:A1)` copied down: O(n²) → O(n)), lookup indices built once
(at compile time when the table is constant, at run time otherwise),
unboxed `f64` for numeric groups, columnar storage per group.

Back ends from the same IR:

- **Runtime**: compact bytecode + columnar value store, loaded by a small
  VM (also usable from WASM and the Python binding).
- **Compiled**: generated Rust source per slice, built into a library or
  binary. *Open:* whether a JIT is wanted (not available on WASM).

Execution: full evaluation of the slice in static topological order per
input set. Incremental re-evaluation (dirty groups) is an addition for
interactive use, not the first target.

The engine needs its own AST and IR: neither the lint token stream nor the
transform's reference stream carries precedence, types or resolution, and
the engine's structures are built at compile time, where their cost does not
affect linting or editing.

## 6. What to share

| Candidate | Lint | Transform | Engine | Decision |
| --------- | ---- | --------- | ------ | -------- |
| Stored AST | not needed | not needed | compile-time only, own form | **not shared** |
| Dependency graph | cells + ranges, no types | not needed | groups, typed, scheduled | **not shared** |
| Shape interning | hash only | not needed | interned ASTs per group | **not shared** |
| Dialect lexer (tokens with spans, decoded references) | yes | yes | yes (parser input) | **shared code**, conditional |
| Function name registry (`_xlfn.`, `COM.MICROSOFT.`, ODF names → canonical id) | yes | warnings only | yes | **shared code** |
| Reference types (cell, area, flags, sheet) | yes | yes | yes | **shared code** |

What is shared is code, not data: each process builds its own structures
from the token stream and pays only for what it consumes. The lexer is the
part that handles the two dialects (`[.A1]`, `$Sheet.A1`, `;`, `~`, `!`,
external references, structured references, `_xlfn.` prefixes), which is
where XLSX/ODS parity bugs already occur (#68); three implementations of it
would triple that surface.

Condition: the shared lexer is kept for linting only if it is at least as
fast as the regex scans it replaces on the production file. If a
specialised scan is faster for a rule family (for example function names
with `aho-corasick`), that rule keeps it.

## 7. Phases

The three processes have no dependency on each other beyond the shared
lexer.

| Phase | Deliverable | Exit criterion |
| ----- | ----------- | -------------- |
| **0. Measurement** | On the production corpus: time of the current regex scans per rule; distinct relative shapes and group sizes; function inventory (`sheetstats`); LibreOffice headless and Excel recalculation time and memory for input → output runs | Numbers recorded here |
| **1. Lexer** | Dialect lexer, reference types, function registry | Every corpus formula tokenizes; `minimal_test.xlsx` and `minimal_test.ods` give identical canonical tokens; #68 fixed; faster than the regex scans it replaces |
| **2a. Lint** | Token stream in `LinterContext`; lint dependency graph with range edges; rules migrated by profile share | Violations diff reviewed; each migration benchmarked (`scripts/README.md`) |
| **2b. Transform** | Reference rewriting in both writers; `sheetcli` row/column insert and delete; all sites of §4.1 | Output reopened in sheetrs and LibreOffice/Excel: references correct, cached values unchanged after recalculation |
| **2c. Engine** | Reader support for DV, autofilter, CF rules, calculation settings; parser; IR; graph; L1 runtime back end; functions by measured usage | Match rate against cached values per function |
| **3. Engine levels** | L2/L3 slicing, folding, compiled back end | Faster and smaller than the Phase 0 baselines on the corpus |

## 8. Open decisions

1. Inputs: explicit list, every constant cell, or cells with data validation?
2. Autofilter: faithful (as saved), re-apply after calculation, or both?
3. Compiled back end: generated Rust only, or also a JIT? Must the runtime
   back end run in WASM?
4. Transforms touching charts, pivot tables or external links: refuse, warn
   or rewrite?
5. Baseline for "faster than Excel": which Excel interface (COM
   `CalculateFull`, Office JS) and which LibreOffice path (headless UNO
   `calculateAll`), and is process start-up included?
