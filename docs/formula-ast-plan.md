# Formula AST Plan (XLSX/ODS)

Status: **proposal** — no code yet. Decisions marked *Open* need an answer
before the phase that depends on them starts.

## 1. Goals

The AST serves three independent processes:

1. **Semantic linting** — rules that need the structure of a formula and the
   relations of the workbook (what a reference points to, which names, sheets
   and ranges feed which cells), not a regex view of the text.
2. **Structural transformations** — edits that rewrite formulas: deleting or
   inserting rows/columns (renumbering every dependent reference), moving
   ranges, renaming or deleting sheets.
3. **Calculation engine** (interpreted or compiled) covering computed cells,
   data validation and conditional formatting, with three optimization levels
   and three access modes (§6).

## 2. Critical flags

These change the validity of the rest of the plan:

- **The three consumers want incompatible things from one tree.** Linting
  wants a cheap, tolerant parse; transformations want lossless output in the
  *source* dialect; the engine wants resolved, typed references. One shared
  tree works only if it stops at the syntactic + reference layer and each
  consumer lowers it to its own form (§3). Three separate ASTs would triple
  the dialect handling, which is where parity bugs already appear (#68).
- **Today's stored formula is lossy.** `Cell.formula` holds an Excel-like
  string: ODS references are rewritten (`normalize_ods_reference`), local
  sheet qualifiers stripped, `A1:A1` collapsed to `A1`, `;` turned into `,`,
  and XLSX shared formulas are expanded by regex (`translate_shared_formula`,
  4.6% of samples on the production file). Neither the transform nor the
  engine can be built on that string: the transform must emit the source
  dialect, and the shared-formula regex does not know about strings, names or
  function names. The parser must read the **raw dialect text**.
- **Scale.** The production file has 1.86 M formulas. A heap-allocated tree
  per cell is not affordable. The plan depends on interning formulas by
  their relative (R1C1) shape so copied formulas share one tree. The dedup
  ratio is an assumption until measured (Phase 0); the production corpus is
  git-ignored, so that measurement runs on the main checkout.
- **Formulas live outside cells.** Defined names, conditional formatting,
  data validation, tables, charts, pivot caches, print areas and autofilters
  all hold references. The reader parses only cell formulas and names today
  (CF is counted, DV is not read). Both the transform and the engine are
  incorrect if they ignore these sites; #56 is the lint-side symptom.
- **A calculation engine is mostly a function library and semantic profile,
  not a parser.** Excel and LibreOffice differ on empty-vs-`""`, text-to-number
  coercion, error propagation, implicit intersection, dynamic arrays and the
  1904 date system. Without a declared semantic profile the engine produces
  values that disagree with the file's own cached results.
- **Explicit outputs alone do not enable constant folding.** Folding a
  constant requires knowing it is not an input. Mode C (§6) needs declared
  inputs as well as outputs, or it degrades to Mode B.
- **Dynamic references cannot be rewritten or statically scheduled.**
  `INDIRECT`, `OFFSET` with computed arguments, `ADDRESS`, text built into
  references: the transform can only warn, and the engine must treat them as
  dynamic dependencies (no static slicing through them).
- **Build vs. reuse.** Rust projects already ship an Excel parser and
  evaluator (e.g. IronCalc, formualizer). Their grammar coverage, ODS support,
  license and memory model must be checked in Phase 0 before committing to an
  in-house parser. The recommendation below (in-house parser, reuse
  function semantics as test oracles) holds only if that check confirms they
  lack ODF dialect support and shape interning.

## 3. Recommended architecture: one front end, three lowerings

```text
 raw XLSX formula ─┐                         ┌─► Lint view   (AST visitors, bound refs, RangeGraph)
                   ├─► Lexer+Parser(dialect) │
 raw ODS formula ──┘          │              ├─► Transform   (ref rewrite → span splice / emitter(dialect))
                              ▼              │
                    FormulaStore (interned   │
                    shapes + sites + anchors)┤
                              │              │
                              ▼              │
                         Binder ─────────────┴─► Engine IR   (typed, resolved) → CalcGraph → O0/O1/O2
```

The shared part is the expensive and error-prone part: dialect grammar,
reference model, function identity, formula sites. Divergence starts after
binding, where each consumer has different needs.

### 3.1 Crate layout

| Crate | Content | Depends on |
| ----- | ------- | ---------- |
| `sheetrs-formula` (new) | lexer, parser, AST arena, dialects, emitter, function registry, reference rewriting | nothing heavy (no zip/xml), WASM-friendly |
| `sheetrs` | reader builds the `FormulaStore`; rules; writer applies transforms | `sheetrs-formula` |
| `sheetrs-calc` (new, Phase 5) | binder → IR, `CalcGraph`, interpreter, compiler, function library | `sheetrs-formula`, `sheetrs` reader model |

A separate formula crate keeps the linter free of engine weight and gives the
Python binding (#14) and the debugger/diff ideas (#8, #6) one reusable core.
Requires updating `ARCHITECTURE.md`, `docs/README.md` and the README project
structure (CONTRIBUTING documentation table).

### 3.2 Lexer and parser

- Hand-written lexer + Pratt parser, one grammar with a `Dialect` parameter:
  - `Ooxml`: `A1`, `Sheet!A1`, `'S n'!A1:B2`, `[1]Sheet!A1`, `,` arguments,
    space intersection, structured references `Table[[#This Row],[Col]]`,
    `_xlfn.`/`_xlws.` prefixes, array constants `{1,2;3,4}`.
  - `OpenFormula`: `of:=` prefix, `[.A1]`, `[$Sheet.A1:.B2]`, `;` arguments,
    `~` union, `!` intersection, `'file'#$Sheet.A1` external refs,
    `COM.MICROSOFT.*` names, `;`/`|` array separators.
- Unparseable input yields an `Unparsed { raw }` node plus a diagnostic, never
  a panic; rules keep working on the text fallback during migration.
- Every node keeps a byte span into the raw text (needed by §5).

### 3.3 AST representation

- Arena per shape: `Vec<Node>` with `u32` indices, no `Box` per node.
- Node kinds: `Number`, `Text`, `Bool`, `Error`, `Array`, `Ref`, `Range`,
  `WholeRows`, `WholeCols`, `Name`, `StructuredRef`, `ExternalRef`,
  `Unary`, `Binary` (incl. `:`, union, intersection), `Percent`, `Call`,
  `Unparsed`.
- **References are stored relative to an anchor**: each axis is
  `Rel(i32)` or `Abs(u32)` (the R1C1 form already used informally by
  INT401–403). The cell, CF range origin or DV range origin provides the anchor.
- **Functions are interned to a canonical `FnId`** through one registry that
  maps `SUM`, `_xlfn.XLOOKUP`, `COM.MICROSOFT.XLOOKUP`, ODF-only names, etc.
  This is the single fix point for #68 and for VUL602/604/606 function lists.

### 3.4 FormulaStore and sites

- `ShapeId` = interned AST of the relative form. Identical copied formulas
  share it; XLSX shared formulas map to the master's shape **without**
  textual translation (removes `translate_shared_formula` from the hot path).
- `FormulaSite` enum: `Cell { sheet, row, col }`, `DefinedName { scope, name }`,
  `ConditionalFormat { sheet, rule, sqref }`, `DataValidation { sheet, sqref }`,
  `TableColumn`, later `Chart`/`Pivot` as read-only reference sites.
- Each site stores `(ShapeId, anchor)`; the cell model gains a `ShapeId`
  (4 bytes) while `Cell.formula` stays during migration for API compatibility.
- Runs of the same shape over contiguous cells (**formula groups**) are a
  first-class index: INT rules, VUL601, transforms (rewrite once per group)
  and the engine's vector kernels (O2) all consume it.

### 3.5 Binder

Resolves, per site: sheet names → sheet index, names → defined-name id
(respecting sheet scope), structured refs → ranges, external refs →
`ExternalWorkbook` index. Output is `BoundRef` (absolute coordinates for a
given anchor). Binding is lazy per consumer: lint needs only refs; the engine
binds everything.

## 4. Consumer 1 — semantic linting

- `LinterContext` exposes the store: `ctx.ast(sheet, cell)` returns the
  shape and anchor; rules visit the AST instead of regex-scanning the text.
- **Dependency graph with range nodes.** CALC202 today keeps only range
  corners (documented false negatives in the middle of ranges). Replace with a
  `RangeGraph`: cell → cell edges plus cell → range edges, and an interval
  index (per sheet, by column then row) to answer "which ranges contain this
  cell". The same graph is the engine's `CalcGraph` skeleton.
- Migration candidates by gain (profile share from `sheetlint/docs/PROFILING.md`):
  CALC202 (20.1%), INT401–403 (11.9%, shape equality replaces regex
  normalization), REF310 (8.7%), CALC201, VUL601 (shape identity), CPX504–508
  (depth/operator/reference counts are AST metrics), REF301 (sites cover names
  used in names, CF and DV — #56), CALC204/205 (argument structure instead of
  `extract_args`).
- New rules enabled by binding (examples): range that skips its total row,
  reference to a cell whose type disagrees with the operator, lookups whose
  table is not sorted for approximate match, CF/DV formulas referencing
  `#REF!`.
- Constraint: the AST must not slow down a full lint run. Parse cost is paid
  once per distinct shape; each migrated rule is benchmarked per
  `scripts/README.md` and recorded in its Comparisons entry.

## 5. Consumer 2 — structural transformations

### 5.1 Model

- `StructuralEdit`: `InsertRows`, `DeleteRows`, `InsertCols`, `DeleteCols`,
  `MoveRange`, `RenameSheet`, `DeleteSheet` (already present without
  reference rewriting).
- `rewrite(shape, anchor, edit) -> RewriteResult`:
  - single reference into deleted area → `#REF!`;
  - range partly deleted → shrink; fully deleted → `#REF!`;
  - references after the edit → shift; absolute references shift too (Excel
    semantics: `$` fixes copy behaviour, not structural edits);
  - the site's own anchor moves, so relative offsets are recomputed in
    absolute space and re-relativized; a group can split into several shapes.
- Dynamic references (`INDIRECT`, `OFFSET` with computed args, `ADDRESS`)
  produce warnings with their sites, never silent edits.

### 5.2 Emission

- **Span splice, not re-printing**: only the reference tokens that changed are
  replaced in the raw text using the AST spans. Untouched formulas are written
  byte-for-byte; edited ones keep their whitespace, casing and prefixes.
  A full emitter per dialect exists for new formulas (e.g. splitting a shared
  group, as `--convert-to-values` already does).
- The writer stays streaming: it already re-reads each sheet's XML, so it
  parses each formula **from the XML at write time** and applies the rewrite.
  The store is used only for the up-front analysis (which sites break,
  warnings), so raw text need not stay in memory.

### 5.3 Sites the writer must update

Cell formulas; `r`/`ref` attributes and row/column indices (XLSX renumbering,
ODS `number-rows-repeated` splitting as in `ods_writer::split_repeats`);
shared formula `ref` ranges and masters; defined names and print areas;
merged cells; hidden row/column ranges; CF `sqref` and formulas; DV `sqref`
and formulas; tables and autofilter ranges; `calcChain.xml` (drop it, as
today). **Charts, pivot caches and external links**: Phase 4 refuses the edit
(or warns, behind a flag) when they reference the edited area — rewriting
them is a later step.

## 6. Consumer 3 — calculation engine

### 6.1 Pipeline

`FormulaStore` → binder → **IR** (typed expression per shape with resolved
function ids and reference descriptors) → `CalcGraph` (cells, ranges, CF/DV
nodes; SCC detection for cycles with an iterative-calculation setting) →
schedule → execution. Value model: `Number | Text | Bool | Error | Empty |
Array`, ranges passed as lazy views. *Open:* semantic profile — Excel,
LibreOffice, or the source format's application.

**Oracle**: every file carries cached results. Evaluating the workbook and
diffing against cached values (per function, per sheet) is the primary test
harness, and it runs on the production corpus.

### 6.2 Access modes

| Mode | Addressable | Kept in the graph | Constants |
| ---- | ----------- | ----------------- | --------- |
| **A — Full** | every cell, formula or constant, plus CF/DV state | everything | mutable inputs |
| **B — Computed only** | formula cells and CF/DV results | formulas + constants they reach | reached ones kept as inputs; unreached dropped |
| **C — Explicit outputs** | declared outputs only | backward slice from outputs | declared inputs mutable; everything else folded |

Mode C requires declared inputs (see §2). The production-run use case behind
`--convert-to-values` is the motivating example: a master workbook compiled
into a function `inputs → outputs`.

CF and DV are output nodes per (rule, range) evaluated lazily per cell:
materializing a boolean per cell per CF rule is too large on real files.

### 6.3 Optimization levels

| Level | Technique | Legal in modes |
| ----- | --------- | -------------- |
| **O0** | Tree-walking interpreter over shared shapes + anchor; topological schedule; dirty propagation for recalculation | A, B, C |
| **O1** | Per-shape bytecode; constant folding of non-input constants; dead-node elimination; cached lookup indices (`VLOOKUP`/`MATCH` tables) and prefix sums for repeated aggregates | folding: C (B if constants are declared immutable) |
| **O2** | Vector kernels per formula group (columnar loops); level-parallel scheduling; AOT code generation of the slice to Rust source, compiled into a standalone function | AOT: C only; kernels/parallel: all |

*Open:* "compiled" target — AOT Rust codegen (portable, works for WASM) vs.
a JIT (Cranelift; not available on WASM). AOT is recommended.

### 6.4 Scope limits for the first engine release

Dynamic arrays/spill, `LAMBDA`/`LET`, external workbooks, RTD/cube
functions, and volatile time functions as anything but inputs are out of
scope. The function library grows by measured usage: `sheetstats` reports
the function inventory of the corpus, and functions are implemented in
descending frequency.

## 7. Phases

| Phase | Deliverable | Exit criterion |
| ----- | ----------- | -------------- |
| **0. Measurement & spike** | Distinct-shape ratio, function inventory, grammar coverage on the corpus; build-vs-reuse check of existing Rust parsers/engines | Numbers recorded here; go/no-go on in-house parser |
| **1. `sheetrs-formula`** | Lexer, parser (both dialects), arena AST, `FnId` registry, emitter, span splice | 100% of corpus formulas parse or yield `Unparsed`; `minimal_test.xlsx` and `minimal_test.ods` produce identical shapes (parity); round-trip tests; fixes #68 |
| **2. FormulaStore in the reader** | Shapes, sites (cells, names, CF, DV), formula groups; shared formulas without text translation | Parse phase not slower than today; memory measured on the production file |
| **3. Lint migration** | `RangeGraph`; CALC202, INT, VUL601, CPX, REF301 on the AST | Same or fewer violations diff reviewed; each rule benchmarked; #56 closed |
| **4. Structural transforms** | `sheetcli --delete-rows/--insert-rows/--delete-cols/--insert-cols` | Re-read output with sheetrs and LibreOffice/Excel: no `#REF!` beyond expected, cached values unchanged after recalculation |
| **5. Engine O0, Mode A** | Binder, IR, `CalcGraph`, interpreter, top functions by usage | Oracle match rate per function on the corpus |
| **6. Modes B/C, O1** | Slicing, folding, bytecode, lookup indices | Speed-up vs O0 recorded in Comparisons |
| **7. O2** | Group kernels, parallel schedule, AOT codegen | Generated function matches O0 on the corpus |

Phases 3, 4 and 5 depend only on Phases 1–2 and can proceed in any order.

## 8. Open decisions

1. "Validation of filters": data validation only, autofilter criteria, or
   both?
2. Semantic profile of the engine: Excel, LibreOffice, or per source format?
3. Mode C: are inputs declared explicitly, or are all constants inputs unless
   marked otherwise?
4. Compiled target: AOT Rust codegen, JIT, or both? Must the engine run in
   WASM?
5. Transform scope for charts, pivots and external links: refuse, warn, or
   rewrite?
6. Is `Cell.formula` (normalized text) kept in the public API after the
   migration, or deprecated per the library conventions?
