# Spike: interpreter, JIT and LLVM for ODS/XLSX

Question: what do a direct interpreter (no build), a JIT and an LLVM compiler add to SheetRS, and are
they viable? Three goals, not simultaneous: (1) cut the time of `calculate` (and `edit`) by skipping the
build; (2) run a workbook directly and speed it up progressively in a service that evaluates the same
template again and again; (3) cut time and peak memory of the run of a compiled workbook.

Status: evaluation with measurements on one generated workbook and one profile of small sample size
(limits in [What this does not show](#what-this-does-not-show)). Nothing here is implemented.

## Findings that change the question

1. **The engine is already an interpreter.** `Engine::eval` walks `Expr` trees over a demand-driven
   schedule ([DESIGN.md](../sheetrs-compiler/docs/DESIGN.md#hybrid-compile--run-time): "the JIT is not machine
   code: it is lazy resolution on the same interpreter"). `sheetc calculate` already builds in memory
   and writes no executable. "Omitting build" therefore means skipping *compile stages* (cone, graph,
   SCC, taint, fold, assembly), not skipping an executable.
2. **The evaluation cost is address resolution, not dispatch.** In the profile of a run, 59 % of
   `Engine::calculate` is `cell_value`; the hash lookup `CellRef → slot` alone is 38 %, the binary search in
   the blocks of constants 20 %. Arithmetic and function bodies (`ROUND`, `MAX`, `IF`) are marginal.
   A JIT that compiles the tree but keeps `CellRef` operands removes the part that costs little.
3. **"No double execution engine" is an explicit requirement** (DESIGN.md, *One IR, one engine*), and
   fidelity to LibreOffice is the acceptance criterion ([CONTRIBUTING.md](../CONTRIBUTING.md#parity-and-fidelity)).
   `eval` carries that fidelity in its control flow: `logical_fmt`, ghost arrays, element-wise lifting,
   propagated errors, the `notes` violations, blank results. A second backend must reproduce all of it
   or bail out to the interpreter; either way it is a decision for the maintainer, not a refactor.
4. **Lazy functions take `Expr`, not values** (`Imp::Lazy(f)`: `IF`, `IFERROR`…). Compiled code cannot
   evaluate them without calling back into the interpreter on a tree that must still exist.

## Measurements

`gen_large_fixture.py … 100000`: 800 005 formulas, 9.7 MB, one declared input and five declared outputs.
`dist` build, 4-core VM, `SHEETC_TRACE=1`, one machine and a few runs: indicative, not a benchmark
(the benchmark protocol is in [performance.md](performance.md#method)).

`sheetc calculate -i` takes 2.49 s (JSON) and 2.67 s (`-o`):

| Stage | Time | Share | Skipped by a direct mode? |
|---|--:|--:|---|
| read workbook into the `Model` | 0.69 s | 28 % | no |
| parse and lower formulas | 0.27 s | 11 % | no (lazy parse only of what outputs demand) |
| bindings, constants that start a path, index | 0.06 s | 2 % | partly |
| cone + CSR graph | 0.17 s | 7 % | yes |
| SCC + taint | 0.04 s | 2 % | yes (needed only for cycles) |
| fold (evaluates input-independent cells) | 0.35 s | 14 % | no: the same evaluation happens in the run |
| assembly of the `Program` and the rest | 0.51 s | 20 % | in part (about 0.3 s) |
| run | ≈0.40 s | 16 % | no |

Profile of the run of the compiled executable (`--features profile`, 126 samples at 100 Hz, so about ±8 %):
`calculate` 67 %, decoding the `postcard` payload 21 %, building the `CellRef → slot` index in
`Engine::build` 9 %. Inside `calculate`: `cell_value` 59 % (hash lookup 38 %, constant blocks 20 %).

Formula shapes (cells whose formula is the same after R1C1 normalization):

| Workbook | Formulas | Shapes |
|---|--:|--:|
| `gen_large_fixture.py` 100 k rows | 800 005 | 14 |
| `loan.xlsx` | 1 086 | 9 |
| `semantics.xlsx` (the oracle: distinct by design) | 34 295 | 33 998 |

Only the first two say anything about real models, and both are synthetic or tiny. **The shape census of
real workbooks is missing and is the data that decides the memory proposals below.**

## Option 1: direct mode (interpret without compile stages)

For `calculate` and for the recalculation that ends an `edit` plan. In the latter every formula is an
output and the inputs are deduced, so the cone prunes nothing and nothing folds: the graph, the fold and
the assembly are pure cost.

- **Gain.** Skipping cone, graph, SCC, taint and part of the assembly: about 0.4-0.55 s of 2.49 s
  (16-22 %) here. The floor is reading plus parsing (about 40 %). For `edit` on 2.4 M formulas, graph +
  fold + assembly are 2.95 s of a 6.85 s build ([performance.md](performance.md#sheetc-calculate-and-sheetc-build)),
  while the edit scenarios take 17-44 s: **at most 7-17 % of `edit`**, because its time is in the
  deletion pass, the second read and the write. The goal "benefits `edit`" holds, but the gain is bounded.
- **Peak memory: little.** The graph is already built into compressed rows and freed before assembly;
  the peak is the cells and the `Expr` trees (about 480 B per formula at 2.4 M: 1152 MB).
- **One engine.** It is the same `Engine` with a different schedule: lowered cells in row-major order,
  demand-driven for the rest (the engine's `Pending/Busy/Done` states). Dynamic references (`INDIRECT`,
  `OFFSET`) become simpler than today, where one makes the whole workbook reachable.
- **Risks, with their guard.**
  - Recursion depth: a chain of 300 000 cells evaluated on demand overflows the stack. Row-major order
    avoids it for backward references; forward chains need an explicit stack or a depth limit.
  - Cycles and iterative calculation need the SCC and give LibreOffice's `Err:522`/`Err:523`. On the first
    `Busy` read or a depth beyond the limit, **fall back to the compiled path** (same engine, so fidelity
    is kept).
  - `RAND`/`RANDBETWEEN` draw from one generator in evaluation order, so a different order gives other
    draws: acceptable for a non-deterministic function, but the equivalence test must exclude it.
- **Verdict: viable, low risk, modest gain.** Do it second, after the index of the next section, because
  that is cheaper and larger.

## Option 2: the cheap structural step that every other option needs

Not in the request, but the profile points at it, and a JIT would need it anyway.

- **Dense index instead of the hash.** `Engine` resolves `CellRef` through `FastMap<CellRef, u32>` for every
  operand and builds it for every cell at start. The project already has a dense row-indexed `Grid`
  (`cellstore.rs`). Replace the map by per-sheet grids of slots, or resolve the slot at lowering time
  (`Expr::Cell(slot)`). Upper bound from the profile: up to 38 % of the evaluation and most of the 9 % of
  start-up. This is the experiment to run first, with the alternating-run protocol of DESIGN.md,
  *Measuring a change*.
- **Compact program (memory and decode).** 21 % of an executable's run is decoding `Expr` trees and a node
  is 32 B plus a heap `Vec`/`Box` per call. If shapes repeat, store one tree per R1C1 shape plus
  `(shape, anchor)` per cell, and evaluate with offsets: this is how XLSX itself stores shared formulas.
  It attacks goals (3) and (1) together and does not touch semantics (same `eval`, relative operands).
  It stands or falls with the shape census.

## Option 3: JIT for services that evaluate the same template

Where it matters: `--serve` and the Python/WebAssembly use, not the cold run. In a cold `calculate`
the evaluation is 16 % of the time, so even infinitely fast code saves at most 16 %. In `--serve`
evaluation is 100 % of each request: 1.14 s per call on `formulas.xlsx`, because the input moves a column
of 300 000 rows and a running total ([performance.md](performance.md#the-model-use-inputs-calculate-ranges)).

- **What compilation can win.** Formula shapes repeat, so the unit of compilation is the shape (tens),
  not the cell (millions): compile latency stops being the problem of other JITs. A numeric shape such
  as `=Data!A2*Data!B2` or `=ROUND(B2*C2,2)` becomes a loop over the rows: nanoseconds per cell against
  hundreds in the tree walker. Shapes that call `VLOOKUP`, text or `COUNTIF` stay bound by the runtime
  functions.
- **What it costs.** Cell types are dynamic (number, text, error, empty, boolean), so compiled code needs
  type guards and a deoptimization path to the interpreter, and must reproduce `blank`, error propagation
  and `logical_fmt` in the fast path or bail out. That is the double engine of finding 3. Containment:
  compile only shapes whose operands are all numeric and whose functions are on an allowed list,
  guard on the input types, fall back per shape.
- **Where it cannot run.** WebAssembly (no JIT in the module), and any build that must stay a single
  file without a toolchain. A JIT is a feature of the native binary and the Python wheel only.
- **Backend.** Cranelift (pure Rust, no system library, compile in milliseconds) over LLVM for a JIT:
  LLVM needs a system installation of a matching version (a CI burden for the linux x86_64 and aarch64
  wheels) and its latency is wasted on shapes of two operations. Size impact of either on `sheetc` and
  the wheel: to measure.
- **Expected value.** Unknown until measured. The per-request profile of `--serve` has not been taken;
  the constants stay decoded there, but the address resolution of finding 2 stays too. Resolving the slot
  (Option 2) may already take a large part of what a JIT would show.

## Option 4: LLVM compiler (ahead of time)

`sheetc build --backend llvm`: machine code for the static part of the program.

- It conflicts with four properties the project states: no toolchain to build and one file as artifact
  ([DESIGN.md](../sheetrs-compiler/docs/DESIGN.md#executable-format)), the WebAssembly target (a stub plus a payload; a
  native object cannot be one), the wheels that build in CI, and the single engine.
- The runtime functions (hundreds, with LibreOffice quirks, ICU4X collation, text formats) remain Rust
  and are called through an ABI; LLVM can inline only the numeric glue. Peak memory does not drop: values
  stay in the same stores.
- For `--serve` it can do what Option 3 does, at a build step the user must run with a toolchain.
- **Verdict: no-go for the main line.** Revisit only if the prototype of Option 3 passes its criterion and
  Cranelift's code is the measured limit (column vectorization), and then as an optional,
  feature-gated artifact.

## Comparison

| | Goal 1: `calculate`/`edit` time | Goal 2: services | Goal 3: run time and memory | One engine | Cost | Verdict |
|---|---|---|---|---|---|---|
| Direct mode | 16-22 % of `calculate`; ≤ 7-17 % of `edit` | none | none | yes | low | do, after the index |
| Dense index / compact program | start-up and run | per request | memory if shapes repeat | yes | low-medium | do first, measure |
| JIT (Cranelift) | none | large on numeric shapes, unmeasured | request time only | no (guarded) | high | prototype with a stop criterion |
| LLVM AOT | none | as the JIT | request time only | no | very high | no-go |

The three are not "complementary and orthogonal" as proposed: the goals overlap on the same lever (how a
cell address becomes a value), and the cheap lever serves all of them before any code generation.

## Proposed sequence, with exit criteria

1. **Measure (days).** The index experiment (dense grid) on `formulas.xlsx`: go if the run drops 15 % or
   more with identical output. The shape census over real models of the maintainer: go on compact
   programs if the formulas-per-shape ratio is 10 or more. A per-request profile of `--serve`.
2. **Direct mode.** Fallback to the compiled path on a cycle or on depth; equivalence tests: every
   fixture of `tests/semantics` and the oracle suites give identical values by both paths (`RAND`
   excluded). Report the stage table above again, and update [performance.md](performance.md).
3. **Compact program**, if step 1 allows it. Goal: memory per formula and decode time.
4. **JIT prototype** (feature `jit`, native only, numeric shapes only), only if step 1 shows that address
   resolution is not already most of the gain. Stop criterion: less than 3× on the per-request time of
   `formulas.xlsx --serve`, or any difference from the interpreter on the oracle suites or on a
   differential fuzz of the shapes.
5. **LLVM:** not started.

## What this does not show

- One generated workbook of 800 k formulas; real models differ in function mix, shape repetition and
  size. The shape census on the private workbooks is the main gap.
- A flamegraph of 126 samples: the order of magnitude of the shares is solid, the exact percentages are not.
- The gains of Option 1 are estimates from the stage trace, not from a prototype; the ones of Options 2-4
  are bounds or hypotheses, marked as such.
- The cold `calculate` and the run of an executable are different programs (the second decodes a payload);
  their profile shares are not additive.

Reproduce: `SHEETC_TRACE=1 sheetc calculate -i FILE`; the profile with `cargo build --release -p
sheetrs-compiler --features profile`, `sheetc build --runtime` that binary, and `SHEETC_FLAMEGRAPH=out.svg`
on the executable.
