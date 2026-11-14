# Spike: interpreter, JIT and LLVM for ODS/XLSX

Question: what do a direct interpreter (no build), a JIT and an LLVM compiler add to SheetRS, and are
they viable? Three goals, not simultaneous: (1) cut the time of `calculate` (and `edit`) by skipping the
build; (2) run a workbook directly and speed it up progressively in a service that evaluates the same
template again and again; (3) cut time and peak memory of the run of a compiled workbook.

Status: the data of the matrix below is collected. Four small changes found on the way are
implemented (they are not one of the three options and pay for themselves); the three options are not.
Limits of the data: [What this does not show](#what-this-does-not-show).

## The matrix

| | Goal 1: `calculate` / `edit` | Goal 2: services (`--serve`) | Goal 3: run of a compiled workbook | One engine | Cost | Verdict |
|---|---|---|---|---|---|---|
| **Direct interpreter** (skip compile stages) | 9-11 % of a full recalculation at 2.4 M formulas; 16-22 % of a declared one-shot calculation at 0.8 M | none | none (peak: the graph is already freed) | yes | low-medium; needs a fallback on depth and cycles | worth doing, modest; not first |
| **JIT** (Cranelift, native only) | none | ceiling about 3-4× per request on numeric shapes, and only with `ROUND` fixed first; unproven | request time only | no (guarded) | high | defer: does not meet its own stop criterion today |
| **LLVM, ahead of time** | none | as the JIT | request time only | no | very high | no-go |
| *Implemented, not an option:* dense index, constants by row, index kept in the program, no copy of the writes | `calculate -i` -11 %; `calculate -o` -19 % time, -12 % peak | request -60 % (1.5 → 0.65 s) | request -32 %, first request -35 % | yes | done | keep |

The three options are not independent: they share one lever, how a cell address becomes a value, and that
lever was served first with no code generation. After it, what is left per request is the exact `ROUND`
and the setup of an engine per request, not tree dispatch.

## Findings that change the question

1. **The engine is already an interpreter.** `Engine::eval` walks `Expr` trees over a demand-driven
   schedule ([DESIGN.md](../sheetrs-compiler/docs/DESIGN.md#hybrid-compile--run-time): "the JIT is not machine
   code: it is lazy resolution on the same interpreter"). `sheetc calculate` already builds in memory
   and writes no executable. "Omitting build" means skipping *compile stages* (cone, graph, SCC, taint,
   fold, assembly), not an executable.
2. **The cost of an evaluation was address resolution, not dispatch.** Before the changes below, 59 % of
   `Engine::calculate` was `cell_value`: the hash lookup `CellRef → slot` 38 %, the search in the blocks
   of constants 20 %. Arithmetic and the bodies of `IF`, `MAX` are marginal.
3. **In a full recalculation the engine is not the bottleneck: writing the workbook is.** In
   `calculate -o` of 2.4 M formulas (every formula an output), the profile of the run (inclusive shares,
   which overlap) gives: the engine (`calculate` 14 %, `Engine::build` 3 %) about 17 %, and the writer of
   the workbook (XML transform, quick-xml and deflate) about 60 %. A copy of the whole map of writes
   sat in that path; removing it took 7 % of the time (12.17 → 11.28 s in a pair measured together) and
   231 MB of the peak. A faster engine, an interpreter or a JIT, moves the smaller part.
4. **`ROUND` is the exact cost of the numeric kernel.** `round_dec` reproduces the 15-digit decimal
   rounding of spreadsheets by formatting the number as text and parsing it back
   (`functions/math.rs`). It is 12-14 % of a request on the benchmark workbook, and 97 % of the floor of
   a hand-written kernel (next section).
5. **"No double execution engine" is an explicit requirement** (DESIGN.md, *One IR, one engine*), and
   fidelity to LibreOffice is the acceptance criterion ([CONTRIBUTING.md](../CONTRIBUTING.md#parity-and-fidelity)).
   `eval` carries it in its control flow: `logical_fmt`, ghost arrays, element-wise lifting, propagated
   errors, `notes`, blank results. A second backend must reproduce all of it or bail out to the interpreter:
   a decision for the maintainer, not a refactor. Lazy functions (`IF`, `IFERROR`) take `Expr`, not values,
   so compiled code still needs the trees.

## Changes implemented (same engine, same results)

All of them format-agnostic (the engine is the one of XLSX and ODS). Measured with the executables run
alternately (median of 5-7 runs), the stubs built with the same profile, on the workbook of
`gen_large_fixture.py … 300000` (2.4 M formulas, the one of [performance.md](performance.md#fixtures)), against
the state before this branch, in one session. The values and the written workbooks are identical
(`cmp` of the executable's output; every entry of the zip of the workbook equal). 1081 tests pass and
clippy is clean.

| Change | Where |
|---|---|
| A table per sheet and row replaces `FastMap<CellRef, u32>`; a row whose cells are far apart (A and XFD) keeps sorted pairs, so it does not reserve a slot per column | `SlotIndex`, `engine.rs` |
| A decoded block of constants knows where each row starts, and is found by sheet and block without a hash | `Decoded`, `consts.rs` |
| The index is kept in the program (`Program::layout`, not serialized, ignored by `==`) for the engines that follow: a service builds one engine per request | `Layout`, `engine.rs`, `ir.rs` |
| The writer takes the map of writes instead of a copy of it | `into_sheets`, `template.rs` |

| Flow (2.4 M formulas) | Before | Now | Change |
|---|--:|--:|--:|
| Executable, one request (declared) | 2.50 s, 496 MB | 1.69 s, 489 MB | -32 % |
| `--serve`, first request | 2.29-2.34 s | 1.49-1.53 s | -35 % |
| `--serve`, each next request | 1.51-1.63 s | 0.65 s | -60 % |
| `sheetc calculate -i` (declared, JSON) | 9.65 s, 1153 MB | 8.57 s, 1153 MB | -11 % |
| Executable `-o`, every formula recalculated and written | 13.49 s, 1537 MB | 10.99 s, 1352 MB | -19 %, peak -12 % |

At 0.8 M formulas the dense index alone took the run of the executable from 0.735 s to 0.557 s (-24 %;
quartiles 0.714-0.761 and 0.553-0.573, with a second copy of the baseline at 0.737 s as the noise floor).
A variant that also streamed the values of the answer cell by cell, to avoid the JSON matrices, saved 4 % of
the peak and no time, and was dropped. `performance.md` is not updated: the full `scripts/bench.py` run
was not done, and its tables are of another session.

## Measurements behind the options

### Stages of `calculate`

`gen_large_fixture.py … 100000` (800 005 formulas, declared input and outputs), `SHEETC_TRACE=1`, before the
changes: `calculate -i` 2.49 s.

| Stage | Time | Share | Skipped by a direct mode? |
|---|--:|--:|---|
| read workbook into the `Model` | 0.69 s | 28 % | no |
| parse and lower formulas | 0.27 s | 11 % | no |
| bindings, constants that start a path, index | 0.06 s | 2 % | partly |
| cone + CSR graph | 0.17 s | 7 % | yes |
| SCC + taint | 0.04 s | 2 % | yes (needed for cycles only) |
| fold (evaluates input-independent cells) | 0.35 s | 14 % | no: the same evaluation happens in the run |
| assembly of the `Program` and the rest | 0.51 s | 20 % | in part (about 0.3 s) |
| run | ≈0.40 s | 16 % | no |

For `edit`, the recalculation is the mode that matters: every formula is an output and the inputs are
deduced, so nothing is pruned and nothing folds. The same fixture without the settings sheet, 2.4 M
formulas, `sheetc build --embed-template` (stage trace): read the workbook 2.62 s; compile 4.16 s, of which
parse and constants 1.9 s, index 0.2 s, cone + CSR 0.6 s, SCC + taint 0.1 s, assembly 0.9 s. The stages a
direct mode skips (index, graph, SCC, taint and most of the assembly) are 1.5-1.9 s, 9-11 % of a
`sheetc calculate -o` that takes about 17 s. The rest of that time is the workbook writer (finding 3).

### Risk of the direct mode: depth

A direct mode evaluates on demand, which is recursive for a chain the schedule did not order. The engine
already runs on a thread with a 512 MB stack (`cli::with_big_stack`). A chain of `A1 = A2 + 1 … = An + 1`
evaluated with no schedule: overflows at 10 000 levels with 8 MB, 100 000 with 64 MB and 1 000 000 with
512 MB; completes 300 000 with 512 MB. That is about 0.8-1.7 KB per level: a workbook of 1 048 576 rows
can overflow, so the direct mode needs a depth limit with fallback to the compiled path (which handles
both the forward and the reversed chain of 300 000 in 0.6-0.7 s). Cycles and iterative calculation
need the SCC the same way.

### Ceiling of a JIT

Profile of one `--serve` request on 2.4 M formulas (after the dense index, before the other changes; share of
self time): engine glue 66 %, constants 13.5 %, `ROUND` and number formatting 12 %, other functions
3 %, decode and start 6 %.

A hand-written typed kernel for the five columns that depend on the input (1.5 M live formulas: `IF`,
product, `ROUND(…,2)`, running total, `IFERROR` of a division, `MAX - MIN`), `rustc -O`, best of 7:

| Kernel | Time | Per formula |
|---|--:|--:|
| rounding as `(x*100).round()/100` | 3.7 ms | 2.5 ns |
| rounding with the project's `round_dec` (copied as is) | 113 ms | 76 ns |

A request takes 0.65 s today. A numeric JIT cannot go below the exact rounding: the ceiling on this workload
is about 0.15-0.2 s per request (3-4×), reached only if everything else in the request (engine setup,
reading the answer) also disappears, and for numeric shapes only; shapes with `VLOOKUP`, text or `COUNTIF`
stay bound by their functions. The ceiling is not a prototype. The stop criterion proposed before (3× per
request over the interpreter) is therefore met, if at all, at the ceiling: the next gains are in
`round_dec` and in not building an engine for every request, both cheaper than code generation.

### Cost of the backends

| | Size | Notes |
|---|--:|---|
| Cranelift (`cranelift-jit` 0.120, pure Rust) | +2.85 MB on a stripped binary with one JIT function | compiles and sets up a function with an `IF` in 0.4 ms; no system library |
| LLVM 18 (`libLLVM.so`) | 123 MB shared library | a matching system installation for every wheel (linux x86_64 and aarch64) and for the build of an executable; no wasm host JIT |

Neither works in the WebAssembly modules (no JIT in a module), and a build of an executable with LLVM would
need a toolchain where today it needs none ([DESIGN.md](../sheetrs-compiler/docs/DESIGN.md#executable-format)).

## Options

**Direct interpreter.** Same `Engine` with a different schedule: lowered cells in row-major order and
demand-driven for the rest, a fallback to the compiled path on a cycle or beyond the depth limit (so
fidelity is kept). `RAND` draws in evaluation order, so its values would differ between paths.
Dynamic references (`INDIRECT`, `OFFSET`) become simpler than today, where one makes the whole workbook
reachable. Gain: 9-11 % of a full recalculation, 16-22 % of a declared one-shot calculation; no gain in
a service. Memory: the graph is freed before the peak already. It is worth doing for what it is, after
the writer.

**JIT.** Formula shapes repeat in a copied-across model, so the unit of compilation is the shape (tens),
not the cell (millions), and compile latency is not the obstacle. The cost is fidelity: dynamic
types need guards and a path back to the interpreter, and blank, error propagation and `logical_fmt` must
be reproduced in the fast path or the shape must bail out. Only native binaries and the Python wheel could
carry it. Defer until `round_dec` and the engine per request are fixed and a profile of `--serve` still
shows the glue.

**LLVM, ahead of time.** Conflicts with no toolchain and a single file, with the WebAssembly target,
with the CI of the wheels, and with the single engine; the runtime functions (hundreds, with quirks of
LibreOffice, ICU4X, formats) stay Rust behind an ABI; peak memory does not drop. No-go.

**Compact program (not an option of the request, and not decided).** One tree per R1C1 shape plus
`(shape, anchor)` per cell would cut memory and decode time (about 480 B per formula at 2.4 M; decode
is 21 % of a cold run of the executable and 6 % of a `--serve` session). It stands or falls with how
much real models repeat, and that census is missing.

| Workbook | Formulas | Shapes |
|---|--:|--:|
| `gen_large_fixture.py` 100 k rows | 800 005 | 14 |
| `loan.xlsx` | 1 086 | 9 |
| `semantics.xlsx` (the oracle: distinct by design) | 34 295 | 33 998 |

Only the first two say anything about real models, and both are synthetic or tiny.

## Sequence

1. **Done** (this branch): the four changes above.
2. **Writer of the workbook**, the largest cost of `calculate -o` and `edit` (finding 3): the XML transform,
   the deflate and the construction of the values to write. Profile first; this is where goal 1 is.
3. **`round_dec` without text**, with a differential test against the current function over many
   values, modes and digits (the rounding is a fidelity rule, so identity is the criterion), and an
   engine kept between the requests of `--serve`. Measure again; this decides the JIT.
4. **Direct mode**, with the depth limit and the fallback, and the equivalence tests of every
   fixture of `tests/semantics` between the two paths (`RAND` excluded).
5. **Shape census** on real workbooks of the maintainer; compact program only if the ratio of formulas
   per shape is 10 or more.
6. **JIT prototype** only if step 3 leaves the glue as the main cost of `--serve`. **LLVM:** not started.

## What this does not show

- One generated workbook family (`gen_large_fixture.py`) in two sizes; real models differ in function mix,
  repetition and size. The shape census is the main gap, and the JIT ceiling depends on how numeric the
  real models are.
- Flamegraphs of 126 to 1700 samples: the order of magnitude of the shares is sound, the exact
  percentages are not. The stubs carry the `profile` feature, the same for both sides of every comparison.
- The ceiling of the JIT comes from a hand-written kernel, not from generated code; the gain of the direct
  mode, from the stage trace, not from a prototype.
- One 4-core virtual machine; times move 10-20 % between sessions, so only the pairs measured together
  are compared.
- The 100 k-row stage table predates the changes; the 2.4 M recalculation trace is after the first of them.

Reproduce: `SHEETC_TRACE=1 sheetc calculate -i FILE` (stages); `cargo build --release -p sheetrs-compiler
--features profile`, `sheetc build --runtime` that binary, and `SHEETC_FLAMEGRAPH=out.svg` on the
executable (profile); the executables of two builds run alternately for the comparison, as in
[DESIGN.md](../sheetrs-compiler/docs/DESIGN.md#measuring-a-change).
