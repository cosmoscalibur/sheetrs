# Spike: interpreter, JIT and LLVM for ODS/XLSX

Question: what do a direct interpreter (no build), a JIT and an LLVM compiler add to SheetRS, and are
they viable? Three goals, not simultaneous: (1) cut the time of `calculate` (and `edit`) by skipping the
build; (2) run a workbook directly and speed it up progressively in a service that evaluates the same
template again and again; (3) cut time and peak memory of the run of a compiled workbook.

Status: the data of the matrix below is collected. Four small changes found on the way are
implemented (they are not one of the three options and pay for themselves); the three options are not.
The LLVM spike (prototype in [spike-llvm/](spike-llvm/README.md)) is done: [what it measured](#llvm-spike).
Limits of the data: [What this does not show](#what-this-does-not-show).

## The matrix

| | Goal 1: `calculate` / `edit` | Goal 2: services (`--serve`) | Goal 3: run of a compiled workbook | One engine | Cost | Verdict |
|---|---|---|---|---|---|---|
| **Direct interpreter** (skip compile stages) | 9-11 % of a full recalculation at 2.4 M formulas; 16-22 % of a declared one-shot calculation at 0.8 M | none | none (peak: the graph is already freed) | yes | low-medium; needs a fallback on depth and cycles | worth doing, modest; not first |
| **JIT** (Cranelift, native only) | none | ceiling about 3-4× per request on numeric shapes, and only with `ROUND` fixed first; unproven | request time only | no (guarded) | high | defer: does not meet its own stop criterion today |
| **LLVM** (JIT or ahead of time) | none | as the JIT; 1.3× over Cranelift on a pure numeric loop, equal with the exact `ROUND` | request time only; the only way to put compiled code in a WebAssembly module (shown) | no (guarded) | very high: a system package, +51 MB, 15-20× the compile time of Cranelift | not now; kept open as a second emitter of the same kernel IR ([LLVM spike](#llvm-spike)) |
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

## Options

**Direct interpreter.** Same `Engine` with a different schedule: lowered cells in row-major order and
demand-driven for the rest, a fallback to the compiled path on a cycle or beyond the depth limit (so
fidelity is kept). `RAND` draws in evaluation order, so its values would differ between paths.
Dynamic references (`INDIRECT`, `OFFSET`) become simpler than today, where one makes the whole workbook
reachable. Gain: 9-11 % of a full recalculation, 16-22 % of a declared one-shot calculation; no gain in
a service. Memory: the graph is freed before the peak already. It is worth doing for what it is, after
the writer.

**JIT.** (Cranelift; LLVM is in its own section.) Formula shapes repeat in a copied-across model, so the unit of compilation is the shape (tens),
not the cell (millions), and compile latency is not the obstacle. The cost is fidelity: dynamic
types need guards and a path back to the interpreter, and blank, error propagation and `logical_fmt` must
be reproduced in the fast path or the shape must bail out. Only native binaries and the Python wheel could
carry it. Defer until `round_dec` and the engine per request are fixed and a profile of `--serve` still
shows the glue.

**LLVM.** See the [LLVM spike](#llvm-spike): the data, what is system-level, and the criteria to include it.

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

## LLVM spike

The question was whether LLVM, as the standard infrastructure for this kind of solution, justifies itself:
less to maintain by having one infrastructure for compiling, and a better result at run time. The prototype
(`docs/spike-llvm/`, see its README) builds the kernel of the five input-dependent columns of the benchmark
workbook with LLVM 23 through `inkwell`, and the same kernel with Cranelift and in Rust, and checks the
values bit by bit against the Rust reference.

### Is it a system dependency?

Yes, in the way you describe: cargo is not enough. `inkwell` (and `llvm-sys` under it) calls `llvm-config` and
links the libraries of a **system installation of the same major version**: on Ubuntu 24.04 the packages
`llvm-N-dev` and `libpolly-N-dev`, plus a C toolchain. A machine with `libLLVM.so`, `clang` and `lld` (this
one had LLVM 18) is not enough: `llvm-config` fails for lack of the static libraries until `-dev` is installed.
The ways to use LLVM, with what each needs:

| Way | Needs at build time of SheetRS | Needs where it runs | Size |
|---|---|---|---|
| Link LLVM (`inkwell`, static) | `llvm-N-dev` of one exact major, in every CI image and wheel builder (manylinux needs it in the container) | nothing | the probe binary: 59.7 MB stripped (x86 only); `sheetc` is 8.8 MB |
| Link LLVM (dynamic) | the same | `libLLVM.so.N`, 123 MB (LLVM 18) | small binary |
| Write IR and call `clang`/`llc`/`lld` | nothing | a toolchain on the machine of whoever compiles a workbook | none |
| Bitcode of the Rust runtime (below) | the LLVM of `rustc` | the same | none |

The third way is the only one with no build dependency, and it gives up "no toolchain to build a workbook",
which the project states as a property. The distribution has LLVM 18 and 20; `rustc` 1.99 bundles
**LLVM 23.1.1**, so the LLVM that matches the compiler of the project is not in the distribution:
`apt.llvm.org` (or building LLVM) is part of every environment. `llvm-sys` 231 and `inkwell` 0.11 (`llvm23-1`)
exist, and the prototype built and ran against 23.1.6 and 20.1.2 with the same results.

### What the measurements say

The kernel, 1.5 M live formulas (`IF`, product, `ROUND(…,2)`, running total, `IFERROR`, `MAX - MIN`):

| | approximate rounding (inline) | exact rounding (`round_dec`) |
|---|--:|--:|
| `rustc -O` | 4.7-5.3 ms | 116-119 ms |
| Cranelift, `speed` | 2.13 ms | 115 ms |
| LLVM, O2 / O3 | 1.60-1.65 ms | 113-115 ms |

All variants give the same values as the Rust reference (checked in the LLVM runs and in Cranelift's exact one;
Cranelift's approximate run uses another rounding instruction and is not compared). So LLVM is 1.3× faster than
Cranelift on a numeric loop with inline rounding and the same with the exact rounding, which is the one the
engine must use. The generated code does not remove the cost that matters.

Compilation of N distinct formulas of 25 nodes on average (one function each; LLVM 23, close to LLVM 20):

| | 1 000 formulas | 10 000 formulas | per formula |
|---|--:|--:|--:|
| LLVM O0 (fast instruction selection) | 65 ms | 0.65 s | 65 µs |
| LLVM O1 | 0.74 s | 8.0 s | 0.8 ms |
| LLVM O2 | 1.05 s | 10.8 s | 1.05 ms |
| Cranelift, no optimization | 55 ms | 0.58 s | 58 µs |
| Cranelift, `speed` | 69 ms | 0.81 s | 80 µs |

The kernel above compiles in 0.5 ms (Cranelift) and 10-14 ms (LLVM O2). Against an interpreter that costs
about 0.43 µs per formula today (0.65 s for 1.5 M), a function pays for itself after about 2 500
evaluations with LLVM O2 and about 190 with Cranelift `speed`. That is a shape that repeats down a column
(300 000 rows), and never a formula that is distinct: in an oracle workbook with 34 k distinct formulas,
LLVM O2 would spend about 36 s and Cranelift about 2.7 s before the first value. The census of shapes of
real models decides how much of a workbook is in the first case.

What is LLVM's alone:

- **Compiled code for the WebAssembly module.** Cranelift does not emit wasm. The prototype generates the
  kernel as a wasm object with LLVM, links it with the Rust runtime (including `round_dec`) with `wasm-ld`
  into a module of 74 KB with **no imports** (the restriction of the project's modules), and in Node 22 it
  gives the same checksum as the Rust version and 152.9 ms against 154.3 ms with exact rounding
  (3.8 against 4.7 ms with approximate rounding).
- **Inlining of the runtime into the generated code**, from the bitcode of `rustc`: it works with LLVM 23
  (a leaf of the runtime was linked, no call is left in the IR, same values) and gains about 5 %
  (1.72 → 1.63 ms). LLVM 20 cannot even read the file (`Unknown attribute kind (105)`, producer
  `LLVM23.1.1-rust-1.99.1`), so it ties the LLVM of the project to the one of its compiler, which changes
  with every Rust release; and only a leaf with no allocation can be a target (`round_dec` formats text).

### Would it reduce what is maintained?

What a compiled path replaces is dispatch, not meaning. The semantics stay in 524 registered functions
(19 lazy, 505 eager, about 19.5 k lines) and about 4 k lines of engine, coercion, value and format that the
generated code still calls: LLVM does not remove one of them. A compiled path adds a lowering (typed
shapes with guards), an ABI for `Value` (an enum with `Arc`, not stable across the boundary), a differential
test between the two paths, the system package, and a pin of the LLVM version. The size of the emitter
is not the issue: the same kernel is about 70 lines with `inkwell` and about 60 with Cranelift.

The reasoning that does hold is the other one: **having both strategies is worth its cost**, because the
interpreter is the reference of the fidelity (the compiled path is tested against it, and it runs the
first requests while a shape compiles, as in any tiered design). What is not worth it is making LLVM the
only path: it would put the 4 k lines of control flow of `eval` into an IR builder, with a compile latency of
1 ms per distinct formula before the first value.

### Criteria to include it, and the way that keeps the option open

LLVM enters when all of these hold, none of which is true today:

1. A typed kernel IR exists (numeric shapes, guards, fallback per shape) and runs with Cranelift or with the
   interpreter as its first backend, so the emitter is one module with one backend trait.
2. `round_dec` stops being the floor of a numeric request: today it is 97 % of the kernel, and any backend ties.
3. A measured need that only LLVM serves: compiled code in the WebAssembly module (the one shown), or a
   numeric workload where its 1.3× over Cranelift on a loop is worth 15-20× the compile time.
4. The maintainer accepts a pinned system LLVM in CI and in the builders of the wheels, under a feature that is off by default.

Until then the decision costs nothing: the kernel IR of criterion 1 is backend-neutral, and an LLVM emitter
of about 300 lines (the prototype) can be added behind a feature later.

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
6. **Typed kernel IR with a backend trait** (interpreter first, Cranelift second) only if step 3 leaves the glue as the main cost of `--serve`. **LLVM** as a third emitter when its criteria are met ([LLVM spike](#llvm-spike)).

## What this does not show

- One generated workbook family (`gen_large_fixture.py`) in two sizes; real models differ in function mix,
  repetition and size. The shape census is the main gap, and the JIT ceiling depends on how numeric the
  real models are.
- Flamegraphs of 126 to 1700 samples: the order of magnitude of the shares is sound, the exact
  percentages are not. The stubs carry the `profile` feature, the same for both sides of every comparison.
- The ceiling of the JIT comes from a hand-written kernel, not from generated code; the gain of the direct
  mode, from the stage trace, not from a prototype.
- LLVM was measured through `inkwell` on x86-64 with 20.1.2 and 23.1.6; the build of wheels for aarch64 and the size of a stripped static LLVM there were not measured.
- One 4-core virtual machine; times move 10-20 % between sessions, so only the pairs measured together
  are compared.
- The 100 k-row stage table predates the changes; the 2.4 M recalculation trace is after the first of them.

Reproduce: `SHEETC_TRACE=1 sheetc calculate -i FILE` (stages); `cargo build --release -p sheetrs-compiler
--features profile`, `sheetc build --runtime` that binary, and `SHEETC_FLAMEGRAPH=out.svg` on the
executable (profile); the executables of two builds run alternately for the comparison, as in
[DESIGN.md](../sheetrs-compiler/docs/DESIGN.md#measuring-a-change).
