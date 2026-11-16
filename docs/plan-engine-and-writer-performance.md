# Plan: dense index, `round_dec` and the writer of the workbook

Work plan taken from the spike in [spike-execution-backends.md](spike-execution-backends.md). It is written to be
read by a session that has not seen the spike and that already carries changes of the behavior of the compiler:
what to integrate, what to check after integrating, and the three workstreams in order. Nothing in it
changes what the engine or the writer answer; every item is accepted by *identical output first, speed second*.

Not in this plan, by decision of the spike: a direct interpreter, a JIT and LLVM. They wait for the conditions of
the spike (the exact `ROUND` without text, an engine kept between requests, the census of shapes of real
workbooks); see [the matrix](spike-execution-backends.md#the-matrix-four-strategies-by-four-scenarios).

## Where things are

| Workstream | State | What it gains (measured; 2.4 M formulas unless said) |
|---|---|---|
| **1. Dense index** | implemented and verified (D1-D4 done; the benchmark of D2 run and performance.md updated) | request of a compiled workbook -32 %; `--serve` request -60 %; `calculate` -11 %; total recalculation with writing -19 % and peak -12 %; ODS the same as XLSX; a regression in the first version (a stray cell in the last row of 100 sheets took 5.2 GB) found by D4 and fixed |
| **2. `round_dec` and the rounding functions** | implemented (R1-R7 done; `u64` divisions after the feedback); `ROUND`, `ROUNDUP`, `ROUNDDOWN` aligned with LibreOffice (0 differences) | 295 → 15 ns per call (20×), 301 → 31 ns in WebAssembly; memory unchanged; 10⁸ cases without a difference; request of a compiled workbook -9 %, `--serve` request -20 % |
| **3. Writer of the workbook** | W0, W1, W2 done; W3-W5 evaluated and not done | the copy of the writes (-7 % time, -231 MB) and compression in a second thread (-16 % XLSX, -11 % ODS) |

Baselines, all on the workbook of `sheetrs-compiler/scripts/gen_large_fixture.py … 300000` (2.4 M formulas, the
`formulas.xlsx` of [performance.md](performance.md#fixtures)), one 4-core virtual machine, executables of two
builds run alternately (median of 5-7), before → now:

| Flow | Before | Now |
|---|--:|--:|
| Executable, one request (declared outputs) | 2.50 s, 496 MB | 1.69 s, 489 MB |
| `--serve`, each request after the first | 1.51-1.63 s | 0.65 s |
| `sheetc calculate -i` (declared, JSON) | 9.65 s, 1153 MB | 8.57 s, 1153 MB |
| Executable `-o` on the same workbook without the settings sheet (every formula recalculated and written) | 13.49 s, 1537 MB | 10.99 s, 1352 MB |

## 0. Integrating with the session that changes the compiler

The two commits to take, in order (the third is documentation; the history of this branch was rewritten into them):

| Commit | What | Files |
|---|---|---|
| `f93566d` | the cell index in flat arrays (the rows of a sparse sheet are listed), constants found by row inside a decoded block, the index kept in the program (`Program::layout`), the 15-digit rounding in integer arithmetic (`round_exact.rs`, with `round_dec_text` kept as the fallback and the reference), `ROUND`, `ROUNDUP` and `ROUNDDOWN` as LibreOffice gives them (`round_calc.rs`) | `sheetrs-engine/**`, `compile.rs`, `FIDELITY.md`, the changelog fragments of the three |
| `8c79918` | the part being rewritten is compressed in another thread (`write/pipe.rs`), the writer takes the map of writes instead of a copy of it, `writer_bench` | `io/write/{pipe,xlsx_writer,ods_writer,mod,cells}.rs`, `answer/template.rs`, `io/examples/writer_bench.rs`, the changelog fragment |

The interfaces they change:

- `Program` has a new field `layout: Layout`, **not serialized** (`#[serde(skip)]`), ignored by `==`, and empty after a
  `clone()`. It is built on the first engine that runs the program and shared with the next ones. Every literal
  `Program { … }` needs `layout: Default::default()`: today there are two (`compile.rs`, the in-memory program of the fold and the
  program of the build). **If the other session adds or builds a `Program` anywhere else, or adds a field, expect a
  compile error or a conflict exactly there.**
- `Engine` takes the index from `prog.layout`; a program whose `cells` change after an engine ran on it is
  detected by the number of cells (the index is rebuilt, uncached). Nothing in the repository does that; a compiler change
  that edits `program.cells` *between* two engine runs on the same `Program` value must build a new one
  (`Layout` starts empty on `clone()`).
- `StoredConsts` keeps the same public interface; the blocks are decoded into `Decoded` (rows located by an offset
  table) and found by sheet and block number.
- `fill_template` / `fill_template_with_answer` / `_to` pass the writes by value to the writer;
  `CellWrites::into_sheets(self)` is new. `CellWrites::sheets(&self)` stays.
- **`modify_workbook_xlsx` and `modify_workbook_ods` (public, generic) now require `W: Write + Seek + Send`**: the archive is written from a second thread. Every caller in the workspace passes a `File` or a `Cursor`; a caller outside it with a writer that is not `Send` would stop compiling. `round_dec` keeps its signature; `round_dec_text` is new and `pub(crate)`.

Conflicts to expect: `engine.rs` around `Engine::build` and the field `index`; `ir.rs` (`Program`);
`compile.rs` where `Program` is built; `template.rs` in `write_cells*`. All resolve mechanically.

**After integrating, before anything else (the compiler of the other session changes the programs, so the
claim "same answers" has to be checked on *its* programs):**

1. `cargo test --release --workspace --exclude sheetrs-py`, clippy and fmt (CONTRIBUTING). The oracle tests against
   LibreOffice are the fidelity criterion; none may change.
2. Same program, two engines, same bytes. Build an executable of the same workbook with the stub of the commit before
   and the stub of the merge, and compare with `scripts/ab_compare.py`:

   ```bash
   cargo build --release -p sheetrs-compiler            # at each of the two commits, into its own --target-dir
   OLD/release/sheetc build -i W.xlsx -o exe_old --runtime OLD/release/sheetc
   NEW/release/sheetc build -i W.xlsx -o exe_new --runtime NEW/release/sheetc
   scripts/ab_compare.py -n 7 --check "old=./exe_old" "old copy=./exe_old" "new=./exe_new"
   ```

   Do it for a declared workbook, for one with nothing declared (`--embed-template`, with `-o OUT`, comparing the
   zip entries of the two workbooks, not the file: dates of the entries), for `--serve` with several requests
   (feed the same lines), for an ODS and for an XLSX. The "old copy" line is the noise floor.
3. A workbook that reads far cells in sparse rows (A and XFD in one row): the sparse path of `SlotIndex`
   is covered by two unit tests and was **not** measured at scale.

## 1. Dense index

Done on the branch. What is left is verification and numbers, not design.

| # | Task | Done when |
|---|---|---|
| D1 | The ODS run: every figure above is from an XLSX fixture. The engine is the same for both formats, but parity is a requirement of the project | the table of baselines repeated with `formulas.ods`, same protocol; no difference in answers |
| D2 | `scripts/bench.py run DIR` (the groups `compiler`, `edit` and the executables), and the tables of [performance.md](performance.md) updated with the new figures (CLAUDE.md: a benchmarked change updates the results table) | the rows of `sheetc calculate`, `sheetc build` (execute a request), the model-use tables and `sheetrs edit` are those of a run after the merge |
| D3 | Memory of `Decoded` (an offset table per block) on `data.xlsx` (600 000 rows of numbers, the case with the most blocks): only the 2.4 M fixture was measured, and its peak did not move (496 → 489 MB) | peak on `data.xlsx` and `data_declared.xlsx` not above the baseline, or the cost written down |
| D4 | The worst case of the index: a program with one row of cells at A and XFD repeated many times, and a program with a cell at row 1 048 575 | time and memory of `Engine::new` stated, with a ceiling that does not grow with the extent of the sheet |
| D5 | Candidate, **only if a profile of `--serve` after D2 still shows `Engine::build` or the allocation of `values` and `state` per request**: keep the engine between requests | a profile that shows it; otherwise dropped |

The profile of the executable's run (126 samples) before the change was `cell_value` 59 % of `calculate`,
the hash lookup 38 %, the constant blocks 20 %; of a `--serve` request, `cell_value` 36 % after the index
and `StoredConsts::get` 28 %, which the second change reduced.

## 2. `round_dec`

`functions/math.rs::round_dec` rounds on the 15 significant decimal digits of the value
(`ROUND(1.005, 2) = 1.01`) by formatting the number as text with `{:.14e}`, rounding the digit string and
parsing it back. It is the rounding of `ROUND`, `ROUNDUP` and `ROUNDDOWN`, and through `round_decimal`
(`functions/mod.rs`) of the number formats (`format.rs`: `TEXT` and the formats of cells), `BAHTTEXT`, `DOLLAR`
(`text.rs`), `office.rs` and `math_ext.rs`. Cost: 273-299 ns per call; in a numeric kernel with one `ROUND` per row it
is 97 % of what is left once the interpreter is out of the way, with any backend.

### Design (exact, with a fast path and the old function as fallback)

The result must be the same `f64`, bit for bit, as today. The old function is the specification and stays,
as `round_dec_text`, for the cases the fast path declines and as the reference of the tests.

1. **The 15 digits as an integer.** For `ax = |x|` normal, `ax = mant · 2^e2` (53-bit `mant`). With `k = 14 - e10`,
   `m = round_half_even(mant · 2^e2 · 10^k)` computed in `u128` (for `0 ≤ k ≤ 22` the numerator is below 2^127; shifts
   up to 126 bits; for `k < 0` a division by `10^-k` with the remainder compared to half). `m` is in
   `[10^14, 10^15)`: it is the digit string `{:.14e}` gives, exactly, with ties to even as the formatter does.
   Decline (return `None`) outside those ranges and for subnormals.
2. **The decimal exponent.** Estimate it with `log10` (the prototype) or from the exponent bits, then correct it by looking at
   `m`: below `10^14` → one less; at or above `10^15` → one more; **exactly `10^14` is ambiguous** (a carry that made
   `1.00000000000000`, or an estimate one too high for a value like `9.99999999999999e11`): compute `m` at the lower
   exponent, and take it if it is below `10^15`, else it was a carry; if that scaling is not representable,
   decline. The prototype got both cases wrong before this rule, and the differential test found them.
3. **Round at the place asked for**, on the integer: `keep = e10 + 1 + digits`; `keep ≥ 15` → `x`; `keep < 0` as today (`Away`
   gives `10^-digits` with the sign, otherwise `0.0`); else `D = m / 10^(15-keep)` and `up` by mode: `Nearest` the digit at
   position `keep` is 5 or more, `Away` any non-zero digit below, `Down` never. A carry gives `D = 10^keep` at the same scale
   (the old code inserts a digit and moves the exponent: the same value). `D = 0` gives `+0.0`, as today (not `copysign`).
4. **Back to `f64` exactly.** `result = D · 10^scale` with `scale = e10 - keep + 1`; if `D < 2^53` and `|scale| ≤ 22`
   one multiplication or division by an exact power of ten is correctly rounded, which is what parsing the digit text
   gives (the fast path of Clinger). Otherwise decline. Apply `copysign(x)`.

Prototype (`/tmp`, not in the repository; to be rewritten in tree): on money-like values (`ROUND(x, 2)`) 273-299 → 51-54 ns, with the
old function behind the `None`. In 120 000 588 cases: **0 differences**, with random bit patterns (21 % declined, nearly
all out of range), money values, ties (`n + 0.5`, `x.xx5`), values just below a power of ten, 40 decades of magnitude,
digits from -6 to 10, the three modes, and a list of known cases (`1.005`, `2.675`, `99.995`, `0.1 + 0.2`, `5e-324`,
`f64::MAX`).

### Tasks

| # | Task | Done when |
|---|---|---|
| R1 | Move the old body to `round_dec_text` (same signature, `pub(crate)`), write `round_dec` as the fast path with that fallback | the tests of `math.rs` (`round_dec(1.005, 2, …)` and the rest) pass untouched |
| R2 | A differential test in the tree: a fixed-seed generator (xorshift, so a failure reproduces) with the families above, a few million cases in the normal run (under a second in release) and an `#[ignore]` run of 10^8; every case compares `to_bits` of both functions, and the failure prints `x`, `digits`, mode and both results | it fails when the fast path is wrong (check by breaking the carry rule of point 2: it must catch it) |
| R3 | The rate of declines on real inputs: count them on the cells of the oracle fixtures (`semantics*.xlsx`) and on the benchmark workbook; the fast path is worth it if it declines rarely | rate stated; if a family of inputs declines often (large magnitudes, `digits` far from zero) widen the ranges or leave it |
| R4 | Fewer cycles in the estimate of the exponent (the prototype's 53 ns still includes `log10`) | measured against the prototype, kept only if it is exact under R2 |
| R5 | The reach: `TEXT`, number formats, `DOLLAR`, `BAHTTEXT`, `ROUND*` all go through the one function. Run the oracle tests (LibreOffice is the reference) and the three locale fixtures | no cell of any oracle changes |
| R6 | WebAssembly: `u128` division is slower in wasm32; time the module (`sheetrs-wasm/tests/module_run.mjs`) and the package | no slower than today in `ROUND`-heavy workbooks, or the fast path is compiled out of wasm with the reason written down |
| R7 | Measure: the kernel of the spike (`docs/spike-llvm/`), `--serve` on `formulas.xlsx`, and the executable request, alternating, with `scripts/ab_compare.py --check` | the figures go to [performance.md](performance.md) and the spike document; expected about 10 % of a `--serve` request here (estimate: 300 000 calls per request, 0.2 µs saved each) |

Fidelity statement: no difference with LibreOffice is introduced or removed by this change, so nothing goes to
[approved-differences.md](approved-differences.md). The differences that were already there are in *Rounding against LibreOffice* below. If R2 or R5 finds one, it is a bug of this change, not an approved
difference.

## 3. Writer of the workbook

The total recalculation (`sheetc calculate -o`, the recalculation that ends a `sheetrs edit`) spends its time here.
Profile of the executable of 2.4 M formulas, `-o OUT` (`SHEETC_FLAMEGRAPH`, inclusive shares, which overlap),
before removing the copy of the writes:

| Part | Share |
|---|--:|
| engine (`Engine::calculate` 14 %, `Engine::build` 3 %) | 17 % |
| the writer of the XLSX (`xlsx_writer::transform_sheet` 59 %) | 60 % |
| · `write_cached_cell` (the cached results) | 31 % |
| · deflate of the output (`zlib_rs`, medium) | 20 % |
| · reading the template back (`quick_xml` reader, `read_element_body_into`) | 11-12 % |
| · `quick_xml::Writer` | 23 % (overlaps `write_cached_cell`) |
| a copy of the whole map of writes (`BTreeMap` clone) | 13 % (removed) |

### Tasks, in order, each with a gate

| # | Task | Gate |
|---|---|---|
| W0 | **Profile the ODS writer and the `sheetrs edit` path with the same flow** (`formulas.ods`, a deletion that ends with the recalculation). Everything below is measured on XLSX; the capabilities of both formats are the same (parity), so each lever is needed or ruled out for both | the table above for ODS, and the list of callers of `CellWrites`/`modify` (`sheetrs-edit`, `sheetrs-py`, `sheetrs-wasm`) that each task reaches |
| W1 | A harness that times only the writer: `Book::modify_to_bytes` with a `CellWrites` of 2.4 M results (no engine), so a change to the writer is judged by the writer alone | it reproduces the 60 % within 10 % |
| W2 | Parallel deflate: the entries of the zip are independent; compress the sheet parts on threads (and, for the big sheet, in blocks that the format allows), and write them in order | the zip has the same entries, order and **decompressed bytes**; the bytes of the file may differ from a single thread's but **not between runs or between thread counts** (the fixtures of the project are deterministic: CONTRIBUTING). Keep if it takes 10 % or more of the total |
| W3 | Format the cached results without going through `quick_xml::Writer` events: numbers with a shortest-round-trip formatter into a reused buffer, escaping only where text needs it | every cell written byte-identical to today's (`cmp` of the decompressed sheet parts over the fixtures and the benchmark workbook) |
| W4 | Copy the untouched bytes between edited cells verbatim instead of parsing and re-emitting them (the reader costs 11-12 %); in a total recalculation every formula cell is edited, so this helps the partial writes (`edit`, `--formula-to-value` of a range) more than `calculate -o` | decide after W1 and W0: it is for the flows where most of the sheet is untouched |
| W5 | Split a sheet by rows (`<row` boundaries) and transform the chunks on threads, concatenating in order | only if W2-W4 leave the writer above 40 % of the flow; the output must be identical to the sequential one |
| W6 | The answer path that builds a JSON matrix of 2.4 M values before writing them: a variant that read each value straight from the engine was measured and saved 4 % of the peak and no time | **dropped**; recorded here so that it is not tried again without a new reason |

Acceptance for every writer task: the written workbook is the same in the content of every zip entry (compare
decompressed), LibreOffice opens it (the existing round-trip tests), the output does not depend on the number
of threads, and the peak memory does not rise. A `scripts/ab_compare.py` run of the flow `-o OUT` before and after,
and the rows of [performance.md](performance.md) for `sheetrs edit` and `sheetc calculate -o` updated.

## Results of the first pass

Measured on the branch after the three workstreams were worked, with executables of two builds run alternately
(`scripts/ab_compare.py`), on the 2.4 M-formula workbook and its ODS conversion (made by LibreOffice 26.8 headless).
Times move between sessions (the machine was faster on the second day), so each pair was measured together.

### Dense index

| Task | Result |
|---|---|
| D1 ODS | ODS source, original → this branch (before the flat index): executable, one request 2.27 → 1.26 s; `--serve` request 1.41 → 0.48 s; `sheetc calculate -i` 12.76 → 11.40 s. Answers identical. The figures include `round_dec` |
| D3 `data.xlsx` | 600 000 rows, nothing declared: executable 2.98 → 1.59 s, peak **809 → 705 MB** (no rise from `Decoded`); `data_declared.xlsx` unchanged (3-4 ms, 10 MB); `sheetc calculate -i` 9.77 → 8.31 s |
| D4 worst case | **A regression.** The first index reserved a table of rows per sheet up to the last row with a cell: one cell at row 1 048 575 took 52 MB, and one such cell in each of 100 sheets took **5.2 GB and 3.8 s** (the hash map it replaced was O(cells)). Rewritten in flat arrays with the rows *listed* for a sparse sheet (`SlotIndex`, `Spans::Listed`), three unit tests (a far row costs bytes, a stray cell keeps the table small, the index answers as a map does over random programs with every layout) |
| D2 | `scripts/bench.py run` of `calculate`, `compiler`, `model`, `python`, `edit` and `libreoffice` (the three runs took 2.5, 37 and 14 minutes; `model` and `python` need the wheel: `maturin develop --profile dist` in a venv). The tables of [performance.md](performance.md) for those flows are replaced. `lint`, `stats`, `query`, WebAssembly and the sizes were not run (the changes do not reach them); the native-executable rows of the WebAssembly table are marked as older. The old Python table said `Book(path)` read the workbook (1.97 s); it only opens the file now (0.01-0.02 s), so that row was already stale |
| D5 | Not done: no profile asks for it |

`Engine::new` before → after the flat index: 2.4 M cells packed 147 → 128 ms and 68 → 53 MB; 2 M cells in wide rows (A and XFD)
511 → 132 ms and 167 → 69 MB; one cell in the last row 28 ms and 52 MB → 0; 100 sheets with one far cell 3 792 ms and 5 200 MB →
0.1 ms and 0 MB; a packed sheet with one stray cell in the last row 101 → 70 MB. The lookup is not slower: `--serve` request
0.46 against 0.46 s (the executable 1.385 against 1.388 s, noise floor 1.344-1.388 s) and the peak 489 → 475 MB.

### `round_dec`

| Task | Result |
|---|---|
| R1, R2 | `round_exact.rs` with `round_dec_text` as the fallback and the reference. Tests: 3 M cases in the normal run (about 1 s), 10⁸ with `--ignored` (33 s): 0 differences. Both mutations tried fail the tests: removing the carry rule of the exponent, and rounding ties up instead of to even |
| R3 | Rate of declines on the cells of the oracle suites (the compiler, engine, answer and edit tests): **26 of 4 761 calls (0.5 %)**, all out of the integer range (`1e300`, `1e-300`, `1e-10` with 1 digit, `9.99999999999999e14`); none with money values |
| R4 | The decimal exponent from the binary one in integer arithmetic (`((biased - 1023) * 1233 + 600) >> 12`) instead of `log10`: 55.7 → 26.7 ns per call, still 0 differences in 3·10⁷ cases |
| R5 | The whole workspace test suite (1085 tests, the LibreOffice oracles and the three locale fixtures): no cell changes |
| R6 | WebAssembly (`wasm32-unknown-unknown`, Node 22, a module with no imports): 321 → 72 ns per call with the first version, same checksum; 301 → **31 ns** after the divisions of the 15 digits moved from `u128` to `u64` (the `u128` of a division is a call in wasm; the multiplication is not the problem) |
| R7 | Executable, one request: 1.461 → 1.326 s (the pair of copies of the old binary: 1.461 and 1.474). `--serve`, each request: 0.50-0.57 → 0.41-0.45 s. Answers identical. (The estimate was about 10 % of a `--serve` request; it was 20 %.) **Memory: no change**: peak 488 MB before and after in the request, and 488-489 MB against 488 MB after 10 requests of `--serve`; `sheetc` is 46 KB smaller (13 148 056 → 13 102 120 bytes); the tables are about 1 KB of static data, and the code allocates nothing |

### Rounding against LibreOffice (feedback of the maintainer, 2026-11-16)

The feedback asked for the memory of `round_dec` (given in R7), what `u64` would change, and whether a tolerance in the
last bits of the `f64` is acceptable, as it is for the functions of the C library
([FIDELITY.md](../sheetrs-compiler/docs/FIDELITY.md#the-c-library)). The data below answers with LibreOffice as the
reference ([harness](spike-round/README.md)).

**`u64` against `u128`.** An exact digit extraction in `u64` alone does not exist: `mant · 10^k` (53 bits plus 3.3 bits per
decimal place) fits 64 bits only for `k ≤ 3`, that is for numbers of 10¹¹ or more: **1.6 % of the `ROUND` cases** of the
LibreOffice set (4 826 of 300 000). The `u128` is used for one multiplication and one shift (on x86-64 and arm64 a `u128`
product is one instruction); what made the first version slow, mostly in WebAssembly, was the *division* of a `u128`.
The 15 digits are below 10¹⁵, so the steps on them were moved to `u64`:

| Version | Native (`ROUND(x, 2)` on money values) | WebAssembly (no imports) |
|---|--:|--:|
| text definition (before) | 257 ns | 301 ns |
| exact, `u128` divisions (first commit) | 24.8 ns | 72 ns |
| **exact, `u64` divisions (now)** | **14.5 ns** | **31 ns** |
| digits in `f64` only, no `u128` (not identical to the text) | 18.1 ns | not measured |
| the algorithm of LibreOffice as recalled (`floor(v + 0.5 + corr)`) | 11.7 ns | not measured |
| plain `f64` (`(x · 10^d).round() / 10^d`) | 2.3 ns | not measured |

**Tolerance.** 900 000 cases: 300 000 values (money, ties `n/10^k + 5/10^(k+1)`, computed products, sums with
noise, random reals, magnitudes from 10⁻⁸ to 10¹², each with both signs) with `ROUND`, `ROUNDUP` and `ROUNDDOWN` at -2 to 8
digits, calculated by LibreOffice 26.8.2 from the ODS it saves. (**Corrected later:** the inputs were read back from that
ODS, which carries 15 significant digits, so 154 010 of the 300 000 differed from the number written by 1 to 5 ulps; that is
the ODS export and not how LibreOffice reads: see *Adjusting `ROUND`, `ROUNDUP` and `ROUNDDOWN`*.)

| Comparison | Differences |
|---|--:|
| exact integer version against the text definition | **0** in 900 000, and 0 in 10⁸ of the unit test |
| digits in `f64` only against the text definition | 0 in the 900 000; **23 ppm (552 of 23.7 M) in the adversarial families**, of one unit of the rounded place or more, never of one ulp |
| plain `f64` against LibreOffice | 16 437 |
| the recalled algorithm of LibreOffice against LibreOffice (`ROUND`) | 413 of 300 000: worse than the engine on the large magnitudes, so it is **not** adopted |
| **the engine (text definition) against LibreOffice, value** | **10 566 (1.2 %)**, all of them already there before this change |

The 10 566 are: `ROUND` 152 (all in magnitudes of 10¹² or 10⁻⁶ and beyond), `ROUNDUP` 8 883 (8 686 in those magnitudes, 197 in
random reals of 15 digits), `ROUNDDOWN` 1 531 (794 in those magnitudes; and money 517, computed products 210 and ties 10).
In the typical families (money, products, ties, sums) `ROUND` and `ROUNDUP` have **0 differences in 200 000**, and
`ROUNDDOWN` has none at 4 digits or fewer (0 of 49 737 at 2) and **255 of 16 557 at 6 and 482 of 16 727 at 8**: LibreOffice truncates
the binary value (`36243.59` at 8 digits gives `36243.58999999`), where the engine, which rounds on 15 digits, gives
`36243.59`. Besides, 18 775 results differ only in the sign of zero (`-0` in LibreOffice for a negative number that rounds to
nothing, `+0` in the engine).

What it means:

- **The tolerance is not needed for speed.** The exact version is the fastest of the identical ones and reproduces the
  text definition; the `f64`-only version is slower than it and not identical. The change stays as it is.
- **The tolerance of the last bits of the `f64` does not reach these differences.** The approved difference is a result that
  changes in its last bit (about 10⁻¹⁶ relative). Here a difference is a whole unit of the rounded place, decided at the
  15th digit: it is a different number, not a different ulp. None of the 10 566 is created or removed by this change (the exact
  version equals the text definition), and none is an approved difference: [CONTRIBUTING.md](../CONTRIBUTING.md#parity-and-fidelity) calls
  it a bug or a pending decision.
- **Adjusted, not a decision** (maintainer's criterion, 2026-11-16: a difference with LibreOffice is adjusted unless something
  prevents it): see *Adjusting `ROUND`, `ROUNDDOWN` and the sign of zero* below.

### Adjusting `ROUND`, `ROUNDUP` and `ROUNDDOWN` (resolved, 2026-11-17)

**Method.** The ODS that LibreOffice saves carries 15 significant digits (`1/3` is `0.333333333333333`), so reading inputs and results back
from it loses the last digits. The inputs were therefore built *inside* LibreOffice as `(a·2²⁶ + b)·2^e` with whole numbers, which is exact, and
the results are compared by the decimal quantity they stand for. (An earlier version of this section said that LibreOffice reads a literal of 16
or 17 digits with 15 of them. **That was wrong**: the loss was in the export. Checked with the bits, 215 of 215 numbers, 17-digit ones included,
have the correctly rounded mantissa in a cell value and in a formula literal.) 887 976 + 426 204 cases (six families, both signs, -14 to 14
digits), 6 669 results of extreme values (10⁻³⁰⁰ to 10³⁰⁸, digits -300 to 300), a probe of 1 020 formulas for the forms of the call and the sign of a zero.

**The rule.** Two steps decide the result, and the first one explains what the fits of the rounding alone could not (`ROUND` 0 differences and
`ROUNDDOWN` 286 after twelve models; `ROUNDUP` unexplained). First, `ROUNDUP` and `ROUNDDOWN` of a value with a fractional part, with fewer than 12
digits, bring the value to 12 significant digits (`ROUNDUP(8.06-8; 2)` is `0.06`, `ROUNDUP(573.9951249000158; 8)` is `573.9951249`); the count of digits
is cut to a whole number with a 15-digit tolerance and clamped to 16 bits. Second, the number is scaled by `10^digits`, the product is brought to 15
significant digits, the floor or the ceiling is taken, and it is scaled back; the digits are limited to the fractional bits the number has, and a
whole number is left alone. It is `round_calc.rs`.

| | Differences with LibreOffice before | After |
|---|--:|--:|
| `ROUND`, 127 031 cases (-14 to 14 digits) | 75 | **0** |
| `ROUNDUP`, the same | 4 882 | **0** |
| `ROUNDDOWN`, the same | 1 843 | **0** |
| the same three over the 293 627 cases of -2 to 8 digits | 7 540 | **0** |
| 6 669 results of extreme values | not measured | **0** |
| probe of 1 020 formulas, through the whole engine (forms of the call, digits that are not whole, 16-bit limits) | 122 | **0** |
| sign of a zero result, 202 cases | 44 wrong | **0** |

The tests: `tests/assets/round_oracle.tsv` (10 500 results of LibreOffice, three functions) and `round_forms.tsv` (the 1 020 forms) in
`round_calc.rs`; the mutation that changes the 12 digits to 13 breaks them, and the three that survive are equivalent (above 2^41 a
number has at most 11 fractional bits, so the threshold and the guard of the bits are the same test; the limit of the digits only acts where
the product is no longer rounded). The 15-digit rule (`round_dec`, now 15 ns) stays for the number formats.

**What remains.** Nothing of the rounding, and the reading of numbers needs no adjustment (a mistaken finding, corrected above).

### Writer of the workbook

| Task | Result |
|---|---|
| W0 | The profile of the flow, XLSX and ODS (inclusive shares of the executable with `-o OUT`, after W2 and the copy removed). **XLSX** (8.6 s): the writer 65 % (`transform_sheet` 63 %: reading and parsing the template 36 % of it, `write_cached_cell` 30 %, `inspect_cell` 9 %, `next_cell` 7 %; deflate 13 %), `fill` 22 %, the engine 12 %, decoding the program 11 %. **ODS** (11.0 s): the writer 75 % (`edited_rows`/`edited_cells` 27-29 %, `write_element` 24 %, deflate 21 %, number formatting 6 %), `fill` 17 %, the engine 9.5 %, decoding 7 % |
| W1 | `sheetrs-io/examples/writer_bench.rs`: `modify_to_bytes` with 2.4 M cached numbers, no engine. XLSX 4.25 s (1 772 ns per cell), ODS 6.28 s (2 615 ns per cell). **The gate was missed**: the writer is about 5.6 s inside the flow (65 %) and 4.25 s alone (49 %), 16 points apart, because inside the flow the heap is larger and the writes come from the engine. The harness follows the direction of a change; the figures of the flow decide |
| W2 | `write/pipe.rs`: the part being rewritten is compressed in another thread through a bounded channel of 256 KB chunks (the same chunking where there are no threads, so the stream is the same on every target; `wasm32` builds). Harness: XLSX 4.25 → 3.40 s, ODS 6.28 → 5.57 s. **Flow**: XLSX 8.46 → 7.09 s (**-16 %**), ODS 11.17 → 9.99 s (**-11 %**), same peak (1 327 and 1 319 MB). Every entry of the zip has the same name, order, method and decompressed bytes; the file is the same on every run; it differs by 6 bytes in size from the single-thread file. Two tests: a part of 3 MB reads back whole, and an output that refuses is an error and not a hang |
| W3 | Not done. After W2 the compression thread is not the bottleneck (about 15 % of the CPU, in its own thread): the thread that transforms is. In it, nothing is above the gate alone: formatting the number is 13 % of `write_cached_cell` and 5 % of the flow, and Rust's shortest formatting is the specification of the bytes (a faster formatter such as `ryu` writes `1e21` where `Display` writes the digits, so it is not byte-identical) |
| W4 | Not done, for the reason in the plan: in a total recalculation every formula cell is edited. It stays for the partial flows (`edit` of a few columns); it needs W0 for those flows, which was done only for the total recalculation |
| W5 | Not done. The part is read through a `BufRead` and the transform keeps the cursor of the row, the arrays and the shared-formula masters, so a split by rows needs either the whole part in memory (about 210 MB for 2.4 M formulas, with a peak that rises: the gate says it must not) or a reader that cuts at `</row>` and workers that return the chunks in order, with the state handed over. It is the lever with the largest size (the transform thread is the bottleneck: up to about -40 % of the flow with 4 cores, an estimate) and the largest cost; it is a project of its own |
| W6 | Dropped, as planned |

## Order and dependencies

```text
0 integrate + verify ──► 1 dense (D1-D4) ──┐
                                           ├──► D2 bench run + performance.md (once, at the end of the three)
0 integrate + verify ──► 2 round_dec (R1-R7)┤
0 integrate + verify ──► 3 writer (W0, W1 ──► W2, W3 ──► W4/W5 by gate)
```

- 1, 2 and 3 are independent of each other (different files); 0 must come first.
- **Do not run the full benchmark until the three are in**: it takes about an hour and a half, and the
  tables of `performance.md` should change once.
- Each workstream ends with a changelog fragment in `changelog.d/` (`dense-cell-index.md` exists; one for
  `round_dec` is *Changed* with no change of results; one for the writer is *Changed*).
- Stop criteria: a task whose gate fails is dropped and written in the spike document with its figure, as W6 is.

## What to measure and how

- Fixtures: `python3 scripts/bench.py fixtures DIR` (about 4 minutes), or only the one that is needed:
  `python3 sheetrs-compiler/scripts/gen_large_fixture.py FILE 300000` (2.4 M formulas); for a recalculation with
  nothing declared, the same script without the `settings` sheet (delete the settings lines and the sheet
  from a copy of the script).
- Two builds, alternately, with the noise floor and the answers checked: `scripts/ab_compare.py` (above).
- Stage times: `SHEETC_TRACE=1 sheetc build|calculate …`. Profile: `cargo build --release -p sheetrs-compiler
  --features profile`, `sheetc build --runtime` that binary, `SHEETC_FLAMEGRAPH=out.svg ./exe …` (126 samples
  for a short run: take at least a few hundred).
- Times move 10-20 % between sessions: compare only what was measured together.

## Open questions for the maintainer

1. **ODS and XLSX writers.** *Answered by the maintainer (2026-11-16): an optimization may depend on the format.* W2 was done for both writers; the levers that remain (W3-W5) may be built for one format only, with each format's speed in [performance.md](performance.md). The parity of CONTRIBUTING stays about capabilities, not about speed.
2. **`round_dec` at 10^8 cases in CI**: the normal run is a few million cases; is the long one run by hand before a
   release, or in the workflow `ci.yml` (which runs only when asked)?
3. The census of shapes of real workbooks is still missing and decides the compact program (memory); it is not part
   of this plan, but each of the workstreams above reports figures that it will need.
