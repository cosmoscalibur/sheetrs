# sheetrs-wasm

WebAssembly build of the SheetRS tools for the browser and Node (`--target web` ES module). The package
has **one object, `Book`**: a workbook made from the bytes of an XLSX, XLSM or ODS file, which lints,
reads, edits, calculates, inspects and compiles in memory: no file, no thread, no network. The answers
are those of the command-line tools (`sheetrs lint`, `sheetrs stats`, `sheetrs query`, `sheetrs edit`,
`sheetc`); the smoke test compares them with the commands on real workbooks. The API is experimental in
every layer, this one included: names and shapes may change.

```js
import init, { Book, initPanicHook } from "sheetrs-wasm";

await init();            // loads the .wasm next to the module
initPanicHook();

const book = new Book(new Uint8Array(await file.arrayBuffer()));
book.sheetNames;                                   // read once, when the Book is made
book.lint().toSarif(file.name);                    // SARIF 2.1.0 text
book.range("Data!A1:C9", { properties: ["value", "formula"] });
book.deleteRows("Data!3:10");                      // the operations accumulate...
book.deleteSheets(["Tmp"]);
const edited = book.save();                        // ...and `save` runs them: a Uint8Array
```

The format (XLSX, XLSM or ODS) is the one of the bytes: it is not an argument. Errors are thrown as
strings. Methods that take options take an object with the keys in `camelCase`; an unknown key is an
error that names it. The results are objects with explicit conversions (`toJSON()`, `toSarif()`, ...),
so `JSON.stringify(result)` works. A `Book` holds memory of the module until it is released: the
runtime releases it by itself when nothing references it any more (a `FinalizationRegistry`, which the
package build checks it has), and `book.free()` releases it at once, which is worth doing with a large
workbook.

## `Book`

`new Book(bytes)` detects the format from the content and reads nothing else. The first property that
describes the workbook (below) **reads it once**, in a single streaming pass that counts what it holds (a
cost proportional to the size of the workbook, paid then), and keeps only those numbers; a workbook that
cannot be read throws there. Everything else is read when it is asked for.

| | Member | Gives |
| - | ------ | ----- |
| **first read** | `format`, `dateSystem`, `hasMacros`, `sheetNames`, `definedNames` | `"xlsx"`, `"xlsm"` or `"ods"`; the date of serial 0 (`"1899-12-30"`, `"1900-01-01"` or `"1904-01-01"`; another is an error that names it); macros found, never run; the sheet names in order; `{name: refersTo}` |
| | `sheets` | `book.sheets["Data"]`: `name`, `usedRange` (`"A1:C9"` or `null`), `state` (`visible`, `hidden`, `very_hidden`), `usedCells`, `formulas`, `errors`, `emptyCells`, `comments`, `charts`, `tables`, `xmlSize` |
| | `stats` | `Stats`: `toJSON()` is `sheetrs stats --format json` (`total_file_size` is 0: bytes carry no size) |
| | `totalSheets`, `totalNamedRanges`, `totalFileSize`, `totalUsedCells`, `totalFormulas`, `totalErrors`, `totalEmptyCells`, `totalComments`, `totalCharts`, `totalTables`, `totalXmlSize` | the eleven totals of the statistics |
| **reads** | `lint(configToml?)` | `LintReport`: `count`, `records` (`{rule_id, rule_name, severity, message, scope, sheet, cell, row, col}`), `toText(file?)`, `toJSON()` (the document of `sheetrs lint --format json` as an object), `toJsonText(file?)` (the same, as text), `toSarif(file?)` (SARIF 2.1.0 text); `file` is the name the documents give the workbook |
| | `query(query?)` | `{cells: [...]}`, as `sheetrs query` prints it |
| | `range(ranges, options?)` | one block per range: `{range, values, formulas, numberFormats, displays, comments}` (the grids asked for, as lists of rows or of columns) |
| | `find(text, options?)` | the cells that match: `[{sheet, cell, row, col, value, ...}]` |
| | `inspect(options?)` | `Inspection`: `toJSON()` is `{sheets, date_system, cells, formulas, circular_groups, run_time_references, unsupported_functions, functions, diagnostics}` (`sheetc inspect`) |
| **operations** | `deleteSheets(names)`, `deleteNames(names)`, `formulaToValue(scopes)`, `deleteRows(ranges)`, `deleteColumns(ranges)`, `replace(replaces, options?)`, `calculate(request?)` | accumulate in the plan; nothing runs until `save` |
| | `pending`, `explain()` | how many operations are pending; the passes over the file that `save` will make, one operation per line |
| **the cut** | `save()` | the workbook the operations make, as a `Uint8Array` in the format of the `Book` (the plan is not consumed: saving again gives the same bytes). With none pending, the bytes as they are |
| | `valueRanges()` | `ValueRanges` of the calculation (`toJSON()` is the answer of the protocol, `{valueRanges: [{range, name, values}]}`, `toObject()` is `{name: rows}`); runs the operations in memory and writes nothing |
| | `report` | `Report` of the last `save` or `valueRanges`: `replaced`, `written`, `skipped` (`{sheet, cell, reason}`), `diagnostics`, `valueRanges`; `undefined` before |
| **compile** | `compile(options?)` | `{module, cells, reachable, formulas, folded, kept, cycles, diagnostics}`: `module` is a `Uint8Array`, the WebAssembly module of the workbook (`sheetc build --target wasm`) |
| **settings** | `dates`, `onDiagnostic` | `dates = true` gives dates and times as `{type: "date" \| "time" \| "datetime", value: "<ISO 8601>"}` in `query`, `range` and `find` instead of serial numbers (the default); `onDiagnostic = fn` receives each diagnostic record `{level, code, message, context}` of a run after it ends (an exception it throws ends the call) |

Free functions: `lintRules()` (array of `{id, name, category, is_default}`), `evaluateFormula(formula,
locale, options)` (the value of one formula, `sheetc evaluate`) and `initPanicHook()`.

The operations of a `Book` cannot be chained (`wasm-bindgen` does not return the object itself): call
them one after another. A read (`query`, `range`, `find`, `lint`, `inspect`, `compile`) on a `Book` with
operations pending throws: the answer would be about a workbook that does not exist yet; `save`, make a
`Book` of the result and ask that one.

### The operations are not a sequence

Each operation is written against the workbook **as it is** (`deleteRows("Data!3")` twice is the row 3,
once, not the rows 3 and 4), and the order the calls are made in changes nothing: the planner decides the
order they run in (replacements, then the calculation, then the conversions and the deletions in one
pass) and runs them in the fewest passes over the workbook. Step by step is a different edit of what the
first one wrote: `save`, make a `Book` of the result, and edit that. Every operation takes the **array**
of its targets (a sweep is one operation); where a list is expected a single text is accepted too.

| Operation | Argument |
| --------- | -------- |
| `deleteSheets`, `deleteNames` | sheet names, defined names |
| `formulaToValue` | scopes: `"Sheet"`, `"Sheet!B:D"`, `"Sheet!3:10"` |
| `deleteRows`, `deleteColumns` | `"Sheet!3:10"`, `"Sheet!3"`, `"Sheet!B:D"`, `"Sheet!B"` |
| `replace(replaces, options?)` | `replaces` is a list of `[text, with]` pairs or an object `{text: with}`; each pair is a replacement of its own, in order, that sees what the one before it left, so their order matters. `options`: `mode` (`literal`, `regex`, `wildcards`), `matchCase`, `entireCells`, `searchIn` (`value`, `formula`, `comment`), `formatted`, `filtered`, `scope` (sheets and ranges, the whole workbook when omitted). It looks inside formulas, as `sheetrs edit --replace` does |
| `calculate(request?)` | `{inputs, ranges, majorDimension}`, the request of `sheetc calculate`: the engine writes the `inputs` (a cell, a range or a defined name each) and calculates every formula of the workbook as it is, before the conversions and the deletions; `ranges` are the ranges `valueRanges()` answers (the workbook's own outputs when omitted; with them `save` also keeps the answer in `report.valueRanges`); `majorDimension` is `"ROWS"` (the default) or `"COLUMNS"` |

A `calculate` is a step of the plan: the workbook is written with fresh results and without the mark that
makes the application that opens it recalculate every formula; so is a workbook whose results a deletion or
a replacement left stale (the engine calculates them again as the last step). Only if the engine cannot
calculate the workbook (a function or syntax it does not have) is it written marked to be recalculated
when it is opened, with a diagnostic with the code `RECALC_NOT_SUPPORTED` in `report.diagnostics`
(and in `onDiagnostic`): the recalculation that is needed will be done by the spreadsheet program that
opens it later. `valueRanges()` has no workbook to mark: there it is an error that gives the reason.

A `formulaToValue` that follows a change which leaves results stale (rows, columns, sheets or names
deleted) converts the results calculated again by the engine, as an application that calculates as it goes
shows them. If the engine cannot calculate the workbook, the call throws with the reason.

### `query`

`query` is the object of the protocol, the JSON `sheetrs query` builds from its flags: `scope` (a scope
or a list: `"workbook"`, `{sheet: "Data"}`, `{range: "Data!A1:C9"}`), `properties` (`value`,
`formula`, `number_format`, `display`, `comment`; the value when omitted), `find` (`text`, `mode`
`literal`/`regex`/`wildcards`, `match_case`, `entire_cells`, `search_in` `value`/`formula`/`comment`,
`include_formulas`, `formatted`, `filtered`), `locale` and `dates`. `range` and `find` are the same
reads in the shape they are used in: `range(ranges, {properties, major: "rows" | "columns", dates})`
and `find(text, {mode, matchCase, entireCells, searchIn, includeFormulas, formatted, filtered, scope,
properties, by, dates})`.

### `calculate` and `valueRanges`

The clock of `NOW()`, `TODAY()` and `RAND()` is `Date.now()`. The options of a `sheetc.toml` (`fold`,
`strict`, `now`, `seed`, `host`, `locale`, ...) and `declarations` (the JSON of a `settings_file`: the
inputs and outputs of a workbook with no settings sheet) apply to `inspect(options)`, `compile(options)`
and `evaluateFormula`; the calculation of a plan takes the request only.

### `compile`

Compiles the workbook into a **WebAssembly module of its own**: the program of the workbook
(inputs, outputs, the formulas that remain) added to a runtime that holds the engine. The module
needs no JavaScript glue and no imports, so a browser or Node instantiates it and calls its
exports, as the executable of `sheetc build` is run by a shell. Those two are the hosts the tests
run it in; as it imports nothing another host could instantiate it, but none is tested. `options` are the
keys of a `sheetc.toml`, `declarations`, and the mode `embed_template`, except that:

- **there is one target here**, WebAssembly: `target` is not an option and a request for one is an
  error (the compiler of the command line has `--target native|wasm`);
- `json` is not an option: whether a module answers JSON, a workbook or both is chosen when it runs.

Without a mode the module answers JSON (the base of the engine only). With `embed_template` it carries the workbook as its template and can write it with the answer, and the
base that does it also holds the readers and writers of XLSX and ODS (it is a third larger).

The package carries the two base modules inside, so there is nothing to fetch and one package to
install. Its size, raw and compressed, is measured in
[docs/performance.md](../docs/performance.md#sizes).

#### The module of a workbook

Exports of the module (all `extern "C"`; sizes are in bytes of its linear memory, `usize` is a
32-bit integer):

| Export | |
| ------ | - |
| `memory` | the linear memory |
| `sheetrs_abi() -> i32` | the version of this interface (`1`) |
| `sheetrs_has_writer() -> i32` | `1` when the module can write a workbook (built with `embed_template`) |
| `sheetrs_alloc(len) -> ptr`, `sheetrs_free(ptr, len)` | memory for a request |
| `sheetrs_set_clock(seconds: f64)` | the clock of `NOW()` and `TODAY()` (seconds since the Unix epoch). A module has no clock of its own: until the host sets one, the time is the epoch |
| `sheetrs_set_seed(seed: f64)` | the seed of `RAND()` (an integer up to 2^53) |
| `sheetrs_meta() -> status` | the metadata, inputs and outputs of the program (`--meta` of an executable), in the result buffer |
| `sheetrs_info() -> status` | the `--info` of an executable |
| `sheetrs_calculate(ptr, len) -> status` | the JSON answer (`{"valueRanges": [...]}`) to a request (`{"inputs": {...}, "ranges": [...]}`; empty for the declared outputs), in the result buffer |
| `sheetrs_write(ptr, len, with_json) -> status` | the workbook with the answer written into it, in the workbook buffer; the JSON answer in the result buffer when `with_json` is not 0 |
| `sheetrs_result_ptr() -> ptr`, `sheetrs_result_len() -> len` | the result buffer: UTF-8 text |
| `sheetrs_workbook_ptr() -> ptr`, `sheetrs_workbook_len() -> len` | the workbook buffer |

A call returns `0` when it worked, and `1` when it failed, with the message in the result buffer
(a request the program does not take, a mistake of the JSON): the module stays usable. The buffers
hold the last answer; read them before the next call and before anything that grows the memory moves
them (create the view after the call). A trap (`RuntimeError: unreachable`) is an internal error.

```js
const { module } = new Book(bytes).compile({ embed_template: true });
const { instance } = await WebAssembly.instantiate(module, {});   // no imports
const x = instance.exports;
x.sheetrs_set_clock(Date.now() / 1000);

const request = new TextEncoder().encode(JSON.stringify({ inputs: { rate: 0.05 } }));
const ptr = x.sheetrs_alloc(request.length);
new Uint8Array(x.memory.buffer, ptr, request.length).set(request);
if (x.sheetrs_write(ptr, request.length, 1) !== 0) throw new Error(/* read the result buffer */);
x.sheetrs_free(ptr, request.length);
const json = new TextDecoder().decode(new Uint8Array(x.memory.buffer, x.sheetrs_result_ptr(), x.sheetrs_result_len()));
const workbook = new Uint8Array(x.memory.buffer, x.sheetrs_workbook_ptr(), x.sheetrs_workbook_len()).slice();
```

(A page has to instantiate asynchronously: a browser refuses to compile a big module on the main
thread.) The same module is what `sheetc build -i FILE --target wasm -o book.wasm` writes, from the base
modules `sheetrs-rt.wasm` and `sheetrs-rt-template.wasm` that sit next to `sheetc` (see the
README of `sheetrs-compiler`). Its size is that of its base plus the program (the sizes are in
[docs/performance.md](../docs/performance.md#sizes)); for a big workbook the program is most of it.

How the module finds its program, since it cannot read its own bytes: the base has a slot (a marker,
an address and a length, all zero); the compiler writes the program as one more data segment at the
end of the memory the base starts with, enlarges that memory to hold it, and writes the address and
the length into the slot. The allocator of the module only grows the memory, so nothing it allocates
overlaps the program. See `sheetrs-compiler/src/wasm_target.rs`.

## What is not here

- **The check of links** (`EXT802`, an HTTP request) is not in the linter of this package.
- **`sheetc verify`** is not in this package: the check against a reference workbook is a command-line tool.
- **Other targets.** The compiler of this package makes WebAssembly modules; the native executables are
  made by `sheetc build` on a machine that has them.
- **Depth of dependent formulas.** The engine orders the formulas of a workbook, so a column of tens
  of thousands of cells that read one another is calculated in a loop. When a reference is computed
  at run time (`INDIRECT`, `OFFSET`) the chain is followed recursively, and in a module the depth is
  what the host's call stack allows: about a thousand levels per MiB (Node: 1 MiB by default, so
  about 1 000; `node --stack-size=7000` gives about 5 000). A longer chain throws a `RangeError`.
  The command-line tools run the same recursion on a stack of 512 MiB.

## Install

```bash
npm install https://github.com/cosmoscalibur/sheetrs/releases/download/vX.Y.Z/sheetrs-wasm-X.Y.Z.tgz
```

The tarball is an asset of each release; npm publishing is pending ([#108](https://github.com/cosmoscalibur/sheetrs/issues/108) covers crates.io and PyPI). More in the [repository README](../README.md#installation).

## Build

`scripts/package-wasm.sh` builds the npm package into `target/npm/` (see the repository README):
first the two base modules of the WebAssembly target (`scripts/build-wasm-rt.sh`), then the package
with them inside; `--demo` also
copies the package into `docs/pkg/`, where the web demo reads it. The Cargo features choose what the
module contains: `query`, `edit` and `calculate` (all on by default; `query` adds `query`, `range` and `find`; `edit` adds the operations, `save` and `valueRanges`, and carries the engine, for the `calculate` step and the
recalculation before a conversion; `calculate` adds `inspect` and `evaluateFormula`); `bundled-runtime` adds `compile` with the base modules inside (the script sets it, with
`SHEETRS_RT_DIR` naming the folder of the bases); the `Book`, the linter and the statistics are always in.

- `sheetrs-wasm/tests/smoke.mjs` loads the package in Node and calls every tool on real workbooks,
  comparing each answer with the command-line tool's; it compiles workbooks and runs the modules.
- `sheetrs-wasm/tests/demo.mjs` drives the web demo (`docs/index.html`) in Chromium (Playwright):
  every tool on real workbooks, and screenshots on a desktop and a phone screen.
- `sheetrs-wasm/tests/timing.mjs` times one call of each tool.
