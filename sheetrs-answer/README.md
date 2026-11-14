# sheetrs-answer

The answer of a compiled `SheetRS` program written into the workbook it was made from: the
workbook is the template of its own answer. It runs the program of `sheetrs-engine` over a
`Book` and writes the results of the formulas and the inputs into the cells, in the formats
the cells have, with the streaming writers of `sheetrs-io` (feature `write`).

```rust,no_run
use sheetrs_answer::fill_template;
# fn run(
#     template: &sheetrs_io::Book,
#     program: &sheetrs_engine::ir::Program,
#     request: &serde_json::Value,
# ) -> sheetrs_answer::Result<()> {
let written = fill_template(
    template,
    program,
    None,
    request,
    &Default::default(),
    std::path::Path::new("out.xlsx"),
)?;
println!("{} cells written", written.cells);
# Ok(()) }
```

- `fill_template` writes a file, `fill_template_to` any `io::Write` (the bytes of a workbook, for
  a host with no file system), and the `_with_answer` forms also give the JSON answer of the
  program to the same request, from the same calculation.
- In the embedded mode a range of the request that lies outside the outputs of the program is read,
  in memory and with no recalculation, from the workbook just written, with `sheetrs-query` (the
  answer keeps the order of the request); a compiled program that writes no workbook refuses it
  (`sheetrs_engine::Error::OutsideOutputs`). The table is in the README of `sheetrs-compiler`.
- It depends on `sheetrs-engine`, `sheetrs-io` and `sheetrs-query` only. `sheetc` (`calculate -o OUT`,
  `--embed-template`), the editor (`BookEdit::calculate`, the recalculation before a conversion of
  formulas to values) and the WebAssembly base module that writes workbooks
  (`sheetrs-rt-wasm`, feature `template`) are its callers.

See [`ARCHITECTURE.md`](../ARCHITECTURE.md) for where it sits in the dependency graph.
