### Changed

- Inputs and outputs are each **declared or not**. The outputs: only the declared form exists, and
  without a declaration the outputs are deduced: the formula cells, as the rectangles that hold them. The inputs: declared, or, without a declaration, deduced: once the workbook is pruned to what
  the outputs read, every constant a path starts from is an input, as the fewest rectangles that cover
  them (named by their range); nothing that reads one is folded. `"inputs": []` declares that there are
  none. Defined names and tables are aliases in both modes: a name inside the outputs reads that output,
  and one inside the inputs gives those inputs; a key of the `inputs` of a request is an input, an alias
  or a range inside the inputs. `--meta` lists the inputs and outputs that were exposed, declared or not.
- Removed: the levels `declared`, `named` and `any` of the outputs (`--outputs`, the `outputs` key of
  `sheetc.toml`, `QueryLevel`, `Queryable::Declared|Named|Any`). `Spec` has `inputs_declared` and
  `mode`; `ir::Queryable` is the list of aliases; `BuildOptions` says `outputs` and
  `inputs` (`declared` or `deduced`). Programs written by earlier versions with `--emit-ir` do not load.
- The convention that `--outputs any` with no inputs recalculates a workbook is gone: recalculating is
  `--recalculate`.
