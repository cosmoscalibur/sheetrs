# Recipes

Short procedures for tasks that take more than one tool. Every command is one the tools have.

## Lint a workbook whose cached results are stale or are not what LibreOffice shows

A workbook keeps the results its last application calculated. If the file was written by another
program, edited without recalculating, or saved by LibreOffice as an XLSX (which has no codes for
its own errors: `=SQRT(-1)` is cached as `#VALUE!` where LibreOffice shows `Err:502` when it opens
the file), the rules that look at values see the cache. **Recalculate first, then lint**:

```bash
sheetc calculate -i book.xlsx -o recalculated.xlsx   # nothing declared or given: every formula computed again
sheetrs lint -i recalculated.xlsx                                     # the rules see fresh results
sheetrs query -i recalculated.xlsx --workbook -p value,display        # or read them
```

- There is no `--recalculate`: `calculate` with an output and no inputs, no `--range` and no declarations
  calculates every formula again. The output is a copy of the workbook with the same formulas,
  styles and formats, and the results of this calculation.
- The results are the ones the engine computes with LibreOffice's semantics, not the cache. The
  codes only LibreOffice has are written as the format can: an XLSX has none (`#VALUE!`), an ODS
  keeps the code LibreOffice gave (`Err:502`, `Err:522`…). The engine carries that identity for the
  functions where LibreOffice always answers the same code; where one function gives several
  codes it answers the most frequent one (a technical limit, issues #90 and #91).
- A workbook that uses a function the engine does not have is refused; `sheetc inspect -i book.xlsx`
  lists them.
- `-o` never overwrites the input; to keep the result in place use `--in-place`
  (`sheetc calculate -i book.xlsx --in-place`), which writes beside the file and renames it over the
  original, so a failure leaves the original as it was.

## Get the answer of a model as a workbook, in the formats of its cells

```bash
sheetc build -i model.xlsx --embed-template -o model       # the workbook travels inside the executable
./model --request '{"inputs": {"Data!B2": 1500}}' --output answer.xlsx
```

`answer.xlsx` is a copy of the workbook with the outputs that are formulas calculated (every formula
when no output is declared) and the inputs given written in their cells, so dates, percentages and currencies come out as the template shows them. Any range or
defined name inside the inputs is accepted as an input. To read values from the answer, use
`sheetrs query` (`-p value` for the numbers, `-p display` for what a person sees).

The same executable gives JSON when it is run without `--output` (`./model --request @request.json`; `--request -` reads the standard input), and
both from one calculation with `--output answer.xlsx --json`. Declare inputs and outputs (settings
sheet or `settings_file`) to narrow the JSON to the outputs you name and to accept only the inputs
you name; the workbook answer then writes the declared outputs, and the formulas outside them keep
the result the template stored.

## Get the answer of a model as JSON

```bash
sheetc build -i model.xlsx -o model       # inputs and outputs from the settings sheet, or deduced
./model --request '{"inputs": {"rate": 0.05}}'
./model --meta                          # what the executable exposes
```

Without a declaration the outputs are the formula cells (the cells that are calculated, as the
rectangles that hold them) and the inputs are deduced from the dependencies; `--meta` lists them. Declare them (settings sheet or `settings_file`) to expose less, to name them,
and to let everything that does not depend on an input be calculated once, at build time.

## Find where something is, and replace it

```bash
sheetrs query -i book.xlsx --find 'Q1' --entire-cells -p value,formula --jsonl
sheetrs query -i book.ods  --find '$Data.' --in formula --jsonl      # in the notation of LibreOffice
sheetrs edit -i book.xlsx --replace 'Q1' --with 'T1' --entire-cells --output renamed.xlsx
sheetrs edit -i book.xlsx --replace 'Data' --with 'Datos' --in formula --in-place
```

A search sees what LibreOffice puts in the input line of a cell (`1234.5` for a cell that shows
`1,234.50`), and a formula in LibreOffice's notation of that file (see
`sheetrs-query/README.md`). `--in-place` writes to a temporary file and renames it over the
original, so a failure leaves the original as it was.
