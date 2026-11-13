// One call of a tool of the WebAssembly package, as a whole process, to read its peak memory (the
// process loads the package, makes the call once and exits; `none` loads the package and stops):
//
//   node sheetrs-wasm/tests/call_run.mjs none|lint|stats|query|edit|calculate
import { readFileSync } from "node:fs";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const pkg = resolve(process.env.PACKAGE ?? join(root, "target/npm"));
const wasm = await import(pathToFileURL(join(pkg, "sheetrs_wasm.js")).href);
wasm.initSync({ module: readFileSync(join(pkg, "sheetrs_wasm_bg.wasm")) });

const lintBook = join(root, "tests/lint_book.xlsx");
const loan = join(root, "sheetrs-compiler/tests/assets/loan.xlsx");
const book = join(root, "sheetrs-edit/tests/assets/book/book.xlsx");
const open = (file) => new wasm.Book(new Uint8Array(readFileSync(file)));
const request = { inputs: { principal: 250000, rate: 0.045, years: 15, month: 24 } };
const scenarios = {
  none: () => {},
  lint: () => open(lintBook).lint().toSarif(lintBook),
  stats: () => open(lintBook).stats.toJSON(),
  query: () => open(book).query({ scope: { range: "Data!A1:C9" } }),
  edit: () => {
    const edited = open(book);
    edited.formulaToValue("Data!A:A");
    edited.deleteRows("Data!10:11");
    edited.deleteColumns("Data!B");
    return edited.save();
  },
  calculate: () => {
    const calculating = open(loan);
    calculating.calculate(request);
    return calculating.valueRanges().toJSON();
  },
};
scenarios[process.argv[2]]();
