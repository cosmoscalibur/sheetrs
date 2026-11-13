// The time of one call of each tool in the WebAssembly package (median of 20, after 3 warm-ups)
// against the command-line tool on the same workbook (median of 20 runs, process start included: a
// command has to start, a call in a loaded module does not).
//
//   scripts/package-wasm.sh && cargo build --release -p sheetrs-linter -p sheetrs-stats \
//     -p sheetrs-query -p sheetrs-edit -p sheetrs-compiler && BIN=target/release node sheetrs-wasm/tests/timing.mjs

import { execFileSync } from "node:child_process";
import { readFileSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const pkg = resolve(process.env.PACKAGE ?? join(root, "target/npm"));
const bin = resolve(process.env.BIN ?? join(root, "target/release"));
const wasm = await import(pathToFileURL(join(pkg, "sheetrs_wasm.js")).href);
wasm.initSync({ module: readFileSync(join(pkg, "sheetrs_wasm_bg.wasm")) });

const lintBook = join(root, "tests/lint_book.xlsx");
const loan = join(root, "sheetrs-compiler/tests/assets/loan.xlsx");
const book = join(root, "sheetrs-edit/tests/assets/book/book.xlsx");
const read = (file) => new Uint8Array(readFileSync(file));
const work = mkdtempSync(join(tmpdir(), "sheetrs-timing-"));
const request = JSON.stringify({ inputs: { principal: 250000, rate: 0.045, years: 15, month: 24 } });

const median = (f) => {
  for (let i = 0; i < 3; i++) f();
  const times = [];
  for (let i = 0; i < 20; i++) {
    const t = process.hrtime.bigint();
    f();
    times.push(Number(process.hrtime.bigint() - t) / 1e6);
  }
  return times.sort((a, b) => a - b)[10];
};
const SUBCOMMANDS = ["query", "edit", "lint", "stats"];
const cli = (tool, args, input) => () => {
  try {
    const [file, argv] = SUBCOMMANDS.includes(tool) ? [join(bin, "sheetrs"), [tool, ...args]] : [join(bin, tool), args];
    execFileSync(file, argv, { input, stdio: ["pipe", "pipe", "pipe"], maxBuffer: 1 << 30 });
  } catch (error) {
    if (!(tool === "lint" && error.status === 1)) throw error;
  }
};

const book_ = (file) => new wasm.Book(read(file));
const cases = [
  ["lint, SARIF (lint_book.xlsx)", () => book_(lintBook).lint().toSarif(lintBook), cli("lint", ["-i", lintBook, "--format", "sarif"])],
  ["statistics (lint_book.xlsx)", () => book_(lintBook).stats.toJSON(), cli("stats", ["-i", lintBook, "--format", "json"])],
  ["query, 9 rows (book.xlsx)", () => book_(book).query({ scope: { range: "Data!A1:C9" } }), cli("query", ["-i", book, "--range", "Data!A1:C9"])],
  ["edit, 3 operations (book.xlsx)", () => {
    const edited = book_(book);
    edited.formulaToValue("Data!A:A");
    edited.deleteRows("Data!10:11");
    edited.deleteColumns("Data!B");
    return edited.save();
  }, cli("edit", ["-i", book, "--delete-rows", "Data!10:11", "--formula-to-value", "Data!A:A", "--delete-columns", "Data!B", "-o", join(work, "out.xlsx")])],
  ["calculate (loan.xlsx)", () => {
    const calculating = book_(loan);
    calculating.calculate(JSON.parse(request));
    return calculating.valueRanges().toJSON();
  }, cli("sheetc", ["calculate", "-i", loan, "--request", "-"], request)],
];
console.log("| Call | WebAssembly (ms) | Command (ms) |\n| ---- | ---: | ---: |");
for (const [name, call, command] of cases) {
  console.log(`| ${name} | ${median(call).toFixed(1)} | ${median(command).toFixed(1)} |`);
}
rmSync(work, { recursive: true, force: true });
