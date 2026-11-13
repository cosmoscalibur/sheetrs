// A smoke test of the WebAssembly package: every tool is called in a real WebAssembly runtime (Node)
// on real workbooks, and its answer is compared with the one the command-line tool gives for the
// same file.
//
//   scripts/package-wasm.sh && cargo build -p sheetrs-linter -p sheetrs-stats -p sheetrs-query \
//     -p sheetrs-edit -p sheetrs-compiler && node sheetrs-wasm/tests/smoke.mjs
//
// The package is `target/npm` and the commands are `target/debug` unless `PACKAGE` and `BIN` say
// otherwise. Exit code 0 when everything matches.

import { execFileSync } from "node:child_process";
import { readFileSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import assert from "node:assert/strict";
import zlib from "node:zlib";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const pkg = resolve(process.env.PACKAGE ?? join(root, "target/npm"));
const bin = resolve(process.env.BIN ?? join(root, "target/debug"));

const wasm = await import(pathToFileURL(join(pkg, "sheetrs_wasm.js")).href);
wasm.initSync({ module: readFileSync(join(pkg, "sheetrs_wasm_bg.wasm")) });
wasm.initPanicHook();

const work = mkdtempSync(join(tmpdir(), "sheetrs-wasm-"));
const lintBook = join(root, "tests/lint_book.xlsx");
const lintBookOds = join(root, "tests/lint_book.ods");
const loan = join(root, "sheetrs-compiler/tests/assets/loan.xlsx");
const loanOds = join(root, "sheetrs-compiler/tests/assets/loan.ods");
const book = join(root, "sheetrs-edit/tests/assets/book/book.xlsx");
const bookOds = join(root, "sheetrs-edit/tests/assets/book/book.ods");

const ext = (file) => file.split(".").pop();
const read = (file) => new Uint8Array(readFileSync(file));
const open = (file) => new wasm.Book(read(file));
// A linter exits with 1 when it finds violations: what it wrote is its answer all the same.
// `lint`, `query`, `edit` and `stats` are the subcommands of `sheetrs`; any other tool is an executable.
const SUBCOMMANDS = ["query", "edit", "lint", "stats"];
const command = (tool, args) =>
  SUBCOMMANDS.includes(tool) ? [join(bin, "sheetrs"), [tool, ...args]] : [join(bin, tool), args];
const run = (tool, args, input) => {
  try {
    const [file, argv] = command(tool, args);
    return execFileSync(file, argv, {
      input,
      stdio: ["pipe", "pipe", "pipe"],
      maxBuffer: 1 << 30,
    });
  } catch (error) {
    if (tool === "lint" && error.status === 1 && error.stdout?.length) return error.stdout;
    throw error;
  }
};
const text = (buffer) => buffer.toString("utf8");
// Bytes are compared with `equals`: a failed `deepEqual` of two big buffers computes a diff that exhausts the memory.
const sameBytes = (a, b, message) => assert.ok(Buffer.from(a).equals(Buffer.from(b)), message ?? "the bytes differ");
// The values of a range of two workbooks, equal up to the digits a writer keeps.
const sameValues = (a, b, range) => {
  const [x, y] = [a, b].map((data) => new wasm.Book(data).range(range)[0].values.flat());
  assert.equal(x.length, y.length);
  x.forEach((value, i) => {
    if (typeof value === "number") assert.ok(Math.abs(value - y[i]) <= 1e-9 * Math.max(1, Math.abs(value)), `${range}[${i}]: ${value} and ${y[i]}`);
    else assert.deepEqual(value, y[i]);
  });
};

let checks = 0;
const check = (name, fn) => {
  try {
    fn();
    checks += 1;
    console.log(`ok   ${name}`);
  } catch (error) {
    console.error(`FAIL ${name}\n${error.message ?? error}`);
    process.exitCode = 1;
  }
};

// -- lint ------------------------------------------------------------------------------------
for (const file of [lintBook, lintBookOds]) {
  check(`lint (SARIF) ${ext(file)}`, () => {
    const cli = text(run("lint", ["-i", file, "--format", "sarif"]));
    const sarif = open(file).lint().toSarif(file);
    assert.equal(sarif, cli);
    const log = JSON.parse(sarif);
    assert.equal(log.version, "2.1.0");
    assert.ok(log.runs[0].results.length > 20);
  });
  check(`lint (JSON) ${ext(file)}`, () => {
    const cli = text(run("lint", ["-i", file, "--format", "json"]));
    assert.equal(open(file).lint().toJsonText(file), cli);
    assert.deepEqual(open(file).lint().toJSON(), JSON.parse(open(file).lint().toJsonText(undefined)));
  });
  check(`lint (records and rules) ${ext(file)}`, () => {
    const report = open(file).lint();
    const sarif = JSON.parse(report.toSarif(file));
    assert.equal(report.records.length, sarif.runs[0].results.length);
    assert.equal(report.count, report.records.length);
    assert.ok(wasm.lintRules().length >= 50);
    assert.ok(report.toText().includes("Summary"));
  });
}

// -- stats -----------------------------------------------------------------------------------
check("the format is the one of the bytes", () => {
  assert.throws(() => new wasm.Book(new Uint8Array([1, 2, 3])), /Reader error|not an XLSX/);
  for (const file of [lintBook, lintBookOds]) {
    assert.ok(open(file).stats().toJSON().total_sheets > 0);
  }
  assert.equal(open(lintBook).format, "xlsx");
  assert.equal(open(lintBookOds).format, "ods");
});
check("stats", () => {
  const cli = JSON.parse(text(run("stats", ["-i", lintBook, "--format", "json"])));
  const stats = open(lintBook).stats().toJSON();
  assert.equal(stats.total_file_size, 0);
  cli.total_file_size = 0;
  assert.deepEqual(stats, cli);
});

// What a Book says about the workbook is what `stats()` says: a Book itself reads nothing.
check("stats gives the totals, the sheets, the macros and the names", () => {
  const cli = JSON.parse(text(run("stats", ["-i", lintBook, "--format", "json"])));
  const stats = open(lintBook).stats();
  const json = stats.toJSON();
  assert.equal(json.total_sheets, cli.total_sheets);
  assert.equal(json.total_used_cells, cli.total_used_cells);
  assert.equal(json.total_formulas, cli.total_formulas);
  assert.equal(json.total_named_ranges, cli.total_named_ranges);
  assert.deepEqual(json.sheets.map((sheet) => sheet.name), cli.sheets.map((sheet) => sheet.name));
  assert.equal(typeof stats.hasMacros, "boolean");
  assert.equal(typeof stats.definedNames, "object");
});

// -- query -----------------------------------------------------------------------------------
for (const file of [book, bookOds]) {
  check(`query (range and properties) ${ext(file)}`, () => {
    const cli = JSON.parse(
      text(run("query", ["-i", file, "--range", "Data!A1:C9", "-p", "value,formula,number_format,display"])),
    );
    const cells = open(file).query({
      scope: { range: "Data!A1:C9" },
      properties: ["value", "formula", "number_format", "display"],
    });
    assert.deepEqual(cells, cli);
    assert.ok(cells.cells.length > 5);
  });
  check(`query (find) ${ext(file)}`, () => {
    const cli = JSON.parse(text(run("query", ["-i", file, "--find", "a", "--match-case", "--workbook"])));
    const found = open(file).query({
      find: { text: "a", match_case: true },
    });
    assert.deepEqual(found, cli);
    // `find` is the same search, as a list of cells.
    const cells = open(file).find("a", { matchCase: true });
    assert.equal(cells.length, cli.cells.length);
    assert.deepEqual(cells.map((c) => c.cell), cli.cells.map((c) => c.cell));
  });
  check(`range ${ext(file)}`, () => {
    const cli = JSON.parse(text(run("query", ["-i", file, "--range", "Data!A1:C9", "-p", "value,formula"])));
    const [block] = open(file).range("Data!A1:C9", { properties: ["value", "formula"] });
    assert.equal(block.range, "Data!A1:C9");
    assert.equal(block.values.length, 9);
    assert.equal(block.formulas.length, 9);
    const withValue = cli.cells.filter((c) => c.value !== null && c.value !== undefined);
    const found = block.values.flat().filter((v) => v !== null);
    assert.equal(found.length, withValue.length);
    const [byColumns] = open(file).range(["Data!A1:C9"], { major: "columns" });
    assert.equal(byColumns.values.length, 3);
    assert.equal(byColumns.values[0].length, 9);
  });
}

// -- dates as dates ---------------------------------------------------------------------------
const assets = join(root, "sheetrs-io/tests/assets");
for (const name of ["dates.xlsx", "dates.ods", "dates_1904.xlsx", "dates_1904.ods", "dates_1900.ods"]) {
  const file = join(assets, name);
  check(`query (dates as dates) ${name}`, () => {
    const cli = JSON.parse(text(run("query", ["-i", file, "--workbook", "--dates"])));
    const cells = open(file).query({ dates: true });
    const flagged = open(file);
    flagged.dates = true;
    assert.deepEqual(flagged.query({}), cells);
    assert.deepEqual(cells, cli);
    assert.ok(cells.cells.some((c) => c.value?.type === "datetime"));
    assert.ok(cells.cells.some((c) => c.value?.type === "time"));
    const serial = open(file).query({});
    // An ODS stores a date before the year 1 as text (`-0001-12-31`): the reader does not take it for a date.
    assert.ok(serial.cells.every((c) => typeof c.value === "number" || String(c.value).startsWith("-")));
  });
}

// A date system that is none of the three of LibreOffice Calc is an error that names the date, from the
// package and from the command alike.
check("query (an ODS with another null date is an error that names it)", () => {
  const file = join(assets, "other_null_date.ods");
  assert.throws(() => open(file).query({}), (error) => {
    const message = String(error);
    return ["1950-06-15", "1899-12-30", "1900-01-01", "1904-01-01"].every((d) => message.includes(d));
  });
  assert.throws(
    () => run("query", ["-i", file, "--workbook"]),
    (error) => String(error.stderr).includes("1950-06-15"),
  );
  assert.throws(() => open(file).stats(), (error) => String(error).includes("1950-06-15"));
});

// -- edit ------------------------------------------------------------------------------------
for (const file of [book, bookOds]) {
  check(`edit (sequence of passes) ${ext(file)}`, () => {
    const out = join(work, `edited.${ext(file)}`);
    // No replacement before the conversion: text replaced leaves results stale, and the command has no
    // engine to calculate them again (the plan here has one), so it would refuse; the sweep test
    // below covers the replacements.
    run("edit", [
      "-i", file,
      "--formula-to-value", "Data!A:A",
      "--delete-rows", "Data!10:11",
      "--delete-columns", "Data!B",
      "-o", out,
    ]);
    const edited = open(file);
    edited.formulaToValue("Data!A:A");
    edited.deleteRows("Data!10:11");
    edited.deleteColumns("Data!B");
    assert.equal(edited.pending, 3);
    const saved = edited.save();
    assert.ok(saved instanceof Uint8Array);
    sameBytes(saved, readFileSync(out));
    assert.ok(edited.report.replaced >= 0);
    assert.ok(edited.explain().includes("delete rows"));
    // The plan is not consumed: saved again it gives the same bytes.
    sameBytes(edited.save(), saved);
  });
}
// A sweep: arrays of targets and a replace with pairs, the same as the command with its flags repeated.
for (const file of [book, bookOds]) {
  check(`edit (arrays: a sweep is one operation) ${ext(file)}`, () => {
    const out = join(work, `swept.${ext(file)}`);
    run("edit", [
      "-i", file, "--replace", "a", "--with", "A", "--replace", "A", "--with", "aa",
      "--delete-rows", "Data!10:11", "Data!15", "--delete-columns", "Data!B", "Data!D",
      "-o", out,
    ]);
    const swept = open(file);
    swept.replace([["a", "A"], ["A", "aa"]]);
    swept.deleteRows(["Data!10:11", "Data!15"]);
    swept.deleteColumns(["Data!B", "Data!D"]);
    sameBytes(swept.save(), readFileSync(out));
    // The single-text forms are the same operations.
    const single = open(file);
    single.deleteRows("Data!10:11");
    const array = open(file);
    array.deleteRows(["Data!10:11"]);
    sameBytes(single.save(), Buffer.from(array.save()));
    assert.equal(swept.explain().split("replace").length - 1, 2);
    // The order the operations are given in changes nothing: each is written against the workbook as it is.
    const reversed = open(file);
    reversed.deleteColumns(["Data!B", "Data!D"]);
    reversed.deleteRows(["Data!10:11", "Data!15"]);
    reversed.replace({ a: "A" });
    reversed.replace({ A: "aa" });
    sameBytes(reversed.save(), Buffer.from(swept.save()));
    assert.throws(() => open(file).replace("a"), /not valid/);
  });
}
check("edit (a mistake is an error, not a crash)", () => {
  const nope = open(book);
  nope.deleteRows("Nope!3");
  assert.throws(() => nope.save(), /Edit error|Nope/);
  assert.throws(() => open(book).deleteRows("bogus"), /Edit error/);
  // A query about a workbook that does not exist yet is refused, not answered about the old one.
  assert.throws(() => nope.range("Data!A1"), /Pending operations/);
  // With nothing to do `save` gives the bytes as they are.
  sameBytes(open(book).save(), readFileSync(book));
});

// -- calculate as a step of the plan -----------------------------------------------------------
const structure = (extension) => join(root, `sheetrs-edit/tests/assets/structure/book.${extension}`);
// One entry of a zip (the central directory is read, so entries written with a data descriptor
// are found too).
function zipEntry(data, name) {
  const buffer = Buffer.from(data);
  let end = buffer.length - 22;
  while (end >= 0 && buffer.readUInt32LE(end) !== 0x06054b50) end--;
  let at = buffer.readUInt32LE(end + 16);
  for (let n = buffer.readUInt16LE(end + 10); n > 0; n--) {
    const method = buffer.readUInt16LE(at + 10);
    const size = buffer.readUInt32LE(at + 20);
    const nameLength = buffer.readUInt16LE(at + 28);
    const extra = buffer.readUInt16LE(at + 30) + buffer.readUInt16LE(at + 32);
    const local = buffer.readUInt32LE(at + 42);
    const entry = buffer.toString("utf8", at + 46, at + 46 + nameLength);
    if (entry === name) {
      const start = local + 30 + buffer.readUInt16LE(local + 26) + buffer.readUInt16LE(local + 28);
      const packed = buffer.subarray(start, start + size);
      return (method === 8 ? zlib.inflateRawSync(packed) : packed).toString("utf8");
    }
    at += 46 + nameLength + extra;
  }
  throw new Error(`no entry ${name}`);
}
const marked = (workbook) => /fullCalcOnLoad="(1|true)"/.test(zipEntry(workbook, "xl/workbook.xml"));
const valueOf = (workbook, range) =>
  new wasm.Book(workbook).query({ scope: { range } }).cells[0].value;
for (const extension of ["xlsx", "ods"]) {
  check(`edit (calculate: fresh results, no warning) ${extension}`, () => {
    const edited = open(structure(extension));
    edited.deleteRows("T!2");
    edited.calculate();
    const workbook = edited.save();
    assert.deepEqual(edited.report.diagnostics, []);
    assert.equal(edited.report.valueRanges, undefined);
    // F!A5 is SUM(T!A2:A3): the row 2 went, so it adds the one row that is left.
    assert.equal(valueOf(workbook, "F!A5"), 3);
    if (extension === "xlsx") {
      assert.equal(marked(workbook), false);
      // The same deletion with no calculation: the engine calculates what it left stale, so it is not
      // marked either (the mark is only for a workbook the engine cannot calculate).
      const left = open(structure(extension));
      left.deleteRows("T!2");
      assert.equal(marked(left.save()), false);
      assert.deepEqual(left.report.diagnostics, []);
    }
    const planned = open(structure(extension));
    planned.calculate();
    assert.ok(planned.explain().includes("calculate every formula"));
  });
}
// `now` and `seed` of a calculation fix NOW(), TODAY() and RAND(), and the recalculation of the same plan.
const volatileBook = (extension) => join(root, `sheetrs-edit/tests/assets/volatile/volatile.${extension}`);
for (const extension of ["xlsx", "ods"]) {
  check(`edit (calculate with now and seed fixes NOW, TODAY and RAND) ${extension}`, () => {
    const row = (options) => {
      const edited = open(volatileBook(extension));
      edited.calculate({ ranges: ["V!A1:C1"] }, options);
      return edited.valueRanges().toObject()["V!A1:C1"][0];
    };
    const [now, today, rand] = row({ now: 45366.75, seed: 42 });
    assert.equal(now, 45366.75);
    assert.equal(today, 45366);
    assert.ok(rand >= 0 && rand < 1);
    assert.equal(row({ seed: 42 })[2], rand);
    assert.notEqual(row({ seed: 44 })[2], rand);
    assert.ok(row(undefined)[0] > 45292, "without options the clock is the system's");
    assert.throws(() => row({ nowish: 1 }), /unknown key `nowish`/);
  });
  check(`edit (the recalculation after a deletion runs with now and seed) ${extension}`, () => {
    const edited = open(volatileBook(extension));
    edited.calculate(undefined, { now: 45366.75, seed: 7 });
    edited.deleteRows("V!3");
    const saved = edited.save();
    const cells = new wasm.Book(saved).range("V!A1:C1")[0].values[0];
    assert.equal(cells[0], 45366.75);
    assert.equal(cells[1], 45366);
    const alone = open(volatileBook(extension));
    alone.calculate({ ranges: ["V!A1:C1"] }, { now: 45366.75, seed: 7 });
    assert.equal(cells[2], alone.valueRanges().toObject()["V!A1:C1"][0][2]);
  });
}
check("edit (calculate with inputs and the ranges asked for)", () => {
  const edited = open(structure("xlsx"));
  edited.calculate({ inputs: { "T!A1": 100 }, ranges: ["T!C1"] });
  const answer = edited.valueRanges();
  assert.deepEqual(edited.report.diagnostics, []);
  assert.deepEqual(answer.toJSON().valueRanges[0].values, [[100]]);
  assert.deepEqual(answer.toObject(), { "T!C1": [[100]] });
  assert.equal(valueOf(edited.save(), "T!C1"), 100);
});
check("edit (calculate on a workbook the engine cannot calculate: marked and a warning)", () => {
  const edited = open(book);
  edited.deleteRows("Data!10:11");
  edited.calculate();
  const seen = [];
  edited.onDiagnostic = (d) => seen.push(d);
  const workbook = edited.save();
  const records = edited.report.diagnostics;
  assert.equal(records.length, 1);
  const [record] = records;
  assert.equal(record.level, "warning");
  assert.equal(record.code, "RECALC_NOT_SUPPORTED");
  assert.match(record.message, /recalculation it needs will be done by the spreadsheet program/);
  assert.equal(record.context.operation, "calculate");
  assert.ok(record.context.reason.length > 0);
  // onDiagnostic receives the same records as the run reports them.
  assert.deepEqual(seen, records);
  assert.equal(marked(workbook), true);
  // With the ranges asked for there is no answer to give: an error, not a warning.
  const asked = open(book);
  asked.calculate();
  assert.throws(() => asked.valueRanges(), /could not be calculated/);
  assert.throws(() => open(book).valueRanges(), /no calculate step/);
});

// -- calculate -------------------------------------------------------------------------------
const request = { inputs: { principal: 250000, rate: 0.045, years: 15, month: 24 } };
// The answer and the workbook of a calculation, from one Book.
const calculated = (file, req) => {
  const calculating = open(file);
  calculating.calculate(req);
  return { answer: calculating.valueRanges().toJSON(), workbook: calculating.save() };
};
for (const file of [loan, loanOds]) {
  check(`calculate (values) ${ext(file)}`, () => {
    const cli = JSON.parse(text(run("sheetc", ["calculate", "-i", file, "--request", "-"], JSON.stringify(request))));
    assert.deepEqual(calculated(file, request).answer, cli);
  });
  check(`calculate (the workbook with the results) ${ext(file)}`, () => {
    const out = join(work, `embedded.${ext(file)}`);
    run("sheetc", ["calculate", "-i", file, "--request", "-", "--output", out], JSON.stringify(request));
    sameBytes(calculated(file, request).workbook, readFileSync(out));
  });
}
check("calculate (nothing given: the workbook as it is, with every formula calculated again)", () => {
  const out = join(work, "recalculated.xlsx");
  run("sheetc", ["calculate", "-i", loan, "--output", out]);
  const recalculating = open(loan);
  recalculating.calculate();
  sameValues(recalculating.save(), read(out), "Loan!B1:B6");
});
check("calculate (an input the workbook does not have is an error; the layout of the answer)", () => {
  const wrong = open(loan);
  wrong.calculate({ inputs: { bogus: 1 } });
  assert.throws(() => wrong.valueRanges(), /bogus/);
  assert.throws(() => open(loan).calculate({ unknown: 1 }), /not valid/);
  const rows = new wasm.Book(chainWorkbook(3, (n) => `A${n - 1}+1`));
  rows.calculate({ ranges: ["Chain!A2:A3"] });
  const columns = new wasm.Book(chainWorkbook(3, (n) => `A${n - 1}+1`));
  columns.calculate({ ranges: ["Chain!A2:A3"], majorDimension: "COLUMNS" });
  assert.deepEqual(rows.valueRanges().toJSON().valueRanges[0].values, [[2], [3]]);
  assert.deepEqual(columns.valueRanges().toJSON().valueRanges[0].values, [[2, 3]]);
});
check("calculate (the clock is Date.now(): NOW() is not the epoch)", () => {
  const before = Date.now() / 1000;
  const now = wasm.evaluateFormula("=NOW()", undefined, undefined);
  // Days since 1899-12-30, a serial number: the epoch would be 25569.
  assert.ok(now > 25569 + before / 86400 - 1, `NOW() gave ${now}`);
  const fixed = wasm.evaluateFormula("=NOW()", undefined, { now: 45000.5 });
  assert.equal(fixed, 45000.5);
});
// -- the locale of a Book: `new Book(bytes, locale)` --------------------------------------------------
// `semantics_es-CO` was made by a LibreOffice running with the es-CO regional setting. Neither format
// carries it (the language an ODS stores in its default cell style is not read, as LibreOffice does not),
// and WebAssembly has no system to take it from: without a locale a Book is en-US.
const esCO = join(root, "sheetrs-compiler/tests/assets/semantics_es-CO");
const formulaLocale = 'TEXT(350*49799,"[$$-en-US]#.###.##0")';
for (const [extension, locale, expected] of [
  ["xlsx", "es-CO", "$17.429.650"],
  ["xlsx", "en-US", "$17429650.000.000"],
  ["xlsx", undefined, "$17429650.000.000"],
  ["ods", "es-CO", "$17.429.650"],
  ["ods", "en-US", "$17429650.000.000"],
  ["ods", undefined, "$17429650.000.000"],
]) {
  check(`locale of the Book: ${extension}, ${locale ?? "the default of the format"}`, () => {
    const data = read(`${esCO}.${extension}`);
    const labels = new wasm.Book(data).range("Calc!B1:B1500")[0].values.flat();
    const asked = `Calc!A${labels.indexOf(formulaLocale) + 1}`;
    const calculating = new wasm.Book(data, locale);
    assert.equal(calculating.locale, locale);
    calculating.calculate({ ranges: [asked] });
    assert.deepEqual(calculating.valueRanges().toJSON().valueRanges[0].values, [[expected]]);
  });
}
check("locale of the Book: a text that is not a locale tag is refused", () => {
  assert.throws(() => new wasm.Book(read(`${esCO}.xlsx`), "not a locale"), /not a locale tag/);
});
check("evaluate", () => {
  for (const formula of ["=SUM(1,2,3)", '=TEXT(0.25,"0%")', "=1/0"]) {
    const cli = text(run("sheetc", ["evaluate", formula])).trim();
    assert.equal(JSON.stringify(wasm.evaluateFormula(formula, undefined, undefined)), cli, formula);
  }
});
check("inspect", () => {
  const cli = text(run("sheetc", ["inspect", "-i", loan]));
  const report = open(loan).inspect().toJSON();
  assert.ok(cli.includes(`sheets: ${report.sheets.join(", ")}`));
  assert.ok(cli.includes(`cells: ${report.cells} (${report.formulas} formulas)`));
  assert.deepEqual(report.unsupported_functions, []);
});

// A workbook made here, with a chain of formulas as long as the stack of the module has to take: the
// engine evaluates a chain recursively when the schedule does not order it (a reference computed at
// run time), which a program runs on a thread of 512 MiB and a module on the stack it was linked with.
function crc32(data) {
  let crc = ~0;
  for (const byte of data) {
    crc ^= byte;
    for (let k = 0; k < 8; k++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return ~crc >>> 0;
}
function zip(entries) {
  const { deflateRawSync } = zlib;
  const parts = [];
  const central = [];
  let offset = 0;
  for (const [name, content] of entries) {
    const data = Buffer.from(content);
    const packed = deflateRawSync(data);
    const nameBytes = Buffer.from(name);
    const header = Buffer.alloc(30);
    header.writeUInt32LE(0x04034b50, 0);
    header.writeUInt16LE(20, 4);
    header.writeUInt16LE(8, 8);
    header.writeUInt32LE(crc32(data), 14);
    header.writeUInt32LE(packed.length, 18);
    header.writeUInt32LE(data.length, 22);
    header.writeUInt16LE(nameBytes.length, 26);
    parts.push(header, nameBytes, packed);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0);
    entry.writeUInt16LE(20, 4);
    entry.writeUInt16LE(20, 6);
    entry.writeUInt16LE(8, 10);
    entry.writeUInt32LE(crc32(data), 16);
    entry.writeUInt32LE(packed.length, 20);
    entry.writeUInt32LE(data.length, 24);
    entry.writeUInt16LE(nameBytes.length, 28);
    entry.writeUInt32LE(offset, 42);
    central.push(entry, nameBytes);
    offset += header.length + nameBytes.length + packed.length;
  }
  const size = central.reduce((n, b) => n + b.length, 0);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(size, 12);
  end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...parts, ...central, end]));
}
function chainWorkbook(length, cell) {
  const xml = (rows) =>
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${rows}`;
  const rows = [`<row r="1"><c r="A1"><v>1</v></c></row>`];
  for (let n = 2; n <= length; n++) {
    rows.push(`<row r="${n}"><c r="A${n}"><f>${cell(n)}</f></c></row>`);
  }
  return zip([
    ["[Content_Types].xml", xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`)],
    ["_rels/.rels", xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`)],
    ["xl/workbook.xml", xml(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Chain" sheetId="1" r:id="rId1"/></sheets></workbook>`)],
    ["xl/_rels/workbook.xml.rels", xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`)],
    ["xl/worksheets/sheet1.xml", xml(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.join("")}</sheetData></worksheet>`)],
  ]);
}
for (const [name, length, cell] of [
  ["direct references", 50000, (n) => `A${n - 1}+1`],
  ["references computed at run time", Number(process.env.RUNTIME_CHAIN ?? 1000), () => `INDIRECT("A"&amp;ROW()-1)+1`],
]) {
  check(`a chain of ${length} formulas with ${name} does not overflow the stack of the module`, () => {
    const chain = new wasm.Book(chainWorkbook(length, cell));
    chain.calculate({ ranges: [`Chain!A${length}`] });
    assert.deepEqual(chain.valueRanges().toJSON().valueRanges, [{ range: `Chain!A${length}`, values: [[length]] }]);
  });
}

// -- compile: the WebAssembly module of a workbook ----------------------------------------------
// The module is the base runtime with the program added; it is instantiated with no imports and its
// exports are called as the README of the package says.
const rt = resolve(process.env.RT ?? join(root, "target/npm-rt"));
function instantiate(module) {
  assert.ok(module instanceof Uint8Array);
  assert.ok(WebAssembly.validate(module), "the module is valid WebAssembly");
  const x = new WebAssembly.Instance(new WebAssembly.Module(module), {}).exports;
  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  const result = () => decoder.decode(new Uint8Array(x.memory.buffer, x.sheetrs_result_ptr(), x.sheetrs_result_len()));
  const call = (fn, request, ...rest) => {
    const bytes = encoder.encode(request);
    const ptr = x.sheetrs_alloc(bytes.length);
    new Uint8Array(x.memory.buffer, ptr, bytes.length).set(bytes);
    const status = fn(ptr, bytes.length, ...rest);
    x.sheetrs_free(ptr, bytes.length);
    return status;
  };
  x.sheetrs_set_clock(Date.now() / 1000);
  return {
    x,
    meta() {
      assert.equal(x.sheetrs_meta(), 0);
      return JSON.parse(result());
    },
    calculate(request) {
      const status = call(x.sheetrs_calculate, JSON.stringify(request));
      return { status, text: result() };
    },
    write(request, withJson) {
      const status = call(x.sheetrs_write, JSON.stringify(request), withJson ? 1 : 0);
      return {
        status,
        text: result(),
        workbook: new Uint8Array(x.memory.buffer, x.sheetrs_workbook_ptr(), x.sheetrs_workbook_len()).slice(),
      };
    },
  };
}
for (const file of [loan, loanOds]) {
  check(`compile (values) ${ext(file)}: the module answers as the executable does`, () => {
    const { module, cells, diagnostics } = open(file).compile();
    assert.ok(cells > 1000 && Array.isArray(diagnostics));
    const book = instantiate(module);
    assert.equal(book.x.sheetrs_abi(), 1);
    assert.equal(book.x.sheetrs_has_writer(), 0);
    const cli = text(run("sheetc", ["calculate", "-i", file, "--request", "-"], JSON.stringify(request))).trim();
    const answer = book.calculate(request);
    assert.equal(answer.status, 0);
    assert.equal(answer.text, cli);
    assert.equal(book.meta().mode, "values");
    // A module that is not given a workbook cannot write one, and says so.
    assert.equal(book.write(request, false).status, 1);
    assert.match(book.write(request, false).text, /cannot write a workbook/);
    // A request the program does not take is an error in the result, not a trap.
    const wrong = book.calculate({ inputs: { bogus: 1 } });
    assert.equal(wrong.status, 1);
    assert.match(wrong.text, /bogus/);
    // The module still answers after a mistake.
    assert.equal(book.calculate(request).text, cli);
  });
  check(`compile (embed_template) ${ext(file)}: the workbook and the JSON from one calculation`, () => {
    const { module } = open(file).compile({ embed_template: true });
    const book = instantiate(module);
    assert.equal(book.x.sheetrs_has_writer(), 1);
    assert.equal(book.meta().mode, "embedded");
    const both = book.write(request, true);
    assert.equal(both.status, 0);
    const direct = calculated(file, request);
    assert.deepEqual(JSON.parse(both.text), direct.answer);
    sameBytes(both.workbook, direct.workbook);
    const only = book.write(request, false);
    assert.equal(only.text, "");
    sameBytes(only.workbook, direct.workbook);
    // And as any executable, the JSON.
    assert.deepEqual(JSON.parse(book.calculate(request).text), direct.answer);
  });
}
check("compile (embed_template, nothing given): the workbook with every formula calculated", () => {
  const { module } = open(loan).compile({ embed_template: true });
  const book = instantiate(module);
  assert.equal(book.meta().mode, "embedded");
  assert.ok(!JSON.stringify(book.meta()).includes("recalculate"));
  const done = book.write({}, false);
  assert.equal(done.status, 0);
  const recalculated = open(loan);
  recalculated.calculate();
  sameValues(done.workbook, recalculated.save(), "Loan!B1:B6");
});
check("compile (declarations and configuration as in the build)", () => {
  const options = {
    declarations: {
      inputs: [{ name: "rate", range: "Loan!B2" }],
      outputs: [{ name: "payment", range: "Loan!B6" }],
    },
  };
  const { module } = open(loan).compile(options);
  const book = instantiate(module);
  const answer = JSON.parse(book.calculate({ inputs: { rate: 0.05 } }).text);
  assert.equal(answer.valueRanges[0].name, "payment");
  assert.equal(typeof answer.valueRanges[0].values[0][0], "number");
  assert.equal(book.calculate({ inputs: { years: 3 } }).status, 1);
});
check("compile (the only target is WebAssembly; json is chosen when the module runs)", () => {
  assert.throws(() => open(loan).compile({ target: "native" }), /one target here, WebAssembly/);
  assert.throws(() => open(loan).compile({ target: "wasm" }), /one target here/);
  assert.throws(() => open(loan).compile({ json: true }), /chosen when the module runs/);
});
check("compile (sheetc build --target wasm makes a module that answers the same)", () => {
  const out = join(work, "loan.wasm");
  run("sheetc", ["build", "-i", loan, "--target", "wasm", "--runtime", join(rt, "sheetrs-rt.wasm"), "-o", out]);
  const fromCli = instantiate(new Uint8Array(readFileSync(out)));
  const { module } = open(loan).compile();
  assert.equal(fromCli.calculate(request).text, instantiate(module).calculate(request).text);
  const outTemplate = join(work, "loan_t.wasm");
  run("sheetc", ["build", "-i", loan, "--target", "wasm", "--embed-template", "--runtime", join(rt, "sheetrs-rt-template.wasm"), "-o", outTemplate]);
  sameBytes(instantiate(new Uint8Array(readFileSync(outTemplate))).write(request, false).workbook, calculated(loan, request).workbook);
});
check("compile (the module has no clock of its own: the host sets it)", () => {
  const nowBook = chainWorkbook(2, () => "NOW()");
  const { module } = new wasm.Book(nowBook).compile();
  const book = instantiate(module);
  book.x.sheetrs_set_clock(1_700_000_000);
  const answer = JSON.parse(book.calculate({ ranges: ["Chain!A2"] }).text);
  assert.ok(Math.abs(answer.valueRanges[0].values[0][0] - (25569 + 1_700_000_000 / 86400)) < 1e-6);
});
check("compile (a big program: the memory the module grows never overlaps its program)", () => {
  const { module } = new wasm.Book(chainWorkbook(30000, (n) => `A${n - 1}+1`)).compile();
  const book = instantiate(module);
  const first = book.calculate({ ranges: ["Chain!A2:A30000"] }).text;
  const again = book.calculate({ ranges: ["Chain!A2:A30000"] }).text;
  assert.equal(first, again);
  assert.equal(JSON.parse(first).valueRanges[0].values[29998][0], 30000);
});

rmSync(work, { recursive: true, force: true });
console.log(`${checks} checks passed`);
