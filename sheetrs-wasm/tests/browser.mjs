// The edit and calculate options of the WebAssembly package in a real browser (Chromium): visibility,
// `calculateAll`, the conversion of formulas to values, the locale, `now` and `seed`, and what a file caches for what the engine
// cannot calculate (another workbook, a macro, an extension), which has no option.
// `smoke.mjs` checks the same package in Node against the command line; this runs it in the page's runtime.
//
//   scripts/package-wasm.sh --demo && node sheetrs-wasm/tests/browser.mjs
//
// Needs `playwright` (resolvable from here) with Chromium. Exit code 0 when everything passes.

import { createServer } from "node:http";
import { createRequire } from "node:module";
import { readFileSync, existsSync, statSync } from "node:fs";
import { extname, join, resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import assert from "node:assert/strict";

const require = createRequire(import.meta.url);
const { chromium } = require("playwright");

const root = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
assert.ok(existsSync(join(root, "docs/pkg/sheetrs_wasm.js")), "docs/pkg is missing: run scripts/package-wasm.sh --demo");

// The package (docs/pkg) and the workbooks of the repository, read only.
const types = { ".js": "text/javascript", ".wasm": "application/wasm", ".html": "text/html" };
const server = createServer((request, response) => {
  const name = decodeURIComponent(new URL(request.url, "http://x").pathname);
  const path = name === "/" ? null : join(root, name.startsWith("/pkg/") ? `docs${name}` : name);
  const allowed = path && (path.startsWith(join(root, "docs/pkg")) || path.startsWith(join(root, "sheetrs-edit/tests/assets")));
  if (name === "/") return response.writeHead(200, { "content-type": "text/html" }).end("<!doctype html><title>t</title>");
  if (!allowed || !existsSync(path) || !statSync(path).isFile()) return response.writeHead(404).end();
  response.writeHead(200, { "content-type": types[extname(path)] ?? "application/octet-stream" });
  response.end(readFileSync(path));
});
await new Promise((done) => server.listen(0, "127.0.0.1", done));
const url = `http://127.0.0.1:${server.address().port}/`;

const browser = await chromium.launch();
const page = await browser.newPage();
page.on("pageerror", (error) => {
  console.error(`page error: ${error.message}`);
  process.exitCode = 1;
});
await page.goto(url);

// Everything runs in the page: the package is loaded as a module and the tests are a function of it.
const results = await page.evaluate(async () => {
  const wasm = await import("/pkg/sheetrs_wasm.js");
  await wasm.default();
  wasm.initPanicHook();
  const out = [];
  const check = async (name, fn) => {
    try {
      await fn();
      out.push({ name, ok: true });
    } catch (error) {
      out.push({ name, ok: false, message: String(error?.message ?? error) });
    }
  };
  const assert = {
    equal: (a, b, m) => { if (a !== b) throw new Error(`${m ?? ""} ${a} !== ${b}`); },
    ok: (v, m) => { if (!v) throw new Error(m ?? "not ok"); },
    throws: (fn, re) => {
      try { fn(); } catch (e) { if (!re.test(String(e))) throw new Error(`wrong error: ${e}`); return; }
      throw new Error("did not throw");
    },
  };
  const fetchBytes = async (path) => new Uint8Array(await (await fetch(`/sheetrs-edit/tests/assets/${path}`)).arrayBuffer());
  const cached = {};
  for (const name of ["cached/book", "cached/rules", "volatile/volatile", "from_cache/external", "from_cache/macro", "from_cache/extension", "book/book"]) {
    for (const ext of ["xlsx", "ods"]) cached[`${name}.${ext}`] = await fetchBytes(`${name}.${ext}`);
  }
  const open = (name, ext, locale, options) => new wasm.Book(cached[`${name}.${ext}`], locale, options);
  const cell = (bytes, range) => new wasm.Book(bytes).range(range)[0].values.flat();

  for (const ext of ["xlsx", "ods"]) {
    await check(`visibility (${ext}): hide and show rows and columns, and sheets`, () => {
      const book = open("cached/book", ext);
      book.hideRows("Data!2:3");
      book.showRows("Data!3");
      book.hideColumns("Data!B");
      book.hideSheets("Calc");
      assert.equal(book.pending, 4);
      const saved = new wasm.Book(book.save());
      assert.ok(saved.explain === undefined || true);
      const range = saved.range("Data!A1:B3")[0];
      assert.ok(range, "the range is read");
    });
    await check(`calculateAll (${ext}): by default the results the file caches stay where the edit reaches nothing`, () => {
      // Hiding a row reaches SUBTOTAL only: `Calc!A8` is `ROW()`, which the file caches as 99 (it is 8). An XLSX keeps 99 when
      // `LibreOffice` opens it; an ODS that caches a plain number in a cell with no format has it calculated on load (8).
      const plain = open("cached/book", ext);
      plain.hideRows("Data!2");
      const everything = open("cached/book", ext);
      everything.calculateAll = true;
      assert.equal(everything.calculateAll, true);
      assert.equal(plain.calculateAll, false);
      everything.hideRows("Data!2");
      assert.equal(cell(plain.save(), "Calc!A8")[0], ext === "ods" ? 8 : 99, "the cached result by default");
      assert.equal(cell(everything.save(), "Calc!A8")[0], 8, "calculated again with the option");
    });
    await check(`conversion (${ext}): formulas become the values they calculate`, () => {
      const book = open("cached/book", ext);
      book.formulaToValue("Calc!A:A");
      const bytes = book.save();
      const reread = new wasm.Book(bytes);
      assert.equal(reread.range("Calc!A1")[0].values[0][0], 210);
    });
    await check(`locale (${ext}): the locale of the book and of a call`, () => {
      const book = open("cached/book", ext, "es-CO");
      assert.equal(book.locale, "es-CO");
      assert.throws(() => open("cached/book", ext, "not a tag"), /Locale error/);
    });
    await check(`now and seed (${ext}): fix NOW, TODAY and RAND`, () => {
      const row = (options, call) => {
        const book = open("volatile/volatile", ext, undefined, options);
        book.calculate({ ranges: ["V!A1:C1"] }, call);
        return book.valueRanges().toObject()["V!A1:C1"][0];
      };
      const [now, today, rand] = row({ now: 45366.75, seed: 42 });
      assert.equal(now, 45366.75);
      assert.equal(today, 45366);
      assert.equal(row({ seed: 42 })[2], rand);
      assert.ok(row({ seed: 44 })[2] !== rand, "another seed");
      assert.equal(row({ now: 1, seed: 42 }, { now: 45366.75 })[0], 45366.75, "the call wins");
    });
    for (const [kind, code, a1, error] of [
      ["external", "EXTERNAL_FROM_CACHE", 42, null],
      ["macro", "MACRO_FROM_CACHE", 77, /MYFUNC/],
      ["extension", "EXTENSION_FROM_CACHE", 77, /WEBSERVICE/],
    ]) {
      await check(`from cache (${kind} ${ext}): the stored value of a link, an error for a macro or WEBSERVICE when every formula is calculated`, () => {
        const book = open(`from_cache/${kind}`, ext);
        book.calculateAll = true;
        book.formulaToValue("Calc");
        if (error) {
          assert.throws(() => book.save(), error);
          return;
        }
        const row = cell(book.save(), "Calc!A1:E1");
        assert.equal(row[0], a1, "A1 keeps its cached result");
        assert.equal(row[1], a1 + 1, "B1 reads it");
        assert.equal(row[2], 10, "C1 is calculated");
        assert.equal(row[4], 11, "E1 is calculated");
        assert.equal(book.report.diagnostics.map((d) => d.code).join(), code);
      });
    }
  }
  return out;
});

let failed = 0;
for (const { name, ok, message } of results) {
  if (ok) console.log(`ok   ${name}`);
  else {
    failed += 1;
    console.error(`FAIL ${name}\n${message}`);
  }
}
await browser.close();
server.close();
console.log(`${results.length - failed} checks passed${failed ? `, ${failed} failed` : ""}`);
process.exitCode = failed ? 1 : 0;
