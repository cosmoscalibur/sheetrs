// A test of the web demo (docs/index.html) in a real browser: Chromium drives every tool on real
// workbooks, the answers are checked, and screenshots of the page are taken (desktop and phone).
//
//   scripts/package-wasm.sh --demo && node sheetrs-wasm/tests/demo.mjs
//
// Needs `playwright` (resolvable from here) with Chromium. Screenshots go to target/demo-screenshots/
// (or $SCREENSHOTS). Exit code 0 when everything passes.

import { createServer } from "node:http";
import { createRequire } from "node:module";
import { mkdirSync, readFileSync, existsSync, statSync } from "node:fs";
import { extname, join, resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import assert from "node:assert/strict";

const require = createRequire(import.meta.url);
const { chromium } = require("playwright");

const root = resolve(dirname(fileURLToPath(import.meta.url)), "../..");
const docs = join(root, "docs");
const shots = resolve(process.env.SCREENSHOTS ?? join(root, "target/demo-screenshots"));
mkdirSync(shots, { recursive: true });
assert.ok(existsSync(join(docs, "pkg/sheetrs_wasm.js")), "docs/pkg is missing: run scripts/package-wasm.sh --demo");

const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".wasm": "application/wasm" };
const server = createServer((request, response) => {
  const path = join(docs, decodeURIComponent(new URL(request.url, "http://x").pathname).replace(/^\/$/, "/index.html"));
  if (!path.startsWith(docs) || !existsSync(path) || !statSync(path).isFile()) {
    response.writeHead(404).end();
    return;
  }
  response.writeHead(200, { "content-type": types[extname(path)] ?? "application/octet-stream" });
  response.end(readFileSync(path));
});
await new Promise((done) => server.listen(0, "127.0.0.1", done));
const url = `http://127.0.0.1:${server.address().port}/`;

const lintBook = join(root, "tests/lint_book.xlsx");
const loan = join(root, "sheetrs-compiler/tests/assets/loan.xlsx");
const loanOds = join(root, "sheetrs-compiler/tests/assets/loan.ods");

const browser = await chromium.launch();
const errors = [];
let checks = 0;

async function open(viewport, scale) {
  const context = await browser.newContext({ viewport, deviceScaleFactor: scale, acceptDownloads: true });
  const page = await context.newPage();
  page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
  page.on("console", (message) => {
    if (message.type() === "error") errors.push(`console: ${message.text()}`);
  });
  await page.goto(url);
  await page.waitForFunction(() => document.getElementById("status").textContent === "WASM Loaded");
  return { page, context };
}

const ok = (name) => {
  checks += 1;
  console.log(`ok   ${name}`);
};

async function load(page, file) {
  await page.setInputFiles("#file-input", file);
  await page.waitForFunction(() => document.getElementById("status").textContent.startsWith("✓ File loaded"));
}
async function runTool(page) {
  await page.click("#play-btn");
  await page.waitForFunction(() => /^(✓ Success|❌)/.test(document.getElementById("status").textContent), null, { timeout: 120000 });
}
const status = (page) => page.textContent("#status");
const out = (page) => page.textContent("#output");
async function tool(page, name) {
  await page.click(`.tool-btn[data-tool="${name}"]`);
  await page.waitForSelector(`#${name}-setup`, { state: "visible" });
}
const shot = (page, name) => page.screenshot({ path: join(shots, `${name}.png`) });
async function download(page, selector) {
  const [file] = await Promise.all([page.waitForEvent("download"), page.click(selector)]);
  const path = await file.path();
  return { name: file.suggestedFilename(), bytes: readFileSync(path) };
}

// -- desktop -------------------------------------------------------------------------------------
const { page, context } = await open({ width: 1280, height: 800 }, 2);

// The evaluate tool needs no workbook.
await tool(page, "evaluate");
assert.equal(await page.isDisabled("#play-btn"), false, "evaluate runs with no file");
await runTool(page);
assert.match(await out(page), /^6/);
await page.fill("#evaluate-formula", '=TEXT(0.25,"0%")');
await runTool(page);
assert.match(await out(page), /"25%"/);
await page.fill("#evaluate-formula", "=1/0");
await runTool(page);
assert.match(await out(page), /#DIV\/0!/);
await shot(page, "desktop-evaluate");
ok("evaluate (no workbook)");

await tool(page, "lint");
assert.equal(await page.isDisabled("#play-btn"), true, "a tool that reads a workbook needs one");
await load(page, lintBook);
await runTool(page);
assert.match(await out(page), /Summary/);
ok("lint (report)");
await page.selectOption("#lint-format", "table");
await runTool(page);
assert.ok((await page.locator("#table-output tbody tr").count()) > 20);
assert.match(await page.textContent("#table-output"), /CALC|ERR|DATA|REF/);
await shot(page, "desktop-lint-table");
ok("lint (table)");
await page.selectOption("#lint-format", "json");
await runTool(page);
assert.match(await out(page), /^JSON: /);
const json = await download(page, ".download-link");
assert.match(json.name, /\.json$/);
const doc = JSON.parse(json.bytes.toString("utf8"));
assert.ok(Array.isArray(doc.violations) && doc.violations.length > 0 && doc.summary);
await shot(page, "desktop-lint-json");
ok("lint (JSON, with the file to download)");
await page.selectOption("#lint-format", "sarif");
await runTool(page);
assert.match(await out(page), /^SARIF 2\.1\.0: \d+ results/);
const sarif = await download(page, ".download-link");
assert.match(sarif.name, /\.sarif$/);
assert.equal(JSON.parse(sarif.bytes.toString("utf8")).version, "2.1.0");
await shot(page, "desktop-lint-sarif");
ok("lint (SARIF, with the file to download)");

await tool(page, "stats");
await runTool(page);
assert.match(await out(page), /sheet/i);
await page.check("#stats-json");
await runTool(page);
assert.ok(JSON.parse((await out(page)).split("\n━")[0]));
ok("stats (text and JSON)");

await load(page, loan);
await tool(page, "query");
await page.fill("#query-scope", "Loan!A1:B3");
await runTool(page);
assert.match(await page.textContent("#table-output"), /Principal/);
assert.match(await out(page), /^6 cells/);
await shot(page, "desktop-query");
ok("query (a range)");
await page.selectOption("#query-scope-kind", "workbook");
await page.fill("#query-find", "Payment");
await runTool(page);
assert.match(await page.textContent("#table-output"), /Monthly payment/);
ok("query (find)");

await tool(page, "edit");
await runTool(page);
assert.match(await out(page), /the plan is empty/);
await page.selectOption("#edit-op", "delete_rows");
await page.fill("#edit-arg", "Loan!10:11");
await page.click("#edit-add");
await page.selectOption("#edit-op", "formula_to_value");
await page.fill("#edit-arg", "Loan!B:B");
await page.click("#edit-add");
await page.selectOption("#edit-op", "replace");
await page.fill("#edit-arg", "Monthly");
await page.fill("#edit-with", "Per month");
await page.click("#edit-add");
assert.equal(await page.locator("#edit-plan li").count(), 3);
await runTool(page);
assert.match(await out(page), /delete rows/);
const edited = await download(page, ".download-link");
assert.match(edited.name, /\.edited\.xlsx$/);
assert.equal(edited.bytes.subarray(0, 2).toString(), "PK");
await shot(page, "desktop-edit");
ok("edit (a plan of three operations, the edited workbook to download)");

await tool(page, "calculate");
await page.click("#calculate-read");
await page.waitForSelector("#calculate-inputs input[data-input]");
const names = await page.locator("#calculate-inputs input[data-input]").evaluateAll((fields) => fields.map((f) => f.dataset.input));
assert.deepEqual(names.sort(), ["month", "principal", "rate", "years"]);
await page.fill('#calculate-inputs input[data-input="principal"]', "250000");
await page.fill('#calculate-inputs input[data-input="rate"]', "0.045");
await page.fill('#calculate-inputs input[data-input="years"]', "15");
await page.fill('#calculate-inputs input[data-input="month"]', "24");
await runTool(page);
const answer = JSON.parse((await out(page)).split("\n━")[0]);
const payment = answer.valueRanges.find((entry) => entry.name === "payment").values[0][0];
assert.ok(Math.abs(payment - 1912.4832220336577) < 1e-6, JSON.stringify(answer));
await shot(page, "desktop-calculate");
ok("calculate (the inputs of the model as a form)");
await page.selectOption("#calculate-mode", "embed_template");
await runTool(page);
assert.match(await out(page), /The workbook with the results/);
const calculated = await download(page, ".download-link");
assert.match(calculated.name, /\.calculated\.xlsx$/);
assert.equal(calculated.bytes.subarray(0, 2).toString(), "PK");
ok("calculate (the workbook with the answer)");

// -- compile: the module of the workbook, run in the page ---------------------------------------------
await tool(page, "compile");
await runTool(page);
assert.match(await out(page), /loan\.wasm: [\d.]+ (KB|MB)/);
assert.match(await out(page), /"mode": "values"/);
const module = await download(page, ".download-link");
assert.match(module.name, /\.wasm$/);
assert.ok(await WebAssembly.validate(module.bytes), "the module that was downloaded is valid");
assert.equal(await page.isVisible("#compile-run"), true);
await page.fill("#compile-request", JSON.stringify({ inputs: { principal: 250000, rate: 0.045, years: 15, month: 24 } }));
await page.click("#compile-run-btn");
await page.waitForFunction(() => document.getElementById("status").textContent === "✓ The module ran");
assert.match(await out(page), /"name":"payment","range":"Loan!B4","values":\[\[1912\.48322203365/);
await shot(page, "desktop-compile");
ok("compile (values: the module answers in the page)");
await page.selectOption("#compile-mode", "embed_template");
await runTool(page);
assert.match(await out(page), /"mode": "embedded"/);
await page.click("#compile-run-btn");
await page.waitForFunction(() => document.getElementById("status").textContent === "✓ The module ran");
assert.match(await out(page), /bytes of workbook/);
const written = await download(page, ".answer-link");
assert.match(written.name, /\.answer\.xlsx$/);
assert.equal(written.bytes.subarray(0, 2).toString(), "PK");
ok("compile (embedded: the module writes the workbook in the page)");
await page.selectOption("#compile-mode", "values");
await page.fill("#compile-declarations", JSON.stringify({ inputs: [{ name: "rate", range: "Loan!B2" }], outputs: [{ name: "payment", range: "Loan!B6" }] }));
await runTool(page);
await page.fill("#compile-request", '{"inputs":{"rate":0.05}}');
await page.click("#compile-run-btn");
await page.waitForFunction(() => document.getElementById("status").textContent === "✓ The module ran");
assert.match(await out(page), /"payment"/);
await page.fill("#compile-request", '{"inputs":{"bogus":1}}');
await page.click("#compile-run-btn");
await page.waitForFunction(() => document.getElementById("status").textContent.startsWith("❌"));
assert.match(await out(page), /bogus/);
ok("compile (declarations, and a mistake of the request is said, not a crash)");

await load(page, loanOds);
await tool(page, "compile");
await runTool(page);
assert.match(await out(page), /loan\.wasm/);
ok("compile (an ODS)");
await context.close();

// -- a phone ---------------------------------------------------------------------------------------------
const phone = await open({ width: 390, height: 844 }, 2);
await load(phone.page, lintBook);
await phone.page.selectOption("#lint-format", "table");
await runTool(phone.page);
const overflow = await phone.page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
assert.ok(overflow <= 1, `the page does not scroll sideways on a phone (${overflow}px)`);
await shot(phone.page, "phone-lint");
await tool(phone.page, "query");
await phone.page.fill("#query-scope", "Data!A1:C5");
await load(phone.page, loan);
await phone.page.fill("#query-scope", "Loan!A1:B3");
await runTool(phone.page);
await shot(phone.page, "phone-query");
await tool(phone.page, "compile");
await runTool(phone.page);
await shot(phone.page, "phone-compile");
const overflowCompile = await phone.page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
assert.ok(overflowCompile <= 1, `the compile tool does not scroll sideways on a phone (${overflowCompile}px)`);
ok("a phone: the layout stacks and the page does not scroll sideways");
await phone.context.close();

await browser.close();
server.close();
assert.deepEqual(errors, [], "no error in the console of the page");
console.log(`${checks} checks passed; screenshots in ${shots}`);
