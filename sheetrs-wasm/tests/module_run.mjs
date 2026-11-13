// Runs the WebAssembly module of a workbook once, as a whole process (the cost of a command: Node
// starts, compiles and instantiates the module, calculates, and writes the answer to the standard
// output), to time it against the native executable of the same workbook:
//
//   node sheetrs-wasm/tests/module_run.mjs MODULE.wasm '{"ranges":["Summary!A1"]}'
import { readFileSync } from "node:fs";

const [file, request = "{}"] = process.argv.slice(2);
const x = new WebAssembly.Instance(new WebAssembly.Module(readFileSync(file)), {}).exports;
const bytes = new TextEncoder().encode(request);
const ptr = x.sheetrs_alloc(bytes.length);
new Uint8Array(x.memory.buffer, ptr, bytes.length).set(bytes);
const status = x.sheetrs_calculate(ptr, bytes.length);
x.sheetrs_free(ptr, bytes.length);
const answer = Buffer.from(x.memory.buffer, x.sheetrs_result_ptr(), x.sheetrs_result_len());
process.stdout.write(answer);
process.exit(status);
