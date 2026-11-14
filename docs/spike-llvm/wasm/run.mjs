import fs from 'fs';
const m = await WebAssembly.compile(fs.readFileSync('target/wasm32-unknown-unknown/release/wk.wasm'));
console.log('imports:', WebAssembly.Module.imports(m).length);
const t0 = performance.now(); const i = await WebAssembly.instantiate(m, {}); const x = i.exports;
x.setup(300000);
const best = (f) => { let b = 1e9, r; for (let k = 0; k < 7; k++) { const t = performance.now(); r = f(); b = Math.min(b, performance.now() - t); } return [b, r]; };
for (const [name, f] of [['LLVM exact round (wasm obj)', () => x.run_llvm(0.05, 1)], ['rustc exact round', () => x.run_rust(0.05, 1)], ['LLVM approx round', () => x.run_llvm(0.05, 0)], ['rustc approx round', () => x.run_rust(0.05, 0)]]) {
  const [b, r] = best(f); console.log(name.padEnd(30), b.toFixed(2), 'ms  checksum', r);
}
