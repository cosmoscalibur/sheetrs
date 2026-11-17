# LLVM spike: prototypes

Throwaway code behind [docs/spike-execution-backends.md](../spike-execution-backends.md). None of it is part
of the workspace or of a release. `rd.rs` is a copy of `round_dec` from
`sheetrs-engine/src/functions/math.rs` (the project's exact rounding), taken as it was when the spike ran.

System packages (Ubuntu 24.04; `apt.llvm.org` for 23, which is the LLVM of `rustc` 1.99):
`llvm-23-dev libpolly-23-dev libclang-common-23-dev`, and `wasm-ld`; Node 22 for the wasm run; the target
`wasm32-unknown-unknown` for rustup.

```bash
# kernel, latency at scale, Cranelift loop, bitcode link (LLVM 23 and the same kernel in Cranelift)
cd inkwell && LLVM_SYS_231_PREFIX=/usr/lib/llvm-23 cargo build --release
./target/release/ll kernel   # LLVM O0/O2/O3 against the Rust reference, exact and approximate rounding
./target/release/ll scale    # N distinct formulas: LLVM O0/O1/O2 and Cranelift, microseconds per formula
./target/release/ll cl       # the same kernel in Cranelift
(cd ../bitcode && rustc --edition 2024 --crate-type=lib --emit=llvm-bc -C opt-level=3 -C panic=abort rt.rs -o rt.bc)
./target/release/ll bc       # rustc bitcode linked into the generated module (needs ../bitcode/rt.bc)

# the same kernel as a WebAssembly object, linked with the Rust runtime, run in Node (no imports)
./target/release/ll wasmobj  # writes kernel_e.o and kernel_a.o
cd ../wasm && RUSTFLAGS="-C link-arg=$PWD/../inkwell/kernel_e.o -C link-arg=$PWD/../inkwell/kernel_a.o" \
  cargo build --release --target wasm32-unknown-unknown && node run.mjs
```

With LLVM 20, change the feature to `llvm20-1` and `LLVM_SYS_201_PREFIX`: the results are the same, but
`ll bc` fails because LLVM 20 cannot read the bitcode of `rustc` 1.99.
