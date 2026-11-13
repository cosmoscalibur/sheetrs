#!/usr/bin/env bash
# Builds the npm package of `sheetrs-wasm` into target/npm/ and packs it (target/npm/*.tgz).
#
#   scripts/package-wasm.sh [--demo]
#
#   --demo   also copies the package into docs/pkg/, where the demo page (docs/index.html) reads it.
#
# Needs the wasm32-unknown-unknown target (rust-toolchain.toml installs it) and `wasm-bindgen`
# CLI at the version `Cargo.lock` pins for the `wasm-bindgen` crate:
#   cargo install wasm-bindgen-cli --version <that version> --locked
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root"

demo=false
for arg in "$@"; do
  case "$arg" in
    --demo) demo=true ;;
    *) echo "unknown option $arg" >&2; exit 2 ;;
  esac
done

version="$(cargo metadata --no-deps --format-version 1 \
  | python3 -c 'import json,sys; print(next(p["version"] for p in json.load(sys.stdin)["packages"] if p["name"]=="sheetrs-wasm"))')"
locked="$(awk '/^name = "wasm-bindgen"$/ {getline; gsub(/[^0-9.]/, "", $0); print; exit}' Cargo.lock)"
installed="$(wasm-bindgen --version 2>/dev/null | awk '{print $2}' || true)"
if [ "$installed" != "$locked" ]; then
  echo "wasm-bindgen CLI is '${installed:-missing}', Cargo.lock pins $locked:" >&2
  echo "  cargo install wasm-bindgen-cli --version $locked --locked" >&2
  exit 1
fi

out="target/npm"
rm -rf "$out"
mkdir -p "$out"

# The base modules of the WebAssembly target go inside the package (`Book.compile` completes one
# with the program of a workbook): they are built first, in the same profile.
rt="$root/target/npm-rt"
scripts/build-wasm-rt.sh --profile dist --out "$rt"
SHEETRS_RT_DIR="$rt" cargo build -p sheetrs-wasm --target wasm32-unknown-unknown --release --locked \
  --features bundled-runtime
wasm-bindgen --target web --out-dir "$out" --out-name sheetrs_wasm \
  target/wasm32-unknown-unknown/release/sheetrs_wasm.wasm
# The objects release their memory by themselves when nothing references them (the `FinalizationRegistry`
# that wasm-bindgen emits; `free()` stays for releasing at once): a package without it would leak.
grep -q FinalizationRegistry "$out/sheetrs_wasm.js" || {
  echo "the package does not release its objects by itself (no FinalizationRegistry): check the wasm-bindgen version" >&2
  exit 1
}

sed "s/\"version\": \"0.0.0\"/\"version\": \"$version\"/" sheetrs-wasm/npm/package.json > "$out/package.json"
cp sheetrs-wasm/README.md LICENSE "$out/"
(cd "$out" && npm pack --silent)
ls -1 "$out"/*.tgz
if $demo; then
  rm -rf docs/pkg
  mkdir -p docs/pkg
  cp "$out"/sheetrs_wasm.js "$out"/sheetrs_wasm_bg.wasm docs/pkg/
  echo "demo package in docs/pkg"
fi
