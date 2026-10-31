#!/usr/bin/env bash
# Builds the npm package of `sheetrs-wasm` into target/npm/ and packs it (target/npm/*.tgz).
#
#   scripts/package-wasm.sh
#
# Needs the wasm32-unknown-unknown target (rust-toolchain.toml installs it) and `wasm-bindgen`
# CLI at the version `Cargo.lock` pins for the `wasm-bindgen` crate:
#   cargo install wasm-bindgen-cli --version <that version> --locked
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root"

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

cargo build -p sheetrs-wasm --target wasm32-unknown-unknown --release --locked
wasm-bindgen --target web --out-dir "$out" --out-name sheetrs_wasm \
  target/wasm32-unknown-unknown/release/sheetrs_wasm.wasm

sed "s/\"version\": \"0.0.0\"/\"version\": \"$version\"/" sheetrs-wasm/npm/package.json > "$out/package.json"
cp sheetrs-wasm/README.md LICENSE "$out/"
(cd "$out" && npm pack --silent)
ls -1 "$out"/*.tgz
