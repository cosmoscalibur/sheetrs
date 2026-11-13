#!/usr/bin/env bash
# Builds the two base modules of the WebAssembly target of `sheetc` (`sheetc build --target wasm`):
#   sheetrs-rt.wasm           the runtime with the engine: a workbook that answers JSON
#   sheetrs-rt-template.wasm  the same, with the readers and writers of XLSX and ODS: a workbook that
#                             answers with a workbook (`--embed-template`)
# and puts them next to `sheetc`, where it looks for them (nothing replaces them if they are missing).
#
#   scripts/build-wasm-rt.sh [--profile dist|release|dev] [--out DIR]
#
# Needs the wasm32-unknown-unknown target (rust-toolchain.toml installs it).
set -euo pipefail

root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root"

profile=dist
out=""
while [ $# -gt 0 ]; do
  case "$1" in
    --profile) profile="$2"; shift ;;
    --out) out="$2"; shift ;;
    *) echo "unknown option $1" >&2; exit 2 ;;
  esac
  shift
done
dir="$profile"
[ "$profile" = dev ] && dir=debug
out="${out:-$root/target/$dir}"
mkdir -p "$out"

built="target/wasm32-unknown-unknown/$dir/sheetrs_rt_wasm.wasm"
cargo build -p sheetrs-rt-wasm --target wasm32-unknown-unknown --profile "$profile" --locked
cp "$built" "$out/sheetrs-rt.wasm"
cargo build -p sheetrs-rt-wasm --target wasm32-unknown-unknown --profile "$profile" --locked --features template
cp "$built" "$out/sheetrs-rt-template.wasm"
ls -l "$out"/sheetrs-rt.wasm "$out"/sheetrs-rt-template.wasm
