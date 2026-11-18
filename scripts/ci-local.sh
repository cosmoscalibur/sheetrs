#!/usr/bin/env bash
# Runs the CI checks in a Linux container. The working tree is mounted
# read-only; build output and caches live under ~/.cache/sheetrs-ci.
set -euo pipefail

cd "$(dirname "$0")/.."

image=sheetrs-ci:1.99
cache="${XDG_CACHE_HOME:-$HOME/.cache}/sheetrs-ci"
mkdir -p "$cache/cargo" "$cache/target"

docker build --quiet --tag "$image" --file ci/linux.Dockerfile ci >/dev/null

docker run --rm --user "$(id -u):$(id -g)" \
  --volume "$PWD":/work:ro \
  --volume "$cache/cargo":/cache/cargo \
  --volume "$cache/target":/cache/target \
  --env CARGO_HOME=/cache/cargo \
  --env CARGO_TARGET_DIR=/cache/target \
  --env HOME=/tmp \
  --workdir /work \
  "$image" bash -euxc '
    cargo fmt --all --check
    cargo clippy --locked --workspace --all-targets -- -D warnings
    cargo test --locked --workspace
    cargo check --locked --workspace --target x86_64-pc-windows-msvc
    cargo check --locked --workspace --target aarch64-apple-darwin
  '
