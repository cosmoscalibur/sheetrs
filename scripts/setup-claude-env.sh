#!/usr/bin/env bash
# Setup script of the Claude Code environment for SheetRS. It is self-contained and takes no
# parameters and no path, because the environment settings give a text box to paste a script into (or
# a file to upload), not a command line: paste this whole file there.
#
# What it leaves, as root, on a Debian or Ubuntu image:
#   - Rust 1.99 (keep CHANNEL equal to the one of rust-toolchain.toml), with clippy, rustfmt and the
#     wasm32 target, as the default toolchain, so that it is the one used in any folder;
#   - LibreOffice 26.8.2.2 (the reference of the project) from the packages of The Document Foundation,
#     checksum verified, `soffice` on the PATH, and the LibreOffice of the image removed;
#   - `perf` and `valgrind`, to profile (`scripts/bench.py`, `docs/performance.md`).
# Idempotent: what is already in place is left alone. The only domain that must be allowed in the network access
# of the environment besides the package managers is downloadarchive.documentfoundation.org (docs/environment.md).
set -euo pipefail

CHANNEL=1.99
LO_VERSION=26.8.2.2
LO_SERIES=26.8

# The proxy of the environment signs with its own CA: use it when there is one.
CURL=(curl -fsSL)
[ -f /root/.ccr/ca-bundle.crt ] && CURL+=(--cacert /root/.ccr/ca-bundle.crt)

# ---- Rust -------------------------------------------------------------------------------------
if ! command -v rustup >/dev/null; then
  "${CURL[@]}" https://sh.rustup.rs | sh -s -- -y --default-toolchain none
  export PATH="$HOME/.cargo/bin:$PATH"
fi
rustup toolchain install "$CHANNEL" --profile minimal -c clippy -c rustfmt -t wasm32-unknown-unknown
rustup default "$CHANNEL"
echo "rust: $(rustc --version)"

# ---- Profiling tools ---------------------------------------------------------------------------
if ! command -v perf >/dev/null || ! command -v valgrind >/dev/null; then
  apt-get update -qq
  apt-get install -y linux-tools-common linux-tools-generic valgrind
fi
if ! command -v perf >/dev/null; then
  # The wrapper of Ubuntu wants the tools of the running kernel; use the ones that were installed.
  ln -sf "$(ls /usr/lib/linux-tools/*/perf | sort -V | tail -1)" /usr/local/bin/perf
fi
echo "perf: $(perf --version 2>&1 | head -1); valgrind: $(valgrind --version)"

# ---- LibreOffice ------------------------------------------------------------------------------
if [ ! -x "/opt/libreoffice${LO_SERIES}/program/soffice" ]; then
  work="$(mktemp -d)"
  base="https://downloadarchive.documentfoundation.org/libreoffice/old/${LO_VERSION}/deb/x86_64"
  tar_name="LibreOffice_${LO_VERSION}_Linux_x86-64_deb.tar.gz"
  "${CURL[@]}" -o "$work/$tar_name" "$base/$tar_name"
  "${CURL[@]}" -o "$work/$tar_name.sha256" "$base/$tar_name.sha256"
  (cd "$work" && sha256sum -c "$tar_name.sha256")
  mkdir "$work/x" && tar -xzf "$work/$tar_name" -C "$work/x"
  # Without the integrations with the desktop: a server needs none of them.
  # shellcheck disable=SC2046
  dpkg -i $(ls "$work"/x/*/DEBS/*.deb | grep -v -E 'desktop-integration|gnome|kde|onlineupdate')
  rm -rf "$work"
fi
apt-get remove -y --purge 'libreoffice-*' python3-uno || true
rm -rf /usr/lib/libreoffice
ln -sf "/opt/libreoffice${LO_SERIES}/program/soffice" /usr/local/bin/soffice
echo "libreoffice: $(soffice --version)"
