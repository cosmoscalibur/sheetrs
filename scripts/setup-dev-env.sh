#!/usr/bin/env bash
# Prepares a Debian/Ubuntu machine (root) to develop and test SheetRS:
#   - Rust: the channel of rust-toolchain.toml (1.99), with clippy, rustfmt and the wasm32 target,
#     made the default toolchain so that it is used from the first command, in any folder;
#   - LibreOffice 26.8 (the reference of the project), from the packages of The Document Foundation,
#     with the checksum verified, and `soffice` on the PATH. A LibreOffice of another version from the
#     system packages is removed (`--replace-system`);
#   - profiling tools: `perf` and `valgrind` (callgrind), for docs/performance.md and the profiles
#     of sheetrs-linter/docs/PROFILING.md. A virtual machine often has no performance counters: then `perf
#     stat -e instructions` says "not supported", and `perf record -e cpu-clock` (sampling) and
#     `valgrind --tool=callgrind` (instruction counts) are what is left.
# Idempotent: what is already in place is left alone.
#
#   scripts/setup-dev-env.sh [--replace-system] [--libreoffice 26.8.2]
set -euo pipefail

LO_VERSION=26.8.2
REPLACE_SYSTEM=false
while [ $# -gt 0 ]; do
  case "$1" in
    --replace-system) REPLACE_SYSTEM=true ;;
    --libreoffice) LO_VERSION="$2"; shift ;;
    *) echo "unknown option $1" >&2; exit 2 ;;
  esac
  shift
done
LO_SERIES="${LO_VERSION%.*}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
# The proxy of a managed environment signs with its own CA: use it when there is one.
CURL=(curl -fsSL)
[ -f /root/.ccr/ca-bundle.crt ] && CURL+=(--cacert /root/.ccr/ca-bundle.crt)

# ---- Rust -------------------------------------------------------------------------------------
if ! command -v rustup >/dev/null; then
  "${CURL[@]}" https://sh.rustup.rs | sh -s -- -y --default-toolchain none
  export PATH="$HOME/.cargo/bin:$PATH"
fi
CHANNEL="$(sed -n 's/^channel *= *"\(.*\)"/\1/p' "$ROOT/rust-toolchain.toml")"
(cd "$ROOT" && rustup toolchain install)           # channel, components and targets of the file
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
if ! perf stat -e instructions:u true 2>&1 | grep -q "not supported"; then
  echo "perf: hardware counters available"
else
  echo "perf: no hardware counters in this machine; use 'perf record -e cpu-clock' and callgrind"
fi

# ---- LibreOffice ------------------------------------------------------------------------------
if [ -x /opt/libreoffice${LO_SERIES}/program/soffice ]; then
  echo "libreoffice: already in /opt/libreoffice${LO_SERIES}"
else
  work="$(mktemp -d)"
  base="https://download.documentfoundation.org/libreoffice/stable/${LO_VERSION}/deb/x86_64"
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
if $REPLACE_SYSTEM; then
  apt-get remove -y --purge 'libreoffice-*' python3-uno || true
  rm -rf /usr/lib/libreoffice
fi
ln -sf /opt/libreoffice${LO_SERIES}/program/soffice /usr/local/bin/soffice
echo "libreoffice: $(soffice --version)"
echo "uno scripts: /opt/libreoffice${LO_SERIES}/program/python (the Python that ships with LibreOffice)"
echo "server for the oracles: scripts/libreoffice-server.sh"
