# Claude Code environment

The Claude Code environment of SheetRS needs two things in its settings:

1. **Allowed domain:** `downloadarchive.documentfoundation.org`, the fixed host of the LibreOffice
   packages (no redirect to a mirror). The Rust, apt and Python package managers are already
   permitted.
2. **Setup script:** paste the whole content of
   [`scripts/setup-claude-env.sh`](../scripts/setup-claude-env.sh) into the script box of the
   environment. It is self-contained, takes no parameters and no path. It leaves Rust 1.99 (the
   channel of `rust-toolchain.toml`), LibreOffice 26.8.2.2 (the reference of the project, in
   `/opt/libreoffice26.8`, `soffice` on the `PATH`), `perf` and `valgrind`.

No variables or secrets are needed.

The scripts that talk to LibreOffice (`sheetrs-io/tests/assets/oracle/make_oracle.py`, the fixture
generators) run with the Python of LibreOffice, `/opt/libreoffice26.8/program/python`, against the
server of `scripts/libreoffice-server.sh` (port 2002; `LO_PORT` and `LO_PROFILE` change the port and
the profile folder).

For a local machine, `sudo scripts/setup-dev-env.sh` (Debian or Ubuntu; `--replace-system` also
removes the system LibreOffice).
