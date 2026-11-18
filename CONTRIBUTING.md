# Contributing

## Branching

`next/0.5` is the integration branch of the 0.5 series and the base of all work
on it. New versions of the series are developed against that branch until the
series is promoted to stable.

## Lints

`clippy::all` and `clippy::pedantic` are `deny` for the whole workspace and
`unsafe_code` is `forbid`. A lint suppression (an `allow` attribute or a
configuration that disables a rule) needs prior approval from the maintainer.

## Checks

Checks run locally in a Linux container (Docker). The working tree is mounted
read-only and build output and caches live under `~/.cache/sheetrs-ci`.

```bash
scripts/ci-local.sh
```

It runs format, clippy and tests on Linux, and compiles the workspace for the
`x86_64-pc-windows-msvc` and `aarch64-apple-darwin` targets.

## GitHub CI

A push does not start CI. It starts only when the head commit message contains
`[ci]`. Pull requests that are not drafts run it as the merge requirement, on
Linux, Windows and macOS, and it can also be started manually.

Open the pull request as a draft and keep it a draft while the work is in
progress: draft pull requests do not run CI. Mark it ready for review only when
the state is final, since that is what starts the merge-requirement run.

Behaviour that is not stated in a public specification follows the rules in
[docs/clean-room.md](docs/clean-room.md).
