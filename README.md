# sheetrs (next/0.5)

A headless spreadsheet engine in Rust, written clean-room with respect to
LibreOffice Calc 26.8. The target is that, for an input workbook and a set of
operations, the output workbook (values, errors and structure) matches the one
LibreOffice produces. Fidelity is measured end to end; the intermediate state
is free to differ.

## Workspace

| Crate | Purpose |
| --- | --- |
| `sheetrs-core` | Core types: cell coordinates with A1 and R1C1 views |
| `sheetrs-testkit` | Differential test harness that runs cases against LibreOffice as the oracle |

## Development

Requires Rust 1.99 or newer, and Docker for the local checks. See [CONTRIBUTING.md](CONTRIBUTING.md) for the
branching and lint rules and [docs/clean-room.md](docs/clean-room.md) for the
clean-room rules.

```bash
scripts/ci-local.sh
```

## License

MIT, see [LICENSE](LICENSE).
