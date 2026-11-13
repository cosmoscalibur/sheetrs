### Fixed

- `gen_native_fixture.py` copies every sheet of `semantics.xlsx` (it dropped `Wide`, so the formulas over several sheets were written back garbled) and `native.ods` is regenerated; the native-ODS test lists the three rows where LibreOffice reads the typed formula otherwise than the import did and requires them to keep differing.
- CI checks that the tracked workbooks are deterministic (`scripts/fixture_determinism.py --check`).
