### Added

- `sheetc` implements `CHISQDIST`, `CHISQINV`, `WEEKNUM_OOO`, `RAND.NV` and `RANDBETWEEN.NV` with oracle rows computed by LibreOffice 26.8 (see `FIDELITY.md`, *Names of OpenDocument and LibreOffice*).

### Fixed

- `RANDBETWEEN` rounds both limits away from zero before comparing them, as LibreOffice does: `RANDBETWEEN(1.2, 1.9)` is 2 (the engine answered `#VALUE!`).
