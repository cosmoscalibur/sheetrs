### Changed

- The generator of `RAND` mixes the `seed` (the finalizer of SplitMix64) and never starts at zero: distinct seeds now give distinct streams, where an even seed and the next odd one (`seed | 1`) gave the same, and the seed `0` is a seed like the others. The stream for a given seed differs from before. `RAND` of LibreOffice cannot be seeded, so `seed` is an extension of SheetRS with no LibreOffice counterpart (`FIDELITY.md`).
