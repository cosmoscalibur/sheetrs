### Added

- `sheetc` implements more functions of LibreOffice, each with oracle rows computed by LibreOffice 26.8: `GAMMALN`, `GAMMALN.PRECISE`, `GAMMA`, `FISHER`, `FISHERINV`; the gamma, chi-square (`CHISQ.*`, `CHIDIST`, `CHIINV`), t (`T.*`, `TDIST`, `TINV`), F (`F.*`, `FDIST`, `FINV`), beta, log-normal, Weibull, hypergeometric and negative binomial distributions and their inverses; `BINOM.INV`/`CRITBINOM`, `BINOM.DIST.RANGE`, `CONFIDENCE`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, and the tests `T.TEST`, `F.TEST`, `CHISQ.TEST`, `Z.TEST` and `PROB`. The special functions (log-gamma, the incomplete gamma and beta, inversion) are in `functions/special.rs`.
- `FDIST` and `FINV` are registered as the two functions they are: the OpenDocument ones (an ODS) and the legacy right-tail ones (an XLSX).
- The incomplete gamma function gives up after the 10,000 terms LibreOffice does (`#VALUE!`, its `Err:523`). `F.INV(1, …)` and the t, F and beta functions with more than 10 million degrees of freedom or shape stop the build: LibreOffice's results there are artifacts of its algorithm (see `FIDELITY.md`, *Distributions*).

### Fixed

- `NORMSINV(0.5)`, `NORM.S.INV(0.5)` and every function that reached `erf` at minus zero never returned (the series test compared against a negative sum).
