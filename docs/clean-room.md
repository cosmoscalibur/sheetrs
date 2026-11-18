# Clean-room rules

The goal is both methodological (behaviour is specified and verified from the
outside) and legal (no work derived from LibreOffice source code, which is
MPL-2.0).

## Allowed sources

- ECMA-376 / ISO 29500 (OOXML) and OASIS ODF 1.3 / 1.4, including OpenFormula.
- Public papers and documentation of the linters whose rules are reproduced
  (PerfectXL and ExceLint).
- Unicode public data (for example CLDR) for locale tables.
- The LibreOffice binary, as a black-box oracle driven through its public UNO
  API. Its source code is never read.

## Not allowed

- Reading, copying or porting LibreOffice source code, including snippets
  found in forums, issues or code review tools.
- Committing private workbooks or content derived from them. Private
  workbooks are referenced by an external path and used only to characterize
  behaviour locally; reports contain aggregates, never cell content.

## Provenance

Every behavioural rule that is not stated in a public specification records
where it came from: the specification section, or the oracle case that showed
it. Oracle observations and their open defects are in [specs](specs/README.md), grouped
as the crates are.

Numeric differences against the oracle start at zero tolerance. Each exception
is recorded in [numeric-exceptions.md](numeric-exceptions.md) with its
function, case, observed difference and cause, and needs maintainer approval.

## Oracle configuration

The oracle runs with a dedicated LibreOffice user profile so that results do
not depend on the developer's settings. Any LibreOffice option that can change
a result (locale, separators, iteration, precision as shown, case sensitivity,
search criteria, wildcards or regular expressions) is set explicitly by the
test case.
