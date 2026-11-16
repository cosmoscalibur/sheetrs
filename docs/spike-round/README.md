# `round_dec` against LibreOffice: harness

Throwaway code behind the section *Rounding against LibreOffice* of
[../plan-engine-and-writer-performance.md](../plan-engine-and-writer-performance.md). It compares the
engine's `ROUND`, `ROUNDUP` and `ROUNDDOWN` (the text definition `round_dec_text`, which the integer
version reproduces bit for bit) with what LibreOffice 26.8 calculates, and with the alternatives that were
tried. `old.rs` is a copy of that definition (`sheetrs-engine/src/functions/math.rs`), taken when the
harness ran.

The ODS that LibreOffice saves carries 15 significant digits, so the inputs and the results read back from it lose the last digits.
`make_exact.py` and `make_wide_d.py` build every input *inside* LibreOffice as `(a·2²⁶ + b)·2^e` with whole numbers, which is exact, and
the results are compared by the decimal quantity they stand for (`result · 10^digits`, rounded to an integer: `by_digits_fit.rs`,
`sweep_threshold.rs`). (Reading the numbers is not the problem: LibreOffice reads cell values and formula literals correctly rounded.)

```bash
python3 make_xlsx.py 25000                      # cases_in.xlsx and meta.tsv: 300 000 values, 3 functions each
mkdir -p prof/user out
# a profile that recalculates on load (OOXMLRecalcMode 0), as sheetrs-compiler/src/recalc.rs does
soffice -env:UserInstallation=file://$PWD/prof --headless --convert-to ods --outdir out cases_in.xlsx
python3 parse_ods.py meta.tsv out/cases_in.ods cases.tsv
rustc --edition 2024 -O compare.rs -o compare && ./compare cases.tsv        # by family and function
rustc --edition 2024 -O by_digits.rs -o by_digits && ./by_digits            # the typical families, by digits
rustc --edition 2024 -O -C target-cpu=native adversarial.rs -o adversarial && ./adversarial   # 3·10^7 cases
```

`compare.rs` treats `+0` and `-0` as equal (LibreOffice gives `-0` where the engine gives `+0`; the two are
the same number). The UNO route (`setFormulaArray` of 900 000 formulas) takes longer than twenty minutes:
the file route takes ten seconds.

```bash
# with exact inputs (use these)
python3 make_exact.py 25000      # cases_exact.xlsx, meta_exact.tsv (digits -2 to 8)
python3 make_wide_d.py 12000     # cases_wide_d.xlsx, meta_wide_d.tsv (digits -14 to 14)
# convert each with LibreOffice to ODS as above; the results are the columns G, H, I (ROUND, ROUNDUP, ROUNDDOWN)
rustc --edition 2024 -O by_digits_fit.rs -o by_digits_fit && ./by_digits_fit cases_wide_d.tsv   # per digits: engine before | fit
rustc --edition 2024 -O sweep_threshold.rs -o sweep && ./sweep cases_exact.tsv                    # the 2^41 threshold against others
```

The permanent oracle that came out of it is `sheetrs-engine/tests/assets/round_oracle.tsv` (7 500 results of
LibreOffice 26.8.2) and `round_oracle_pending.tsv` (the `ROUNDDOWN` cases where the fit still differs).
