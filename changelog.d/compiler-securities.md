### Added

- `sheetc` implements the coupon functions (`COUPPCD`, `COUPNCD`, `COUPNUM`, `COUPDAYBS`, `COUPDAYS`, `COUPDAYSNC`) and the securities built on them: `PRICE`, `YIELD`, `DURATION`, `MDURATION`, `DISC`, `PRICEDISC`, `YIELDDISC`, `RECEIVED`, `INTRATE`, `PRICEMAT`, `YIELDMAT`, `ACCRINT`, `ACCRINTM`, `TBILLPRICE`, `TBILLYIELD` and `TBILLEQ`, with oracle rows computed by LibreOffice 26.8 (see `FIDELITY.md`, *Securities*): LibreOffice counts the days of a 30/360 period differently in each family, and the engine has each count.
- `TBILLYIELD` over a settlement and a maturity that are both the last day of February stops the build.
