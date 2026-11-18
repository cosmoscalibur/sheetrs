# Input and output

Pages here describe behaviour common to several formats; the lexical rules they share are in
`sheetrs-io`. Behaviour of one format is in its folder: [`xlsx/`](xlsx/), [`ods/`](ods/) and
[`csv/`](csv/).

Whatever two formats read differently stays in each folder: the same text is a number in one
format and not in another, and a year before 1 counts differently in each.
