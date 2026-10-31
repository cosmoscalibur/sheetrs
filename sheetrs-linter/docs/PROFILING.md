# Sheetlint Profiling & Architecture

This document captures performance profiling data and architectural insights for optimizing sheetlint.

---

## Profiling Results

> **Code state**: after #73 (REF310 compares formula patterns computed once per cell)  
> **Test file**: `TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.xlsx` (37 MB)  
> **Workbook stats**: 51 sheets, 4,453,591 non-empty cells, 1,855,268 formulas  
> **Build**: `profiling` Cargo profile with frame pointers

Shares are inclusive `perf` samples (a function plus its callees) over one run:

```bash
RUSTFLAGS="-C force-frame-pointers=yes" cargo build --profile profiling -p sheetrs-linter
perf record -F 499 --call-graph fp ./target/profiling/sheetlint <file> > /dev/null
perf report --children --sort symbol -g none
```

The hooks of INT401–403 and VUL601 have moved since this profile was taken: they read the
sheet's cells in `on_sheet_end` instead of copying them in `on_cell`.

Benchmark metrics per change live in
[Comparisons](../../scripts/README.md#comparisons).

### Phase Breakdown

| Phase | Share |
|-------|-------|
| **Parsing** (ZIP + XML → Workbook) | 6.9% |
| **Rule execution** (walker and violation sorting) | 86.5% |
| **Output formatting** (`print_human`) | 6.6% |

Parsing includes `translate_shared_formula` at 1.4%.

### Per-Rule Share (Descending)

| Rule | Share | Hooks |
|------|-------|-------|
| **REF310** | 22.3% | `on_workbook_end` |
| **CALC202** | 21.3% | `on_cell` 9.5%, `on_workbook_end` 11.8% |
| **INT401–403** | 7.5% | `on_sheet_end` |
| CALC201 | 4.7% | `on_cell` |
| REF307 | 4.0% | `on_cell` |
| REF309 | 3.9% | `on_workbook_end` |
| VUL606 | 2.5% | `on_cell` |
| CPX504 | 2.5% | `on_cell` |
| VUL601 | 1.7% | `on_sheet_end` |
| DAT706 | 1.4% | `on_workbook_end` |
| ERR102 | 1.2% | `on_cell` |
| CALC203 | 1.0% | `on_cell` |
| DAT703 | 1.0% | `on_cell` |
| CALC204 | 0.9% | `on_cell` |
| CPX507 | 0.8% | `on_cell` |
| CALC205 | 0.8% | `on_cell` |
| REF301 | 0.7% | `on_cell` |
| VUL602 | 0.5% | `on_cell` |
| DAT704 | 0.5% | `on_cell` |
| EXT801 | 0.5% | `on_cell` |
| VUL604 | 0.5% | `on_cell` |
| VUL603 | 0.5% | `on_cell`, `on_sheet_end` |
| Other rules | <0.4% each | — |

REF310 and CALC202 account for 44% of samples. `PatternIds::of` (the formula pattern of a
cell, kept by cell) is 4.2% of all samples.

## Current Architecture

### Single-Pass Walker

Every rule implements `WalkerRule` (see [coding patterns](../../docs/coding-patterns.md#the-walker-pattern)).
`WorkbookWalker` traverses the parsed workbook once and calls each rule's hooks:

```mermaid
flowchart LR
    subgraph Parse["Parsing"]
        A[ZIP Archive] --> B[XML Parsing<br/>quick-xml]
        B --> C[Build Workbook<br/>sheets, cells, names]
    end

    subgraph Walk["Rule execution: WorkbookWalker"]
        D[on_workbook_start<br/>all rules] --> E
        subgraph Sheets["for each sheet"]
            E[on_sheet_start] --> F[on_cell<br/>each cell × all rules]
            F --> G[on_sheet_end]
        end
        G --> H[on_workbook_end<br/>all rules]
    end

    C --> D
    H --> I[Sort violations] --> J[Format output]
```
