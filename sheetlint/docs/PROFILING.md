# Sheetlint Profiling & Architecture

This document captures performance profiling data and architectural insights for optimizing sheetlint.

---

## Profiling Results

> **Code state**: after #37 (REF301 single formula scan)  
> **Test file**: `TC2025 2025.12.05 Motor Tributi 1402 - MY- textos sin j.xlsx` (37 MB)  
> **Workbook stats**: 51 sheets, 4,453,591 non-empty cells, 1,855,268 formulas  
> **Build**: `profiling` Cargo profile with frame pointers

Shares are inclusive `perf` samples (a function plus its callees) over one run:

```bash
RUSTFLAGS="-C force-frame-pointers=yes" cargo build --profile profiling -p sheetlint
perf record -F 499 --call-graph fp ./target/profiling/sheetlint <file> > /dev/null
perf report --children --sort symbol -g none
```

Benchmark metrics per change live in
[Comparisons](../../scripts/README.md#comparisons).

### Phase Breakdown

| Phase | Share |
|-------|-------|
| **Parsing** (ZIP + XML → Workbook) | 14.3% |
| **Rule execution** (walker and violation sorting) | 79.6% |
| **Output formatting** (`print_human`) | 6.1% |

Parsing includes `translate_shared_formula` at 4.6%. A run with only FILE1001
enabled, a workbook-level rule, takes 4.2 s user time against the 39.5 s of
all rules in the #31 benchmark.

### Per-Rule Share (Descending)

| Rule | Share | Hooks |
|------|-------|-------|
| **CALC202** | 20.1% | `on_cell` 8.9%, `on_workbook_end` 11.3% |
| **INT401–403** | 11.9% | `on_cell` 7.1%, `on_sheet_end` 4.8% |
| **REF310** | 8.7% | `on_workbook_end` |
| CALC201 | 5.7% | `on_cell` |
| VUL601 | 3.5% | `on_cell` 2.2%, `on_sheet_end` 1.2% |
| REF309 | 3.0% | `on_workbook_end` |
| VUL606 | 2.6% | `on_cell` |
| CALC205 | 2.1% | `on_cell` |
| CPX504 | 2.1% | `on_cell` |
| ERR102 | 1.9% | `on_cell` |
| DAT706 | 1.3% | `on_workbook_end` |
| DAT703 | 1.1% | `on_cell` |
| VUL602 | 1.1% | `on_cell` |
| CALC204 | 1.0% | `on_cell` |
| CALC203 | 1.0% | `on_cell` |
| REF305 | 0.9% | `on_cell` |
| CPX507 | 0.7% | `on_cell` |
| REF301 | 0.7% | `on_cell` |
| VUL604 | 0.6% | `on_cell` |
| EXT801 | 0.5% | `on_cell` |
| VUL603 | 0.5% | `on_cell` |
| Other rules | <0.4% each | — |

CALC202, INT401–403, and REF310 account for 41% of samples.

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
