//! Where `LibreOffice` puts a shape when rows or columns are deleted.
//!
//! A shape anchored to a cell has its position and size from the sizes of the rows and columns, so
//! where it ends up depends on them: they are in the file (the height of each row and the width of
//! each column), and this is the model of what `LibreOffice` does with them, read from its source
//! (`ScDrawLayer::DeleteObjectsInArea`, `MoveCells`, `RecalcPos`, `ResizeLastRectFromAnchor`) and
//! checked against `LibreOffice` 26.8 on every band of the fixtures:
//!
//! 1. A shape that lies wholly inside the deleted band is deleted with it.
//! 2. The cell it is anchored to (its start, and its end when it has one) moves up by the size of the
//!    band when it is below it; when it is inside the band it keeps its index (whatever is at that
//!    index now is its cell). A shape that is a rectangle, a picture or a chart has its start and
//!    end put in order when they cross; a custom shape does not.
//! 3. Its top is the top of its start cell plus its offset in the cell, where the offset cannot be
//!    larger than the cell. Its size: when it resizes with its cells and its end is not in its start
//!    cell, from its end cell and the offset in it (the distance, when the end is above the start);
//!    when both are in one cell, the size it had scaled by what the cell changed (nothing, unless the
//!    rows or columns have different sizes); otherwise the size it had.
//!
//! The unit is the EMU (1/914400 inch), in which the sizes of the rows of a sheet (a point is
//! 12700, a twip 635) are whole numbers.

use super::structure::Bands;

/// How much wider than the shape its bounding rectangle is: the outline of a hairline, a hundredth
/// of a millimetre (checked against `LibreOffice`: a shape that begins where the deleted band begins
/// is not deleted, one that begins a row after it is).
const OUTLINE: i64 = 360;

/// A run of rows (or columns) of one size.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Run {
    first: u32,
    last: u32,
    size: i64,
}

/// The size of each row (or column) of a sheet: the default one and the runs that differ, hidden
/// ones counting for nothing.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct Sizes {
    default: i64,
    runs: Vec<Run>,
}

impl Sizes {
    pub(super) fn new(default: i64) -> Self {
        Self {
            default,
            runs: Vec::new(),
        }
    }

    /// Sets the size of `index` (0 for a hidden one).
    pub(super) fn set(&mut self, index: u32, size: i64) {
        self.set_run(index, index, size);
    }

    /// Sets the size of the rows (or columns) `first..=last`, which are set in order of position or
    /// not at all (a run that is not after the last one is dropped).
    pub(super) fn set_run(&mut self, first: u32, last: u32, size: i64) {
        if size == self.default || first > last {
            return;
        }
        if self.runs.last().is_some_and(|run| run.last >= first) {
            return;
        }
        match self.runs.last_mut() {
            Some(run) if run.size == size && run.last + 1 == first => run.last = last,
            _ => self.runs.push(Run { first, last, size }),
        }
    }

    pub(super) fn size(&self, index: u32) -> i64 {
        self.runs
            .iter()
            .find(|run| (run.first..=run.last).contains(&index))
            .map_or(self.default, |run| run.size)
    }

    /// Where the row (or column) `index` starts.
    pub(super) fn top(&self, index: u32) -> i64 {
        let extra: i64 = self
            .runs
            .iter()
            .filter(|run| run.first < index)
            .map(|run| {
                (i64::from(run.last.min(index - 1)) - i64::from(run.first) + 1)
                    * (run.size - self.default)
            })
            .sum();
        i64::from(index) * self.default + extra
    }

    /// The index at `position` and how far into it that is. A position past the sizes known is in
    /// the default size; a hidden one is skipped.
    pub(super) fn cell_at(&self, position: i64) -> (u32, i64) {
        let (mut index, mut start) = (0u32, 0i64);
        let within = |size: i64, start: i64, first: u32| {
            let skipped = (position - start) / size;
            (
                first + u32::try_from(skipped).unwrap_or(0),
                position - start - skipped * size,
            )
        };
        for run in &self.runs {
            let gap = i64::from(run.first - index) * self.default;
            if self.default > 0 && position < start + gap {
                return within(self.default, start, index);
            }
            start += gap;
            let length = i64::from(run.last - run.first + 1) * run.size;
            if run.size > 0 && position < start + length {
                return within(run.size, start, run.first);
            }
            start += length;
            index = run.last + 1;
        }
        if self.default <= 0 || position < start {
            return (index, 0);
        }
        within(self.default, start, index)
    }

    /// The sizes once the bands are gone.
    pub(super) fn without(&self, bands: &Bands) -> Self {
        let mut out = Self::new(self.default);
        for run in &self.runs {
            let mut current = run.first;
            let mut segments = Vec::new();
            for (first, last) in bands.ranges() {
                if last < current || first > run.last {
                    continue;
                }
                if first > current {
                    segments.push((current, first - 1));
                }
                current = current.max(last.saturating_add(1));
            }
            if current <= run.last {
                segments.push((current, run.last));
            }
            for (a, b) in segments {
                out.set_run(bands.shifted(a), bands.shifted(b), run.size);
            }
        }
        out
    }
}

/// A shape along one axis: its start and end cells, where in them, and its size.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct Span {
    pub start: u32,
    pub end: u32,
    /// From the start of the start cell to the shape.
    pub start_offset: i64,
    /// From the start of the end cell to the end of the shape.
    pub end_offset: i64,
    pub size: i64,
}

impl Span {
    /// A span from where it starts and how big it is, with its end where that falls.
    pub(super) fn of_size(sizes: &Sizes, start: u32, start_offset: i64, size: i64) -> Self {
        let top = sizes.top(start) + start_offset;
        let (end, end_offset) = sizes.cell_at(top + size);
        Self {
            start,
            end,
            start_offset,
            end_offset,
            size,
        }
    }

    /// A span from its two ends.
    pub(super) fn of_ends(
        sizes: &Sizes,
        start: u32,
        start_offset: i64,
        end: u32,
        end_offset: i64,
    ) -> Self {
        let size = sizes.top(end) + end_offset - (sizes.top(start) + start_offset);
        Self {
            start,
            end,
            start_offset,
            end_offset,
            size,
        }
    }
}

/// What a deletion does to a shape.
pub(super) struct Kind {
    /// It resizes with its cells (it has an end cell and its anchor says so).
    pub resizes: bool,
    /// A rectangle, a picture or a chart: its start and end are put in order.
    pub rectangle: bool,
}

/// The span of a shape along an axis from which `bands` are deleted (`old` and `new` are the sizes
/// before and after), `None` when the shape is deleted with them. `other_in_one_cell` says that the
/// shape starts and ends in the same cell along the other axis.
pub(super) fn deleted(
    span: Span,
    old: &Sizes,
    new: &Sizes,
    bands: &Bands,
    kind: &Kind,
    other_in_one_cell: bool,
) -> Option<Span> {
    let top = old.top(span.start) + span.start_offset;
    let bottom = top + span.size;
    for (first, last) in bands.ranges() {
        // The shape inside one band goes with it: its bounding rectangle, which is its outline
        // wider (a hairline is a hundredth of a millimetre), inside the deleted one, so a shape that
        // starts where the band starts stays.
        if top - OUTLINE >= old.top(first) && bottom + OUTLINE <= old.top(last + 1) {
            return None;
        }
    }
    let moved = |index: u32| bands.moved(index).unwrap_or(index);
    let (mut start, mut end) = (moved(span.start), moved(span.end));
    if kind.rectangle && start > end {
        std::mem::swap(&mut start, &mut end);
    }
    let start_offset = span.start_offset.min(new.size(start));
    let top = new.top(start) + start_offset;
    let size = if kind.resizes && (start != end || !other_in_one_cell) {
        let end_offset = span.end_offset.min(new.size(end));
        (new.top(end) + end_offset - top).abs()
    } else if kind.resizes {
        let before = old.size(span.start);
        if before == 0 {
            span.size
        } else {
            span.size * new.size(start) / before
        }
    } else {
        span.size
    };
    let (end, end_offset) = new.cell_at(top + size);
    Some(Span {
        start,
        end,
        start_offset,
        end_offset,
        size,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const ROW: i64 = 256 * 635;

    /// The shape of `book.xlsx` that resizes with its cells: from D2 to D4 + 106920 EMU, rows of
    /// 12.8 points.
    fn resizing() -> (Span, Sizes) {
        let sizes = Sizes::new(ROW);
        (Span::of_ends(&sizes, 1, 0, 3, 106_920), sizes)
    }

    fn after(span: Span, first: u32, count: u32, kind: &Kind) -> Option<Span> {
        let old = Sizes::new(ROW);
        let bands = Bands::new(vec![(first, first + count - 1)]);
        let new = old.without(&bands);
        deleted(span, &old, &new, &bands, kind, true)
    }

    #[test]
    fn the_sizes_find_a_position_and_its_offset() {
        let mut sizes = Sizes::new(100);
        sizes.set(2, 300);
        sizes.set(4, 0);
        assert_eq!(sizes.top(3), 500);
        assert_eq!(sizes.top(5), 600);
        for position in [0, 99, 100, 250, 499, 500, 599, 600, 699, 700, 1234] {
            let (index, offset) = sizes.cell_at(position);
            assert_eq!(sizes.top(index) + offset, position, "{position}");
            assert!(offset < sizes.size(index).max(1) || index > 5, "{position}");
        }
    }

    #[test]
    fn a_shape_that_resizes_follows_its_end_and_keeps_its_start_cell_when_that_goes() {
        let (span, _) = resizing();
        // A row above it: both move up one.
        let up = after(
            span,
            0,
            1,
            &Kind {
                resizes: true,
                rectangle: false,
            },
        )
        .unwrap();
        assert_eq!((up.start, up.end, up.size), (0, 2, span.size));
        // Its start row goes: it stays at index 1 and loses that row of its height.
        let shorter = after(
            span,
            1,
            1,
            &Kind {
                resizes: true,
                rectangle: false,
            },
        )
        .unwrap();
        assert_eq!((shorter.start, shorter.end), (1, 2));
        assert_eq!(shorter.size, span.size - ROW);
        // Its start and the row before it go (a band of two): both ends are in one cell, so it keeps
        // its size.
        let kept = after(
            span,
            0,
            2,
            &Kind {
                resizes: true,
                rectangle: false,
            },
        )
        .unwrap();
        assert_eq!((kept.start, kept.end, kept.size), (1, 3, span.size));
    }

    #[test]
    fn a_shape_wholly_inside_the_band_goes_with_it() {
        let (span, _) = resizing();
        let kind = Kind {
            resizes: true,
            rectangle: false,
        };
        assert!(after(span, 0, 4, &kind).is_none());
        // The band starts where the shape does: the outline is outside it, so it stays.
        assert!(after(span, 1, 3, &kind).is_some());
        assert!(after(span, 0, 3, &kind).is_some());
    }

    #[test]
    fn a_rectangle_whose_ends_cross_has_them_put_in_order_and_a_custom_shape_does_not() {
        let (span, _) = resizing();
        // Rows 0 to 2 go: its start (1) stays, its end (3) moves to 0.
        let custom = after(
            span,
            0,
            3,
            &Kind {
                resizes: true,
                rectangle: false,
            },
        )
        .unwrap();
        assert_eq!((custom.start, custom.end), (1, 1));
        let rectangle = after(
            span,
            0,
            3,
            &Kind {
                resizes: true,
                rectangle: true,
            },
        )
        .unwrap();
        assert_eq!(rectangle.start, 0);
        assert_eq!(rectangle.start_offset, 0);
    }

    #[test]
    fn a_shape_that_does_not_resize_keeps_its_size_and_follows_its_start() {
        let sizes = Sizes::new(ROW);
        let span = Span::of_size(&sizes, 9, 0, 360_000);
        let kind = Kind {
            resizes: false,
            rectangle: true,
        };
        let moved = after(span, 8, 2, &kind).unwrap();
        assert_eq!((moved.start, moved.size), (9, 360_000));
        let up = after(span, 0, 1, &kind).unwrap();
        assert_eq!((up.start, up.size), (8, 360_000));
    }

    /// `LibreOffice` 26.8 on the two shapes of `sheetrs-edit/tests/assets/structure/book.*` after
    /// every band of rows and of columns up to five wide
    /// (`sheetrs-edit/scripts/make_shapes_oracle.py`): where each ends up is what `deleted` says,
    /// to a hundredth of a millimetre or two.
    #[test]
    fn the_shapes_end_up_where_libreoffice_puts_them() {
        #[derive(serde::Deserialize)]
        struct Case {
            format: String,
            axis: String,
            first: u32,
            count: u32,
            shapes: Vec<[i64; 6]>,
        }
        let cases: Vec<Case> =
            serde_json::from_str(include_str!("../../tests/assets/oracle/shapes_delete.json"))
                .unwrap();
        assert_eq!(cases.len(), 190);
        // The rows are 12.8 points and the columns 2.258 cm, in twips as `LibreOffice` has them.
        let rows = Sizes::new(256 * 635);
        let cols = Sizes::new(1280 * 635);
        // (row, row offset, end row, end offset, column, offset, end column, end offset, resizes,
        // rectangle), in EMU: the shapes of the fixtures.
        type Shape = (u32, i64, u32, i64, u32, i64, u32, i64, bool, bool);
        let shapes = |format: &str| -> [Shape; 2] {
            if format == "xlsx" {
                [
                    (1, 0, 3, 106_920, 3, 0, 3, 792_000, true, false),
                    (9, 0, 11, 34_920, 2, 0, 2, 720_000, false, false),
                ]
            } else {
                // The ODS rectangles: 0.8661 in by 0.4724 in ending in D4, and 0.7874 in by 0.3937 in.
                [
                    (1, 0, 3, 106_920, 3, 0, 3, 792_000, true, true),
                    (9, 0, 11, 34_880, 2, 0, 2, 720_000, false, true),
                ]
            }
        };
        let mm100 = |emu: i64| emu as f64 / 360.0;
        let mut compared = 0;
        for case in &cases {
            let band = Bands::new(vec![(case.first, case.first + case.count - 1)]);
            let mut ours: Vec<[f64; 6]> = Vec::new();
            for (r, ro, er, eo, c, co, ec, eco, resizes, rectangle) in shapes(&case.format) {
                let (mut row, mut col) = (
                    Span::of_ends(&rows, r, ro, er, eo),
                    Span::of_ends(&cols, c, co, ec, eco),
                );
                let kind = Kind { resizes, rectangle };
                let gone = if case.axis == "rows" {
                    let new = rows.without(&band);
                    deleted(row, &rows, &new, &band, &kind, col.start == col.end).map(|s| row = s)
                } else {
                    let new = cols.without(&band);
                    deleted(col, &cols, &new, &band, &kind, row.start == row.end).map(|s| col = s)
                };
                if gone.is_none() {
                    continue;
                }
                let (new_rows, new_cols) = if case.axis == "rows" {
                    (rows.without(&band), cols.clone())
                } else {
                    (rows.clone(), cols.without(&band))
                };
                ours.push([
                    f64::from(row.start),
                    f64::from(col.start),
                    mm100(new_cols.top(col.start) + col.start_offset),
                    mm100(new_rows.top(row.start) + row.start_offset),
                    mm100(col.size),
                    mm100(row.size),
                ]);
            }
            ours.sort_by(|a, b| (a[2], a[3]).partial_cmp(&(b[2], b[3])).unwrap());
            let what = format!(
                "{} {} {}+{}",
                case.format, case.axis, case.first, case.count
            );
            assert_eq!(
                ours.len(),
                case.shapes.len(),
                "{what}: {ours:?} / {:?}",
                case.shapes
            );
            for (got, want) in ours.iter().zip(&case.shapes) {
                // The anchor cell exactly, the geometry to two hundredths of a millimetre.
                assert_eq!(
                    (got[0], got[1]),
                    (want[0] as f64, want[1] as f64),
                    "{what}: {ours:?} / {:?}",
                    case.shapes
                );
                for k in 2..6 {
                    assert!(
                        (got[k] - want[k] as f64).abs() <= 2.5,
                        "{what}: {ours:?} / {:?}",
                        case.shapes
                    );
                }
                compared += 1;
            }
        }
        assert!(compared > 250, "{compared}");
    }
}
