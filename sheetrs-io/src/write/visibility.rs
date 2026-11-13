//! Hiding and showing rows, columns and sheets: what a change is, and the intervals the writers
//! of both formats read it through.

use super::{ColumnRange, RowRange};
use std::collections::BTreeMap;
use std::fmt;

/// One change of visibility, in the coordinates of the workbook as it is.
///
/// A later change of the same row, column or sheet wins over an earlier one.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum VisibilityChange {
    /// Hide (`visible` false) or show the rows of a range.
    Rows { range: RowRange, visible: bool },
    /// Hide or show the columns of a range.
    Columns { range: ColumnRange, visible: bool },
    /// Hide or show a sheet. Showing makes it visible whatever hid it; hiding leaves a very hidden
    /// sheet (XLSX) as it is.
    Sheet { name: String, visible: bool },
}

impl fmt::Display for VisibilityChange {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let verb = |visible: bool| if visible { "show" } else { "hide" };
        match self {
            Self::Rows { range, visible } => write!(f, "{} rows {range}", verb(*visible)),
            Self::Columns { range, visible } => write!(f, "{} columns {range}", verb(*visible)),
            Self::Sheet { name, visible } => write!(f, "{} sheet {name}", verb(*visible)),
        }
    }
}

/// What a set of changes asks of one sheet.
#[derive(Debug, Default, Clone)]
pub(crate) struct SheetVisibility {
    pub rows: Intervals,
    pub columns: Intervals,
}

/// The changes of a pass, by sheet and by kind.
#[derive(Debug, Default, Clone)]
pub(crate) struct Visibilities {
    pub sheets: BTreeMap<String, SheetVisibility>,
    /// The state each sheet asked for: `true` is hidden.
    pub states: BTreeMap<String, bool>,
}

impl Visibilities {
    pub(crate) fn new(changes: &[VisibilityChange]) -> Self {
        let mut found = Self::default();
        for change in changes {
            match change {
                VisibilityChange::Rows { range, visible } => found
                    .sheets
                    .entry(range.sheet.clone())
                    .or_default()
                    .rows
                    .assign(range.first, range.last, !visible),
                VisibilityChange::Columns { range, visible } => found
                    .sheets
                    .entry(range.sheet.clone())
                    .or_default()
                    .columns
                    .assign(range.first, range.last, !visible),
                VisibilityChange::Sheet { name, visible } => {
                    found.states.insert(name.clone(), !visible);
                }
            }
        }
        found
    }

    /// The names of the sheets the changes address that are not in `names`.
    pub(crate) fn unknown<'a>(&'a self, names: &[String]) -> Option<&'a str> {
        self.sheets
            .keys()
            .chain(self.states.keys())
            .find(|name| !names.contains(name))
            .map(String::as_str)
    }
}

/// Disjoint intervals of positions (0-based, inclusive) with the state they were given (`true` is
/// hidden); positions outside every interval were not mentioned. A later assignment wins.
#[derive(Debug, Default, Clone)]
pub(crate) struct Intervals {
    /// Start of an interval -> (end, hidden).
    map: BTreeMap<u32, (u32, bool)>,
}

impl Intervals {
    pub(crate) fn is_empty(&self) -> bool {
        self.map.is_empty()
    }

    pub(crate) fn assign(&mut self, first: u32, last: u32, hidden: bool) {
        // The intervals that overlap `first..=last`, which keep what lies outside it.
        let overlapping: Vec<u32> = self
            .map
            .range(..=last)
            .rev()
            .take_while(|(_, (end, _))| *end >= first)
            .map(|(&start, _)| start)
            .collect();
        for start in overlapping {
            let (end, was) = self.map.remove(&start).expect("found a moment ago");
            if start < first {
                self.map.insert(start, (first - 1, was));
            }
            if end > last {
                self.map.insert(last + 1, (end, was));
            }
        }
        self.map.insert(first, (last, hidden));
    }

    /// The state given to `at`, `None` when it was not mentioned.
    pub(crate) fn state_at(&self, at: u32) -> Option<bool> {
        self.map
            .range(..=at)
            .next_back()
            .filter(|(_, (end, _))| *end >= at)
            .map(|(_, (_, hidden))| *hidden)
    }

    /// The first hidden position that is `from` or after it.
    pub(crate) fn next_hidden(&self, from: u32) -> Option<u32> {
        if self.state_at(from) == Some(true) {
            return Some(from);
        }
        self.map
            .range(from..)
            .find(|(_, (_, hidden))| *hidden)
            .map(|(&start, _)| start)
    }

    /// `first..=last` cut into runs of one state each, in order: `None` where nothing was said.
    pub(crate) fn pieces(&self, first: u32, last: u32) -> Vec<(u32, u32, Option<bool>)> {
        let mut pieces = Vec::new();
        let mut at = first;
        while at <= last {
            let (end, state) = match self.state_at(at) {
                Some(hidden) => {
                    let (_, (end, _)) = self.map.range(..=at).next_back().expect("state found");
                    (*end, Some(hidden))
                }
                None => (
                    self.map
                        .range(at..)
                        .next()
                        .map_or(u32::MAX, |(&start, _)| start - 1),
                    None,
                ),
            };
            let end = end.min(last);
            pieces.push((at, end, state));
            if end == u32::MAX {
                break;
            }
            at = end + 1;
        }
        pieces
    }

    /// The runs of hidden positions at or after `from`, clipped to start there, and not past
    /// `limit`.
    pub(crate) fn hidden_runs(&self, from: u32, limit: u32) -> Vec<(u32, u32)> {
        let mut runs: Vec<(u32, u32)> = Vec::new();
        for (&start, &(end, hidden)) in &self.map {
            if !hidden || end < from || start > limit {
                continue;
            }
            let (start, end) = (start.max(from), end.min(limit));
            match runs.last_mut() {
                Some(last) if last.1 + 1 == start => last.1 = end,
                _ => runs.push((start, end)),
            }
        }
        runs
    }

    /// Every position where the state may change: the start of each interval and the one after its
    /// end.
    pub(crate) fn boundaries(&self) -> impl Iterator<Item = u32> + '_ {
        self.map
            .iter()
            .flat_map(|(&start, &(end, _))| [start, end.saturating_add(1)])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_later_assignment_wins_over_the_part_it_overlaps() {
        let mut intervals = Intervals::default();
        intervals.assign(2, 9, true);
        intervals.assign(4, 5, false);
        assert_eq!(intervals.state_at(1), None);
        assert_eq!(intervals.state_at(3), Some(true));
        assert_eq!(intervals.state_at(4), Some(false));
        assert_eq!(intervals.state_at(6), Some(true));
        assert_eq!(intervals.hidden_runs(0, 100), [(2, 3), (6, 9)]);
    }

    #[test]
    fn pieces_cover_the_range_with_one_state_each() {
        let mut intervals = Intervals::default();
        intervals.assign(3, 4, true);
        intervals.assign(5, 6, false);
        assert_eq!(
            intervals.pieces(0, 8),
            [
                (0, 2, None),
                (3, 4, Some(true)),
                (5, 6, Some(false)),
                (7, 8, None)
            ]
        );
        assert_eq!(
            intervals.pieces(4, 5),
            [(4, 4, Some(true)), (5, 5, Some(false))]
        );
    }

    #[test]
    fn the_next_hidden_position_skips_the_shown_ones() {
        let mut intervals = Intervals::default();
        intervals.assign(2, 3, false);
        intervals.assign(7, 8, true);
        assert_eq!(intervals.next_hidden(0), Some(7));
        assert_eq!(intervals.next_hidden(8), Some(8));
        assert_eq!(intervals.next_hidden(9), None);
    }
}
