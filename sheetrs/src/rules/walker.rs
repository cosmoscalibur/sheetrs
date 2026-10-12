//! Workbook walker to execute optimized rules in a single pass

use crate::reader::Workbook;
use crate::rules::{LinterContext, WalkerRule};
use crate::violation::Violation;

/// Executor that traverses the workbook once and invokes rules
pub struct WorkbookWalker<'a> {
    workbook: &'a Workbook,
    rules: Vec<Box<dyn WalkerRule>>,
    context: LinterContext,
}

impl<'a> WorkbookWalker<'a> {
    /// Create a new walker for the given workbook and rules
    pub fn new(workbook: &'a Workbook, rules: Vec<Box<dyn WalkerRule>>) -> Self {
        Self {
            workbook,
            rules,
            context: LinterContext::default(),
        }
    }

    /// Walk the workbook and return all violations found
    pub fn walk(mut self) -> Vec<Violation> {
        let mut violations = Vec::new();

        // 0. Pre-populate LinterContext
        self.context.name_to_index = self
            .workbook
            .sheets
            .iter()
            .map(|s| (s.name.clone(), s.sheet_index))
            .collect();

        // 1. Workbook Start
        for rule in &mut self.rules {
            violations.extend(rule.on_workbook_start(self.workbook, &mut self.context));
        }

        // 2. Iterate Sheets
        for sheet in &self.workbook.sheets {
            // Sheet Start
            for rule in &mut self.rules {
                violations.extend(rule.on_sheet_start(sheet, &mut self.context));
            }

            // Iterate Cells
            for cell in sheet.cells.values() {
                for rule in &mut self.rules {
                    violations.extend(rule.on_cell(sheet, cell, &mut self.context));
                }
            }

            // Sheet End
            for rule in &mut self.rules {
                violations.extend(rule.on_sheet_end(sheet, &mut self.context));
            }
        }

        // 3. Workbook End
        for rule in &mut self.rules {
            violations.extend(rule.on_workbook_end(self.workbook, &mut self.context));
        }

        violations
    }

    /// Get a reference to the accumulated context (useful for debugging/testing)
    pub fn context(&self) -> &LinterContext {
        &self.context
    }
}
