//! The parsed VBA project of each open workbook (design Decision 3). The model keeps only the
//! `vbaProject.bin` bytes (`Workbook::vba_project`, saved back unchanged); this cache reads them
//! into a [`VbaProject`] the first time a command asks (Req 6.1) and reuses that while the same
//! bytes are alive (Req 6.4). Opening a workbook never touches it (Req 6.3).
//!
//! Slots are keyed on the identity of the bytes' `Arc` and on `Workbook::vba_signature_part`.
//! Undo and redo snapshots share the `Arc`, so they reuse the parse; replacing the bytes makes a
//! new `Arc` and parses again. A slot holds only a `Weak`, so it never keeps a closed workbook's
//! bytes alive, and while the `Weak` lives the allocation can't be reused for other bytes: a
//! stale slot can never alias a new buffer. Dead slots are dropped on each call, and at most
//! [`MAX_SLOTS`] are kept (least recently used first out).
//!
//! A slot also keeps the project's [`CompatibilityReport`] once `vba.analyze` asks for it. The
//! report depends only on the project (and the member catalog, which is fixed at build time), so
//! it lives and dies with the slot. Analysing parses every module again: measured on a 64-module
//! project of 5.9 MB of source (25,600 procedures), it took 230 ms in a release build (361 ms in
//! a debug build), roughly 40 ms per MB. A View Code window that refreshes the report on every
//! open would pay that each time, so the report is kept; it is small next to the source.
use std::sync::{Arc, Weak};

use gridcraft_model::Workbook;
use gridcraft_vba::{CompatibilityReport, EmptyCatalog, Limits, VbaProject};

/// Most parsed projects kept: a handful of open workbooks, with room for edits in between.
pub(crate) const MAX_SLOTS: usize = 16;

struct Slot {
    bytes: Weak<Vec<u8>>,
    signature_part: bool,
    project: Arc<VbaProject>,
    report: Option<Arc<CompatibilityReport>>,
}

#[derive(Default)]
pub(crate) struct VbaCache {
    /// Least recently used first.
    slots: Vec<Slot>,
    /// How many times a project was read, for tests.
    #[cfg(test)]
    pub(crate) parses: usize,
    /// How many times a project was analysed, for tests.
    #[cfg(test)]
    pub(crate) analyses: usize,
}

/// The analyzer run the engine uses: the checked-in catalog once it exists (task 12), until then
/// [`EmptyCatalog`].
fn analyze(project: &VbaProject) -> CompatibilityReport {
    gridcraft_vba::analyze(project, &EmptyCatalog)
}

impl VbaCache {
    /// `wb`'s VBA project, read on the first request and reused while the same bytes are alive.
    /// A workbook without a project gives the empty project (no diagnostics) and reads nothing
    /// (Req 6.2); bytes that can't be read give a project carrying only diagnostics (Req 6.5).
    pub(crate) fn project(&mut self, wb: &Workbook) -> Arc<VbaProject> {
        match self.slot(wb) {
            Some(slot) => slot.project.clone(),
            None => Arc::new(VbaProject::default()),
        }
    }

    /// The compatibility report of `wb`'s VBA project, analysed on the first request and reused
    /// with the project. A workbook without a project gives the report of the empty project
    /// (verdict "no macros", Req 13.13).
    pub(crate) fn report(&mut self, wb: &Workbook) -> Arc<CompatibilityReport> {
        let Some(slot) = self.slot(wb) else {
            return Arc::new(analyze(&VbaProject::default()));
        };
        if let Some(report) = &slot.report {
            return report.clone();
        }
        let report = Arc::new(analyze(&slot.project));
        slot.report = Some(report.clone());
        #[cfg(test)]
        {
            self.analyses = self.analyses.saturating_add(1);
        }
        report
    }

    /// `wb`'s slot, made (and its project read) on the first request and moved to the back as
    /// the most recently used; `None` when `wb` has no project.
    fn slot(&mut self, wb: &Workbook) -> Option<&mut Slot> {
        self.slots.retain(|s| s.bytes.strong_count() > 0);
        let bytes = wb.vba_project.as_ref()?;
        let signature_part = wb.vba_signature_part;
        if let Some(i) = self.slots.iter().position(|s| std::ptr::eq(s.bytes.as_ptr(), Arc::as_ptr(bytes)) && s.signature_part == signature_part) {
            let slot = self.slots.remove(i);
            self.slots.push(slot);
            return self.slots.last_mut();
        }
        let project = Arc::new(crate::vba_source::read_workbook_vba(wb, &Limits::default()).unwrap_or_default());
        #[cfg(test)]
        {
            self.parses = self.parses.saturating_add(1);
        }
        if self.slots.len() >= MAX_SLOTS {
            self.slots.remove(0);
        }
        self.slots.push(Slot { bytes: Arc::downgrade(bytes), signature_part, project, report: None });
        self.slots.last_mut()
    }

    /// How many parsed projects are kept, for tests.
    #[cfg(test)]
    pub(crate) fn len(&self) -> usize {
        self.slots.len()
    }
}

#[cfg(test)]
mod tests;
