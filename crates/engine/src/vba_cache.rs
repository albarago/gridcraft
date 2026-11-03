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
use std::sync::{Arc, Weak};

use gridcraft_model::Workbook;
use gridcraft_vba::{Limits, VbaProject};

/// Most parsed projects kept: a handful of open workbooks, with room for edits in between.
pub(crate) const MAX_SLOTS: usize = 16;

struct Slot {
    bytes: Weak<Vec<u8>>,
    signature_part: bool,
    project: Arc<VbaProject>,
}

#[derive(Default)]
pub(crate) struct VbaCache {
    /// Least recently used first.
    slots: Vec<Slot>,
    /// How many times a project was read, for tests.
    #[cfg(test)]
    pub(crate) parses: usize,
}

impl VbaCache {
    /// `wb`'s VBA project, read on the first request and reused while the same bytes are alive.
    /// A workbook without a project gives the empty project (no diagnostics) and reads nothing
    /// (Req 6.2); bytes that can't be read give a project carrying only diagnostics (Req 6.5).
    pub(crate) fn project(&mut self, wb: &Workbook) -> Arc<VbaProject> {
        self.slots.retain(|s| s.bytes.strong_count() > 0);
        let Some(bytes) = wb.vba_project.as_ref() else {
            return Arc::new(VbaProject::default());
        };
        let signature_part = wb.vba_signature_part;
        if let Some(i) = self.slots.iter().position(|s| std::ptr::eq(s.bytes.as_ptr(), Arc::as_ptr(bytes)) && s.signature_part == signature_part) {
            let slot = self.slots.remove(i);
            let project = slot.project.clone();
            self.slots.push(slot);
            return project;
        }
        let project = Arc::new(crate::vba_source::read_workbook_vba(wb, &Limits::default()).unwrap_or_default());
        #[cfg(test)]
        {
            self.parses = self.parses.saturating_add(1);
        }
        if self.slots.len() >= MAX_SLOTS {
            self.slots.remove(0);
        }
        self.slots.push(Slot { bytes: Arc::downgrade(bytes), signature_part, project: project.clone() });
        project
    }

    /// How many parsed projects are kept, for tests.
    #[cfg(test)]
    pub(crate) fn len(&self) -> usize {
        self.slots.len()
    }
}

#[cfg(test)]
mod tests;
