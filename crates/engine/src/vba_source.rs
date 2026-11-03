//! Glue between the xlsx compound file reader and the VBA project reader (design Decision 1):
//! `gridcraft-vba` never parses CFB itself, it reads streams through
//! [`gridcraft_vba::StreamSource`]. [`CompoundSource`] implements that trait over
//! [`gridcraft_xlsx::Compound`], and [`read_vba_project`] reads a whole `vbaProject.bin`.
//! [`read_workbook_vba`] reads a workbook's project and adds what only the package knows: a VBA
//! signature part (`Workbook::vba_signature_part`, Req 7.5) marks the project "signed (not
//! verified)" as a signature stream in the bin does (Req 7.4), with one `Signed` diagnostic.
//!
//! Nothing here runs on open; [`crate::vba_cache`] calls it on the first request.
use gridcraft_model::Workbook;
use gridcraft_vba::{Diag, DiagCode, Limits, StreamSource, VbaProject, mark_signed, read_project};
use gridcraft_xlsx::Compound;

/// A [`StreamSource`] over an opened compound file. Paths are relative to its root storage, which
/// for `vbaProject.bin` is the project root. Names match case-insensitively, as in CFB.
pub(crate) struct CompoundSource<'a>(pub(crate) Compound<'a>);

impl StreamSource for CompoundSource<'_> {
    fn stream(&self, path: &[&str]) -> Result<Option<Vec<u8>>, String> {
        self.0.stream_at_path(path)
    }

    fn children(&self, path: &[&str]) -> Result<Vec<String>, String> {
        self.0.children_at_path(path)
    }
}

/// Reads the VBA project in a `vbaProject.bin`. Never fails: a file that isn't a readable
/// compound file gives an empty project with one `DirCorrupt` diagnostic, since its `VBA/dir`
/// stream can't be reached; every other problem is reported by [`read_project`].
pub(crate) fn read_vba_project(bytes: &[u8], limits: &Limits) -> VbaProject {
    match Compound::open(bytes) {
        Ok(cfb) => read_project(&CompoundSource(cfb), limits),
        Err(e) => VbaProject {
            diagnostics: vec![Diag::project(DiagCode::DirCorrupt, format!("vbaProject.bin can't be opened: {e}"))],
            ..VbaProject::default()
        },
    }
}

/// Reads `wb`'s VBA project, or `None` when it has none. A VBA signature part in the package
/// marks the project signed even when the bin has no signature stream; both together still give
/// one `Signed` diagnostic. Nothing is verified.
pub(crate) fn read_workbook_vba(wb: &Workbook, limits: &Limits) -> Option<VbaProject> {
    let bytes = wb.vba_project.as_deref()?;
    let mut project = read_vba_project(bytes, limits);
    if wb.vba_signature_part {
        mark_signed(&mut project);
    }
    Some(project)
}
#[cfg(test)]
mod props;
#[cfg(test)]
mod tests;
