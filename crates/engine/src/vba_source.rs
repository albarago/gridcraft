//! Glue between the xlsx compound file reader and the VBA project reader (design Decision 1):
//! `gridcraft-vba` never parses CFB itself, it reads streams through
//! [`gridcraft_vba::StreamSource`]. [`CompoundSource`] implements that trait over
//! [`gridcraft_xlsx::Compound`], and [`read_vba_project`] reads a whole `vbaProject.bin`.
//!
//! Nothing here runs on open; the VBA cache (VBA-1.5) calls it on the first request.
use gridcraft_vba::{Diag, DiagCode, Limits, StreamSource, VbaProject, read_project};
use gridcraft_xlsx::Compound;

/// A [`StreamSource`] over an opened compound file. Paths are relative to its root storage, which
/// for `vbaProject.bin` is the project root. Names match case-insensitively, as in CFB.
#[cfg_attr(not(test), expect(dead_code, reason = "used by the VBA cache in VBA-1.5"))]
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
#[cfg_attr(not(test), expect(dead_code, reason = "used by the VBA cache in VBA-1.5"))]
pub(crate) fn read_vba_project(bytes: &[u8], limits: &Limits) -> VbaProject {
    match Compound::open(bytes) {
        Ok(cfb) => read_project(&CompoundSource(cfb), limits),
        Err(e) => VbaProject {
            diagnostics: vec![Diag::project(DiagCode::DirCorrupt, format!("vbaProject.bin can't be opened: {e}"))],
            ..VbaProject::default()
        },
    }
}

#[cfg(test)]
mod props;
#[cfg(test)]
mod tests;
