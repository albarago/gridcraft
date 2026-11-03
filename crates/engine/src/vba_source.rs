//! Glue between the compound file reader (`craft-cfb`) and the VBA project reader (design
//! Decision 1): `gridcraft-vba` never parses CFB itself, it reads streams through
//! [`gridcraft_vba::StreamSource`]. [`CompoundSource`] implements that trait over
//! [`craft_cfb::CompoundFile`], and [`read_vba_project`] reads a whole `vbaProject.bin`.
//! [`read_workbook_vba`] reads a workbook's project and adds what only the package knows: a VBA
//! signature part (`Workbook::vba_signature_part`, Req 7.5) marks the project "signed (not
//! verified)" as a signature stream in the bin does (Req 7.4), with one `Signed` diagnostic.
//!
//! Nothing here runs on open; the VBA cache (VBA-1.5) calls it on the first request.
use craft_cfb::{CompoundFile, EntryKind};
use gridcraft_model::Workbook;
use gridcraft_vba::{Diag, DiagCode, Limits, StreamSource, VbaProject, mark_signed, read_project};

/// The most storages a path may go through: a VBA project's own paths are two names deep
/// (`VBA/dir`), so a longer one is a mistake or a hostile caller, and is refused.
#[cfg_attr(not(test), expect(dead_code, reason = "used by the VBA cache in VBA-1.5"))]
const MAX_PATH_DEPTH: usize = 8;

/// A [`StreamSource`] over an opened compound file. Paths are relative to its root storage, which
/// for `vbaProject.bin` is the project root. Names match case-insensitively, as in CFB.
#[cfg_attr(not(test), expect(dead_code, reason = "used by the VBA cache in VBA-1.5"))]
pub(crate) struct CompoundSource<'a>(pub(crate) CompoundFile<'a>);

impl CompoundSource<'_> {
    /// The directory entry at `path`, one `child` step per name; `None` when a name is missing or
    /// a step would go through a stream. `&[]` is the root.
    #[cfg_attr(not(test), expect(dead_code, reason = "used by the VBA cache in VBA-1.5"))]
    fn entry_at(&self, path: &[&str]) -> Result<Option<usize>, String> {
        if path.len() > MAX_PATH_DEPTH {
            return Err(format!("the path goes through more than {MAX_PATH_DEPTH} storages"));
        }
        let mut at = 0;
        for name in path {
            if !self.0.entry(at).is_some_and(|e| matches!(e.kind, EntryKind::Root | EntryKind::Storage)) {
                return Ok(None);
            }
            match self.0.child(at, name) {
                Some(next) => at = next,
                None => return Ok(None),
            }
        }
        Ok(Some(at))
    }
}

impl StreamSource for CompoundSource<'_> {
    fn stream(&self, path: &[&str]) -> Result<Option<Vec<u8>>, String> {
        if path.is_empty() {
            return Err("a stream needs a path".into());
        }
        let Some(id) = self.entry_at(path)? else { return Ok(None) };
        if self.0.entry(id).is_none_or(|e| e.kind != EntryKind::Stream) {
            return Ok(None);
        }
        self.0.read(id).map(Some).map_err(|e| e.to_string())
    }

    fn children(&self, path: &[&str]) -> Result<Vec<String>, String> {
        let at = self.entry_at(path)?.ok_or_else(|| "no such storage".to_string())?;
        let storage =
            self.0.entry(at).filter(|e| matches!(e.kind, EntryKind::Root | EntryKind::Storage)).ok_or_else(|| "not a storage".to_string())?;
        let mut names: Vec<String> = storage.children.iter().filter_map(|&c| self.0.entry(c)).map(|c| c.name.clone()).collect();
        names.sort();
        Ok(names)
    }
}

/// Reads the VBA project in a `vbaProject.bin`. Never fails: a file that isn't a readable
/// compound file gives an empty project with one `DirCorrupt` diagnostic, since its `VBA/dir`
/// stream can't be reached; every other problem is reported by [`read_project`].
#[cfg_attr(not(test), expect(dead_code, reason = "used by the VBA cache in VBA-1.5"))]
pub(crate) fn read_vba_project(bytes: &[u8], limits: &Limits) -> VbaProject {
    match CompoundFile::open(bytes) {
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
#[cfg_attr(not(test), expect(dead_code, reason = "used by the VBA cache in VBA-1.5"))]
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
