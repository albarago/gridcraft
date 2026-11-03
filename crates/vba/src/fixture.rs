//! The Fixture_Builder: VBA projects built in code for tests (Requirement 10), so the repository
//! never holds an Office-produced file. Compiled only for tests and the `test-support` feature,
//! so release and web builds don't contain it. Fixtures live in memory; nothing here writes to
//! disk.
//!
//! A [`Fixture`] lists modules as (name, kind, source) and gives either the project's streams as
//! a [`MemSource`] ([`Fixture::to_streams`]) or a whole `vbaProject.bin` compound file
//! ([`Fixture::to_vba_project_bin`], written by the test-only [`cfb`] writer). Both hold the same
//! streams, from [MS-OVBA] §2.2–§2.3:
//!
//! - `PROJECT`: the project text (§2.3.1), CRLF lines in the code page: `ID=`, one
//!   `Module=`/`Class=`/`BaseClass=`/`Document=…/&H00000000` line per module, `Name=`,
//!   `HelpContextID=`, `VersionCompatible32=`, then `[Host Extender Info]` and `[Workspace]`.
//!   **Deviation:** the ABNF's `CMG=`, `DPB=` and `GC=` lines are left out. They hold data
//!   encrypted per §2.4.3, which nothing in this crate reads; writing them is follow-up work for
//!   whoever first needs them.
//! - `VBA/_VBA_PROJECT`: the 7-byte header of §2.3.4.1 (Reserved1 0x61CC, Version 0xFFFF as the
//!   spec requires on write, Reserved2 0x00, Reserved3, undefined, written as 0x0000) and no
//!   performance cache.
//! - `VBA/dir`: a [`DirFixture`] with one module per entry, compressed with [`ovba::compress`].
//!   MODULETYPE is procedural for [`ModuleKind::Standard`] and non-procedural otherwise, and
//!   MODULEOFFSET is the length of the module's `cache_prefix`.
//! - `VBA/<name>`: each module's `cache_prefix` bytes, then its source in the code page,
//!   compressed. The stream name is the module name.
//!
//! **Text.** Source and the `PROJECT` text are encoded with [`encode`]: UTF-8 for code page
//! 65001, otherwise the inverse of [`encoding::decode`] for that code page, so Windows-1252 and
//! Mac OS Roman text round-trips. A character the code page can't hold is written as `?`, so a
//! test that uses one sees its source change. The `dir` stream writes MBCS names as the `&str`'s
//! bytes ([`DirFixture`]), so module and project names should be ASCII unless the code page is
//! 65001.
//!
//! **Corruption.** [`Fixture::corrupt`] lists [`Corruption`]s, applied in order. A path or module
//! that names nothing is ignored. In the [`MemSource`], whose streams have no declared length or
//! sector chain, [`Corruption::OversizedLength`] and [`Corruption::CyclicChain`] make reading
//! that stream fail, which is what a compound file reader reports for them.
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
pub mod cfb;

use std::collections::BTreeMap;

use crate::dir::DirModuleType;
use crate::dir::fixture::{DirFixture, ModuleFixture};
use crate::encoding::{self, CP_UTF8};
use crate::ovba;
use crate::project::MemSource;
use crate::project_stream::ModuleKind;

pub use cfb::CfbError;

/// The fixed project GUID written to `ID=`.
pub const PROJECT_ID: &str = "{0F1C7E00-0000-4000-8000-000000000001}";

/// One module of a [`Fixture`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FixtureModule {
    pub name: String,
    pub kind: ModuleKind,
    pub source: String,
    /// Bytes stored before MODULEOFFSET (the compiled-code cache). Empty by default.
    pub cache_prefix: Vec<u8>,
}

impl FixtureModule {
    /// A module with no cache prefix.
    pub fn new(name: &str, kind: ModuleKind, source: &str) -> Self {
        Self { name: name.to_owned(), kind, source: source.to_owned(), cache_prefix: Vec::new() }
    }
}

/// Damage applied to a built fixture.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Corruption {
    /// Keep only the first `keep` bytes of the stream at `path` (for example `["VBA", "dir"]`).
    TruncateStream { path: Vec<String>, keep: usize },
    /// Set the module's MODULEOFFSET one byte past the end of its stream.
    OffsetPastEnd { module: String },
    /// Declare the stream larger than its data (see [`cfb::Damage::OversizedLength`]).
    OversizedLength { path: Vec<String> },
    /// Make the stream's sector chain loop (see [`cfb::Damage::CyclicChain`]).
    CyclicChain { path: Vec<String> },
}

/// A VBA project to build. [`Default`] is an empty project named `VBAProject` in Windows-1252.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Fixture {
    pub project_name: String,
    pub code_page: u16,
    pub modules: Vec<FixtureModule>,
    pub corrupt: Vec<Corruption>,
}

impl Default for Fixture {
    fn default() -> Self {
        Self { project_name: "VBAProject".to_owned(), code_page: encoding::CP_WINDOWS_1252, modules: Vec::new(), corrupt: Vec::new() }
    }
}

impl Fixture {
    /// A Windows-1252 project named `VBAProject` with these modules.
    pub fn new(modules: Vec<FixtureModule>) -> Self {
        Self { modules, ..Self::default() }
    }

    /// The project's streams, corruptions applied, as a [`MemSource`].
    pub fn to_streams(&self) -> MemSource {
        let mut src = MemSource::new();
        for (path, bytes) in self.streams() {
            let p: Vec<&str> = path.iter().map(String::as_str).collect();
            src.insert(&p, bytes);
        }
        for c in &self.corrupt {
            if let Corruption::OversizedLength { path } | Corruption::CyclicChain { path } = c {
                let p: Vec<&str> = path.iter().map(String::as_str).collect();
                src = src.failing(&p);
            }
        }
        src
    }

    /// The project as a `vbaProject.bin` compound file, corruptions applied.
    ///
    /// # Errors
    ///
    /// A [`CfbError`] when a module name isn't a valid compound file name (1 to 31 UTF-16
    /// units, no `/ \ : !`), two modules share a name, the file would pass 128 MiB, or a
    /// [`Corruption::CyclicChain`] targets a stream that fits in one sector.
    pub fn to_vba_project_bin(&self) -> Result<Vec<u8>, CfbError> {
        let mut specs: Vec<cfb::StreamSpec> =
            self.streams().into_iter().map(|(path, data)| cfb::StreamSpec { path, data, damage: cfb::Damage::None }).collect();
        for c in &self.corrupt {
            let (path, damage) = match c {
                Corruption::OversizedLength { path } => (path, cfb::Damage::OversizedLength),
                Corruption::CyclicChain { path } => (path, cfb::Damage::CyclicChain),
                Corruption::TruncateStream { .. } | Corruption::OffsetPastEnd { .. } => continue,
            };
            if let Some(spec) = specs.iter_mut().find(|s| same_path(&s.path, path)) {
                spec.damage = damage;
            }
        }
        cfb::write(&specs)
    }

    /// The decompressed `dir` stream's fields for this project, before corruption.
    pub fn dir_fixture(&self) -> DirFixture {
        let modules = self
            .modules
            .iter()
            .map(|m| {
                let module_type = if m.kind == ModuleKind::Standard { DirModuleType::Procedural } else { DirModuleType::NonProcedural };
                ModuleFixture { text_offset: u32::try_from(m.cache_prefix.len()).unwrap_or(u32::MAX), ..ModuleFixture::new(&m.name, module_type) }
            })
            .collect();
        DirFixture { code_page: self.code_page, project_name: self.project_name.clone(), modules, ..DirFixture::default() }
    }

    /// The `PROJECT` stream text, encoded in the code page.
    pub fn project_text(&self) -> Vec<u8> {
        let mut text = format!("ID=\"{PROJECT_ID}\"\r\n");
        for m in &self.modules {
            let line = match m.kind {
                ModuleKind::Standard => format!("Module={}", m.name),
                ModuleKind::Class => format!("Class={}", m.name),
                ModuleKind::Form => format!("BaseClass={}", m.name),
                ModuleKind::Document => format!("Document={}/&H00000000", m.name),
            };
            text.push_str(&line);
            text.push_str("\r\n");
        }
        text.push_str(&format!("Name=\"{}\"\r\nHelpContextID=\"0\"\r\nVersionCompatible32=\"393222000\"\r\n\r\n", self.project_name));
        text.push_str("[Host Extender Info]\r\n&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n\r\n[Workspace]\r\n");
        for m in &self.modules {
            text.push_str(&format!("{}=0, 0, 0, 0, C\r\n", m.name));
        }
        encode(self.code_page, &text)
    }

    /// A module's stream: its cache prefix, then its compressed source.
    pub fn module_stream(&self, module: &FixtureModule) -> Vec<u8> {
        let mut bytes = module.cache_prefix.clone();
        bytes.extend_from_slice(&ovba::compress(&encode(self.code_page, &module.source)));
        bytes
    }

    /// Every stream by path, with [`Corruption::TruncateStream`] and
    /// [`Corruption::OffsetPastEnd`] applied.
    fn streams(&self) -> Vec<(Vec<String>, Vec<u8>)> {
        let path = |p: &[&str]| p.iter().map(|s| (*s).to_owned()).collect::<Vec<_>>();
        let mut streams = vec![(path(&["PROJECT"]), self.project_text()), (path(&["VBA", "_VBA_PROJECT"]), vba_project_header().to_vec())];
        for m in &self.modules {
            streams.push((path(&["VBA", &m.name]), self.module_stream(m)));
        }
        self.truncate(&mut streams);
        // The `dir` stream is built last, so MODULEOFFSET can point past the final stream length.
        let mut dir = self.dir_fixture();
        for c in &self.corrupt {
            let Corruption::OffsetPastEnd { module } = c else { continue };
            let Some(m) = dir.modules.iter_mut().find(|m| m.name.eq_ignore_ascii_case(module)) else { continue };
            let stream_path = path(&["VBA", &m.stream_name]);
            let len = streams.iter().find(|(p, _)| same_path(p, &stream_path)).map_or(0, |(_, b)| b.len());
            m.text_offset = u32::try_from(len.saturating_add(1)).unwrap_or(u32::MAX);
        }
        let mut with_dir = vec![(path(&["VBA", "dir"]), ovba::compress(&dir.to_bytes()))];
        self.truncate(&mut with_dir);
        streams.extend(with_dir);
        streams
    }

    /// Applies every [`Corruption::TruncateStream`] to the matching streams.
    fn truncate(&self, streams: &mut [(Vec<String>, Vec<u8>)]) {
        for c in &self.corrupt {
            let Corruption::TruncateStream { path, keep } = c else { continue };
            for (p, bytes) in streams.iter_mut() {
                if same_path(p, path) {
                    bytes.truncate(*keep);
                }
            }
        }
    }
}

/// The `_VBA_PROJECT` stream ([MS-OVBA] §2.3.4.1) with no performance cache.
///
/// [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
pub const fn vba_project_header() -> [u8; 7] {
    [0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]
}

/// `text` as MBCS bytes in `code_page`: UTF-8 for 65001, else the inverse of
/// [`encoding::decode`]. Characters the code page can't hold become `?`.
pub fn encode(code_page: u16, text: &str) -> Vec<u8> {
    if code_page == CP_UTF8 {
        return text.as_bytes().to_vec();
    }
    let mut table = BTreeMap::new();
    for b in 0..=u8::MAX {
        if let Some(c) = encoding::decode(code_page, &[b]).0.chars().next() {
            table.entry(c).or_insert(b);
        }
    }
    text.chars().map(|c| table.get(&c).copied().unwrap_or(b'?')).collect()
}

/// Whether two paths are equal, matching names ASCII case-insensitively.
fn same_path(a: &[String], b: &[String]) -> bool {
    a.len() == b.len() && a.iter().zip(b).all(|(x, y)| x.eq_ignore_ascii_case(y))
}

#[cfg(test)]
mod tests;
