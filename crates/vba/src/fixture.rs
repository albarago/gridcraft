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
//! - `\u{5}DigitalSignature`, `…Ex`, `…Agile` in the root: one per [`Fixture::signatures`] entry
//!   ([`SignatureKind`], names from [`SIGNATURE_STREAMS`]), holding [`SIGNATURE_PLACEHOLDER`].
//!
//! **P_Code region.** The task and Requirement 10.6 call it a P_Code region with an identifier
//! table; here it is a module's `cache_prefix`, the bytes before MODULEOFFSET where
//! [MS-OVBA] §2.3.4.3 puts the compiled-code cache. [`FixtureModule::with_pcode`] fills it with
//! [`PCODE_PLACEHOLDER`] repeated. The bytes are arbitrary on purpose: the cache's layout,
//! identifier table included, is undocumented, the reader only counts the bytes (through
//! MODULEOFFSET) and never decodes them, and the clean-room rule forbids copying real compiled
//! code. So no identifier table is written; that waits for public documentation of the cache,
//! like Requirement 7 criterion 2 (see the crate docs).
//!
//! **Signatures.** A signature stream holds [`SIGNATURE_PLACEHOLDER`], plain text that is not a
//! PKCS #7 blob, and never a real signature or certificate: the reader only looks at stream
//! names, and nothing signed is ever committed. The names start with U+0005, a control
//! character. [MS-CFB] §2.6.1 forbids only `/`, `\`, `:` and `!` in names (plus the 31 UTF-16
//! unit limit), so control characters are allowed and the [`cfb`] writer accepts them as they
//! are.
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
//! [`Corruption::RemoveStream`] and [`Corruption::PadProjectStream`] change the stream list
//! itself, so they act the same on both outputs.
//!
//! **One fixture per diagnostic.** [`diag_fixture`] gives, for every [`DiagCode`], a fixture
//! and [`Limits`] whose reading emits that code, through either output. Its `match` is
//! exhaustive, so a new code doesn't compile until it has a fixture; add the code to
//! [`ALL_DIAG_CODES`] as well, which the tests walk.
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
//! [MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
pub mod cfb;

use std::collections::BTreeMap;

use crate::diag::DiagCode;
use crate::dir::DirModuleType;
use crate::dir::fixture::{DirFixture, ModuleFixture};
use crate::encoding::{self, CP_UTF8};
use crate::ovba;
use crate::project::{Limits, MemSource, SIGNATURE_STREAMS};
use crate::project_stream::{MAX_PROJECT_STREAM, ModuleKind};

pub use cfb::CfbError;

/// The fixed project GUID written to `ID=`.
pub const PROJECT_ID: &str = "{0F1C7E00-0000-4000-8000-000000000001}";

/// The bytes [`FixtureModule::with_pcode`] repeats. Not compiled code: see the module docs.
pub const PCODE_PLACEHOLDER: &[u8] = b"fixture p-code placeholder, never decoded. ";

/// The most bytes [`FixtureModule::with_pcode`] writes (the 16 MiB stream cap).
pub const MAX_PCODE: usize = ovba::MAX_DECOMPRESSED;

/// The bytes of every fixture signature stream. Not a PKCS #7 blob: see the module docs.
pub const SIGNATURE_PLACEHOLDER: &[u8] = b"fixture signature, not a real PKCS#7 blob";

/// The filler line [`Corruption::PadProjectStream`] repeats: neither a property nor a section.
const PROJECT_FILLER: &[u8] = b"; fixture padding\r\n";

/// A VBA signature stream in the project root. A real one holds the PKCS #7 structure of
/// [MS-OSHARED] §2.3.2; a fixture one holds [`SIGNATURE_PLACEHOLDER`].
///
/// [MS-OSHARED]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SignatureKind {
    /// `\u{5}DigitalSignature`.
    Legacy,
    /// `\u{5}DigitalSignatureEx`.
    Ex,
    /// `\u{5}DigitalSignatureAgile`.
    Agile,
}

impl SignatureKind {
    /// Every kind, in [`SIGNATURE_STREAMS`] order.
    pub const ALL: [Self; 3] = [Self::Legacy, Self::Ex, Self::Agile];

    /// The stream name, from [`SIGNATURE_STREAMS`].
    #[must_use]
    pub const fn stream_name(self) -> &'static str {
        match self {
            Self::Legacy => SIGNATURE_STREAMS[0],
            Self::Ex => SIGNATURE_STREAMS[1],
            Self::Agile => SIGNATURE_STREAMS[2],
        }
    }
}

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
    /// This module with a P_Code region of `len` placeholder bytes before MODULEOFFSET
    /// ([`PCODE_PLACEHOLDER`] repeated, at most [`MAX_PCODE`]). The same `len` always gives the
    /// same bytes. `0` removes the region.
    #[must_use]
    pub fn with_pcode(self, len: usize) -> Self {
        let cache = PCODE_PLACEHOLDER.iter().copied().cycle().take(len.min(MAX_PCODE)).collect();
        self.with_cache(cache)
    }
    /// This module with `bytes` as its P_Code region (the cache before MODULEOFFSET).
    #[must_use]
    pub fn with_cache(mut self, bytes: Vec<u8>) -> Self {
        self.cache_prefix = bytes;
        self
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
    /// Leave the stream at `path` out (for example `["PROJECT"]` or `["VBA", "Module1"]`).
    RemoveStream { path: Vec<String> },
    /// Insert more than [`MAX_PROJECT_STREAM`] bytes of filler lines right after the `PROJECT`
    /// stream's `ID=` line, so its module lines fall past the reader's cap. The stream is then
    /// over 1 MiB, so it is written in regular sectors.
    PadProjectStream,
}

/// A VBA project to build. [`Default`] is an empty, unsigned project named `VBAProject` in
/// Windows-1252.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Fixture {
    pub project_name: String,
    pub code_page: u16,
    pub modules: Vec<FixtureModule>,
    /// Signature streams to add to the root, each holding [`SIGNATURE_PLACEHOLDER`]. A kind
    /// listed twice is written once.
    pub signatures: Vec<SignatureKind>,
    pub corrupt: Vec<Corruption>,
}

impl Default for Fixture {
    fn default() -> Self {
        Self {
            project_name: "VBAProject".to_owned(),
            code_page: encoding::CP_WINDOWS_1252,
            modules: Vec::new(),
            signatures: Vec::new(),
            corrupt: Vec::new(),
        }
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
    /// units, no `/ \ : !`; control characters such as the signature streams' U+0005 are
    /// allowed), two modules share a name, the file would pass 128 MiB, or a
    /// [`Corruption::CyclicChain`] targets a stream that fits in one sector.
    pub fn to_vba_project_bin(&self) -> Result<Vec<u8>, CfbError> {
        let mut specs: Vec<cfb::StreamSpec> =
            self.streams().into_iter().map(|(path, data)| cfb::StreamSpec { path, data, damage: cfb::Damage::None }).collect();
        for c in &self.corrupt {
            let (path, damage) = match c {
                Corruption::OversizedLength { path } => (path, cfb::Damage::OversizedLength),
                Corruption::CyclicChain { path } => (path, cfb::Damage::CyclicChain),
                Corruption::TruncateStream { .. }
                | Corruption::OffsetPastEnd { .. }
                | Corruption::RemoveStream { .. }
                | Corruption::PadProjectStream => continue,
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

    /// Every stream by path, with [`Corruption::TruncateStream`], [`Corruption::OffsetPastEnd`],
    /// [`Corruption::PadProjectStream`] and [`Corruption::RemoveStream`] applied.
    fn streams(&self) -> Vec<(Vec<String>, Vec<u8>)> {
        let path = |p: &[&str]| p.iter().map(|s| (*s).to_owned()).collect::<Vec<_>>();
        let mut project = self.project_text();
        if self.corrupt.contains(&Corruption::PadProjectStream) {
            project = pad_after_first_line(&project);
        }
        let mut streams = vec![(path(&["PROJECT"]), project), (path(&["VBA", "_VBA_PROJECT"]), vba_project_header().to_vec())];
        for kind in SignatureKind::ALL.into_iter().filter(|k| self.signatures.contains(k)) {
            streams.push((path(&[kind.stream_name()]), SIGNATURE_PLACEHOLDER.to_vec()));
        }
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
        for c in &self.corrupt {
            if let Corruption::RemoveStream { path } = c {
                streams.retain(|(p, _)| !same_path(p, path));
            }
        }
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

/// `text` with filler lines inserted after its first line, adding more than
/// [`MAX_PROJECT_STREAM`] bytes ([`Corruption::PadProjectStream`]).
fn pad_after_first_line(text: &[u8]) -> Vec<u8> {
    let split = text.iter().position(|&b| b == b'\n').map_or(text.len(), |i| i.saturating_add(1));
    let (head, tail) = text.split_at(split);
    let lines = MAX_PROJECT_STREAM.div_ceil(PROJECT_FILLER.len()).saturating_add(1);
    let mut out = Vec::with_capacity(text.len().saturating_add(lines.saturating_mul(PROJECT_FILLER.len())));
    out.extend_from_slice(head);
    for _ in 0..lines {
        out.extend_from_slice(PROJECT_FILLER);
    }
    out.extend_from_slice(tail);
    out
}

/// A fixture whose reading emits one diagnostic code: see [`diag_fixture`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagFixture {
    pub fixture: Fixture,
    /// The limits to read it with.
    pub limits: Limits,
    /// The module the diagnostic is attached to, or `None` for a project-level diagnostic.
    pub module: Option<&'static str>,
}

/// Every [`DiagCode`], once each. [`DiagCode::UnsupportedCodePage`] carries code page 1251,
/// which [`diag_fixture`] uses.
pub const ALL_DIAG_CODES: [DiagCode; 15] = [
    DiagCode::DirMissing,
    DiagCode::DirCorrupt,
    DiagCode::ProjectStreamMissing,
    DiagCode::ProjectStreamTruncated,
    DiagCode::KindFromModuleType,
    DiagCode::ModuleStreamMissing,
    DiagCode::ModuleOffsetOutOfRange,
    DiagCode::ModuleTruncated,
    DiagCode::ModuleCorrupt,
    DiagCode::UnsupportedCodePage(1251),
    DiagCode::ModuleLimit,
    DiagCode::SourceLimit,
    DiagCode::PossibleStomping,
    DiagCode::CompiledCodeNotChecked,
    DiagCode::Signed,
];

/// The module every [`diag_fixture`] damages: `Module1`, a standard module whose source spans
/// several mini sectors once compressed (so a cyclic chain and a cut in the middle both apply).
pub const DIAG_MODULE: &str = "Module1";

/// A fixture whose reading with [`DiagFixture::limits`] emits `code` (Requirement 10.6), from
/// either [`Fixture::to_streams`] or [`Fixture::to_vba_project_bin`] through a compound file
/// reader. The `match` is exhaustive on purpose: a new [`DiagCode`] doesn't compile until it has
/// a fixture here. [`DiagCode::UnsupportedCodePage`] uses the code page it carries.
#[must_use]
pub fn diag_fixture(code: DiagCode) -> DiagFixture {
    let source: String = (0..40).map(|i| format!("Public V{i} As Long ' {}\r\n", i * 7919)).collect();
    let module1 = FixtureModule::new(DIAG_MODULE, ModuleKind::Standard, &source);
    let class1 = FixtureModule::new("Class1", ModuleKind::Class, "Attribute VB_Name = \"Class1\"\r\nPublic X As Long\r\n");
    let base = Fixture::new(vec![module1.clone(), class1.clone()]);
    let vba = |name: &str| vec!["VBA".to_owned(), name.to_owned()];
    let damaged = |c: Corruption| Fixture { corrupt: vec![c], ..base.clone() };
    let half = base.module_stream(&module1).len() / 2;
    let project = |fixture: Fixture| DiagFixture { fixture, limits: Limits::default(), module: None };
    let module = |fixture: Fixture| DiagFixture { fixture, limits: Limits::default(), module: Some(DIAG_MODULE) };
    match code {
        DiagCode::DirMissing => project(damaged(Corruption::RemoveStream { path: vba("dir") })),
        DiagCode::DirCorrupt => project(damaged(Corruption::TruncateStream { path: vba("dir"), keep: 40 })),
        DiagCode::ProjectStreamMissing => project(damaged(Corruption::RemoveStream { path: vec!["PROJECT".to_owned()] })),
        DiagCode::ProjectStreamTruncated => project(damaged(Corruption::PadProjectStream)),
        // With the module lines past the cap, no module is listed in the PROJECT stream.
        DiagCode::KindFromModuleType => module(damaged(Corruption::PadProjectStream)),
        DiagCode::ModuleStreamMissing => module(damaged(Corruption::RemoveStream { path: vba(DIAG_MODULE) })),
        DiagCode::ModuleOffsetOutOfRange => module(damaged(Corruption::OffsetPastEnd { module: DIAG_MODULE.to_owned() })),
        DiagCode::ModuleTruncated => module(damaged(Corruption::TruncateStream { path: vba(DIAG_MODULE), keep: half })),
        DiagCode::ModuleCorrupt => module(damaged(Corruption::CyclicChain { path: vba(DIAG_MODULE) })),
        DiagCode::UnsupportedCodePage(code_page) => project(Fixture { code_page, ..base }),
        DiagCode::ModuleLimit => DiagFixture { limits: Limits { max_modules: 1, ..Limits::default() }, ..project(base) },
        DiagCode::SourceLimit => DiagFixture { limits: Limits { max_total_source: source.len(), ..Limits::default() }, ..project(base) },
        DiagCode::PossibleStomping => {
            let stomped = FixtureModule::new(DIAG_MODULE, ModuleKind::Standard, "Attribute VB_Name = \"Module1\"\r\n").with_pcode(256);
            module(Fixture::new(vec![stomped, class1]))
        }
        DiagCode::CompiledCodeNotChecked => module(Fixture::new(vec![module1.with_pcode(256), class1])),
        DiagCode::Signed => project(Fixture { signatures: SignatureKind::ALL.to_vec(), ..base }),
    }
}

/// Whether two paths are equal, matching names ASCII case-insensitively.
fn same_path(a: &[String], b: &[String]) -> bool {
    a.len() == b.len() && a.iter().zip(b).all(|(x, y)| x.eq_ignore_ascii_case(y))
}

#[cfg(test)]
mod tests;
