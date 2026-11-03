//! Reads a whole VBA project from its streams: [`read_project`].
//!
//! The caller opens the compound file and hands the streams over through [`StreamSource`], so
//! this crate never parses CFB. The flow ([MS-OVBA] §2.3, Req 5):
//!
//! 1. `VBA/dir` is decompressed and parsed. If it is missing or bad, the result has no modules
//!    and one `DirMissing` or `DirCorrupt` diagnostic (Req 5.3).
//! 2. The `PROJECT` stream (plain MBCS text, not compressed) gives each module's kind. When it is
//!    missing, or doesn't list a module, the kind comes from MODULETYPE (Req 5.7–5.8).
//! 3. Each module stream `VBA/<stream name>` is read from its MODULEOFFSET, decompressed, cut
//!    before any trailing 0x00 bytes (a short raw last chunk's padding, §2.4.1.3.10) and
//!    decoded in the project code page. A failure leaves that module listed with empty source
//!    and a diagnostic; the other modules are still read (Req 5.9).
//! 4. A module that was read, has compiled code before its MODULEOFFSET and whose source is only
//!    blank and `Attribute` lines gets a `PossibleStomping` diagnostic (Req 7.1, see
//!    `possible_stomping`). Any other module that was read and has compiled code gets a
//!    `CompiledCodeNotChecked` diagnostic (Req 7.3), because the compiled code is never decoded
//!    and the identifier comparison of Req 7.2 is not implemented (see the crate docs). A
//!    stomped module gets only `PossibleStomping`: it is the stronger statement, and adding
//!    "not checked" would only repeat that the compiled code wasn't read. A module with no
//!    compiled code (`text_offset == 0`) has nothing to check, and a module whose source couldn't
//!    be read gets only its read diagnostic.
//! 5. The project root is listed with [`StreamSource::children`]. A stream named
//!    `\u{5}DigitalSignature`, `\u{5}DigitalSignatureEx` or `\u{5}DigitalSignatureAgile`
//!    ([`SIGNATURE_STREAMS`]) marks the project signed, with one `Signed` diagnostic (Req 7.4,
//!    see [`mark_signed`]). Only the names are looked at: the signature streams are never read,
//!    parsed or verified, so the status is "signed (not verified)" (Req 7.7-7.8). This step runs
//!    even when `VBA/dir` can't be read, so a damaged project is still reported as signed. If the
//!    root can't be listed, the project is treated as unsigned and no diagnostic is added: the
//!    signature check is advisory, and a root that can't be listed already breaks the reads above,
//!    which report their own diagnostics. A VBA signature part in the workbook package (Req 7.5)
//!    isn't visible here; the caller that sees the package calls [`mark_signed`] for it.
//!
//! Limits ([`Limits`]): at most 1,024 modules (Req 5.10), at most 64 MiB of decoded source in
//! total (Req 5.11), and at most 16 MiB per decompressed stream. [`read_project`] never fails and
//! never panics: every problem becomes a [`Diag`].
//!
//! Names. MODULENAME (§2.3.4.2.3.2.1) is the module's VBA identifier in the project code page;
//! MODULENAMEUNICODE (§2.3.4.2.3.2.2) is the same identifier in UTF-16. [`VbaModule::name`] is
//! for display, so it prefers MODULENAMEUNICODE when present and non-empty and falls back to
//! MODULENAME. [`VbaModule::code_name`] is always MODULENAME decoded in the project code page:
//! that is the identifier the `PROJECT` stream's `Document=`/`Module=` lines and the module's
//! `Attribute VB_Name` use, so later stages match against it. In a conforming file both are equal;
//! they differ only when the file is inconsistent or the code page isn't supported.
//! [`VbaModule::stream_name`] prefers StreamNameUnicode over the MBCS StreamName, because
//! compound file names are UTF-16.
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/

#[cfg(any(test, feature = "test-support"))]
mod mem;
#[cfg(any(test, feature = "test-support"))]
pub use mem::MemSource;

use crate::diag::{Diag, DiagCode};
use crate::dir::{self, DirInfo, DirModule, ReferenceKind};
use crate::encoding;
use crate::ovba;
use crate::project_stream::{self, ModuleKind, ProjectStreamInfo};

/// The most decoded source text kept across all modules (64 MiB, Req 5.11).
pub const MAX_TOTAL_SOURCE: usize = 64 * 1024 * 1024;

/// Streams of a VBA project storage (`vbaProject.bin`, or the `_VBA_PROJECT_CUR` storage of a
/// legacy workbook). Paths are storage names ending in a stream name, relative to the project
/// root, for example `["VBA", "dir"]` or `["PROJECT"]`. Implementations match names the way the
/// compound file does (case-insensitively).
pub trait StreamSource {
    /// The stream at `path`, or `Ok(None)` when it doesn't exist.
    ///
    /// # Errors
    ///
    /// A message when the stream exists but can't be read.
    fn stream(&self, path: &[&str]) -> Result<Option<Vec<u8>>, String>;
    /// The names directly under the storage at `path` (`&[]` is the project root).
    ///
    /// # Errors
    ///
    /// A message when the storage can't be listed.
    fn children(&self, path: &[&str]) -> Result<Vec<String>, String>;
}

/// Caps that keep a hostile project's cost bounded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// The most modules kept (default 1,024).
    pub max_modules: usize,
    /// The most decoded source bytes (UTF-8) kept across all modules (default 64 MiB).
    pub max_total_source: usize,
    /// The most bytes one stream may decompress to (default 16 MiB).
    pub max_stream: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self { max_modules: dir::MAX_MODULES, max_total_source: MAX_TOTAL_SOURCE, max_stream: ovba::MAX_DECOMPRESSED }
    }
}

/// A read VBA project. [`Default`] is the empty project.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct VbaProject {
    /// PROJECTNAME, decoded; `None` when the `dir` stream couldn't be read or the name is empty.
    pub name: Option<String>,
    /// PROJECTCODEPAGE; `None` when the `dir` stream couldn't be read.
    pub code_page: Option<u16>,
    pub references: Vec<Reference>,
    /// The modules in `dir` stream order.
    pub modules: Vec<VbaModule>,
    /// Whether the project carries a signature (a signature stream, or a package signature part
    /// reported through [`mark_signed`]). Never verified: "signed (not verified)".
    pub signed: bool,
    /// Workbook-level diagnostics.
    pub diagnostics: Vec<Diag>,
}

/// One project reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reference {
    /// REFERENCENAME, decoded; empty when the reference has no name record.
    pub name: String,
    pub kind: ReferenceKind,
}

/// One module.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaModule {
    /// The display name: MODULENAMEUNICODE, else MODULENAME (see the module docs).
    pub name: String,
    /// MODULENAME in the project code page: the VBA identifier.
    pub code_name: String,
    /// The module stream's name in the `VBA` storage.
    pub stream_name: String,
    pub kind: ModuleKind,
    /// MODULEOFFSET: where the compressed source starts in the module stream. Bytes before it are
    /// the compiled-code cache.
    pub text_offset: u32,
    /// The decoded source; empty when the module couldn't be read.
    pub source: String,
    pub diagnostics: Vec<Diag>,
}

/// Names of the VBA signature streams in the project root. Each starts with U+0005 and compares
/// case-insensitively, like every compound file name. The signature inside is the PKCS #7
/// structure of [MS-OSHARED] §2.3.2; GridCraft never reads it.
///
/// [MS-OSHARED]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/
pub const SIGNATURE_STREAMS: [&str; 3] = ["\u{5}DigitalSignature", "\u{5}DigitalSignatureEx", "\u{5}DigitalSignatureAgile"];

/// The message of the `Signed` diagnostic.
const SIGNED_MESSAGE: &str = "The VBA project is signed. GridCraft doesn't verify signatures (signed, not verified).";

/// Reads a VBA project. Never fails: problems become diagnostics (see the module docs).
#[must_use]
pub fn read_project(src: &dyn StreamSource, limits: &Limits) -> VbaProject {
    let mut project = read_contents(src, limits);
    if has_signature_stream(src) {
        mark_signed(&mut project);
    }
    project
}

/// Marks `project` "signed (not verified)": sets [`VbaProject::signed`] and adds one
/// project-level `Signed` diagnostic. Calling it again changes nothing, so a project with both a
/// signature stream and a package signature part (Req 7.4-7.5) gets one diagnostic.
pub fn mark_signed(project: &mut VbaProject) {
    project.signed = true;
    if !project.diagnostics.iter().any(|d| d.code == DiagCode::Signed) {
        project.diagnostics.push(Diag::project(DiagCode::Signed, SIGNED_MESSAGE));
    }
}

/// Whether the project root lists a signature stream. Only names are compared; no stream is read.
/// A root that can't be listed counts as unsigned (see the module docs).
fn has_signature_stream(src: &dyn StreamSource) -> bool {
    src.children(&[]).is_ok_and(|names| names.iter().any(|n| SIGNATURE_STREAMS.iter().any(|s| s.eq_ignore_ascii_case(n))))
}

/// Steps 1-4 of the module docs: the `dir` stream, the `PROJECT` stream and the modules.
fn read_contents(src: &dyn StreamSource, limits: &Limits) -> VbaProject {
    let mut project = VbaProject::default();
    let info = match read_dir(src, limits) {
        Ok(info) => info,
        Err(diag) => {
            project.diagnostics.push(diag);
            return project;
        }
    };
    let code_page = info.code_page;
    project.code_page = Some(code_page);
    if !encoding::is_supported(code_page) {
        project.diagnostics.push(Diag::project(
            DiagCode::UnsupportedCodePage(code_page),
            format!("The macro project uses code page {code_page}, which isn't supported, so its text was read as Windows-1252."),
        ));
    }
    let name = mbcs(code_page, &info.name);
    project.name = (!name.is_empty()).then_some(name);
    project.references = info.references.iter().map(|r| Reference { name: mbcs(code_page, &r.name), kind: r.kind }).collect();
    if let Some(declared) = info.truncated_modules {
        project.diagnostics.push(Diag::project(
            DiagCode::ModuleLimit,
            format!("The macro project declares {declared} modules; only the first {} are listed.", limits.max_modules),
        ));
    }
    let kinds = if info.modules.is_empty() { None } else { read_project_stream(src, &mut project.diagnostics) };
    let mut total = 0usize;
    for dir_module in &info.modules {
        let module = read_module(src, limits, code_page, kinds.as_ref(), dir_module);
        match total.checked_add(module.source.len()).filter(|&t| t <= limits.max_total_source) {
            Some(t) => total = t,
            None => {
                project.diagnostics.push(Diag::project(
                    DiagCode::SourceLimit,
                    format!(
                        "The macro source is larger than the {} limit, so only the first {} modules are listed.",
                        size_text(limits.max_total_source),
                        project.modules.len()
                    ),
                ));
                break;
            }
        }
        project.modules.push(module);
    }
    project
}

/// Reads, decompresses and parses `VBA/dir`, or the diagnostic saying why it couldn't.
fn read_dir(src: &dyn StreamSource, limits: &Limits) -> Result<DirInfo, Diag> {
    let corrupt = |why: String| Diag::project(DiagCode::DirCorrupt, format!("The macro project's module list couldn't be read: {why}."));
    let bytes = match src.stream(&["VBA", "dir"]) {
        Ok(Some(bytes)) => bytes,
        Ok(None) => {
            return Err(Diag::project(DiagCode::DirMissing, "The macro project has no module list (the VBA/dir stream is missing)."));
        }
        Err(e) => return Err(corrupt(e)),
    };
    let decompressed = ovba::decompress(&bytes, limits.max_stream).map_err(|e| corrupt(e.to_string()))?;
    dir::parse_dir(&decompressed, limits.max_modules).map_err(|e| corrupt(e.to_string()))
}

/// Reads the `PROJECT` stream's module kinds, adding a diagnostic when it is missing or capped.
fn read_project_stream(src: &dyn StreamSource, diagnostics: &mut Vec<Diag>) -> Option<ProjectStreamInfo> {
    let missing = |why: &str| {
        Diag::project(DiagCode::ProjectStreamMissing, format!("The macro project's PROJECT stream {why}, so module kinds come from the module list."))
    };
    match src.stream(&["PROJECT"]) {
        Ok(Some(bytes)) => {
            let info = project_stream::parse_project_stream(&bytes);
            if info.truncated {
                diagnostics.push(Diag::project(
                    DiagCode::ProjectStreamTruncated,
                    "The macro project's PROJECT stream is too large to read in full, so some module kinds come from the module list.",
                ));
            }
            Some(info)
        }
        Ok(None) => {
            diagnostics.push(missing("is missing"));
            None
        }
        Err(e) => {
            diagnostics.push(missing(&format!("couldn't be read ({e})")));
            None
        }
    }
}

/// Reads one module. Failures leave the source empty and add a diagnostic to the module.
fn read_module(src: &dyn StreamSource, limits: &Limits, code_page: u16, kinds: Option<&ProjectStreamInfo>, m: &DirModule) -> VbaModule {
    let code_name = mbcs(code_page, &m.name);
    let name = match m.name_unicode.as_deref().map(encoding::decode_utf16le) {
        Some(unicode) if !unicode.is_empty() => unicode,
        _ => code_name.clone(),
    };
    let stream_name = match m.stream_name_unicode.as_deref().map(encoding::decode_utf16le) {
        Some(unicode) if !unicode.is_empty() => unicode,
        _ => mbcs(code_page, &m.stream_name),
    };
    let mut diagnostics = Vec::new();
    let (kind, fallback) = project_stream::resolve_kind(kinds, &m.name, m.module_type);
    if fallback {
        diagnostics.push(Diag::module(
            DiagCode::KindFromModuleType,
            &name,
            format!("Module \"{name}\" isn't listed in the PROJECT stream, so it is shown as a {} module.", kind.as_str()),
        ));
    }
    let source = match module_source(src, limits, code_page, &stream_name, m.text_offset) {
        Ok(source) => {
            if possible_stomping(m.text_offset, &source) {
                diagnostics.push(Diag::module(
                    DiagCode::PossibleStomping,
                    &name,
                    format!(
                        "Module \"{name}\" has compiled code but no source to match it. The source may have been removed to hide what the macro does (VBA stomping)."
                    ),
                ));
            } else if m.text_offset > 0 {
                diagnostics.push(Diag::module(
                    DiagCode::CompiledCodeNotChecked,
                    &name,
                    format!(
                        "Module \"{name}\" also holds compiled code. GridCraft doesn't read compiled code, so it can't confirm the source matches it."
                    ),
                ));
            }
            source
        }
        Err((code, why)) => {
            diagnostics.push(Diag::module(code, &name, format!("The source of module \"{name}\" couldn't be read: {why}.")));
            String::new()
        }
    };
    VbaModule { name, code_name, stream_name, kind, text_offset: m.text_offset, source, diagnostics }
}

/// The decoded source of the module stream `VBA/<stream_name>`, or a diagnostic code and reason.
fn module_source(src: &dyn StreamSource, limits: &Limits, code_page: u16, stream_name: &str, text_offset: u32) -> Result<String, (DiagCode, String)> {
    let bytes = match src.stream(&["VBA", stream_name]) {
        Ok(Some(bytes)) => bytes,
        Ok(None) => return Err((DiagCode::ModuleStreamMissing, format!("its stream \"{stream_name}\" is missing"))),
        Err(e) => return Err((DiagCode::ModuleCorrupt, format!("its stream \"{stream_name}\" couldn't be read ({e})"))),
    };
    let compressed = usize::try_from(text_offset).ok().and_then(|at| bytes.get(at..)).filter(|rest| !rest.is_empty()).ok_or_else(|| {
        (DiagCode::ModuleOffsetOutOfRange, format!("its source offset {text_offset} is at or past the end of its {}-byte stream", bytes.len()))
    })?;
    let decompressed = ovba::decompress(compressed, limits.max_stream).map_err(|e| {
        let code = if ovba::is_cut_short(compressed, &e) { DiagCode::ModuleTruncated } else { DiagCode::ModuleCorrupt };
        (code, e.to_string())
    })?;
    // A raw chunk always decompresses to 4,096 bytes, so a short raw last chunk ends in 0x00
    // padding ([MS-OVBA] §2.4.1.3.10). Source text never holds NUL, so trailing NULs are padding.
    let end = decompressed.iter().rposition(|&b| b != 0).map_or(0, |i| i.saturating_add(1));
    Ok(mbcs(code_page, decompressed.get(..end).unwrap_or_default()))
}

/// Whether a module that was read looks stomped (Req 7.1).
///
/// VBA stomping removes or replaces a module's source while keeping its compiled code (P_Code),
/// which Office runs when the VBA version matches, so the code that runs is not the code shown.
/// Two facts are enough to spot the plain form without decoding anything:
///
/// - `text_offset > 0`: the module stream holds bytes before MODULEOFFSET, which is where
///   [MS-OVBA] §2.3.4.3 puts the compiled-code cache. The caller only asks once the source has
///   been read, so MODULEOFFSET is inside the stream and those bytes really exist. A module with
///   `text_offset == 0` has nothing compiled to hide, so an empty source there is just an empty
///   module.
/// - the source is empty in substance: every line, trimmed, is blank or an `Attribute`
///   statement. Those lines are written by the VBA editor itself (`Attribute VB_Name = …`), so
///   a module that has compiled code and nothing else has lost the code the user wrote.
///
/// This is the reliable, clean-room check: the P_Code bytes are only counted (through
/// MODULEOFFSET), never decoded, so no undocumented format is involved and a valid module is
/// never flagged for what its compiled code holds. A comment line counts as source, so a module
/// whose only line is a comment is not flagged: comments are what a user writes, and flagging
/// them would turn the check into a guess about intent. Replacing the source with different
/// code is not caught here; that needs the P_Code identifier comparison of Req 7.2.
fn possible_stomping(text_offset: u32, source: &str) -> bool {
    text_offset > 0 && source.split(['\r', '\n']).map(str::trim).all(|line| line.is_empty() || is_attribute_line(line))
}

/// Whether a trimmed line is a VBA `Attribute` statement: the keyword, case-insensitively, then
/// a space or tab.
fn is_attribute_line(line: &str) -> bool {
    const KEYWORD: &str = "attribute";
    line.get(..KEYWORD.len()).is_some_and(|word| word.eq_ignore_ascii_case(KEYWORD))
        && line.as_bytes().get(KEYWORD.len()).is_some_and(|&b| b == b' ' || b == b'\t')
}

/// MBCS text in the project code page. An unsupported code page is reported once, by the caller.
fn mbcs(code_page: u16, bytes: &[u8]) -> String {
    encoding::decode(code_page, bytes).0
}

/// A byte count for messages: whole MiB when exact, else bytes.
fn size_text(bytes: usize) -> String {
    const MIB: usize = 1024 * 1024;
    if bytes >= MIB && bytes.is_multiple_of(MIB) { format!("{} MiB", bytes / MIB) } else { format!("{bytes}-byte") }
}

#[cfg(test)]
mod props;
#[cfg(test)]
mod tests;
