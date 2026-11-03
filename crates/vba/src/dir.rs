//! The `dir` stream: version-independent project information ([MS-OVBA] §2.3.4.2).
//!
//! [`parse_dir`] reads the **decompressed** stream: a PROJECTINFORMATION record (§2.3.4.2.1), a
//! PROJECTREFERENCES record (§2.3.4.2.2), a PROJECTMODULES record (§2.3.4.2.3), then the
//! Terminator 0x0010 and four reserved bytes. It keeps what later stages need (the code page,
//! the project name and compilation constants, each reference's name and kind, and each module's
//! names, stream name, text offset, type and flags) and steps over the rest.
//!
//! Strings stay raw bytes: MBCS fields in the project code page, `*_unicode` fields as UTF-16LE.
//! Decoding happens later, once the code page is known.
//!
//! Every read goes through a bounds-checked cursor. Parsing stops at the first record that is cut
//! short, has the wrong id, or declares a bad size; there is no partial result. The rules:
//!
//! - Records the spec requires at a position must carry its id ([`DirError::UnexpectedId`]).
//!   Where the spec allows several ids (the four REFERENCE kinds, the two MODULETYPE ids), any
//!   other id is [`DirError::UnknownRecord`].
//! - Fixed-size records must declare the size the spec gives (for example 4 for PROJECTLCID),
//!   else [`DirError::BadSize`]. The REFERENCE records' outer `Size` fields are skipped, because
//!   the spec says to ignore them on read; their inner sizes are used instead.
//! - Variable fields are capped at the spec's maximum where it gives one (ProjectName 128,
//!   DocString 2000, HelpFile 260, Constants 1015 bytes) and at [`MAX_FIELD`] otherwise. A size
//!   above the cap is [`DirError::BadSize`]; a size past the end of the input is
//!   [`DirError::Truncated`]. The UTF-16 module name and stream name must have an even size.
//! - Field values the spec fixes (LCID 0x0409, reserved words, the second help file path) are
//!   not checked: they "MUST be ignored" or don't change how the stream is read.
//! - Bytes after the final Terminator and its reserved field are ignored, since a raw last
//!   compression chunk pads the decompressed stream.
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/

/// Test-only writer for `dir` streams, shared with the Fixture_Builder.
#[cfg(any(test, feature = "test-support"))]
pub mod fixture;

/// The largest variable-length field read when the spec gives no maximum (64 KiB). Names,
/// stream names, libids and the UTF-16 copies of strings are far shorter in practice.
pub const MAX_FIELD: usize = 64 * 1024;
/// The most modules a reader keeps by default (design: the first 1,024 plus a diagnostic).
pub const MAX_MODULES: usize = 1024;

/// SizeOfProjectName MUST be at most 128 (§2.3.4.2.1.6).
const MAX_PROJECT_NAME: usize = 128;
/// SizeOfDocString MUST be at most 2000 (§2.3.4.2.1.7).
const MAX_DOC_STRING: usize = 2000;
/// SizeOfHelpFile1 MUST be at most 260 (§2.3.4.2.1.8); SizeOfHelpFile2 equals it.
const MAX_HELP_FILE: usize = 260;
/// SizeOfConstants MUST be at most 1015 (§2.3.4.2.1.12).
const MAX_CONSTANTS: usize = 1015;

/// Record ids, with the section that defines each.
pub(crate) mod id {
    /// PROJECTSYSKIND (§2.3.4.2.1.1).
    pub const SYS_KIND: u16 = 0x0001;
    /// PROJECTCOMPATVERSION (§2.3.4.2.1.2), optional.
    pub const COMPAT_VERSION: u16 = 0x004A;
    /// PROJECTLCID (§2.3.4.2.1.3).
    pub const LCID: u16 = 0x0002;
    /// PROJECTLCIDINVOKE (§2.3.4.2.1.4).
    pub const LCID_INVOKE: u16 = 0x0014;
    /// PROJECTCODEPAGE (§2.3.4.2.1.5).
    pub const CODE_PAGE: u16 = 0x0003;
    /// PROJECTNAME (§2.3.4.2.1.6).
    pub const NAME: u16 = 0x0004;
    /// PROJECTDOCSTRING (§2.3.4.2.1.7).
    pub const DOC_STRING: u16 = 0x0005;
    /// PROJECTHELPFILEPATH (§2.3.4.2.1.8).
    pub const HELP_FILE_PATH: u16 = 0x0006;
    /// PROJECTHELPCONTEXT (§2.3.4.2.1.9).
    pub const HELP_CONTEXT: u16 = 0x0007;
    /// PROJECTLIBFLAGS (§2.3.4.2.1.10).
    pub const LIB_FLAGS: u16 = 0x0008;
    /// PROJECTVERSION (§2.3.4.2.1.11).
    pub const VERSION: u16 = 0x0009;
    /// PROJECTCONSTANTS (§2.3.4.2.1.12), optional.
    pub const CONSTANTS: u16 = 0x000C;
    /// REFERENCENAME (§2.3.4.2.2.2).
    pub const REFERENCE_NAME: u16 = 0x0016;
    /// REFERENCECONTROL (§2.3.4.2.2.3).
    pub const REFERENCE_CONTROL: u16 = 0x002F;
    /// REFERENCEORIGINAL (§2.3.4.2.2.4).
    pub const REFERENCE_ORIGINAL: u16 = 0x0033;
    /// REFERENCEREGISTERED (§2.3.4.2.2.5).
    pub const REFERENCE_REGISTERED: u16 = 0x000D;
    /// REFERENCEPROJECT (§2.3.4.2.2.6).
    pub const REFERENCE_PROJECT: u16 = 0x000E;
    /// PROJECTMODULES (§2.3.4.2.3); also ends the reference array (§2.3.4.2.2).
    pub const PROJECT_MODULES: u16 = 0x000F;
    /// PROJECTCOOKIE (§2.3.4.2.3.1).
    pub const PROJECT_COOKIE: u16 = 0x0013;
    /// MODULENAME (§2.3.4.2.3.2.1).
    pub const MODULE_NAME: u16 = 0x0019;
    /// MODULENAMEUNICODE (§2.3.4.2.3.2.2), optional.
    pub const MODULE_NAME_UNICODE: u16 = 0x0047;
    /// MODULESTREAMNAME (§2.3.4.2.3.2.3).
    pub const MODULE_STREAM_NAME: u16 = 0x001A;
    /// MODULEDOCSTRING (§2.3.4.2.3.2.4).
    pub const MODULE_DOC_STRING: u16 = 0x001C;
    /// MODULEOFFSET (§2.3.4.2.3.2.5).
    pub const MODULE_OFFSET: u16 = 0x0031;
    /// MODULEHELPCONTEXT (§2.3.4.2.3.2.6).
    pub const MODULE_HELP_CONTEXT: u16 = 0x001E;
    /// MODULECOOKIE (§2.3.4.2.3.2.7).
    pub const MODULE_COOKIE: u16 = 0x002C;
    /// MODULETYPE for a procedural module (§2.3.4.2.3.2.8).
    pub const MODULE_TYPE_PROCEDURAL: u16 = 0x0021;
    /// MODULETYPE for a document, class or designer module (§2.3.4.2.3.2.8).
    pub const MODULE_TYPE_OTHER: u16 = 0x0022;
    /// MODULEREADONLY (§2.3.4.2.3.2.9), optional.
    pub const MODULE_READ_ONLY: u16 = 0x0025;
    /// MODULEPRIVATE (§2.3.4.2.3.2.10), optional.
    pub const MODULE_PRIVATE: u16 = 0x0028;
    /// MODULE Terminator (§2.3.4.2.3.2).
    pub const MODULE_TERMINATOR: u16 = 0x002B;
    /// dir stream Terminator (§2.3.4.2).
    pub const DIR_TERMINATOR: u16 = 0x0010;
}

/// What the `dir` stream says about a VBA project.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DirInfo {
    /// PROJECTSYSKIND.SysKind: 0 Win16, 1 Win32, 2 Macintosh, 3 Win64.
    pub sys_kind: u32,
    /// PROJECTCOMPATVERSION.CompatVersion, when the record is present.
    pub compat_version: Option<u32>,
    /// PROJECTCODEPAGE.CodePage: the code page of every MBCS string in the project.
    pub code_page: u16,
    /// PROJECTNAME.ProjectName (MBCS).
    pub name: Vec<u8>,
    /// PROJECTCONSTANTS.Constants (MBCS, `name = value : …`); empty when the record is absent.
    pub constants: Vec<u8>,
    /// The project references, in stream order.
    pub references: Vec<DirReference>,
    /// The first `max_modules` modules, in stream order.
    pub modules: Vec<DirModule>,
    /// `Some(declared count)` when PROJECTMODULES declared more than `max_modules` modules. The
    /// rest were still read (to reach the Terminator) but not kept.
    pub truncated_modules: Option<u16>,
}

/// One project reference (§2.3.4.2.2.1).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DirReference {
    /// REFERENCENAME.Name (MBCS); empty when the optional name record is absent.
    pub name: Vec<u8>,
    pub kind: ReferenceKind,
}

/// The ReferenceRecord type of a REFERENCE (§2.3.4.2.2.1).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceKind {
    /// REFERENCECONTROL (0x002F): a twiddled type library and its extended type library.
    Control,
    /// REFERENCEORIGINAL (0x0033): an Automation type library plus the REFERENCECONTROL made
    /// from it.
    Original,
    /// REFERENCEREGISTERED (0x000D): a registered Automation type library.
    Registered,
    /// REFERENCEPROJECT (0x000E): another VBA project.
    Project,
}

/// One MODULE record (§2.3.4.2.3.2).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DirModule {
    /// MODULENAME.ModuleName (MBCS).
    pub name: Vec<u8>,
    /// MODULENAMEUNICODE.ModuleNameUnicode (UTF-16LE); `None` when absent or empty.
    pub name_unicode: Option<Vec<u8>>,
    /// MODULESTREAMNAME.StreamName (MBCS): the module stream's name in the VBA storage.
    pub stream_name: Vec<u8>,
    /// MODULESTREAMNAME.StreamNameUnicode (UTF-16LE); `None` when empty.
    pub stream_name_unicode: Option<Vec<u8>>,
    /// MODULEOFFSET.TextOffset: where the compressed source starts in the module stream.
    pub text_offset: u32,
    pub module_type: DirModuleType,
    /// MODULEREADONLY is present.
    pub read_only: bool,
    /// MODULEPRIVATE is present.
    pub private: bool,
}

/// MODULETYPE (§2.3.4.2.3.2.8). The PROJECT stream tells document, class and designer apart.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DirModuleType {
    /// Id 0x0021: a procedural (standard) module.
    Procedural,
    /// Id 0x0022: a document, class or designer module.
    NonProcedural,
}

/// Why a `dir` stream couldn't be read. `at` is a byte offset into the decompressed stream.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum DirError {
    /// A read starting at `at` runs past the end of the stream.
    #[error("the dir stream is cut short at byte {at}")]
    Truncated { at: usize },
    /// The record at `at` has id `found` where the spec requires `expected`.
    #[error("the dir stream has record 0x{found:04X} at byte {at} where record 0x{expected:04X} belongs")]
    UnexpectedId { at: usize, expected: u16, found: u16 },
    /// The record at `at` has an id the spec doesn't allow there (a REFERENCE kind or MODULETYPE).
    #[error("the dir stream has an unknown record 0x{found:04X} at byte {at}")]
    UnknownRecord { at: usize, found: u16 },
    /// The record `id` at `at` declares a size the spec doesn't allow: a wrong fixed size, a
    /// field above its cap, or an odd UTF-16 size.
    #[error("record 0x{id:04X} at byte {at} declares a size that isn't allowed")]
    BadSize { at: usize, id: u16 },
}

/// Parses a decompressed `dir` stream, keeping at most `max_modules` modules (callers pass
/// [`MAX_MODULES`]).
///
/// # Errors
///
/// Returns a [`DirError`] for the first record that is cut short, has an unexpected or unknown
/// id, or declares a bad size. See the module docs for the exact rules.
pub fn parse_dir(bytes: &[u8], max_modules: usize) -> Result<DirInfo, DirError> {
    let mut c = Cursor { bytes, pos: 0 };
    let info = information(&mut c)?;
    let references = references(&mut c)?;
    let (modules, truncated_modules) = modules(&mut c, max_modules)?;
    c.record(id::DIR_TERMINATOR)?;
    c.skip(4)?; // Reserved: MUST be ignored.
    Ok(DirInfo {
        sys_kind: info.sys_kind,
        compat_version: info.compat_version,
        code_page: info.code_page,
        name: info.name,
        constants: info.constants,
        references,
        modules,
        truncated_modules,
    })
}

/// The PROJECTINFORMATION fields kept.
struct Information {
    sys_kind: u32,
    compat_version: Option<u32>,
    code_page: u16,
    name: Vec<u8>,
    constants: Vec<u8>,
}

/// PROJECTINFORMATION (§2.3.4.2.1).
fn information(c: &mut Cursor<'_>) -> Result<Information, DirError> {
    let sys_kind = c.fixed_u32(id::SYS_KIND)?;
    let compat_version = if c.peek_id() == Some(id::COMPAT_VERSION) { Some(c.fixed_u32(id::COMPAT_VERSION)?) } else { None };
    c.fixed_u32(id::LCID)?;
    c.fixed_u32(id::LCID_INVOKE)?;
    let code_page = c.fixed_u16(id::CODE_PAGE)?;
    let rec = c.record(id::NAME)?;
    let name = c.sized(rec, MAX_PROJECT_NAME)?.to_vec();
    let rec = c.record(id::DOC_STRING)?;
    c.sized(rec, MAX_DOC_STRING)?;
    c.skip(2)?; // Reserved 0x0040: MUST be ignored.
    c.sized(rec, MAX_FIELD)?;
    let rec = c.record(id::HELP_FILE_PATH)?;
    c.sized(rec, MAX_HELP_FILE)?;
    c.skip(2)?; // Reserved 0x003D: MUST be ignored.
    c.sized(rec, MAX_HELP_FILE)?;
    c.fixed_u32(id::HELP_CONTEXT)?;
    c.fixed_u32(id::LIB_FLAGS)?;
    // PROJECTVERSION has no Size field: Id, Reserved (4, ignored), VersionMajor (4), VersionMinor (2).
    c.record(id::VERSION)?;
    c.skip(4 + 4 + 2)?;
    let mut constants = Vec::new();
    if c.peek_id() == Some(id::CONSTANTS) {
        let rec = c.record(id::CONSTANTS)?;
        constants = c.sized(rec, MAX_CONSTANTS)?.to_vec();
        c.skip(2)?; // Reserved 0x003C: MUST be ignored.
        c.sized(rec, MAX_FIELD)?;
    }
    Ok(Information { sys_kind, compat_version, code_page, name, constants })
}

/// PROJECTREFERENCES (§2.3.4.2.2): REFERENCE records until the PROJECTMODULES id 0x000F.
/// Each iteration consumes at least six bytes, so the list is bounded by the input.
fn references(c: &mut Cursor<'_>) -> Result<Vec<DirReference>, DirError> {
    let mut out = Vec::new();
    loop {
        let next = c.peek_id().ok_or(DirError::Truncated { at: c.pos })?;
        if next == id::PROJECT_MODULES {
            return Ok(out);
        }
        let name = if next == id::REFERENCE_NAME { reference_name(c)? } else { Vec::new() };
        let rec = c.any_record()?;
        let kind = match rec.id {
            id::REFERENCE_CONTROL => {
                reference_control(c, rec)?;
                ReferenceKind::Control
            }
            id::REFERENCE_ORIGINAL => {
                reference_original(c, rec)?;
                ReferenceKind::Original
            }
            id::REFERENCE_REGISTERED => {
                reference_registered(c, rec)?;
                ReferenceKind::Registered
            }
            id::REFERENCE_PROJECT => {
                reference_project(c, rec)?;
                ReferenceKind::Project
            }
            found => return Err(DirError::UnknownRecord { at: rec.at, found }),
        };
        out.push(DirReference { name, kind });
    }
}

/// REFERENCENAME (§2.3.4.2.2.2); returns Name.
fn reference_name(c: &mut Cursor<'_>) -> Result<Vec<u8>, DirError> {
    let rec = c.record(id::REFERENCE_NAME)?;
    let name = c.sized(rec, MAX_FIELD)?.to_vec();
    c.skip(2)?; // Reserved 0x003E: MUST be ignored.
    c.sized(rec, MAX_FIELD)?;
    Ok(name)
}

/// The body of REFERENCECONTROL (§2.3.4.2.2.3) after its Id.
fn reference_control(c: &mut Cursor<'_>, rec: Record) -> Result<(), DirError> {
    c.skip(4)?; // SizeTwiddled: MUST be ignored on read.
    c.sized(rec, MAX_FIELD)?; // LibidTwiddled
    c.skip(4 + 2)?; // Reserved1, Reserved2
    if c.peek_id() == Some(id::REFERENCE_NAME) {
        reference_name(c)?; // NameRecordExtended, optional
    }
    c.skip(2 + 4)?; // Reserved3 (0x0030, ignored), SizeExtended (ignored on read)
    c.sized(rec, MAX_FIELD)?; // LibidExtended
    c.skip(4 + 2 + 16 + 4) // Reserved4, Reserved5, OriginalTypeLib, Cookie
}

/// The body of REFERENCEORIGINAL (§2.3.4.2.2.4) after its Id: LibidOriginal, then a
/// REFERENCECONTROL.
fn reference_original(c: &mut Cursor<'_>, rec: Record) -> Result<(), DirError> {
    c.sized(rec, MAX_FIELD)?;
    let control = c.record(id::REFERENCE_CONTROL)?;
    reference_control(c, control)
}

/// The body of REFERENCEREGISTERED (§2.3.4.2.2.5) after its Id.
fn reference_registered(c: &mut Cursor<'_>, rec: Record) -> Result<(), DirError> {
    c.skip(4)?; // Size: MUST be ignored on read.
    c.sized(rec, MAX_FIELD)?; // Libid
    c.skip(4 + 2) // Reserved1, Reserved2
}

/// The body of REFERENCEPROJECT (§2.3.4.2.2.6) after its Id.
fn reference_project(c: &mut Cursor<'_>, rec: Record) -> Result<(), DirError> {
    c.skip(4)?; // Size: MUST be ignored on read.
    c.sized(rec, MAX_FIELD)?; // LibidAbsolute
    c.sized(rec, MAX_FIELD)?; // LibidRelative
    c.skip(4 + 2) // MajorVersion, MinorVersion
}

/// PROJECTMODULES (§2.3.4.2.3): reads all Count modules, keeps the first `max_modules`.
fn modules(c: &mut Cursor<'_>, max_modules: usize) -> Result<(Vec<DirModule>, Option<u16>), DirError> {
    let count = c.fixed_u16(id::PROJECT_MODULES)?;
    c.fixed_u16(id::PROJECT_COOKIE)?;
    let mut out = Vec::new();
    for _ in 0..count {
        let module = module(c)?;
        if out.len() < max_modules {
            out.push(module);
        }
    }
    let truncated = (usize::from(count) > max_modules).then_some(count);
    Ok((out, truncated))
}

/// MODULE (§2.3.4.2.3.2).
fn module(c: &mut Cursor<'_>) -> Result<DirModule, DirError> {
    let rec = c.record(id::MODULE_NAME)?;
    let name = c.sized(rec, MAX_FIELD)?.to_vec();
    let mut name_unicode = None;
    if c.peek_id() == Some(id::MODULE_NAME_UNICODE) {
        let rec = c.record(id::MODULE_NAME_UNICODE)?;
        name_unicode = non_empty(c.even_sized(rec, MAX_FIELD)?);
    }
    let rec = c.record(id::MODULE_STREAM_NAME)?;
    let stream_name = c.sized(rec, MAX_FIELD)?.to_vec();
    c.skip(2)?; // Reserved 0x0032: MUST be ignored.
    let stream_name_unicode = non_empty(c.even_sized(rec, MAX_FIELD)?);
    let rec = c.record(id::MODULE_DOC_STRING)?;
    c.sized(rec, MAX_FIELD)?;
    c.skip(2)?; // Reserved 0x0048: MUST be ignored.
    c.sized(rec, MAX_FIELD)?;
    let text_offset = c.fixed_u32(id::MODULE_OFFSET)?;
    c.fixed_u32(id::MODULE_HELP_CONTEXT)?;
    c.fixed_u16(id::MODULE_COOKIE)?;
    let rec = c.any_record()?;
    let module_type = match rec.id {
        id::MODULE_TYPE_PROCEDURAL => DirModuleType::Procedural,
        id::MODULE_TYPE_OTHER => DirModuleType::NonProcedural,
        found => return Err(DirError::UnknownRecord { at: rec.at, found }),
    };
    c.skip(4)?; // Reserved: MUST be ignored.
    let read_only = c.flag(id::MODULE_READ_ONLY)?;
    let private = c.flag(id::MODULE_PRIVATE)?;
    c.record(id::MODULE_TERMINATOR)?;
    c.skip(4)?; // Reserved: MUST be ignored.
    Ok(DirModule { name, name_unicode, stream_name, stream_name_unicode, text_offset, module_type, read_only, private })
}

fn non_empty(bytes: &[u8]) -> Option<Vec<u8>> {
    (!bytes.is_empty()).then(|| bytes.to_vec())
}

/// A record's Id and where it starts, for error reports.
#[derive(Debug, Clone, Copy)]
struct Record {
    at: usize,
    id: u16,
}

/// A bounds-checked little-endian reader. No read moves `pos` past `bytes.len()`.
struct Cursor<'a> {
    bytes: &'a [u8],
    pos: usize,
}

impl<'a> Cursor<'a> {
    /// The next `n` bytes; `Truncated` at the current position if fewer remain.
    fn bytes(&mut self, n: usize) -> Result<&'a [u8], DirError> {
        let truncated = DirError::Truncated { at: self.pos };
        let end = self.pos.checked_add(n).ok_or_else(|| truncated.clone())?;
        let out = self.bytes.get(self.pos..end).ok_or(truncated)?;
        self.pos = end;
        Ok(out)
    }

    fn skip(&mut self, n: usize) -> Result<(), DirError> {
        self.bytes(n).map(|_| ())
    }

    fn u16(&mut self) -> Result<u16, DirError> {
        let at = self.pos;
        let b: [u8; 2] = self.bytes(2)?.try_into().map_err(|_| DirError::Truncated { at })?;
        Ok(u16::from_le_bytes(b))
    }

    fn u32(&mut self) -> Result<u32, DirError> {
        let at = self.pos;
        let b: [u8; 4] = self.bytes(4)?.try_into().map_err(|_| DirError::Truncated { at })?;
        Ok(u32::from_le_bytes(b))
    }

    /// The next record id, without consuming it; `None` if fewer than two bytes remain.
    fn peek_id(&self) -> Option<u16> {
        let at = self.pos;
        let lo = *self.bytes.get(at)?;
        let hi = *self.bytes.get(at.checked_add(1)?)?;
        Some(u16::from_le_bytes([lo, hi]))
    }

    /// Reads any record id.
    fn any_record(&mut self) -> Result<Record, DirError> {
        let at = self.pos;
        Ok(Record { at, id: self.u16()? })
    }

    /// Reads a record id that must be `expected`.
    fn record(&mut self, expected: u16) -> Result<Record, DirError> {
        let rec = self.any_record()?;
        if rec.id == expected { Ok(rec) } else { Err(DirError::UnexpectedId { at: rec.at, expected, found: rec.id }) }
    }

    /// Reads a record id that must be `id` and a Size that must be `size`.
    fn fixed(&mut self, id: u16, size: u32) -> Result<(), DirError> {
        let rec = self.record(id)?;
        if self.u32()? == size { Ok(()) } else { Err(DirError::BadSize { at: rec.at, id }) }
    }

    /// A fixed record holding one u32 (Size 4).
    fn fixed_u32(&mut self, id: u16) -> Result<u32, DirError> {
        self.fixed(id, 4)?;
        self.u32()
    }

    /// A fixed record holding one u16 (Size 2).
    fn fixed_u16(&mut self, id: u16) -> Result<u16, DirError> {
        self.fixed(id, 2)?;
        self.u16()
    }

    /// A u32 size then that many bytes, at most `max`.
    fn sized(&mut self, rec: Record, max: usize) -> Result<&'a [u8], DirError> {
        let bad = DirError::BadSize { at: rec.at, id: rec.id };
        let n = usize::try_from(self.u32()?).map_err(|_| bad.clone())?;
        if n > max {
            return Err(bad);
        }
        self.bytes(n)
    }

    /// Like [`Cursor::sized`], for UTF-16 fields whose size MUST be even.
    fn even_sized(&mut self, rec: Record, max: usize) -> Result<&'a [u8], DirError> {
        let out = self.sized(rec, max)?;
        if out.len() % 2 == 0 { Ok(out) } else { Err(DirError::BadSize { at: rec.at, id: rec.id }) }
    }

    /// An optional flag record (Id plus four reserved bytes): true and consumed if present.
    fn flag(&mut self, id: u16) -> Result<bool, DirError> {
        if self.peek_id() != Some(id) {
            return Ok(false);
        }
        self.record(id)?;
        self.skip(4)?; // Reserved: MUST be ignored.
        Ok(true)
    }
}

/// Unit tests, in `dir/tests.rs`.
#[cfg(test)]
mod tests;
