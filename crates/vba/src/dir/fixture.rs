//! A test-only writer for decompressed `dir` streams ([MS-OVBA] §2.3.4.2), built from fields so
//! no Office-produced bytes are needed. The Fixture_Builder compresses its output with
//! [`crate::ovba::compress`].
//!
//! MBCS strings are written as the `&str`'s bytes (ASCII in practice) and the UTF-16 copies as
//! UTF-16LE of the same text. Reserved fields get the values the spec gives for writing; cookies
//! are 0xFFFF. Lengths above `u32::MAX` (or more than 65,535 modules) are clamped, never panic.
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/

use super::{DirModuleType, ReferenceKind, id};

/// The fields of a `dir` stream. [`Default`] is a minimal valid project named `VBAProject`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DirFixture {
    pub sys_kind: u32,
    /// Writes PROJECTCOMPATVERSION when `Some`.
    pub compat_version: Option<u32>,
    pub code_page: u16,
    pub project_name: String,
    pub doc_string: String,
    pub help_file: String,
    /// Writes PROJECTCONSTANTS when `Some`.
    pub constants: Option<String>,
    pub references: Vec<ReferenceFixture>,
    pub modules: Vec<ModuleFixture>,
}

impl Default for DirFixture {
    fn default() -> Self {
        Self {
            sys_kind: 1,
            compat_version: None,
            code_page: 1252,
            project_name: "VBAProject".to_owned(),
            doc_string: String::new(),
            help_file: String::new(),
            constants: None,
            references: Vec::new(),
            modules: Vec::new(),
        }
    }
}

/// One REFERENCE. `libid` fills LibidTwiddled (Control), LibidOriginal (Original), Libid
/// (Registered) or both LibidAbsolute and LibidRelative (Project). A named Control (or Original)
/// also gets a NameRecordExtended with the same name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReferenceFixture {
    /// Writes REFERENCENAME when `Some`.
    pub name: Option<String>,
    pub kind: ReferenceKind,
    pub libid: String,
}

impl ReferenceFixture {
    /// A named reference with a placeholder libid.
    pub fn new(name: &str, kind: ReferenceKind) -> Self {
        let libid = match kind {
            ReferenceKind::Project => "*\\CExample.xlsm",
            _ => "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##",
        };
        Self { name: Some(name.to_owned()), kind, libid: libid.to_owned() }
    }
}

/// One MODULE. [`ModuleFixture::new`] uses the name as the stream name and writes
/// MODULENAMEUNICODE.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModuleFixture {
    pub name: String,
    pub stream_name: String,
    /// Writes MODULENAMEUNICODE.
    pub name_unicode: bool,
    pub doc_string: String,
    pub text_offset: u32,
    pub module_type: DirModuleType,
    /// Writes MODULEREADONLY.
    pub read_only: bool,
    /// Writes MODULEPRIVATE.
    pub private: bool,
}

impl ModuleFixture {
    pub fn new(name: &str, module_type: DirModuleType) -> Self {
        Self {
            name: name.to_owned(),
            stream_name: name.to_owned(),
            name_unicode: true,
            doc_string: String::new(),
            text_offset: 0,
            module_type,
            read_only: false,
            private: false,
        }
    }
}

impl DirFixture {
    /// The decompressed `dir` stream for these fields.
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut w = Writer::default();
        w.fixed_u32(id::SYS_KIND, self.sys_kind);
        if let Some(v) = self.compat_version {
            w.fixed_u32(id::COMPAT_VERSION, v);
        }
        w.fixed_u32(id::LCID, 0x0409);
        w.fixed_u32(id::LCID_INVOKE, 0x0409);
        w.fixed_u16(id::CODE_PAGE, self.code_page);
        w.u16(id::NAME);
        w.sized(self.project_name.as_bytes());
        w.u16(id::DOC_STRING);
        w.dual(&self.doc_string, 0x0040);
        w.u16(id::HELP_FILE_PATH);
        w.sized(self.help_file.as_bytes());
        w.u16(0x003D);
        w.sized(self.help_file.as_bytes());
        w.fixed_u32(id::HELP_CONTEXT, 0);
        w.fixed_u32(id::LIB_FLAGS, 0);
        w.u16(id::VERSION);
        w.u32(4); // Reserved
        w.u32(1); // VersionMajor
        w.u16(0); // VersionMinor
        if let Some(constants) = &self.constants {
            w.u16(id::CONSTANTS);
            w.dual(constants, 0x003C);
        }
        for reference in &self.references {
            w.reference(reference);
        }
        w.fixed_u16(id::PROJECT_MODULES, u16::try_from(self.modules.len()).unwrap_or(u16::MAX));
        w.fixed_u16(id::PROJECT_COOKIE, 0xFFFF);
        for module in &self.modules {
            w.module(module);
        }
        w.u16(id::DIR_TERMINATOR);
        w.u32(0);
        w.0
    }
}

#[derive(Default)]
struct Writer(Vec<u8>);

impl Writer {
    fn u16(&mut self, v: u16) {
        self.0.extend_from_slice(&v.to_le_bytes());
    }

    fn u32(&mut self, v: u32) {
        self.0.extend_from_slice(&v.to_le_bytes());
    }

    fn fixed_u32(&mut self, id: u16, v: u32) {
        self.u16(id);
        self.u32(4);
        self.u32(v);
    }

    fn fixed_u16(&mut self, id: u16, v: u16) {
        self.u16(id);
        self.u32(2);
        self.u16(v);
    }

    /// A u32 size then the bytes.
    fn sized(&mut self, bytes: &[u8]) {
        self.u32(u32::try_from(bytes.len()).unwrap_or(u32::MAX));
        self.0.extend_from_slice(bytes);
    }

    /// The MBCS text, a reserved word, then the UTF-16LE text, each sized.
    fn dual(&mut self, text: &str, reserved: u16) {
        self.sized(text.as_bytes());
        self.u16(reserved);
        self.sized(&utf16(text));
    }

    fn reference_name(&mut self, name: &str) {
        self.u16(id::REFERENCE_NAME);
        self.dual(name, 0x003E);
    }

    fn reference(&mut self, r: &ReferenceFixture) {
        if let Some(name) = &r.name {
            self.reference_name(name);
        }
        let libid = r.libid.as_bytes();
        match r.kind {
            ReferenceKind::Control => self.control(r),
            ReferenceKind::Original => {
                self.u16(id::REFERENCE_ORIGINAL);
                self.sized(libid);
                self.control(r);
            }
            ReferenceKind::Registered => {
                self.u16(id::REFERENCE_REGISTERED);
                self.u32(total(&[4, libid.len(), 4, 2]));
                self.sized(libid);
                self.u32(0);
                self.u16(0);
            }
            ReferenceKind::Project => {
                self.u16(id::REFERENCE_PROJECT);
                self.u32(total(&[4, libid.len(), 4, libid.len(), 4, 2]));
                self.sized(libid);
                self.sized(libid);
                self.u32(1);
                self.u16(0);
            }
        }
    }

    /// REFERENCECONTROL, with a NameRecordExtended when the reference is named.
    fn control(&mut self, r: &ReferenceFixture) {
        let twiddled = r.libid.as_bytes();
        let extended = b"*\\G{00000000-0000-0000-0000-000000000000}#2.0#0#Example.exd#Example";
        self.u16(id::REFERENCE_CONTROL);
        self.u32(total(&[4, twiddled.len(), 4, 2]));
        self.sized(twiddled);
        self.u32(0);
        self.u16(0);
        if let Some(name) = &r.name {
            self.reference_name(name);
        }
        self.u16(0x0030); // Reserved3
        self.u32(total(&[4, extended.len(), 4, 2, 16, 4]));
        self.sized(extended);
        self.u32(0);
        self.u16(0);
        self.0.extend_from_slice(&[0; 16]); // OriginalTypeLib
        self.u32(1); // Cookie
    }

    fn module(&mut self, m: &ModuleFixture) {
        self.u16(id::MODULE_NAME);
        self.sized(m.name.as_bytes());
        if m.name_unicode {
            self.u16(id::MODULE_NAME_UNICODE);
            self.sized(&utf16(&m.name));
        }
        self.u16(id::MODULE_STREAM_NAME);
        self.dual(&m.stream_name, 0x0032);
        self.u16(id::MODULE_DOC_STRING);
        self.dual(&m.doc_string, 0x0048);
        self.fixed_u32(id::MODULE_OFFSET, m.text_offset);
        self.fixed_u32(id::MODULE_HELP_CONTEXT, 0);
        self.fixed_u16(id::MODULE_COOKIE, 0xFFFF);
        self.u16(match m.module_type {
            DirModuleType::Procedural => id::MODULE_TYPE_PROCEDURAL,
            DirModuleType::NonProcedural => id::MODULE_TYPE_OTHER,
        });
        self.u32(0);
        if m.read_only {
            self.u16(id::MODULE_READ_ONLY);
            self.u32(0);
        }
        if m.private {
            self.u16(id::MODULE_PRIVATE);
            self.u32(0);
        }
        self.u16(id::MODULE_TERMINATOR);
        self.u32(0);
    }
}

/// UTF-16LE bytes of `text`.
pub fn utf16(text: &str) -> Vec<u8> {
    text.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

/// A sum of field sizes for the ignored outer `Size` fields, saturating.
fn total(parts: &[usize]) -> u32 {
    let sum = parts.iter().fold(0usize, |acc, n| acc.saturating_add(*n));
    u32::try_from(sum).unwrap_or(u32::MAX)
}
