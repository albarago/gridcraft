//! The `PROJECT` stream: project information as text ([MS-OVBA] §2.3.1).
//!
//! The stream is MBCS text in the project code page, one property per line, ended by CRLF. It
//! starts with the ProjectProperties (`ID=`, `Document=`, `Module=`, `Package=`, `Name=`, …) and
//! then has the `[Host Extender Info]` and `[Workspace]` sections. Only the ProjectModules
//! properties matter here, because they give each module's kind (§2.3.1.3–§2.3.1.7):
//!
//! | Property     | [`ModuleKind`]           |
//! |--------------|--------------------------|
//! | `Module=`    | [`ModuleKind::Standard`] |
//! | `Class=`     | [`ModuleKind::Class`]    |
//! | `BaseClass=` | [`ModuleKind::Form`]     |
//! | `Document=`  | [`ModuleKind::Document`] |
//!
//! `Document=` values end in `/&H…` (the DocTlibVer); that suffix is dropped.
//!
//! [`parse_project_stream`] never fails. It is lenient where real files differ from the ABNF:
//!
//! - LF-only line endings are accepted, and ASCII whitespace around keys and values is trimmed.
//! - Property names match ASCII case-insensitively (ABNF string literals are case-insensitive).
//! - Unknown properties and lines without `=` are skipped. Empty module names are skipped.
//! - Reading stops at the first line starting with `[`, so `[Workspace]` entries such as
//!   `ThisWorkbook=23, 23, 911, 280,` are never taken for properties.
//! - Garbage gives an empty result.
//!
//! Names stay raw MBCS bytes; decoding happens later, once the code page is known. They are
//! matched against the `dir` stream's MODULENAME bytes with [`resolve_kind`]: byte for byte,
//! with ASCII letters folded. In a double-byte code page a trail byte can fall in the ASCII
//! letter range, so two names that differ only there would match; VBA itself rejects such
//! near-duplicates, so this is accepted.
//!
//! Caps, so a hostile stream costs bounded time and memory:
//!
//! - [`MAX_PROJECT_STREAM`] (1 MiB): only that many bytes are read. The line cut by the cap is
//!   dropped rather than read as a shorter name.
//! - [`MAX_LINES`] (65,536): only that many lines are read.
//!
//! Hitting either cap with properties still unread sets [`ProjectStreamInfo::truncated`]. Real
//! PROJECT streams are a few kilobytes: about one line per module plus a dozen others.
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/

use crate::dir::DirModuleType;

/// The most bytes of a PROJECT stream that are read (1 MiB).
pub const MAX_PROJECT_STREAM: usize = 1024 * 1024;
/// The most lines of a PROJECT stream that are read.
pub const MAX_LINES: usize = 65_536;

/// What a module is, as shown to users and used by the engine.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ModuleKind {
    /// A procedural (standard) module: `Module=`.
    Standard,
    /// A class module: `Class=`.
    Class,
    /// A document module behind a workbook or sheet: `Document=`.
    Document,
    /// A designer module (a UserForm): `BaseClass=`.
    Form,
}

impl ModuleKind {
    /// A stable lowercase name for JSON and logs: `standard`, `class`, `document` or `form`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Standard => "standard",
            Self::Class => "class",
            Self::Document => "document",
            Self::Form => "form",
        }
    }

    /// The fallback kind from a MODULETYPE record: procedural gives standard, otherwise class.
    #[must_use]
    pub const fn from_module_type(module_type: DirModuleType) -> Self {
        match module_type {
            DirModuleType::Procedural => Self::Standard,
            DirModuleType::NonProcedural => Self::Class,
        }
    }
}

/// The module kinds listed in a PROJECT stream.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ProjectStreamInfo {
    /// Each module property in stream order: the MBCS name and its kind.
    pub modules: Vec<(Vec<u8>, ModuleKind)>,
    /// True when [`MAX_PROJECT_STREAM`] or [`MAX_LINES`] stopped reading before the end of the
    /// properties, so some modules may be missing.
    pub truncated: bool,
}

impl ProjectStreamInfo {
    /// The kind of the first module property whose name equals `name`, ASCII case-insensitively.
    #[must_use]
    pub fn kind_of(&self, name: &[u8]) -> Option<ModuleKind> {
        self.modules.iter().find(|(n, _)| n.eq_ignore_ascii_case(name)).map(|&(_, kind)| kind)
    }
}

/// Parses a PROJECT stream's module properties. Never fails: garbage gives an empty result.
#[must_use]
pub fn parse_project_stream(bytes: &[u8]) -> ProjectStreamInfo {
    let mut info = ProjectStreamInfo::default();
    let text = match bytes.get(..MAX_PROJECT_STREAM) {
        Some(head) if bytes.len() > MAX_PROJECT_STREAM => {
            info.truncated = true;
            // Drop the line the cap cut through; keep everything up to its start.
            match head.iter().rposition(|&b| b == b'\n') {
                Some(end) => head.get(..end).unwrap_or_default(),
                None => &[],
            }
        }
        _ => bytes,
    };
    let mut lines = text.split(|&b| b == b'\n');
    let mut read = 0usize;
    while let Some(line) = lines.next() {
        let line = line.trim_ascii();
        if line.first() == Some(&b'[') {
            // The properties end at the first section header; nothing after it counts, even
            // if a byte cap or line cap was hit.
            info.truncated = false;
            return info;
        }
        if read == MAX_LINES {
            // Truncated only if a non-blank line is left before the next section header.
            info.truncated |= !line.is_empty() || lines.map(<[u8]>::trim_ascii).take_while(|l| l.first() != Some(&b'[')).any(|l| !l.is_empty());
            return info;
        }
        read = read.saturating_add(1);
        if let Some(module) = module_property(line) {
            info.modules.push(module);
        }
    }
    info
}

/// Reads one trimmed line as a ProjectModules property, or `None` for anything else.
fn module_property(line: &[u8]) -> Option<(Vec<u8>, ModuleKind)> {
    let eq = line.iter().position(|&b| b == b'=')?;
    let key = line.get(..eq)?.trim_ascii();
    let value = line.get(eq.checked_add(1)?..)?.trim_ascii();
    let (kind, name) = if key.eq_ignore_ascii_case(b"Module") {
        (ModuleKind::Standard, value)
    } else if key.eq_ignore_ascii_case(b"Class") {
        (ModuleKind::Class, value)
    } else if key.eq_ignore_ascii_case(b"BaseClass") {
        (ModuleKind::Form, value)
    } else if key.eq_ignore_ascii_case(b"Document") {
        // `Document=Name/&H00000000`: the name ends at the last '/' (names never contain one,
        // and no supported code page uses 0x2F as a trail byte).
        let name = match value.iter().rposition(|&b| b == b'/') {
            Some(slash) => value.get(..slash)?.trim_ascii(),
            None => value,
        };
        (ModuleKind::Document, name)
    } else {
        return None;
    };
    (!name.is_empty()).then(|| (name.to_vec(), kind))
}

/// Resolves a `dir` module's kind ([MS-OVBA] §2.3.1, Req 5.7–5.8).
///
/// Uses the PROJECT stream entry whose name equals `dir_name` (the MODULENAME MBCS bytes),
/// ASCII case-insensitively. When the stream is missing (`info` is `None`) or has no entry for
/// the module, falls back to [`ModuleKind::from_module_type`]. The `bool` is true on fallback:
/// the caller then emits the "kind taken from MODULETYPE" diagnostic naming the module.
///
/// [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
#[must_use]
pub fn resolve_kind(info: Option<&ProjectStreamInfo>, dir_name: &[u8], module_type: DirModuleType) -> (ModuleKind, bool) {
    match info.and_then(|i| i.kind_of(dir_name)) {
        Some(kind) => (kind, false),
        None => (ModuleKind::from_module_type(module_type), true),
    }
}

#[cfg(test)]
mod tests;
