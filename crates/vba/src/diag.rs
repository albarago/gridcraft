//! Diagnostics: what went wrong (or is worth knowing) while reading a VBA project.
//!
//! A [`Diag`] is plain data: a stable [`DiagCode`], the module it concerns (if any) and a
//! user-facing sentence. Readers never fail; problems become diagnostics, so the workbook stays
//! open and editable. [`DiagCode::as_str`] gives the stable ids used in JSON. New codes are
//! added as new variants with new ids; existing ids never change. [`DiagCode::severity`] says how
//! much a diagnostic matters, so a view can sort or style it. It follows from the code, so a
//! [`Diag`] has no separate field for it.

/// One diagnostic. `module` is `None` for workbook-level diagnostics.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Diag {
    pub code: DiagCode,
    /// The module's display name, for diagnostics attached to one module.
    pub module: Option<String>,
    /// A plain sentence for users, naming the module when there is one.
    pub message: String,
}

impl Diag {
    /// A workbook-level diagnostic.
    #[must_use]
    pub fn project(code: DiagCode, message: impl Into<String>) -> Self {
        Self { code, module: None, message: message.into() }
    }

    /// A diagnostic about the module named `module`.
    #[must_use]
    pub fn module(code: DiagCode, module: &str, message: impl Into<String>) -> Self {
        Self { code, module: Some(module.to_owned()), message: message.into() }
    }
}

/// What a diagnostic is about. See [`DiagCode::as_str`] for the stable ids.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DiagCode {
    /// The `VBA/dir` stream is missing, so no modules can be listed.
    DirMissing,
    /// The `VBA/dir` stream couldn't be read, decompressed or parsed.
    DirCorrupt,
    /// The `PROJECT` stream is missing or unreadable; kinds come from MODULETYPE.
    ProjectStreamMissing,
    /// The `PROJECT` stream hit a reading cap, so some module kinds may be missing.
    ProjectStreamTruncated,
    /// The module's kind was taken from its MODULETYPE record, not the `PROJECT` stream.
    KindFromModuleType,
    /// The module's stream is missing from the `VBA` storage.
    ModuleStreamMissing,
    /// The module's MODULEOFFSET is at or past the end of its stream.
    ModuleOffsetOutOfRange,
    /// The module's compressed source is cut short.
    ModuleTruncated,
    /// The module's stream couldn't be read or its compressed source is damaged.
    ModuleCorrupt,
    /// The project code page isn't supported; text was decoded as Windows-1252.
    UnsupportedCodePage(u16),
    /// The project declares more modules than the reader keeps.
    ModuleLimit,
    /// The total source text passed the reader's limit; later modules were not read.
    SourceLimit,
    /// The module stream holds compiled code before MODULEOFFSET, but its source is empty or only
    /// `Attribute` lines: the source may have been removed to hide the macro (VBA stomping).
    PossibleStomping,
    /// The module stream holds compiled code before MODULEOFFSET. GridCraft doesn't decode it, so
    /// it can't confirm the source matches it (Req 7.3; the comparison of Req 7.2 is not
    /// implemented, see the crate docs).
    CompiledCodeNotChecked,
    /// The project carries a signature: a `\u{5}DigitalSignature*` stream in the project, or a
    /// VBA signature part in the workbook package. It is found, never read or verified, so the
    /// status is "signed (not verified)" (Req 7.4-7.5).
    Signed,
}

/// How much a diagnostic matters, from least to most.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Severity {
    /// Worth knowing; nothing is missing or wrong.
    Info,
    /// Something is incomplete, approximate or suspicious, but what was read is shown.
    Warning,
    /// Something couldn't be read: the module list, or a module's source.
    Error,
}

impl Severity {
    /// The stable id: `info`, `warning` or `error`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Info => "info",
            Self::Warning => "warning",
            Self::Error => "error",
        }
    }
}

impl DiagCode {
    /// The stable id used in JSON: `dirMissing`, `dirCorrupt`, …. The code page of
    /// [`DiagCode::UnsupportedCodePage`] is not part of the id; it is in the message.
    #[must_use]
    pub const fn as_str(&self) -> &'static str {
        match self {
            Self::DirMissing => "dirMissing",
            Self::DirCorrupt => "dirCorrupt",
            Self::ProjectStreamMissing => "projectStreamMissing",
            Self::ProjectStreamTruncated => "projectStreamTruncated",
            Self::KindFromModuleType => "kindFromModuleType",
            Self::ModuleStreamMissing => "moduleStreamMissing",
            Self::ModuleOffsetOutOfRange => "moduleOffsetOutOfRange",
            Self::ModuleTruncated => "moduleTruncated",
            Self::ModuleCorrupt => "moduleCorrupt",
            Self::UnsupportedCodePage(_) => "unsupportedCodePage",
            Self::ModuleLimit => "moduleLimit",
            Self::SourceLimit => "sourceLimit",
            Self::PossibleStomping => "possibleStomping",
            Self::CompiledCodeNotChecked => "compiledCodeNotChecked",
            Self::Signed => "signed",
        }
    }

    /// How much this diagnostic matters. Failing to read the module list or a module's source is
    /// an error; caps, fallbacks and possible stomping are warnings; compiled code that wasn't
    /// checked and a signature that wasn't verified are information.
    #[must_use]
    pub const fn severity(&self) -> Severity {
        match self {
            Self::DirMissing
            | Self::DirCorrupt
            | Self::ModuleStreamMissing
            | Self::ModuleOffsetOutOfRange
            | Self::ModuleTruncated
            | Self::ModuleCorrupt => Severity::Error,
            Self::ProjectStreamMissing
            | Self::ProjectStreamTruncated
            | Self::KindFromModuleType
            | Self::UnsupportedCodePage(_)
            | Self::ModuleLimit
            | Self::SourceLimit
            | Self::PossibleStomping => Severity::Warning,
            Self::CompiledCodeNotChecked | Self::Signed => Severity::Info,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{DiagCode, Severity};

    #[test]
    fn ids_are_unique_camel_case() {
        let all = [
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
        let mut ids: Vec<&str> = all.iter().map(DiagCode::as_str).collect();
        for id in &ids {
            assert!(id.chars().next().is_some_and(|c| c.is_ascii_lowercase()), "{id}");
            assert!(id.chars().all(|c| c.is_ascii_alphanumeric()), "{id}");
        }
        ids.sort_unstable();
        ids.dedup();
        assert_eq!(ids.len(), all.len());
    }

    #[test]
    fn severities_follow_the_code() {
        assert_eq!(DiagCode::CompiledCodeNotChecked.severity(), Severity::Info);
        assert_eq!(DiagCode::Signed.severity(), Severity::Info);
        assert_eq!(DiagCode::PossibleStomping.severity(), Severity::Warning);
        assert_eq!(DiagCode::UnsupportedCodePage(1251).severity(), Severity::Warning);
        assert_eq!(DiagCode::ModuleTruncated.severity(), Severity::Error);
        assert!(Severity::Info < Severity::Warning && Severity::Warning < Severity::Error);
        assert_eq!([Severity::Info, Severity::Warning, Severity::Error].map(Severity::as_str), ["info", "warning", "error"]);
    }
}
