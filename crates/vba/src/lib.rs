//! Reads VBA projects (`vbaProject.bin`) from their public specifications ([MS-OVBA]) and lexes
//! and parses module source ([MS-VBAL]), prints syntax trees back as source, and reports what
//! the source uses ([`mod@analyze`]) against the checked-in Excel member catalog ([`catalog`]). Read-only: nothing here runs VBA.
//!
//! The crate depends on `gridcraft-core` only, so it stays small, wasm-safe and usable outside
//! the workbook stack. It never opens the compound file itself: callers hand it the streams.
//!
//! Text encoding: Phase 1 decodes only Windows-1252, UTF-8 and Mac OS Roman, with built-in
//! tables. Any other code page decodes as Windows-1252 with an `UnsupportedCodePage` diagnostic,
//! a deviation from the larger set proposed in Requirement 5 criterion 5. Whether to add
//! `encoding_rs` for the rest is a maintainers' decision for later. See [`encoding`].
//!
//! Compiled code is not checked. A module stream may hold a compiled-code cache (P_Code) before
//! its MODULEOFFSET. [MS-OVBA] §2.3.4.3 leaves that cache's layout undefined, and the clean-room
//! rule forbids learning it from Office binaries, so this crate counts those bytes but never
//! decodes them. Requirement 7 criterion 2 (compare the identifiers the P_Code uses with the
//! source and report a "source and compiled code differ" diagnostic) is therefore **not
//! delivered**. Instead, as criterion 3 allows when the identifier table can't be read, every
//! module that was read and has a cache gets a `CompiledCodeNotChecked` diagnostic, except a
//! module flagged `PossibleStomping`, which already says its source doesn't match its compiled
//! code. Criterion 2 can be added only from public documentation of the cache. See
//! [`project`].
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
//! [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
#![forbid(unsafe_code)]
pub mod analyze;
pub mod ast;
pub mod catalog;
pub mod diag;
pub mod dir;
pub mod encoding;
#[cfg(any(test, feature = "test-support"))]
pub mod fixture;
pub mod lexer;
pub mod ovba;
pub mod parser;
pub mod print;
pub mod project;
pub mod project_stream;
#[cfg(any(test, feature = "test-support"))]
pub mod testutil;

pub use analyze::{
    CapSpan, Capability, CapabilityManifest, CategoryEntry, CompatibilityReport, DynamicUse, EmptyCatalog, Flag, FlagKind, IncompleteReason, Mark,
    MemberCatalog, MemberStatus, MemberUse, ParamInfo, ProcInfo, ProgId, Reason, Verdict, VerdictRule, analyze,
};
pub use ast::Module;
pub use catalog::ExcelCatalog;
pub use diag::{Diag, DiagCode, Severity};
pub use dir::ReferenceKind;
pub use lexer::{Keyword, LexError, MAX_SOURCE, SourceTooLarge, Span, Token, TokenKind, lex};
pub use parser::{ParseDiag, ParseDiagKind, parse};
pub use print::print_module;
pub use project::{Limits, MAX_TOTAL_SOURCE, Reference, SIGNATURE_STREAMS, StreamSource, VbaModule, VbaProject, mark_signed, read_project};
pub use project_stream::ModuleKind;
