//! Reads VBA projects (`vbaProject.bin`) from their public specifications ([MS-OVBA], later
//! [MS-VBAL]). Read-only: nothing here runs VBA.
//!
//! The crate depends on `gridcraft-core` only, so it stays small, wasm-safe and usable outside
//! the workbook stack. It never opens the compound file itself: callers hand it the streams.
//!
//! Text encoding: Phase 1 decodes only Windows-1252, UTF-8 and Mac OS Roman, with built-in
//! tables. Any other code page decodes as Windows-1252 with an `UnsupportedCodePage` diagnostic,
//! a deviation from the larger set proposed in Requirement 5 criterion 5. Whether to add
//! `encoding_rs` for the rest is a maintainers' decision for later. See [`encoding`].
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
//! [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]
#![forbid(unsafe_code)]
pub mod diag;
pub mod dir;
pub mod encoding;
#[cfg(any(test, feature = "test-support"))]
pub mod fixture;
pub mod ovba;
pub mod project;
pub mod project_stream;
#[cfg(any(test, feature = "test-support"))]
pub mod testutil;
pub use diag::{Diag, DiagCode};
pub use dir::ReferenceKind;
pub use project::{Limits, MAX_TOTAL_SOURCE, Reference, StreamSource, VbaModule, VbaProject, read_project};
pub use project_stream::ModuleKind;
