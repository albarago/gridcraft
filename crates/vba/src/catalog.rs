//! The checked-in Member_Catalog (Requirement 15): `data/excel-members.tsv`, generated offline by
//! `cargo xtask vba-catalog` from MicrosoftDocs/VBA-Docs (CC-BY-4.0, see `ATTRIBUTION.md`) and
//! compiled in with `include_str!`. Nothing reads VBA-Docs or the network at run time (Req 15.9).
//!
//! [`ExcelCatalog`] is the [`MemberCatalog`] the engine's `vba.analyze` uses. The file is parsed
//! once, on the first lookup ([`OnceLock`], which also works on wasm). A row that doesn't have
//! seven tab-separated columns, a known kind and a known status is skipped, never a panic; a test
//! checks the shipped file has none.
//!
//! # Statuses
//!
//! `supported`, `planned` and `notPlanned` map to the [`MemberStatus`] of the same name.
//! `unreviewed` (nobody has looked at the member yet, the default for new rows) maps to
//! [`MemberStatus::Unknown`], exactly like a member the catalog doesn't list (Req 13.6), so an
//! unreviewed catalog makes a report "analysis only", never "uses unsupported features". The JSON
//! status set is unchanged.
//!
//! # Kinds
//!
//! The file keys rows on (object, member, kind), but a source reference `x.Name` has no kind.
//! An event can't be referenced with `.`, so property and method rows win over event rows (the
//! method `Workbook.Activate`, not the event); event rows answer only when the member has no other
//! kind. Should two rows of the same group remain (the generated file has no property and method
//! of the same name), the best status wins: supported, then planned, then not planned, then
//! unreviewed.
//!
//! # Objects
//!
//! The analyzer names the object as written in the source (`ws`, `Range`, `Cells`, `Me`, …); the
//! catalog lists Excel type names. A lookup compares case-insensitively and tries, in order:
//!
//! 1. the types in [`ALIASES`] for a few unambiguous global and Range-returning names (`Cells`,
//!    `ActiveCell`, `UsedRange`, … are `Range`; `ActiveWorkbook` and `ThisWorkbook` are
//!    `Workbook`; `ActiveSheet` is taken as `Worksheet`, its usual type, though Excel types it as
//!    `Object`). A collection tries itself, then its item type, because the analyzer looks through
//!    calls: `Worksheets("A").Range` is named `Range` on `Worksheets`, while `Worksheets.Add` is a
//!    collection member;
//! 2. otherwise the name itself, when it is a catalog object (`Range`, `Application`, `Font`).
//!
//! Anything else (a variable such as `ws`, `Me`, a `With` with no name) has no type information
//! yet, so the lookup gives `None` and the member is "unknown".
use std::collections::HashMap;
use std::sync::OnceLock;

use crate::analyze::{MemberCatalog, MemberStatus};

/// The generated catalog (see the module docs and `xtask/src/vba_catalog.rs` for the format).
pub static RAW: &str = include_str!("../data/excel-members.tsv");

/// Source names that stand for an Excel type, with the types to try in order (all lower case).
pub const ALIASES: &[(&str, &[&str])] = &[
    ("activecell", &["range"]),
    ("activesheet", &["worksheet"]),
    ("activeworkbook", &["workbook"]),
    ("cells", &["range"]),
    ("columns", &["range"]),
    ("currentregion", &["range"]),
    ("end", &["range"]),
    ("entirecolumn", &["range"]),
    ("entirerow", &["range"]),
    ("offset", &["range"]),
    ("resize", &["range"]),
    ("rows", &["range"]),
    ("sheets", &["sheets", "worksheet"]),
    ("thisworkbook", &["workbook"]),
    ("usedrange", &["range"]),
    ("workbooks", &["workbooks", "workbook"]),
    ("worksheets", &["worksheets", "worksheet"]),
];

/// A row's status as the file spells it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum RowStatus {
    // Declared worst first, so `max` is the best status.
    Unreviewed,
    NotPlanned,
    Planned,
    Supported,
}

impl RowStatus {
    fn parse(s: &str) -> Option<Self> {
        match s {
            "supported" => Some(Self::Supported),
            "planned" => Some(Self::Planned),
            "notPlanned" => Some(Self::NotPlanned),
            "unreviewed" => Some(Self::Unreviewed),
            _ => None,
        }
    }

    const fn member_status(self) -> MemberStatus {
        match self {
            Self::Supported => MemberStatus::Supported,
            Self::Planned => MemberStatus::Planned,
            Self::NotPlanned => MemberStatus::NotPlanned,
            Self::Unreviewed => MemberStatus::Unknown,
        }
    }
}

/// The best status of a member's property/method rows and of its event rows.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct Entry {
    member: Option<RowStatus>,
    event: Option<RowStatus>,
}

impl Entry {
    fn status(self) -> Option<RowStatus> {
        self.member.or(self.event)
    }
}

/// A parsed catalog: `(object, member)` lower-cased to its [`Entry`], plus the objects it lists.
#[derive(Debug, Default)]
pub struct Parsed {
    rows: HashMap<(String, String), Entry>,
    objects: std::collections::HashSet<String>,
    /// Data rows that were skipped as malformed.
    pub bad_rows: usize,
}

impl Parsed {
    /// The number of `(object, member)` pairs.
    #[must_use]
    pub fn len(&self) -> usize {
        self.rows.len()
    }

    /// Whether no row was read.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.rows.is_empty()
    }

    fn get(&self, object: &str, member: &str) -> Option<RowStatus> {
        self.rows.get(&(object.to_owned(), member.to_owned())).and_then(|e| e.status())
    }

    /// The status of `member` on the source name `object` (see the module docs).
    #[must_use]
    pub fn status(&self, object: Option<&str>, member: &str) -> Option<MemberStatus> {
        let object = object?.to_lowercase();
        let member = member.to_lowercase();
        let found = match ALIASES.iter().find(|(name, _)| *name == object) {
            Some((_, types)) => types.iter().find_map(|t| self.get(t, &member)),
            None if self.objects.contains(&object) => self.get(&object, &member),
            None => None,
        };
        found.map(RowStatus::member_status)
    }
}

/// Parses catalog text; comment (`#`) and empty lines are ignored, malformed rows counted.
#[must_use]
pub fn parse(text: &str) -> Parsed {
    let mut out = Parsed::default();
    for line in text.lines() {
        if line.is_empty() || line.starts_with('#') {
            continue;
        }
        let fields: Vec<&str> = line.split('\t').collect();
        let [object, member, kind, _params, _returns, status, _note] = fields.as_slice() else {
            out.bad_rows = out.bad_rows.saturating_add(1);
            continue;
        };
        let event = match *kind {
            "property" | "method" => false,
            "event" => true,
            _ => {
                out.bad_rows = out.bad_rows.saturating_add(1);
                continue;
            }
        };
        let (Some(status), true) = (RowStatus::parse(status), !object.is_empty() && !member.is_empty()) else {
            out.bad_rows = out.bad_rows.saturating_add(1);
            continue;
        };
        let object = object.to_lowercase();
        let entry = out.rows.entry((object.clone(), member.to_lowercase())).or_default();
        let slot = if event { &mut entry.event } else { &mut entry.member };
        *slot = Some(slot.map_or(status, |s| s.max(status)));
        out.objects.insert(object);
    }
    out
}

/// The shipped catalog, parsed on first use.
pub fn shipped() -> &'static Parsed {
    static PARSED: OnceLock<Parsed> = OnceLock::new();
    PARSED.get_or_init(|| parse(RAW))
}

/// The checked-in Excel member catalog.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ExcelCatalog;

impl MemberCatalog for ExcelCatalog {
    fn status(&self, object: Option<&str>, member: &str) -> Option<MemberStatus> {
        shipped().status(object, member)
    }
}

#[cfg(test)]
mod tests;
