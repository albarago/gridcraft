//! `cargo xtask vba-catalog --docs DIR [--out PATH] [--commit SHA]`: build the VBA Member_Catalog
//! (`crates/vba/data/excel-members.tsv`) from a MicrosoftDocs/VBA-Docs checkout (CC-BY-4.0).
//!
//! Offline only: `cargo xtask ci` never runs it, and nothing at run time reads VBA-Docs (Req 15.8,
//! 15.9). It reads every `DIR/api/Excel.*.md` page and keeps the member pages, whose title is
//! `Object.Member property|method|event (Excel)` (`api_name: - Excel.Object.Member` is the fallback
//! for the names). Object, collection and enumeration overview pages, Excel Graph pages and pages it
//! can't parse are skipped and listed with a reason in the run summary (Req 15.2).
//!
//! Only names and signatures are taken from a page: the object, member and kind, the parameter
//! names, required/optional and types (Parameters table, else the Syntax line), and the return type
//! (Return value / Property value section, else the type token after `Read/write`, `Read-only` or a
//! bold run just after `returns` in the summary line). No description, remark or example text is copied
//! (Req 15.7).
//!
//! File format (UTF-8, LF, one member per line, seven tab-separated columns):
//!
//! ```text
//! object  member  kind  params  returns  status  note
//! ```
//!
//! - `kind`: `property`, `method` or `event`.
//! - `params`: `;`-separated `name:req|opt:type` in page order, or empty when there are no
//!   parameters. `type` is empty when the page gives none, `?` when it isn't one type token, and
//!   ends in `()` for an array parameter (`Strings()`). A table row naming a numbered range
//!   (`Arg1 – Arg30`) becomes one parameter per number.
//! - `returns`: a type name, `none` (events, `Nothing`, methods that return nothing), or `?` when
//!   the page doesn't state a type in a form the parser recognises.
//! - `status`: `supported`, `planned`, `notPlanned` or `unreviewed`. The first three are the
//!   `MemberStatus::as_str` spellings; the analyzer reports an `unreviewed` member as `unknown`
//!   (Req 13.6), so an unreviewed catalog never makes a workbook "uses unsupported features".
//! - `note`: free text of at most 500 characters, with no tab, CR, LF or other control character.
//!
//! Lines starting with `#` are comments: the provenance header (VBA-Docs commit and licence) and the
//! column header. Rows are sorted by object, then member, then kind (ASCII case-insensitive), so the
//! same pages and the same status file always give the same bytes (Req 15.5). A re-run keeps the
//! status and note of every row whose (object, member, kind) still exists (case-insensitive); new
//! rows get `unreviewed` and an empty note, and rows that disappeared are listed in the run summary
//! (Req 15.3, 15.4).
//!
//! # Tranche 1 seed
//!
//! [`TRANCHE_1`] is the checked-in list of the Phase 5 first tranche (Req 23.1). After the merge,
//! every property or method row (never an event) that the list names and whose status is still
//! `unreviewed` becomes `planned`, with the note `tranche 1 (Req 23.1)` when its note is empty. A
//! status a person set (`supported`, `planned` or `notPlanned`) is never changed. The run summary
//! counts the seeded rows and lists every seed entry that names no property or method in the docs.
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

/// Where the checked-in catalog lives (relative to the workspace root).
const DEFAULT_OUT: &str = "crates/vba/data/excel-members.tsv";
/// The longest note a row may carry, in characters (Req 15.3).
pub const NOTE_MAX_CHARS: usize = 500;
/// Pages larger than this are skipped (the largest Excel page is far smaller).
const PAGE_MAX_BYTES: u64 = 1 << 20;
/// The most pages read from one checkout.
const PAGES_MAX: usize = 100_000;
/// The largest existing catalog read back for the merge.
const CATALOG_MAX_BYTES: u64 = 16 << 20;
/// The most parameters one member may have.
const PARAMS_MAX: usize = 1024;
/// How far after "returns" the type token may start, in characters.
const RETURNS_WINDOW: usize = 40;
/// The column header line.
const COLUMNS: &str = "# object\tmember\tkind\tparams\treturns\tstatus\tnote";

/// The note a seeded row gets when it has none.
pub const SEED_NOTE: &str = "tranche 1 (Req 23.1)";

/// The Phase 5 first tranche (Req 23.1): `(object, members)`, Excel type names as the docs spell
/// them. "Core" `Workbook(s)`/`Worksheet(s)` members are the ones the tranche's own members need
/// (open, add, index, count, activate, save, name).
pub const TRANCHE_1: &[(&str, &[&str])] = &[
    (
        "Application",
        &["ActiveSheet", "ActiveWorkbook", "ActiveCell", "Selection", "ScreenUpdating", "EnableEvents", "Calculation", "WorksheetFunction", "Run"],
    ),
    ("Workbooks", &["Add", "Open", "Item", "Count", "Close"]),
    ("Workbook", &["Activate", "Close", "Save", "SaveAs", "Sheets", "Worksheets", "Name", "FullName", "Path"]),
    ("Worksheets", &["Add", "Item", "Count"]),
    ("Worksheet", &["Activate", "Select", "Name", "Cells", "Range", "UsedRange", "Rows", "Columns"]),
    (
        "Range",
        &[
            "Cells",
            "Rows",
            "Columns",
            "Value",
            "Value2",
            "Formula",
            "FormulaR1C1",
            "Offset",
            "Resize",
            "End",
            "CurrentRegion",
            "Font",
            "Interior",
            "NumberFormat",
            "ClearContents",
            "Copy",
            "PasteSpecial",
            "Select",
            "Activate",
        ],
    ),
    ("Font", &["Bold", "Italic", "Size", "Name", "Color"]),
    ("Interior", &["Color"]),
];

/// Whether [`TRANCHE_1`] names `object.member` (ASCII case-insensitive).
fn seeded(object: &str, member: &str) -> bool {
    TRANCHE_1.iter().any(|(o, ms)| o.eq_ignore_ascii_case(object) && ms.iter().any(|m| m.eq_ignore_ascii_case(member)))
}

const USAGE: &str = "usage: cargo xtask vba-catalog --docs <VBA-Docs checkout> [--out PATH] [--commit SHA]";

/// A member's kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub enum Kind {
    Event,
    Method,
    Property,
}

impl Kind {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Event => "event",
            Self::Method => "method",
            Self::Property => "property",
        }
    }

    fn parse(s: &str) -> Option<Self> {
        match s {
            "event" => Some(Self::Event),
            "method" => Some(Self::Method),
            "property" => Some(Self::Property),
            _ => None,
        }
    }
}

/// A row's catalog status (the first three spelled as `gridcraft_vba::MemberStatus::as_str`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Status {
    Supported,
    Planned,
    NotPlanned,
    /// Nobody has reviewed the member yet: the default for new rows (Req 15.4).
    Unreviewed,
}

impl Status {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Supported => "supported",
            Self::Planned => "planned",
            Self::NotPlanned => "notPlanned",
            Self::Unreviewed => "unreviewed",
        }
    }

    fn parse(s: &str) -> Option<Self> {
        match s {
            "supported" => Some(Self::Supported),
            "planned" => Some(Self::Planned),
            "notPlanned" => Some(Self::NotPlanned),
            "unreviewed" => Some(Self::Unreviewed),
            _ => None,
        }
    }
}

/// One parameter of a member.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Param {
    pub name: String,
    pub optional: bool,
    /// The declared type, or empty when the page gives none.
    pub ty: String,
}

/// One member page, reduced to names and signature.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Member {
    pub object: String,
    pub member: String,
    pub kind: Kind,
    pub params: Vec<Param>,
    /// A type name, `none`, or `?`.
    pub returns: String,
}

/// One input page: its file name, and its text or the reason it couldn't be read.
pub type Page = (String, Result<String, String>);

/// The merge key: object and member lower-cased, and the kind.
type Key = (String, String, Kind);

fn key(object: &str, member: &str, kind: Kind) -> Key {
    (object.to_ascii_lowercase(), member.to_ascii_lowercase(), kind)
}

/// One row of an existing catalog.
#[derive(Debug, Clone, PartialEq, Eq)]
struct Row {
    object: String,
    member: String,
    kind: Kind,
    status: Status,
    note: String,
}

/// What a run did, for the summary.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct Summary {
    pub pages_read: usize,
    pub members: usize,
    /// Members whose status and note came from the existing catalog.
    pub kept: usize,
    /// Members new to the catalog (`unreviewed`, empty note, before the seed).
    pub added: usize,
    /// Rows the [`TRANCHE_1`] seed moved from `unreviewed` to `planned`.
    pub seeded: usize,
    /// `Object.Member` for every [`TRANCHE_1`] entry that names no property or method in the docs.
    pub seed_unmatched: Vec<String>,
    /// Members whose return type is `?`.
    pub returns_unknown: usize,
    /// `(page, reason)` for every skipped page, in page order.
    pub skipped: Vec<(String, String)>,
    /// `Object.Member (kind)` for every existing row no longer in the docs, sorted.
    pub removed: Vec<String>,
}

// ---------------------------------------------------------------------------------------------
// Page parsing
// ---------------------------------------------------------------------------------------------

/// A valid VBA identifier as used in the docs: ASCII letters, digits and `_`, not starting with a
/// digit.
fn is_ident(s: &str) -> bool {
    let mut chars = s.chars();
    chars.next().is_some_and(|c| c.is_ascii_alphabetic() || c == '_') && chars.all(|c| c.is_ascii_alphanumeric() || c == '_')
}

/// A type token: starts with an ASCII capital, then letters, digits, `_` or `.`.
fn is_type_name(s: &str) -> bool {
    let mut chars = s.chars();
    chars.next().is_some_and(|c| c.is_ascii_uppercase()) && chars.all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
}

/// At most 40 characters of `s`, for a skip reason.
fn short(s: &str) -> String {
    let mut out: String = s.chars().take(40).collect();
    if s.chars().nth(40).is_some() {
        out.push('…');
    }
    out
}

/// The front matter's `title` and `api_name` entries, and the body lines after it.
fn front_matter(text: &str) -> Result<(String, Vec<String>, Vec<&str>), String> {
    let text = text.strip_prefix('\u{feff}').unwrap_or(text);
    let lines: Vec<&str> = text.lines().collect();
    if lines.first().map(|l| l.trim()) != Some("---") {
        return Err("no front matter".into());
    }
    let end = lines.iter().skip(1).position(|l| l.trim() == "---").ok_or("unterminated front matter")?;
    let fm = lines.get(1..=end).unwrap_or_default();
    let body = lines.get(end.saturating_add(2)..).unwrap_or_default().to_vec();
    let mut title = String::new();
    let mut api = Vec::new();
    let mut in_api = false;
    for line in fm {
        if let Some(v) = line.strip_prefix("title:") {
            title = v.trim().trim_matches(|c| c == '"' || c == '\'').trim().to_string();
            in_api = false;
        } else if line.trim_end() == "api_name:" {
            in_api = true;
        } else if in_api && let Some(item) = line.trim_start().strip_prefix("- ") {
            api.push(item.trim().trim_matches(|c| c == '"' || c == '\'').to_string());
        } else {
            in_api = false;
        }
    }
    Ok((title, api, body))
}

/// The kind and the `Object.Member` part of a member page title.
fn title_kind(title: &str) -> Result<(Kind, &str), String> {
    if title.is_empty() {
        return Err("no title".into());
    }
    if title.ends_with("(Excel Graph)") {
        return Err("Excel Graph page (not the Excel object model)".into());
    }
    let head = title.strip_suffix("(Excel)").ok_or_else(|| format!("title `{}` doesn't end in (Excel)", short(title)))?.trim_end();
    let (name, word) = head.rsplit_once(' ').ok_or("no member kind in title")?;
    let word = word.to_ascii_lowercase();
    match word.as_str() {
        "property" => Ok((Kind::Property, name.trim())),
        "method" => Ok((Kind::Method, name.trim())),
        "event" => Ok((Kind::Event, name.trim())),
        "object" | "collection" | "enumeration" | "interface" | "objects" => Err(format!("overview page ({word})")),
        _ => Err(format!("unknown page kind `{}`", short(&word))),
    }
}

/// `Object.Member` split into two identifiers.
fn split_name(name: &str) -> Option<(String, String)> {
    let (object, member) = name.split_once('.')?;
    (is_ident(object) && is_ident(member)).then(|| (object.to_string(), member.to_string()))
}

/// The object and member: from the title, else from the one `Excel.Object.Member` api_name.
fn names(title_name: &str, api: &[String]) -> Result<(String, String), String> {
    if let Some(pair) = split_name(title_name) {
        return Ok(pair);
    }
    let candidates: Vec<(String, String)> = api.iter().filter_map(|a| a.strip_prefix("Excel.")).filter_map(split_name).collect();
    match candidates.as_slice() {
        [one] => Ok(one.clone()),
        [] => Err("no object or member name".into()),
        _ => Err("several api_name entries and no Object.Member title".into()),
    }
}

/// The body split into the summary text (between the H1 and the first `##`) and the `##`
/// sections keyed by lower-cased heading.
fn sections<'a>(body: &[&'a str]) -> (String, BTreeMap<String, Vec<&'a str>>) {
    let mut summary = Vec::new();
    let mut map: BTreeMap<String, Vec<&'a str>> = BTreeMap::new();
    let mut current: Option<String> = None;
    for line in body {
        if let Some(h) = line.strip_prefix("## ") {
            let h = h.trim().trim_end_matches(':').to_ascii_lowercase();
            map.entry(h.clone()).or_default();
            current = Some(h);
        } else if let Some(h) = &current {
            if let Some(v) = map.get_mut(h) {
                v.push(line);
            }
        } else if !line.starts_with("# ") && !line.trim().is_empty() && !line.trim_start().starts_with("[!") {
            summary.push(line.trim());
        }
    }
    (summary.join(" "), map)
}

/// `[Text](link)` → `Text`; strips `**`, `_`, `\`, spaces.
fn clean_cell(cell: &str) -> String {
    let mut s = cell.replace("**", "");
    if let Some(open) = s.find('[')
        && let Some(close) = s.get(open..).and_then(|t| t.find("](")).and_then(|c| c.checked_add(open))
    {
        let inner = s.get(open.saturating_add(1)..close).unwrap_or_default().to_string();
        // The last `)`: link targets such as `Excel.Range(object).md` hold parentheses themselves.
        let after = s.get(close..).and_then(|t| t.rfind(')')).and_then(|p| p.checked_add(close)).and_then(|p| p.checked_add(1));
        let tail = after.and_then(|a| s.get(a..)).unwrap_or_default().to_string();
        s = format!("{}{inner}{tail}", s.get(..open).unwrap_or_default());
    }
    s.trim_matches(|c: char| c == '_' || c == '*' || c == '\\' || c.is_whitespace()).to_string()
}

/// The first type token among the `**bold**` runs of `text` (`True`/`False` mean `Boolean`).
fn bold_type(text: &str) -> Option<String> {
    let mut rest = text;
    while let Some(start) = rest.find("**") {
        let after = rest.get(start.saturating_add(2)..)?;
        let end = after.find("**")?;
        let token = clean_cell(after.get(..end)?);
        if token == "True" || token == "False" {
            return Some("Boolean".into());
        }
        if is_type_name(&token) {
            return Some(token);
        }
        rest = after.get(end.saturating_add(2)..)?;
    }
    None
}

/// A table type cell: a type name, empty, or `?` when it isn't a single type token.
fn param_type(cell: &str) -> String {
    // The cell holds only the type, so a lower-case first letter (`xlSortOrder`) is still a type;
    // `[INT]` (brackets without a link) is `INT`.
    let t = clean_cell(cell);
    let t = t.strip_prefix('[').and_then(|x| x.strip_suffix(']')).unwrap_or(&t);
    let typeish = t.chars().next().is_some_and(|c| c.is_ascii_alphabetic()) && t.chars().all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.');
    if t.is_empty() || typeish { t.to_string() } else { "?".into() }
}

/// The most parameters one table row may stand for (`Arg1 – Arg30`).
const RANGE_MAX: u32 = 255;

/// `Arg7` → `("Arg", 7)`.
fn numbered(s: &str) -> Option<(&str, u32)> {
    let cut = s.trim_end_matches(|c: char| c.is_ascii_digit()).len();
    let (prefix, digits) = (s.get(..cut)?, s.get(cut..)?);
    if digits.is_empty() || !is_ident(prefix) {
        return None;
    }
    digits.parse().ok().map(|n| (prefix, n))
}

/// The names a Parameters-table name cell stands for, and whether it is an array: one
/// identifier, `Name()`, or a numbered range such as `Arg1 – Arg30` (`-`, `–`, `—` or `&ndash;`).
fn row_names(cell: &str) -> Result<(Vec<String>, bool), String> {
    let name = clean_cell(cell);
    if is_ident(&name) {
        return Ok((vec![name], false));
    }
    if let Some(base) = name.strip_suffix("()")
        && is_ident(base)
    {
        return Ok((vec![base.to_string()], true));
    }
    let flat: String = name.replace("&ndash;", "-").replace(['–', '—'], "-").chars().filter(|c| !c.is_whitespace() && *c != '_').collect();
    if let Some((a, b)) = flat.split_once('-')
        && let (Some((pa, from)), Some((pb, to))) = (numbered(a), numbered(b))
        && pa == pb
        && to.checked_sub(from).is_some_and(|d| d < RANGE_MAX)
    {
        return Ok(((from..=to).map(|i| format!("{pa}{i}")).collect(), false));
    }
    Err(format!("bad parameter name `{}`", short(&name)))
}

/// A header (`Name`, `Required/Optional`) or `|:---|` separator row.
fn is_header_row(cells: &[&str]) -> bool {
    let first = cells.first().map(|c| c.trim()).unwrap_or_default();
    let second = clean_cell(cells.get(1).copied().unwrap_or_default());
    first.chars().all(|c| c == ':' || c == '-') || clean_cell(first).eq_ignore_ascii_case("name") || second.eq_ignore_ascii_case("required/optional")
}

/// The Parameters table rows, or `None` when the page has no table.
fn table_params(lines: &[&str]) -> Result<Option<Vec<Param>>, String> {
    let mut params = Vec::new();
    let mut seen_table = false;
    for line in lines {
        let Some(inner) = line.trim().strip_prefix('|') else { continue };
        seen_table = true;
        let cells: Vec<&str> = inner.split('|').collect();
        if is_header_row(&cells) {
            continue;
        }
        let (names, array) = row_names(cells.first().copied().unwrap_or_default())?;
        let optional = match clean_cell(cells.get(1).copied().unwrap_or_default()).to_ascii_lowercase().as_str() {
            "optional" => true,
            "required" => false,
            other => return Err(format!("parameter {}: bad Required/Optional cell `{}`", names.join(","), short(other))),
        };
        let mut ty = param_type(cells.get(2).copied().unwrap_or_default());
        if array && !ty.is_empty() && ty != "?" {
            ty.push_str("()");
        }
        if params.len().saturating_add(names.len()) > PARAMS_MAX {
            return Err(format!("more than {PARAMS_MAX} parameters"));
        }
        params.extend(names.into_iter().map(|name| Param { name, optional, ty: ty.clone() }));
    }
    Ok(seen_table.then_some(params))
}

/// The arguments on the Syntax line `expression.**Member** (A, [B])` (no types).
fn syntax_params(lines: &[&str], member: &str) -> Vec<Param> {
    let needle = member.to_ascii_lowercase();
    for line in lines {
        let lower = line.to_ascii_lowercase();
        let Some(pos) = lower.find(&needle) else { continue };
        let rest = lower.get(pos.saturating_add(needle.len())..).unwrap_or_default();
        let rest = rest.trim_start_matches('*').trim_start();
        let Some(args) = rest.strip_prefix('(') else { return Vec::new() };
        let original = line.get(line.len().saturating_sub(args.len())..).unwrap_or_default();
        let Some(close) = original.rfind(')') else { return Vec::new() };
        let mut out = Vec::new();
        for arg in original.get(..close).unwrap_or_default().split(',') {
            let optional = arg.contains('[');
            let name = arg.trim_matches(|c: char| c == '[' || c == ']' || c == '_' || c == '*' || c == '\\' || c.is_whitespace());
            if name.is_empty() {
                continue;
            }
            if !is_ident(name) {
                return Vec::new();
            }
            out.push(Param { name: name.to_string(), optional, ty: String::new() });
        }
        return out;
    }
    Vec::new()
}

/// The return type: `none`, a type, or `?`.
fn return_type(kind: Kind, summary: &str, map: &BTreeMap<String, Vec<&str>>) -> String {
    if kind == Kind::Event {
        return "none".into();
    }
    let section = ["return value", "return values", "property value"].iter().find_map(|h| map.get(*h));
    if let Some(lines) = section {
        let first = lines.iter().map(|l| l.trim()).find(|l| !l.is_empty()).unwrap_or_default();
        let lower = first.to_ascii_lowercase();
        if ["nothing", "none", "void", "**nothing**", "**void**"].contains(&lower.as_str()) || lower.starts_with("**nothing**") {
            return "none".into();
        }
        let plain = clean_cell(first);
        if is_type_name(&plain) {
            return plain;
        }
        return bold_type(first).unwrap_or_else(|| "?".into());
    }
    let lower = summary.to_ascii_lowercase();
    for marker in ["read/write", "read-only", "read only"] {
        if let Some(p) = lower.find(marker)
            && let Some(t) = summary.get(p..).and_then(bold_type)
        {
            return t;
        }
    }
    // "Returns a **Range** object …", also mid-sentence ("… this property returns an
    // **Application** …"): the type is the first bold run within a few words of "returns".
    if let Some(p) = lower.find("returns") {
        let rest = summary.get(p..).unwrap_or_default();
        let bold = rest.find("**").filter(|b| rest.get(..*b).is_some_and(|gap| gap.chars().count() <= RETURNS_WINDOW));
        return bold.and_then(|b| rest.get(b..)).and_then(bold_type).unwrap_or_else(|| "?".into());
    }
    if kind == Kind::Method { "none".into() } else { "?".into() }
}

/// One page reduced to a [`Member`], or the reason it was skipped.
pub fn parse_page(text: &str) -> Result<Member, String> {
    let (title, api, body) = front_matter(text)?;
    let (kind, title_name) = title_kind(&title)?;
    let (object, member) = names(title_name, &api)?;
    let (summary, map) = sections(&body);
    let syntax = map.get("syntax").map(Vec::as_slice).unwrap_or_default();
    let params = match map.get("parameters").map(|l| table_params(l)).transpose()?.flatten() {
        Some(p) => p,
        None => syntax_params(syntax, &member),
    };
    let returns = return_type(kind, &summary, &map);
    Ok(Member { object, member, kind, params, returns })
}

// ---------------------------------------------------------------------------------------------
// Catalog file
// ---------------------------------------------------------------------------------------------

/// Checks a note against Req 15.3 (length) and the file format (no tab/CR/LF/control).
fn check_note(who: &str, note: &str) -> Result<(), String> {
    let n = note.chars().count();
    if n > NOTE_MAX_CHARS {
        return Err(format!("{who}: note is {n} characters (at most {NOTE_MAX_CHARS})"));
    }
    if note.chars().any(char::is_control) {
        return Err(format!("{who}: note contains a tab, newline or other control character"));
    }
    Ok(())
}

/// Reads an existing catalog's rows, keyed for the merge.
fn parse_catalog(text: &str) -> Result<BTreeMap<Key, Row>, String> {
    let mut rows = BTreeMap::new();
    for (i, line) in text.lines().enumerate() {
        let n = i.saturating_add(1);
        if line.starts_with('#') || line.is_empty() {
            continue;
        }
        let f: Vec<&str> = line.split('\t').collect();
        let [object, member, kind, _params, _returns, status, note] = f.as_slice() else {
            return Err(format!("catalog line {n}: expected 7 tab-separated columns, found {}", f.len()));
        };
        let kind = Kind::parse(kind).ok_or_else(|| format!("catalog line {n}: unknown kind `{}`", short(kind)))?;
        let who = format!("{object}.{member} ({})", kind.as_str());
        let status = Status::parse(status)
            .ok_or_else(|| format!("{who}: unknown status `{}` (supported, planned, notPlanned or unreviewed)", short(status)))?;
        check_note(&who, note)?;
        let row = Row { object: (*object).to_string(), member: (*member).to_string(), kind, status, note: (*note).to_string() };
        if rows.insert(key(object, member, kind), row).is_some() {
            return Err(format!("{who}: listed twice in the catalog"));
        }
    }
    Ok(rows)
}

fn render_params(params: &[Param]) -> String {
    params.iter().map(|p| format!("{}:{}:{}", p.name, if p.optional { "opt" } else { "req" }, p.ty)).collect::<Vec<_>>().join(";")
}

/// The provenance header line.
fn header(commit: &str) -> String {
    format!("# generated by cargo xtask vba-catalog from MicrosoftDocs/VBA-Docs {commit} (CC-BY-4.0); names and signatures only")
}

/// Builds the catalog text from `(page name, page text or read error)` pairs (any order) and the
/// existing catalog, if any.
pub fn build(pages: &[Page], existing: Option<&str>, commit: &str) -> Result<(String, Summary), String> {
    let mut old = existing.map(parse_catalog).transpose()?.unwrap_or_default();
    let mut order: Vec<&Page> = pages.iter().collect();
    order.sort_by(|a, b| a.0.cmp(&b.0));
    let mut summary = Summary { pages_read: order.len(), ..Summary::default() };
    let mut members: BTreeMap<Key, (Member, String)> = BTreeMap::new();
    for (name, text) in order {
        let parsed = text.as_ref().map_err(Clone::clone).and_then(|t| parse_page(t));
        match parsed {
            Err(reason) => summary.skipped.push((name.clone(), reason)),
            Ok(m) => {
                let k = key(&m.object, &m.member, m.kind);
                if let Some((_, first)) = members.get(&k) {
                    summary.skipped.push((name.clone(), format!("duplicate of {first}")));
                } else {
                    members.insert(k, (m, name.clone()));
                }
            }
        }
    }
    let mut out = format!("{}\n{COLUMNS}\n", header(commit));
    for (k, (m, _)) in &members {
        let (status, note) = match old.remove(k) {
            Some(row) => {
                summary.kept = summary.kept.saturating_add(1);
                (row.status, row.note)
            }
            None => {
                summary.added = summary.added.saturating_add(1);
                (Status::Unreviewed, String::new())
            }
        };
        let (status, note) = if status == Status::Unreviewed && m.kind != Kind::Event && seeded(&m.object, &m.member) {
            summary.seeded = summary.seeded.saturating_add(1);
            (Status::Planned, if note.is_empty() { SEED_NOTE.to_string() } else { note })
        } else {
            (status, note)
        };
        if m.returns == "?" {
            summary.returns_unknown = summary.returns_unknown.saturating_add(1);
        }
        let row = [m.object.as_str(), &m.member, m.kind.as_str(), &render_params(&m.params), &m.returns, status.as_str(), &note];
        out.push_str(&row.join("\t"));
        out.push('\n');
    }
    summary.members = members.len();
    for (object, list) in TRANCHE_1 {
        for member in *list {
            if ![Kind::Property, Kind::Method].iter().any(|k| members.contains_key(&key(object, member, *k))) {
                summary.seed_unmatched.push(format!("{object}.{member}"));
            }
        }
    }
    summary.removed = old.values().map(|r| format!("{}.{} ({})", r.object, r.member, r.kind.as_str())).collect();
    Ok((out, summary))
}

// ---------------------------------------------------------------------------------------------
// Command
// ---------------------------------------------------------------------------------------------

/// Reads a file of at most `max` bytes.
fn read_capped(path: &Path, max: u64) -> Result<String, String> {
    let len = std::fs::metadata(path).map_err(|e| format!("{}: {e}", path.display()))?.len();
    if len > max {
        return Err(format!("larger than {max} bytes"));
    }
    let bytes = std::fs::read(path).map_err(|e| format!("{}: {e}", path.display()))?;
    String::from_utf8(bytes).map_err(|_| "not UTF-8".into())
}

/// Every `Excel.*.md` page under `docs/api`.
fn read_pages(docs: &Path) -> Result<Vec<Page>, String> {
    let api = docs.join("api");
    let rd = std::fs::read_dir(&api).map_err(|e| format!("{}: {e} (is --docs a VBA-Docs checkout?)", api.display()))?;
    let mut pages = Vec::new();
    for entry in rd.flatten() {
        let name = entry.file_name().to_string_lossy().into_owned();
        let lower = name.to_ascii_lowercase();
        if !(lower.starts_with("excel.") && lower.ends_with(".md")) || !entry.path().is_file() {
            continue;
        }
        if pages.len() >= PAGES_MAX {
            return Err(format!("{}: more than {PAGES_MAX} Excel pages", api.display()));
        }
        pages.push((name, read_capped(&entry.path(), PAGE_MAX_BYTES)));
    }
    if pages.is_empty() {
        return Err(format!("{}: no Excel.*.md pages", api.display()));
    }
    Ok(pages)
}

/// A commit id: 7 to 40 lower-case hex digits.
fn valid_commit(s: &str) -> bool {
    (7..=40).contains(&s.len()) && s.chars().all(|c| c.is_ascii_digit() || ('a'..='f').contains(&c))
}

fn docs_commit(docs: &Path) -> Result<String, String> {
    let out = std::process::Command::new("git")
        .arg("-C")
        .arg(docs)
        .args(["rev-parse", "HEAD"])
        .output()
        .map_err(|e| format!("git rev-parse in {}: {e} (pass --commit SHA)", docs.display()))?;
    let sha = String::from_utf8_lossy(&out.stdout).trim().to_string();
    if out.status.success() && valid_commit(&sha) { Ok(sha) } else { Err(format!("{}: no git commit found (pass --commit SHA)", docs.display())) }
}

fn print_summary(s: &Summary, out: &Path) {
    println!(
        "vba-catalog: {} pages read, {} members -> {} ({} kept, {} new, {} seeded planned, {} with return type `?`), {} pages skipped, {} removed",
        s.pages_read,
        s.members,
        out.display(),
        s.kept,
        s.added,
        s.seeded,
        s.returns_unknown,
        s.skipped.len(),
        s.removed.len()
    );
    let mut reasons: BTreeMap<&str, usize> = BTreeMap::new();
    for (_, r) in &s.skipped {
        let c = reasons.entry(r.split(" `").next().unwrap_or(r)).or_default();
        *c = c.saturating_add(1);
    }
    for (r, n) in &reasons {
        println!("  skipped {n:>5}  {r}");
    }
    for (page, reason) in &s.skipped {
        println!("  skip    {page}: {reason}");
    }
    for r in &s.removed {
        println!("  removed {r}");
    }
    for m in &s.seed_unmatched {
        println!("  seed    {m}: no such property or method in the docs");
    }
}

pub fn run(root: &Path, args: &[&str]) -> Result<(), String> {
    let mut docs: Option<PathBuf> = None;
    let mut out: PathBuf = root.join(DEFAULT_OUT);
    let mut commit: Option<String> = None;
    let mut it = args.iter();
    while let Some(a) = it.next() {
        match *a {
            "--docs" => docs = Some(PathBuf::from(it.next().ok_or("--docs needs a path")?)),
            "--out" => out = PathBuf::from(it.next().ok_or("--out needs a path")?),
            "--commit" => commit = Some((*it.next().ok_or("--commit needs a SHA")?).to_string()),
            other => return Err(format!("vba-catalog: unknown argument `{other}` ({USAGE})")),
        }
    }
    let docs = docs.ok_or_else(|| format!("vba-catalog: --docs is required ({USAGE})"))?;
    let commit = match commit {
        Some(c) if valid_commit(&c) => c,
        Some(c) => return Err(format!("--commit `{}`: expected 7 to 40 lower-case hex digits", short(&c))),
        None => docs_commit(&docs)?,
    };
    let pages = read_pages(&docs)?;
    let existing = if out.exists() { Some(read_capped(&out, CATALOG_MAX_BYTES).map_err(|e| format!("{}: {e}", out.display()))?) } else { None };
    let (text, summary) = build(&pages, existing.as_deref(), &commit)?;
    if let Some(dir) = out.parent()
        && !dir.as_os_str().is_empty()
    {
        std::fs::create_dir_all(dir).map_err(|e| format!("create {}: {e}", dir.display()))?;
    }
    std::fs::write(&out, text).map_err(|e| format!("write {}: {e}", out.display()))?;
    print_summary(&summary, &out);
    Ok(())
}

#[cfg(test)]
mod tests;
