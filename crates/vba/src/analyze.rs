//! The Compatibility_Analyzer (Requirement 13): what a VBA project's source uses, and how ready
//! it is for GridCraft. It reads the parsed source only; nothing here runs VBA.
//!
//! [`analyze`] parses every module with [`crate::parser::parse`] and reports:
//!
//! - **Procedures** (13.1) with kind, parameters (declared type, a type suffix's type, or
//!   `Variant`) and 1-based start and end lines, plus their [`Mark`]s (13.2–13.5).
//! - **Member references** (13.6, 13.7): every `object.member` access, counted per object and
//!   member and looked up in a [`MemberCatalog`].
//! - **Flags** (13.8) for constructs GridCraft will not support.
//! - **Unparsed modules** (13.11), **unreadable modules** (13.9) and the ordered **verdict** with
//!   its **reasons** (13.9, 13.10).
//! - The **Capability_Manifest** (Requirement 14): what the project could do.
//!
//! # Member references
//!
//! Each [`Expr::Member`] access `x.name` is one reference to `name` on the object `x`, where the
//! object is named after the *last link* of `x`, as written: an identifier (`ws`), `Me`, or a
//! member name (`Range` in `ws.Range("A1").Value`). Calls and indexes are looked through, so
//! `Worksheets("A").Range("B1").Value` counts `Range` on `Worksheets` and `Value` on `Range`.
//! Inside `With target`, a leading `.name` is a member of `target`, named by the same rule; with
//! no enclosing `With`, or a target with no name (`With New Collection`, `With (x)`), the object
//! is `None`. Names group case-insensitively and keep the first spelling seen. `rs!Field`
//! dictionary accesses are not member references (the name after `!` is a string key) and are
//! not counted. Bare global names (`Range("A1")`, `ActiveSheet`) are not resolved here.
//!
//! The report lists members by count (highest first), then object (`None` first), then member,
//! both compared lowercased.
//!
//! # Flags
//!
//! - `Declare`: every `Declare Sub|Function` statement.
//! - `CreateObject`, `GetObject`: calls, bare or qualified as `VBA.`, `Interaction.` or
//!   `VBA.Interaction.`, with or without parentheses, also as a `With`-relative `.CreateObject`.
//!   The ProgID is the `Class` argument (by name, else the first positional argument of
//!   `CreateObject` and the second of `GetObject`): [`ProgId::Literal`] for a string literal
//!   (parentheses looked through), else [`ProgId::NonLiteral`]. A `GetObject(path)` call with no
//!   class has no ProgID (`None`); a `CreateObject` with no class argument counts as non-literal.
//! - `Shell`: bare or `VBA.`/`Interaction.`/`VBA.Interaction.` qualified calls.
//! - `SendKeys`: the same, plus `Application.` and `Excel.Application.`.
//! - `ExecuteExcel4Macro`: bare, `Application.` or `Excel.Application.` qualified calls.
//! - ActiveX (a heuristic): a module-level `WithEvents` variable or an `Implements` whose type
//!   name is dotted (`Lib.Type`) and whose library is not `Excel`, `VBA`, `Office`, `MSForms`
//!   or `Scripting` (compared case-insensitively). Embedded OLE controls on sheets live outside
//!   the VBA source and are not seen.
//!
//! Each flag carries the module and the 1-based start and end lines of the call or statement.
//! Flags are listed in module order, then source order.
//!
//! # Unparsed modules and recovered procedures
//!
//! A module with at least one [`ParseDiag`](crate::parser::ParseDiag) (syntax, nesting too deep,
//! source too large) is unparsed. Its procedures that parsed without error are still listed; a
//! procedure the parser recovered from ([`Proc::is_recovered`]) is not listed, but the statements
//! that parsed before its error are still scanned for members and flags, so a flag is never
//! missed because a later line failed.
//!
//! # Unreadable and empty modules
//!
//! A module with empty source is never parsed and adds no procedures, members, flags or
//! capabilities. A module is *unreadable* (13.9) when its source was removed or couldn't be
//! read: it has [`DiagCode::PossibleStomping`] (which the reader gives only when the source is
//! empty or only `Attribute` lines, so those lines don't count as readable source), or its
//! source is empty and it has [`DiagCode::ModuleStreamMissing`],
//! [`DiagCode::ModuleOffsetOutOfRange`], [`DiagCode::ModuleTruncated`] or
//! [`DiagCode::ModuleCorrupt`]. Other diagnostics (a kind taken from MODULETYPE, an unsupported
//! code page, compiled code not checked, …) don't hide source. An unreadable module isn't parsed
//! either. Its code exists but GridCraft can't see it, so it is listed in
//! [`CompatibilityReport::unreadable`], gets a [`VerdictRule::UnreadableModule`] reason and makes
//! the manifest incomplete. A module with empty source and no such diagnostic is just empty and is
//! ignored.
//!
//! # Verdict and reasons
//!
//! The first matching rule wins (13.9): [`Verdict::NoMacros`] (no modules; the engine passes
//! [`VbaProject::default`] for a workbook without a project), [`Verdict::UsesUnsupportedFeatures`]
//! (a flag or a member [`MemberStatus::NotPlanned`]), [`Verdict::AnalysisOnly`] (an unparsed or
//! unreadable module, or a member [`MemberStatus::Unknown`]), else
//! [`Verdict::WouldRunInFuturePhase`]. Reasons follow the rules in that order: flags, members not
//! planned, unparsed modules, unreadable modules, unknown members. Every
//! matching condition of every rule gets a [`Reason`], in rule order, so a report that is
//! unsupported still shows what would also keep it analysis-only; [`Reason::rule`] tells them
//! apart. `WouldRunInFuturePhase` has one reason with no location. `NoMacros` has none.
//!
//! # Capability manifest
//!
//! [`CapabilityManifest`] always lists the nine [`Capability`] categories in [`Capability::ALL`]
//! order (14.1). The checked-in Trigger_Table (`analyze/triggers.rs`, its module docs give each
//! row's matching rule) says which constructs set which category (14.2). Each occurrence is a
//! [`CapSpan`] with the module and the 1-based line and column (in characters) of its first and
//! last character (14.3); one occurrence that two rows of a category match counts once. Spans
//! are in module order, then source order. The whole procedure is the occurrence of an event
//! handler or auto-exec procedure.
//!
//! Dynamic dispatch the analyzer can't resolve (14.4) is a [`DynamicUse`]: any `CallByName`;
//! `Application.Run` (or the bare `Run`) whose macro name isn't a literal; `Evaluate`, bare or on
//! any object, whose argument isn't a literal; `CreateObject`/`GetObject` with a
//! [`ProgId::NonLiteral`]. The manifest is then incomplete, as it is for every unparsed module
//! (14.5) and unreadable module; [`CapabilityManifest::incomplete_reasons`] says why.
//!
//! # Bounded cost
//!
//! Nothing recurses per syntax node: statements and expressions are walked with a heap stack, so a
//! 200,000-link member chain needs no more host stack than a short one. The tree is never cloned.
use std::collections::BTreeMap;

use self::triggers::{Callee, DYNAMIC, Rule, TRIGGERS, Trigger, VBA_QUALIFIERS};
use crate::ast::{Arg, CaseTest, Decl, Expr, FileOp, FileStmt, Literal, Module, Proc, ProcKind, Stmt, Visibility};
use crate::diag::DiagCode;
use crate::lexer::Span;
use crate::parser::parse;
use crate::project::{VbaModule, VbaProject};
use crate::project_stream::ModuleKind;

mod triggers;

/// What the analyzer found in a project.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CompatibilityReport {
    /// Procedures, in module order then source order.
    pub procedures: Vec<ProcInfo>,
    /// Member references, sorted as the module docs describe.
    pub members: Vec<MemberUse>,
    /// Flagged constructs, in module order then source order.
    pub flags: Vec<Flag>,
    /// The names of unparsed modules, in module order.
    pub unparsed: Vec<String>,
    /// The names of modules with no readable source (see the module docs), in module order.
    pub unreadable: Vec<String>,
    /// What the project could do.
    pub manifest: CapabilityManifest,
    /// The readiness verdict.
    pub verdict: Verdict,
    /// Why, one per matching condition (empty for [`Verdict::NoMacros`]).
    pub reasons: Vec<Reason>,
}

/// One procedure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProcInfo {
    /// The module's display name.
    pub module: String,
    /// The procedure name as written.
    pub name: String,
    /// `Sub`, `Function` or a `Property` kind.
    pub kind: ProcKind,
    /// The parameters, in order.
    pub params: Vec<ParamInfo>,
    /// The 1-based line of the header.
    pub start_line: u32,
    /// The 1-based line of the `End …`.
    pub end_line: u32,
    /// Its marks, in [`Mark`] order.
    pub marks: Vec<Mark>,
}

/// One parameter.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParamInfo {
    /// The name as written, without a type suffix.
    pub name: String,
    /// The declared type as written (`Long`, `Excel.Range`), the type a suffix declares
    /// (`s$` is `String`), or `Variant`; `()` is appended for an array parameter.
    pub ty: String,
}

/// What a procedure is to Excel.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum Mark {
    /// A non-`Private` `Sub` with no parameters in a standard module (13.2).
    EntryPoint,
    /// A non-`Private` `Function` in a standard module (13.3).
    UdfCandidate,
    /// `Workbook_<Event>` in `ThisWorkbook`, or `Worksheet_<Event>` in a sheet module (13.4).
    EventHandler,
    /// `Auto_Open`/`Auto_Close` in a standard module, or `Workbook_Open` in `ThisWorkbook` (13.5).
    AutoExec,
}

impl Mark {
    /// A stable name for JSON: `entryPoint`, `udfCandidate`, `eventHandler` or `autoExec`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::EntryPoint => "entryPoint",
            Self::UdfCandidate => "udfCandidate",
            Self::EventHandler => "eventHandler",
            Self::AutoExec => "autoExec",
        }
    }
}

/// One member, with how often it is referenced.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MemberUse {
    /// The object's name (see the module docs), as first written; `None` when it has none.
    pub object: Option<String>,
    /// The member name, as first written.
    pub member: String,
    /// The number of references (saturating).
    pub count: u32,
    /// The catalog status.
    pub status: MemberStatus,
}

/// A member's status in the Member_Catalog.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MemberStatus {
    /// GridCraft supports it.
    Supported,
    /// Support is planned.
    Planned,
    /// Support is not planned.
    NotPlanned,
    /// Not in the catalog (13.7), or listed as `unreviewed` (13.6).
    Unknown,
}

impl MemberStatus {
    /// A stable name for JSON: `supported`, `planned`, `notPlanned` or `unknown`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Supported => "supported",
            Self::Planned => "planned",
            Self::NotPlanned => "notPlanned",
            Self::Unknown => "unknown",
        }
    }
}

/// The Member_Catalog the analyzer looks members up in.
pub trait MemberCatalog {
    /// The status of `member` on `object` (both as written; compare case-insensitively), or
    /// `None` when the catalog doesn't list it.
    fn status(&self, object: Option<&str>, member: &str) -> Option<MemberStatus>;
}

/// A catalog that lists nothing, so every member is [`MemberStatus::Unknown`]. For tests; the
/// engine uses [`ExcelCatalog`](crate::catalog::ExcelCatalog).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EmptyCatalog;

impl MemberCatalog for EmptyCatalog {
    fn status(&self, _object: Option<&str>, _member: &str) -> Option<MemberStatus> {
        None
    }
}

/// One flagged construct.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Flag {
    /// What was flagged.
    pub kind: FlagKind,
    /// The module's display name.
    pub module: String,
    /// The 1-based first line.
    pub start_line: u32,
    /// The 1-based last line.
    pub end_line: u32,
    /// The ProgID of `CreateObject`/`GetObject`; `None` for other kinds and a class-less
    /// `GetObject(path)`.
    pub prog_id: Option<ProgId>,
}

/// The flagged constructs (13.8).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FlagKind {
    /// `Declare Sub|Function … Lib`.
    Declare,
    /// `CreateObject`.
    CreateObject,
    /// `GetObject`.
    GetObject,
    /// `Shell`.
    Shell,
    /// `SendKeys`.
    SendKeys,
    /// An ActiveX reference (the heuristic in the module docs).
    ActiveX,
    /// `ExecuteExcel4Macro`.
    ExecuteExcel4Macro,
}

impl FlagKind {
    /// A stable name for JSON: the construct's VBA spelling, or `ActiveX`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Declare => "Declare",
            Self::CreateObject => "CreateObject",
            Self::GetObject => "GetObject",
            Self::Shell => "Shell",
            Self::SendKeys => "SendKeys",
            Self::ActiveX => "ActiveX",
            Self::ExecuteExcel4Macro => "ExecuteExcel4Macro",
        }
    }
}

/// A `CreateObject`/`GetObject` ProgID.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ProgId {
    /// A string literal, unquoted.
    Literal(String),
    /// Anything else; JSON writes it as `"non-literal"`.
    NonLiteral,
}

impl ProgId {
    /// The text for JSON: the literal, or `non-literal`.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::Literal(text) => text,
            Self::NonLiteral => "non-literal",
        }
    }
}

/// The readiness verdict (13.9).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Verdict {
    /// No VBA project, or no modules.
    NoMacros,
    /// A flagged construct, or a member that is not planned.
    UsesUnsupportedFeatures,
    /// An unparsed or unreadable module, or a member not in the catalog.
    AnalysisOnly,
    /// Everything found is supported or planned.
    WouldRunInFuturePhase,
}

impl Verdict {
    /// A stable name for JSON: `noMacros`, `usesUnsupportedFeatures`, `analysisOnly` or
    /// `wouldRunInFuturePhase`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::NoMacros => "noMacros",
            Self::UsesUnsupportedFeatures => "usesUnsupportedFeatures",
            Self::AnalysisOnly => "analysisOnly",
            Self::WouldRunInFuturePhase => "wouldRunInFuturePhase",
        }
    }
}

/// The condition a [`Reason`] reports.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum VerdictRule {
    /// A flagged construct ([`Verdict::UsesUnsupportedFeatures`]).
    Flagged(FlagKind),
    /// A member with [`MemberStatus::NotPlanned`] ([`Verdict::UsesUnsupportedFeatures`]).
    MemberNotPlanned,
    /// An unparsed module ([`Verdict::AnalysisOnly`]).
    UnparsedModule,
    /// A module with no readable source ([`Verdict::AnalysisOnly`]).
    UnreadableModule,
    /// A member with [`MemberStatus::Unknown`] ([`Verdict::AnalysisOnly`]).
    MemberUnknown,
    /// Nothing unsupported or unknown ([`Verdict::WouldRunInFuturePhase`]).
    AllKnown,
}

impl VerdictRule {
    /// A stable name for JSON: `flagged` (the [`FlagKind`] goes alongside), `memberNotPlanned`,
    /// `unparsedModule`, `unreadableModule`, `memberUnknown` or `allKnown`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Flagged(_) => "flagged",
            Self::MemberNotPlanned => "memberNotPlanned",
            Self::UnparsedModule => "unparsedModule",
            Self::UnreadableModule => "unreadableModule",
            Self::MemberUnknown => "memberUnknown",
            Self::AllKnown => "allKnown",
        }
    }

    /// The verdict this condition leads to.
    #[must_use]
    pub const fn verdict(self) -> Verdict {
        match self {
            Self::Flagged(_) | Self::MemberNotPlanned => Verdict::UsesUnsupportedFeatures,
            Self::UnparsedModule | Self::UnreadableModule | Self::MemberUnknown => Verdict::AnalysisOnly,
            Self::AllKnown => Verdict::WouldRunInFuturePhase,
        }
    }
}

/// One reason for the verdict.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reason {
    /// The condition.
    pub rule: VerdictRule,
    /// The module, for flags, unparsed and unreadable modules.
    pub module: Option<String>,
    /// The 1-based line: a flag's first line, or an unparsed module's first diagnostic.
    pub line: Option<u32>,
    /// The member, as `object.member` or `member`, for member rules.
    pub member: Option<String>,
}

/// What a VBA project could do (Requirement 14): the Capability_Manifest.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CapabilityManifest {
    /// Exactly nine entries, one per [`Capability`], in [`Capability::ALL`] order.
    pub categories: Vec<CategoryEntry>,
    /// Dynamic dispatch the analyzer can't resolve, in module order then source order.
    pub dynamic: Vec<DynamicUse>,
    /// `true` when the categories may miss something: see [`IncompleteReason`].
    pub incomplete: bool,
    /// Why the manifest is incomplete; empty when it isn't.
    pub incomplete_reasons: Vec<IncompleteReason>,
}

/// One category of the manifest.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CategoryEntry {
    /// The category.
    pub category: Capability,
    /// Whether any construct sets it; `spans` is empty when it doesn't.
    pub present: bool,
    /// Each occurrence, in module order then source order.
    pub spans: Vec<CapSpan>,
}

/// Where a construct occurs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CapSpan {
    /// The module's display name.
    pub module: String,
    /// The 1-based first line.
    pub start_line: u32,
    /// The 1-based column (in characters) of the first character.
    pub start_col: u32,
    /// The 1-based last line.
    pub end_line: u32,
    /// The 1-based column (in characters) of the last character, inclusive.
    pub end_col: u32,
    /// The Trigger_Table construct that matched (`Open For Input`, `Shell`, `CallByName`).
    pub construct: &'static str,
}

/// The Capability_Manifest categories (14.1). See `docs/macros.md` for how each maps to the
/// engine's script sandbox.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum Capability {
    /// Reads files or folders.
    FileRead,
    /// Writes, deletes or creates files or folders.
    FileWrite,
    /// Uses the network.
    Network,
    /// Starts processes or calls native code.
    Process,
    /// Reads or writes the registry.
    Registry,
    /// Uses the clipboard.
    Clipboard,
    /// Shows dialogs.
    UiDialogs,
    /// Has event handlers (Requirement 13.4).
    EventHandlers,
    /// Runs on its own when the workbook opens or closes (Requirement 13.5).
    AutoExec,
}

impl Capability {
    /// Every category, in manifest order.
    pub const ALL: [Self; 9] = [
        Self::FileRead,
        Self::FileWrite,
        Self::Network,
        Self::Process,
        Self::Registry,
        Self::Clipboard,
        Self::UiDialogs,
        Self::EventHandlers,
        Self::AutoExec,
    ];

    /// A stable name for JSON: `fileRead`, `fileWrite`, `network`, `process`, `registry`,
    /// `clipboard`, `uiDialogs`, `eventHandlers` or `autoExec`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::FileRead => "fileRead",
            Self::FileWrite => "fileWrite",
            Self::Network => "network",
            Self::Process => "process",
            Self::Registry => "registry",
            Self::Clipboard => "clipboard",
            Self::UiDialogs => "uiDialogs",
            Self::EventHandlers => "eventHandlers",
            Self::AutoExec => "autoExec",
        }
    }

    /// The position in [`Self::ALL`].
    const fn index(self) -> usize {
        self as usize
    }
}

/// One use of dynamic dispatch: "dynamic: unknown capabilities" (14.4).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DynamicUse {
    /// The module's display name.
    pub module: String,
    /// Where it is.
    pub span: CapSpan,
    /// `CallByName`, `Application.Run`, `Evaluate`, `CreateObject` or `GetObject`.
    pub construct: &'static str,
}

/// Why a manifest is incomplete.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum IncompleteReason {
    /// At least one [`DynamicUse`] (14.4); listed once, first.
    Dynamic,
    /// The module is unparsed (14.5); after `Dynamic`, in module order.
    UnparsedModule(String),
    /// The module has no readable source (Requirement 13.9); after the unparsed ones, in module
    /// order.
    UnreadableModule(String),
}

impl IncompleteReason {
    /// A stable name for JSON: `dynamic`, `unparsedModule` or `unreadableModule`.
    #[must_use]
    pub const fn as_str(&self) -> &'static str {
        match self {
            Self::Dynamic => "dynamic",
            Self::UnparsedModule(_) => "unparsedModule",
            Self::UnreadableModule(_) => "unreadableModule",
        }
    }

    /// The text to show: `dynamic: unknown capabilities`, `unparsed module: capabilities
    /// unknown` or `unreadable module: capabilities unknown`.
    #[must_use]
    pub const fn description(&self) -> &'static str {
        match self {
            Self::Dynamic => "dynamic: unknown capabilities",
            Self::UnparsedModule(_) => "unparsed module: capabilities unknown",
            Self::UnreadableModule(_) => "unreadable module: capabilities unknown",
        }
    }

    /// The module it names, if any.
    #[must_use]
    pub fn module(&self) -> Option<&str> {
        match self {
            Self::Dynamic => None,
            Self::UnparsedModule(name) | Self::UnreadableModule(name) => Some(name),
        }
    }
}

/// A module with no readable source (13.9): possible stomping (the reader only says so when the
/// source is empty or only `Attribute` lines), or empty source with a reader diagnostic saying
/// the source couldn't be read. See the module docs.
fn is_unreadable(module: &VbaModule) -> bool {
    module.diagnostics.iter().any(|d| d.code == DiagCode::PossibleStomping)
        || (module.source.is_empty() && module.diagnostics.iter().any(|d| hides_source(d.code)))
}

/// Whether `code` means a module's source couldn't be read or was removed. The match is
/// exhaustive on purpose: a new code must decide whether it makes an empty module unreadable.
const fn hides_source(code: DiagCode) -> bool {
    match code {
        DiagCode::PossibleStomping
        | DiagCode::ModuleStreamMissing
        | DiagCode::ModuleOffsetOutOfRange
        | DiagCode::ModuleTruncated
        | DiagCode::ModuleCorrupt => true,
        DiagCode::DirMissing
        | DiagCode::DirCorrupt
        | DiagCode::ProjectStreamMissing
        | DiagCode::ProjectStreamTruncated
        | DiagCode::KindFromModuleType
        | DiagCode::UnsupportedCodePage(_)
        | DiagCode::ModuleLimit
        | DiagCode::SourceLimit
        | DiagCode::CompiledCodeNotChecked
        | DiagCode::Signed => false,
    }
}

/// Analyses a project's source. Never fails and never panics; see the module docs.
#[must_use]
pub fn analyze(project: &VbaProject, catalog: &dyn MemberCatalog) -> CompatibilityReport {
    let mut acc = Acc::default();
    let mut unparsed: Vec<(String, u32)> = Vec::new();
    let mut unreadable: Vec<String> = Vec::new();
    for (index, module) in project.modules.iter().enumerate() {
        if is_unreadable(module) {
            unreadable.push(module.name.clone());
            continue;
        }
        if module.source.is_empty() {
            continue;
        }
        let (tree, diags) = parse(&module.source);
        if let Some(first) = diags.iter().map(|d| d.span.line).min() {
            unparsed.push((module.name.clone(), first));
        }
        if let Some(tree) = tree {
            let lines = LineIndex::new(&module.source);
            acc.module(index, module, &tree, &lines);
        }
    }
    let members = acc.members(catalog);
    let manifest = acc.manifest(&unparsed, &unreadable);
    let flags = acc.flags;
    let mut reasons = Vec::new();
    for flag in &flags {
        reasons.push(Reason { rule: VerdictRule::Flagged(flag.kind), module: Some(flag.module.clone()), line: Some(flag.start_line), member: None });
    }
    push_member_reasons(&mut reasons, &members, MemberStatus::NotPlanned, VerdictRule::MemberNotPlanned);
    for (name, line) in &unparsed {
        reasons.push(Reason { rule: VerdictRule::UnparsedModule, module: Some(name.clone()), line: Some(*line), member: None });
    }
    for name in &unreadable {
        reasons.push(Reason { rule: VerdictRule::UnreadableModule, module: Some(name.clone()), line: None, member: None });
    }
    push_member_reasons(&mut reasons, &members, MemberStatus::Unknown, VerdictRule::MemberUnknown);
    let verdict = if project.modules.is_empty() {
        reasons.clear();
        Verdict::NoMacros
    } else if let Some(first) = reasons.first() {
        first.rule.verdict()
    } else {
        reasons.push(Reason { rule: VerdictRule::AllKnown, module: None, line: None, member: None });
        Verdict::WouldRunInFuturePhase
    };
    CompatibilityReport {
        procedures: acc.procedures,
        members,
        flags,
        unparsed: unparsed.into_iter().map(|(name, _)| name).collect(),
        unreadable,
        manifest,
        verdict,
        reasons,
    }
}

/// Adds one reason per member with `status`, in report order.
fn push_member_reasons(reasons: &mut Vec<Reason>, members: &[MemberUse], status: MemberStatus, rule: VerdictRule) {
    for m in members.iter().filter(|m| m.status == status) {
        let member = match &m.object {
            Some(object) => format!("{object}.{}", m.member),
            None => m.member.clone(),
        };
        reasons.push(Reason { rule, module: None, line: None, member: Some(member) });
    }
}

/// The byte offset where each 1-based line starts, with the lexer's terminators (CR, LF, CRLF,
/// U+2028, U+2029).
///
/// Columns count characters, like the lexer's. `wide` lists each multi-byte character's offset
/// with the extra bytes (beyond one per character) up to and including it, so a column is two
/// binary searches whatever the line's length.
struct LineIndex<'a> {
    source: &'a str,
    starts: Vec<usize>,
    wide: Vec<(usize, usize)>,
}

impl<'a> LineIndex<'a> {
    fn new(source: &'a str) -> Self {
        let mut starts = vec![0];
        let mut wide = Vec::new();
        let mut extra = 0usize;
        let mut chars = source.char_indices().peekable();
        while let Some((pos, c)) = chars.next() {
            if c.len_utf8() > 1 {
                extra = extra.saturating_add(c.len_utf8().saturating_sub(1));
                wide.push((pos, extra));
            }
            if c == '\r' && chars.peek().is_some_and(|&(_, next)| next == '\n') {
                chars.next();
                starts.push(pos.saturating_add(2));
            } else if matches!(c, '\r' | '\n' | '\u{2028}' | '\u{2029}') {
                starts.push(pos.saturating_add(c.len_utf8()));
            }
        }
        Self { source, starts, wide }
    }

    /// The extra bytes of the multi-byte characters before byte `pos`.
    fn extra_before(&self, pos: usize) -> usize {
        match self.wide.partition_point(|&(at, _)| at < pos).checked_sub(1) {
            Some(i) => self.wide.get(i).map_or(0, |&(_, extra)| extra),
            None => 0,
        }
    }

    /// The 1-based line and column of the character starting at byte `pos`.
    fn line_col(&self, pos: usize) -> (u32, u32) {
        let line = self.line_of(pos);
        let start = usize::try_from(line).ok().and_then(|l| self.starts.get(l.saturating_sub(1))).copied().unwrap_or(0).min(pos);
        let chars = pos.saturating_sub(start).saturating_sub(self.extra_before(pos).saturating_sub(self.extra_before(start)));
        (line, u32::try_from(chars).unwrap_or(u32::MAX).saturating_add(1))
    }

    /// The 1-based line and column of `span`'s last character; its start for an empty span.
    fn end(&self, span: Span) -> (u32, u32) {
        let mut last = span.end.min(self.source.len());
        if last <= span.start {
            return (span.line, span.col);
        }
        last = last.saturating_sub(1);
        while last > span.start && !self.source.is_char_boundary(last) {
            last = last.saturating_sub(1);
        }
        let (line, col) = self.line_col(last);
        if line < span.line { (span.line, span.col) } else { (line, col) }
    }

    /// The 1-based line holding byte `pos`.
    fn line_of(&self, pos: usize) -> u32 {
        u32::try_from(self.starts.partition_point(|&start| start <= pos)).unwrap_or(u32::MAX).max(1)
    }

    /// The 1-based first and last lines of `span`.
    fn lines(&self, span: Span) -> (u32, u32) {
        let end = if span.end > span.start { self.line_of(span.end.saturating_sub(1)) } else { span.line };
        (span.line, end.max(span.line))
    }
}

/// A member's grouping key: the lowercased object and member.
type MemberKey = (Option<String>, String);

/// What the walk collects across modules.
#[derive(Default)]
struct Acc {
    procedures: Vec<ProcInfo>,
    flags: Vec<Flag>,
    /// Lowercased key → (first spelling of object, first spelling of member, count).
    members: BTreeMap<MemberKey, (Option<String>, String, u32)>,
    /// Trigger_Table hits, with the module's index.
    caps: Vec<(usize, (Capability, CapSpan))>,
    /// Dynamic dispatch, with the module's index.
    dynamic: Vec<(usize, DynamicUse)>,
}

/// A unit of work for the heap-stack walk, with the innermost `With` target.
enum Work<'a> {
    Stmt(&'a Stmt, Option<&'a Expr>),
    Expr(&'a Expr, Option<&'a Expr>),
    /// The target of a call: walked like an expression, but its own name was already matched
    /// against the Trigger_Table with the whole call's span.
    Callee(&'a Expr, Option<&'a Expr>),
}

/// One module's context during the walk.
struct ModCtx<'a> {
    index: usize,
    module: &'a VbaModule,
    lines: &'a LineIndex<'a>,
}

impl ModCtx<'_> {
    /// `span` as a manifest span for `construct`.
    fn cap_span(&self, span: Span, construct: &'static str) -> CapSpan {
        let (end_line, end_col) = self.lines.end(span);
        CapSpan { module: self.module.name.clone(), start_line: span.line, start_col: span.col, end_line, end_col, construct }
    }
}

/// Sorts `(module index, …)` entries into module order then source order (stable).
fn sort_by_source<T>(items: &mut [(usize, T)], span: impl Fn(&T) -> &CapSpan) {
    items.sort_by_key(|(index, item)| {
        let s = span(item);
        (*index, s.start_line, s.start_col, s.end_line, s.end_col)
    });
}

impl Acc {
    fn module(&mut self, index: usize, module: &VbaModule, tree: &Module, lines: &LineIndex<'_>) {
        let ctx = ModCtx { index, module, lines };
        let first_flag = self.flags.len();
        let mut stack: Vec<Work<'_>> = Vec::new();
        for decl in &tree.decls {
            self.decl(&ctx, decl, &mut stack);
        }
        for proc in &tree.procs {
            let proc_marks = marks(module, proc);
            for t in TRIGGERS {
                if matches!(t.rule, Rule::Mark(m) if proc_marks.contains(&m)) {
                    self.cap(&ctx, t, proc.span);
                }
            }
            if !proc.is_recovered() {
                self.procedures.push(proc_info(module, proc, lines, proc_marks));
            }
            for param in &proc.params {
                stack.extend(param.default.iter().map(|e| Work::Expr(e, None)));
            }
            stack.extend(proc.body.iter().rev().map(|s| Work::Stmt(s, None)));
            self.drain(&ctx, &mut stack);
        }
        // The walk visits a statement before the calls inside it, so restore source order (stable).
        if let Some(flags) = self.flags.get_mut(first_flag..) {
            flags.sort_by_key(|f| f.start_line);
        }
    }

    fn decl<'a>(&mut self, ctx: &ModCtx<'_>, decl: &'a Decl, stack: &mut Vec<Work<'a>>) {
        match decl {
            Decl::Declare(d) => {
                self.flag(ctx, FlagKind::Declare, d.span, None);
                self.caps_where(ctx, d.span, |rule| matches!(rule, Rule::Declare));
            }
            Decl::Var(var) => {
                for v in &var.vars {
                    if v.with_events && v.ty.as_ref().is_some_and(|ty| is_activex_type(&ty.name)) {
                        self.flag(ctx, FlagKind::ActiveX, v.span, None);
                    }
                    if let Some(ty) = &v.ty {
                        self.type_caps(ctx, &ty.name, v.span);
                    }
                    for b in v.bounds.iter().flatten() {
                        stack.extend(b.lower.iter().chain(std::iter::once(&b.upper)).map(|e| Work::Expr(e, None)));
                    }
                }
            }
            Decl::Implements { name, span } if is_activex_type(name) => self.flag(ctx, FlagKind::ActiveX, *span, None),
            Decl::Const(c) => stack.extend(c.consts.iter().map(|c| Work::Expr(&c.value, None))),
            Decl::Implements { .. } | Decl::Type(_) | Decl::Enum(_) | Decl::Event(_) | Decl::Def(_) | Decl::Directive { .. } => {}
        }
        self.drain(ctx, stack);
    }

    /// Processes the stack until it is empty.
    fn drain<'a>(&mut self, ctx: &ModCtx<'_>, stack: &mut Vec<Work<'a>>) {
        while let Some(work) = stack.pop() {
            match work {
                Work::Stmt(stmt, with) => self.stmt(ctx, stmt, with, stack),
                Work::Expr(expr, with) => self.expr(ctx, expr, with, stack, false),
                Work::Callee(expr, with) => self.expr(ctx, expr, with, stack, true),
            }
        }
    }

    /// Records a hit of row `t` at `span`.
    fn cap(&mut self, ctx: &ModCtx<'_>, t: &Trigger, span: Span) {
        self.caps.push((ctx.index, (t.category, ctx.cap_span(span, t.construct))));
    }

    /// Records a hit at `span` for every row whose rule passes `test`.
    fn caps_where(&mut self, ctx: &ModCtx<'_>, span: Span, test: impl Fn(&Rule) -> bool) {
        for t in TRIGGERS.iter().filter(|t| test(&t.rule)) {
            self.cap(ctx, t, span);
        }
    }

    /// The [`Rule::Type`] rows for a variable or `New` of type `name`.
    fn type_caps(&mut self, ctx: &ModCtx<'_>, name: &str, span: Span) {
        self.caps_where(ctx, span, |rule| matches!(rule, Rule::Type(ty) if name.eq_ignore_ascii_case(ty)));
    }

    /// The [`Rule::Call`] rows for a call of (or bare use of) `target` spanning `span`.
    fn call_caps(&mut self, ctx: &ModCtx<'_>, target: &Expr, span: Span, with: Option<&Expr>) {
        self.caps_where(ctx, span, |rule| matches!(rule, Rule::Call(callee) if callee_matches(callee, target, with)));
    }

    /// The file-statement rows for `f`.
    fn file_caps(&mut self, ctx: &ModCtx<'_>, f: &FileStmt) {
        let mode = if f.op == FileOp::Open { Some(f.mode.first().map_or("Random", String::as_str)) } else { None };
        // `Access` words run up to `Lock` or `Shared`.
        let access: Vec<&str> = f
            .mode
            .iter()
            .skip_while(|w| !w.eq_ignore_ascii_case("Access"))
            .skip(1)
            .take_while(|w| !w.eq_ignore_ascii_case("Lock") && !w.eq_ignore_ascii_case("Shared"))
            .map(String::as_str)
            .collect();
        self.caps_where(ctx, f.span, |rule| match rule {
            Rule::File(op) => *op == f.op,
            Rule::OpenMode(word) => mode.is_some_and(|m| m.eq_ignore_ascii_case(word)),
            Rule::OpenAccess(word) => f.op == FileOp::Open && access.iter().any(|a| a.eq_ignore_ascii_case(word)),
            _ => false,
        });
    }

    fn dynamic_use(&mut self, ctx: &ModCtx<'_>, construct: &'static str, span: Span) {
        let span = ctx.cap_span(span, construct);
        self.dynamic.push((ctx.index, DynamicUse { module: ctx.module.name.clone(), span, construct }));
    }

    /// The manifest from the hits, with the incomplete reasons.
    fn manifest(&mut self, unparsed: &[(String, u32)], unreadable: &[String]) -> CapabilityManifest {
        let mut caps = std::mem::take(&mut self.caps);
        sort_by_source(&mut caps, |(_, s)| s);
        let mut by_category: [Vec<CapSpan>; 9] = Default::default();
        for (_, (category, span)) in caps {
            if let Some(list) = by_category.get_mut(category.index()) {
                // Two rows of one category on the same occurrence count once.
                let same = |s: &CapSpan| {
                    s.module == span.module
                        && (s.start_line, s.start_col, s.end_line, s.end_col) == (span.start_line, span.start_col, span.end_line, span.end_col)
                };
                if !list.last().is_some_and(same) {
                    list.push(span);
                }
            }
        }
        let categories =
            Capability::ALL.iter().zip(by_category).map(|(&category, spans)| CategoryEntry { category, present: !spans.is_empty(), spans }).collect();
        let mut dynamic = std::mem::take(&mut self.dynamic);
        sort_by_source(&mut dynamic, |d| &d.span);
        let dynamic: Vec<DynamicUse> = dynamic.into_iter().map(|(_, d)| d).collect();
        let mut incomplete_reasons = Vec::new();
        if !dynamic.is_empty() {
            incomplete_reasons.push(IncompleteReason::Dynamic);
        }
        incomplete_reasons.extend(unparsed.iter().map(|(name, _)| IncompleteReason::UnparsedModule(name.clone())));
        incomplete_reasons.extend(unreadable.iter().cloned().map(IncompleteReason::UnreadableModule));
        CapabilityManifest { categories, dynamic, incomplete: !incomplete_reasons.is_empty(), incomplete_reasons }
    }

    fn stmt<'a>(&mut self, ctx: &ModCtx<'_>, stmt: &'a Stmt, with: Option<&'a Expr>, stack: &mut Vec<Work<'a>>) {
        let mut exprs: Vec<&'a Expr> = Vec::new();
        let mut bodies: Vec<&'a [Stmt]> = Vec::new();
        let mut callee: Option<&'a Expr> = None;
        match stmt {
            Stmt::Assign { target, value, .. } => exprs.extend([target, value]),
            Stmt::Call { target, args, span, .. } => {
                self.call(ctx, target, args, *span, with);
                callee = Some(target);
                exprs.extend(args.iter().filter_map(|a| a.value.as_ref()));
            }
            Stmt::Var(var) => {
                for v in &var.vars {
                    if let Some(ty) = &v.ty {
                        self.type_caps(ctx, &ty.name, v.span);
                    }
                }
                for b in var.vars.iter().flat_map(|v| v.bounds.iter().flatten()) {
                    exprs.extend(b.lower.iter().chain(std::iter::once(&b.upper)));
                }
            }
            Stmt::Const(c) => exprs.extend(c.consts.iter().map(|c| &c.value)),
            Stmt::If(s) => {
                for branch in &s.branches {
                    exprs.push(&branch.cond);
                    bodies.push(&branch.body);
                }
                bodies.extend(s.else_body.as_deref());
            }
            Stmt::Select(s) => {
                exprs.push(&s.expr);
                for case in &s.cases {
                    for test in &case.tests {
                        match test {
                            CaseTest::Value(e) | CaseTest::Is { value: e, .. } => exprs.push(e),
                            CaseTest::Range { from, to } => exprs.extend([from, to]),
                        }
                    }
                    bodies.push(&case.body);
                }
                bodies.extend(s.else_body.as_deref());
            }
            Stmt::For(s) => {
                exprs.extend([&s.var, &s.from, &s.to]);
                exprs.extend(s.step.iter().chain(s.next_var.iter()));
                bodies.push(&s.body);
            }
            Stmt::ForEach(s) => {
                exprs.extend([&s.var, &s.collection]);
                exprs.extend(s.next_var.iter());
                bodies.push(&s.body);
            }
            Stmt::Do(s) => {
                exprs.extend(s.pre.iter().chain(s.post.iter()).map(|c| &c.cond));
                bodies.push(&s.body);
            }
            Stmt::While { cond, body, .. } => {
                exprs.push(cond);
                bodies.push(body);
            }
            Stmt::With { object, body, .. } => {
                // The body's `.x` refers to `object`; `object` itself is in the outer `With`.
                stack.extend(body.iter().rev().map(|s| Work::Stmt(s, Some(object))));
                exprs.push(object);
            }
            Stmt::OnJump { expr, .. } => exprs.push(expr),
            Stmt::ReDim { vars, .. } => {
                for v in vars {
                    exprs.push(&v.target);
                    for b in &v.bounds {
                        exprs.extend(b.lower.iter().chain(std::iter::once(&b.upper)));
                    }
                }
            }
            Stmt::Erase { targets, .. } => exprs.extend(targets),
            Stmt::RaiseEvent { args, .. } => exprs.extend(args.iter().filter_map(|a| a.value.as_ref())),
            Stmt::DebugPrint { items, .. } => exprs.extend(items.iter().filter_map(|i| i.expr.as_ref())),
            Stmt::File(f) => {
                self.file_caps(ctx, f);
                exprs.extend(f.exprs());
            }
            Stmt::Exit { .. }
            | Stmt::OnError { .. }
            | Stmt::Resume { .. }
            | Stmt::GoTo { .. }
            | Stmt::GoSub { .. }
            | Stmt::Return { .. }
            | Stmt::Label { .. }
            | Stmt::End { .. }
            | Stmt::Stop { .. }
            | Stmt::Attribute(_)
            | Stmt::Directive { .. }
            | Stmt::Unknown { .. } => {}
        }
        // Pushed in reverse so the stack pops in source order.
        for body in bodies.into_iter().rev() {
            stack.extend(body.iter().rev().map(|s| Work::Stmt(s, with)));
        }
        stack.extend(exprs.into_iter().rev().map(|e| Work::Expr(e, with)));
        stack.extend(callee.map(|e| Work::Callee(e, with)));
    }

    /// Walks `expr`. A `callee` was already matched as a call target, so its own name isn't
    /// matched again here.
    fn expr<'a>(&mut self, ctx: &ModCtx<'_>, expr: &'a Expr, with: Option<&'a Expr>, stack: &mut Vec<Work<'a>>, callee: bool) {
        match expr {
            Expr::Member { object, bang, name, span, .. } => {
                if !*bang {
                    let owner = match object {
                        Some(object) => object_name(object),
                        None => with.and_then(object_name),
                    };
                    self.member(owner, name);
                    if !callee {
                        self.call_caps(ctx, expr, *span, with);
                    }
                }
                stack.extend(object.as_deref().map(|o| Work::Expr(o, with)));
            }
            Expr::Ident { span, .. } => {
                if !callee {
                    self.call_caps(ctx, expr, *span, with);
                }
            }
            Expr::New { class, span } => self.type_caps(ctx, class, *span),
            Expr::CallOrIndex { target, args, span } => {
                self.call(ctx, target, args, *span, with);
                stack.extend(args.iter().rev().filter_map(|a| a.value.as_ref()).map(|e| Work::Expr(e, with)));
                stack.push(Work::Callee(target, with));
            }
            Expr::Paren { inner: child, .. }
            | Expr::Unary { operand: child, .. }
            | Expr::TypeOfIs { object: child, .. }
            | Expr::AddressOf { target: child, .. } => stack.push(Work::Expr(child, with)),
            Expr::Binary { left, right, .. } => {
                stack.push(Work::Expr(right, with));
                stack.push(Work::Expr(left, with));
            }
            Expr::Literal { .. } | Expr::Me { .. } => {}
        }
    }

    /// Matches a call of `target` with `args` against the Trigger_Table, the dynamic-dispatch
    /// rows and the flagged functions.
    fn call(&mut self, ctx: &ModCtx<'_>, target: &Expr, args: &[Arg], span: Span, with: Option<&Expr>) {
        self.call_caps(ctx, target, span, with);
        for d in DYNAMIC.iter().filter(|d| callee_matches(&d.callee, target, with)) {
            let resolved = d.arg.is_some_and(|(name, position)| arg(args, name, position).is_some_and(is_literal));
            if !resolved {
                self.dynamic_use(ctx, d.construct, span);
            }
        }
        let Some(kind) = flagged_callee(target, with) else { return };
        let prog_id = match kind {
            FlagKind::CreateObject => Some(prog_id(args, 0).unwrap_or(ProgId::NonLiteral)),
            FlagKind::GetObject => prog_id(args, 1),
            _ => None,
        };
        match &prog_id {
            Some(ProgId::Literal(text)) => self.caps_where(ctx, span, |rule| matches!(rule, Rule::ProgId(p) if prog_id_matches(text, p))),
            Some(ProgId::NonLiteral) => self.dynamic_use(ctx, kind.as_str(), span),
            None => {}
        }
        self.flag(ctx, kind, span, prog_id);
    }

    fn flag(&mut self, ctx: &ModCtx<'_>, kind: FlagKind, span: Span, prog_id: Option<ProgId>) {
        let (start_line, end_line) = ctx.lines.lines(span);
        self.flags.push(Flag { kind, module: ctx.module.name.clone(), start_line, end_line, prog_id });
    }

    fn member(&mut self, object: Option<&str>, member: &str) {
        let key = (object.map(str::to_lowercase), member.to_lowercase());
        let entry = self.members.entry(key).or_insert_with(|| (object.map(str::to_owned), member.to_owned(), 0));
        entry.2 = entry.2.saturating_add(1);
    }

    /// The members with their catalog status, sorted by count, then object, then member.
    fn members(&mut self, catalog: &dyn MemberCatalog) -> Vec<MemberUse> {
        let mut list: Vec<(MemberKey, MemberUse)> = std::mem::take(&mut self.members)
            .into_iter()
            .map(|(key, (object, member, count))| {
                let status = catalog.status(object.as_deref(), &member).unwrap_or(MemberStatus::Unknown);
                (key, MemberUse { object, member, count, status })
            })
            .collect();
        // The map already orders by key; a stable sort by count keeps that order for ties.
        list.sort_by_key(|m| std::cmp::Reverse(m.1.count));
        list.into_iter().map(|(_, m)| m).collect()
    }
}

/// The procedure's entry for the report.
fn proc_info(module: &VbaModule, proc: &Proc, lines: &LineIndex<'_>, marks: Vec<Mark>) -> ProcInfo {
    let (start_line, end_line) = lines.lines(proc.span);
    let params = proc
        .params
        .iter()
        .map(|p| {
            let base = p.ty.as_ref().map(|t| t.name.clone()).or_else(|| p.suffix.and_then(suffix_type).map(str::to_owned));
            let base = base.unwrap_or_else(|| "Variant".to_owned());
            ParamInfo { name: p.name.clone(), ty: if p.is_array { format!("{base}()") } else { base } }
        })
        .collect();
    ProcInfo { module: module.name.clone(), name: proc.name.clone(), kind: proc.kind, params, start_line, end_line, marks }
}

/// The type a type suffix declares (§5.2.3.1.1 type-suffix).
const fn suffix_type(suffix: char) -> Option<&'static str> {
    match suffix {
        '%' => Some("Integer"),
        '&' => Some("Long"),
        '^' => Some("LongLong"),
        '!' => Some("Single"),
        '#' => Some("Double"),
        '@' => Some("Currency"),
        '$' => Some("String"),
        _ => None,
    }
}

/// The marks of 13.2–13.5, in [`Mark`] order.
fn marks(module: &VbaModule, proc: &Proc) -> Vec<Mark> {
    let public = proc.visibility != Some(Visibility::Private);
    let standard = module.kind == ModuleKind::Standard;
    let this_workbook = module.kind == ModuleKind::Document && module.code_name.eq_ignore_ascii_case("ThisWorkbook");
    let sheet = module.kind == ModuleKind::Document && !this_workbook;
    let mut marks = Vec::new();
    if standard && public && proc.kind == ProcKind::Sub && proc.params.is_empty() {
        marks.push(Mark::EntryPoint);
    }
    if standard && public && proc.kind == ProcKind::Function {
        marks.push(Mark::UdfCandidate);
    }
    if (this_workbook && has_event_prefix(&proc.name, "Workbook_")) || (sheet && has_event_prefix(&proc.name, "Worksheet_")) {
        marks.push(Mark::EventHandler);
    }
    let auto = (standard && (proc.name.eq_ignore_ascii_case("Auto_Open") || proc.name.eq_ignore_ascii_case("Auto_Close")))
        || (this_workbook && proc.name.eq_ignore_ascii_case("Workbook_Open"));
    if auto {
        marks.push(Mark::AutoExec);
    }
    marks
}

/// `name` is `prefix` (ASCII case-insensitive) followed by at least one character.
fn has_event_prefix(name: &str, prefix: &str) -> bool {
    name.len() > prefix.len() && name.get(..prefix.len()).is_some_and(|head| head.eq_ignore_ascii_case(prefix))
}

/// The object name of `expr` (see the module docs): its last link, looking through calls and
/// indexes. A loop, not recursion, so long `f(1)(1)…` chains are fine.
fn object_name(expr: &Expr) -> Option<&str> {
    let mut cur = expr;
    loop {
        match cur {
            Expr::CallOrIndex { target, .. } => cur = target,
            Expr::Ident { name, .. } => return Some(name),
            Expr::Member { name, .. } => return Some(name),
            Expr::Me { .. } => return Some("Me"),
            _ => return None,
        }
    }
}

/// Which flagged function `target` calls, if any.
fn flagged_callee(target: &Expr, with: Option<&Expr>) -> Option<FlagKind> {
    let (name, qualifier) = match target {
        Expr::Ident { name, .. } => (name.as_str(), None),
        Expr::Member { object: Some(object), bang: false, name, .. } => (name.as_str(), Some(&**object)),
        Expr::Member { object: None, bang: false, name, .. } => (name.as_str(), Some(with?)),
        _ => return None,
    };
    let kind = [
        ("CreateObject", FlagKind::CreateObject),
        ("GetObject", FlagKind::GetObject),
        ("Shell", FlagKind::Shell),
        ("SendKeys", FlagKind::SendKeys),
        ("ExecuteExcel4Macro", FlagKind::ExecuteExcel4Macro),
    ]
    .into_iter()
    .find(|(n, _)| name.eq_ignore_ascii_case(n))
    .map(|(_, kind)| kind)?;
    let Some(qualifier) = qualifier else { return Some(kind) };
    let vba = is_path(qualifier, &["VBA"]) || is_path(qualifier, &["Interaction"]) || is_path(qualifier, &["VBA", "Interaction"]);
    let application = is_path(qualifier, &["Application"]) || is_path(qualifier, &["Excel", "Application"]);
    let allowed = match kind {
        FlagKind::CreateObject | FlagKind::GetObject | FlagKind::Shell => vba,
        FlagKind::SendKeys => vba || application,
        FlagKind::ExecuteExcel4Macro => application,
        FlagKind::Declare | FlagKind::ActiveX => false,
    };
    allowed.then_some(kind)
}

/// `expr` is exactly the dotted name `path` (ASCII case-insensitive), e.g. `VBA.Interaction`.
fn is_path(expr: &Expr, path: &[&str]) -> bool {
    let mut cur = expr;
    for (i, part) in path.iter().enumerate().rev() {
        match cur {
            Expr::Member { object: Some(object), bang: false, name, .. } if i > 0 && name.eq_ignore_ascii_case(part) => cur = object,
            Expr::Ident { name, .. } if i == 0 => return name.eq_ignore_ascii_case(part),
            _ => return false,
        }
    }
    false
}

/// The `Class` argument's ProgID: the argument named `Class`, else the positional argument at
/// `position`. `None` when there is no such argument.
fn prog_id(args: &[Arg], position: usize) -> Option<ProgId> {
    Some(match arg(args, "Class", position)? {
        Expr::Literal { value: Literal::String(text), .. } => ProgId::Literal(text.clone()),
        _ => ProgId::NonLiteral,
    })
}

/// The value of the argument named `name`, else of the positional argument at `position`, with
/// parentheses looked through. `None` when there is no such argument or it is missing.
fn arg<'a>(args: &'a [Arg], name: &str, position: usize) -> Option<&'a Expr> {
    let arg = match args.iter().find(|a| a.name.as_deref().is_some_and(|n| n.eq_ignore_ascii_case(name))) {
        Some(named) => named,
        None => args.iter().filter(|a| a.name.is_none()).nth(position)?,
    };
    let mut value = arg.value.as_ref()?;
    while let Expr::Paren { inner, .. } = value {
        value = inner;
    }
    Some(value)
}

/// `expr` is a literal (parentheses already looked through by [`arg`]).
const fn is_literal(expr: &Expr) -> bool {
    matches!(expr, Expr::Literal { .. })
}

/// `prog_id` is `pattern`, or `pattern` followed by `.` and more (ASCII case-insensitive).
fn prog_id_matches(prog_id: &str, pattern: &str) -> bool {
    match prog_id.get(..pattern.len()) {
        Some(head) if head.eq_ignore_ascii_case(pattern) => {
            let rest = prog_id.get(pattern.len()..).unwrap_or("");
            rest.is_empty() || rest.starts_with('.')
        }
        _ => false,
    }
}

/// Whether a call of (or bare use of) `target` matches `callee` (see [`triggers`]).
fn callee_matches(callee: &Callee, target: &Expr, with: Option<&Expr>) -> bool {
    // The name, and the object it is on: `None` for a bare name, `Some(None)` for a `With`
    // member with no enclosing `With`.
    let (name, object) = match target {
        Expr::Ident { name, .. } => (name.as_str(), None),
        Expr::Member { object: Some(object), bang: false, name, .. } => (name.as_str(), Some(Some(&**object))),
        Expr::Member { object: None, bang: false, name, .. } => (name.as_str(), Some(with)),
        _ => return false,
    };
    let on_any = |object: Option<&Expr>, paths: &[&[&str]]| object.is_some_and(|o| paths.iter().any(|p| is_path(o, p)));
    match callee {
        Callee::Vba(n) => name.eq_ignore_ascii_case(n) && object.is_none_or(|o| on_any(o, VBA_QUALIFIERS)),
        Callee::On(n, paths) => name.eq_ignore_ascii_case(n) && object.is_some_and(|o| on_any(o, paths)),
        Callee::Method(n) => name.eq_ignore_ascii_case(n) && object.is_some(),
    }
}

/// The ActiveX heuristic: a dotted type name from a library outside Excel's own object models.
fn is_activex_type(name: &str) -> bool {
    const KNOWN: [&str; 5] = ["Excel", "VBA", "Office", "MSForms", "Scripting"];
    match name.split_once('.') {
        Some((lib, _)) => !KNOWN.iter().any(|k| lib.eq_ignore_ascii_case(k)),
        None => false,
    }
}

#[cfg(test)]
mod manifest_tests;
#[cfg(test)]
mod tests;
