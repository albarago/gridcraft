//! The VBA syntax tree that [`crate::parser::parse`] builds, following the module, procedure,
//! statement and expression grammar of [MS-VBAL] §5.
//!
//! Every node carries the [`Span`] of the source it came from: line and column of its first token
//! and the byte range up to the end of its last token. Identifiers keep their source spelling
//! (VBA compares them case-insensitively, so consumers should too); a `[bracketed]` foreign name
//! keeps its brackets. A type suffix (`s$`, `n%`) is split off into a `suffix` field.
//!
//! The tree records syntax only. It doesn't resolve names, so a call and an array index look the
//! same ([`Expr::CallOrIndex`]), and `Debug.Assert x` is an ordinary call statement.
//!
//! [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/
use crate::lexer::{Keyword, Span};

/// A parsed module (§5.1, §5.2).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Module {
    /// `Attribute name = value` lines outside procedures, in source order.
    pub attrs: Vec<Attribute>,
    /// `Option …` directives, in source order.
    pub options: Vec<OptionStmt>,
    /// Module-level declarations, in source order.
    pub decls: Vec<Decl>,
    /// Procedures, in source order.
    pub procs: Vec<Proc>,
    /// The whole source.
    pub span: Span,
}

/// `Attribute name = value[, value…]` (§5.2.4.1). The name may be dotted (`Foo.VB_UserMemId`);
/// `VB_Ext_KEY` takes two values.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attribute {
    /// The attribute name as written, dots included.
    pub name: String,
    /// The values after `=`.
    pub values: Vec<Expr>,
    /// The whole line.
    pub span: Span,
}

/// An `Option` directive (§5.2.1).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OptionStmt {
    /// Which option.
    pub kind: OptionKind,
    /// The whole directive.
    pub span: Span,
}

/// The `Option` directives.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OptionKind {
    /// `Option Explicit`.
    Explicit,
    /// `Option Base n`, with `n` as written.
    Base(String),
    /// `Option Compare Binary|Text|Database`.
    Compare(CompareMode),
    /// `Option Private Module`.
    PrivateModule,
}

/// The `Option Compare` modes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum CompareMode {
    /// `Binary`.
    Binary,
    /// `Text`.
    Text,
    /// `Database` (Access only).
    Database,
}

/// A declaration's access modifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Visibility {
    /// `Public`.
    Public,
    /// `Private`.
    Private,
    /// `Friend` (procedures in class modules).
    Friend,
    /// `Global` (an old spelling of `Public`).
    Global,
}

/// A module-level declaration (§5.2.3).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Decl {
    /// `Dim`, `Private`, `Public`, `Global` or `Static` variables.
    Var(VarStmt),
    /// `[Public|Private] Const …`.
    Const(ConstStmt),
    /// `Type … End Type`.
    Type(TypeDecl),
    /// `Enum … End Enum`.
    Enum(EnumDecl),
    /// `Declare [PtrSafe] Sub|Function … Lib "…"`.
    Declare(DeclareDecl),
    /// `Event name(params)`.
    Event(EventDecl),
    /// `Implements name`.
    Implements {
        /// The interface name, dots included.
        name: String,
        /// The whole statement.
        span: Span,
    },
    /// `DefInt A-Z` and the other `Def…` statements.
    Def(DefDecl),
    /// A conditional-compilation line (`#If`, `#Const`, …), kept unparsed.
    Directive {
        /// The line's source text from `#` to its last token (trailing comment excluded), so the
        /// printer can write it back.
        text: String,
        /// The whole line.
        span: Span,
    },
}

/// The keyword that starts a variable declaration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum VarKeyword {
    /// `Dim`.
    Dim,
    /// `Static`.
    Static,
    /// `Public`.
    Public,
    /// `Private`.
    Private,
    /// `Global`.
    Global,
}

/// `Dim a, b(1 To 5) As Long` and friends (§5.2.3.1, §5.4.3.1).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VarStmt {
    /// The declaring keyword.
    pub keyword: VarKeyword,
    /// The variables, in order.
    pub vars: Vec<VarDecl>,
    /// The whole statement.
    pub span: Span,
}

/// One declared variable, or one `Type` member.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VarDecl {
    /// The name as written.
    pub name: String,
    /// A type suffix (`$`, `%`, …) after the name.
    pub suffix: Option<char>,
    /// `WithEvents`.
    pub with_events: bool,
    /// Array bounds: `None` for a scalar, `Some(vec![])` for `a()`.
    pub bounds: Option<Vec<Bound>>,
    /// `As …`.
    pub ty: Option<TypeRef>,
    /// The declarator.
    pub span: Span,
}

/// One array dimension: `upper` or `lower To upper`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Bound {
    /// The lower bound, when written.
    pub lower: Option<Expr>,
    /// The upper bound.
    pub upper: Expr,
    /// The dimension.
    pub span: Span,
}

/// The type in an `As` clause.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TypeRef {
    /// The type name as written, dots included (`Long`, `Excel.Range`).
    pub name: String,
    /// `As New …`.
    pub is_new: bool,
    /// The length of a fixed-length `String * n`.
    pub fixed_len: Option<Box<Expr>>,
    /// `As T()`, an array return type.
    pub is_array: bool,
    /// The clause after `As`.
    pub span: Span,
}

/// `[Public|Private] Const a = 1, b As Long = 2` (§5.2.3.2, §5.4.3.2).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConstStmt {
    /// The access modifier, when written.
    pub visibility: Option<Visibility>,
    /// The constants, in order.
    pub consts: Vec<ConstDecl>,
    /// The whole statement.
    pub span: Span,
}

/// One constant.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ConstDecl {
    /// The name as written.
    pub name: String,
    /// A type suffix after the name.
    pub suffix: Option<char>,
    /// `As …`.
    pub ty: Option<TypeRef>,
    /// The value after `=`.
    pub value: Expr,
    /// The declarator.
    pub span: Span,
}

/// `Type name … End Type` (§5.2.3.3).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TypeDecl {
    /// The access modifier, when written.
    pub visibility: Option<Visibility>,
    /// The type name.
    pub name: String,
    /// The members, in order.
    pub members: Vec<VarDecl>,
    /// `Type` to `End Type`.
    pub span: Span,
}

/// `Enum name … End Enum` (§5.2.3.4).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EnumDecl {
    /// The access modifier, when written.
    pub visibility: Option<Visibility>,
    /// The enum name.
    pub name: String,
    /// The members, in order.
    pub members: Vec<EnumMember>,
    /// `Enum` to `End Enum`.
    pub span: Span,
}

/// One `name [= value]` line of an enum.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EnumMember {
    /// The member name.
    pub name: String,
    /// The explicit value.
    pub value: Option<Expr>,
    /// The line.
    pub span: Span,
}

/// Whether a `Declare` names a `Sub` or a `Function`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DeclareKind {
    /// `Declare Sub`.
    Sub,
    /// `Declare Function`.
    Function,
}

/// `Declare [PtrSafe] Sub|Function name Lib "lib" [Alias "alias"] (params) [As T]` (§5.2.3.5): a
/// call into native code, which the capability analyzer flags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DeclareDecl {
    /// The access modifier, when written.
    pub visibility: Option<Visibility>,
    /// `PtrSafe`.
    pub ptr_safe: bool,
    /// `Sub` or `Function`.
    pub kind: DeclareKind,
    /// The VBA-side name.
    pub name: String,
    /// A type suffix after the name.
    pub suffix: Option<char>,
    /// The library, unquoted.
    pub lib: String,
    /// The `Alias` entry point, unquoted.
    pub alias: Option<String>,
    /// The parameters.
    pub params: Vec<Param>,
    /// `As …` of a `Function`.
    pub return_type: Option<TypeRef>,
    /// The whole statement.
    pub span: Span,
}

/// `Event name(params)` (§5.2.3.6).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EventDecl {
    /// The access modifier, when written.
    pub visibility: Option<Visibility>,
    /// The event name.
    pub name: String,
    /// The parameters.
    pub params: Vec<Param>,
    /// The whole statement.
    pub span: Span,
}

/// `DefInt A-Z, I` and the other `Def…` statements (§5.2.2).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefDecl {
    /// Which `Def…` keyword.
    pub keyword: Keyword,
    /// The letter ranges.
    pub ranges: Vec<DefRange>,
    /// The whole statement.
    pub span: Span,
}

/// One `A` or `A-Z` letter range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefRange {
    /// The first letter.
    pub from: String,
    /// The last letter, for a range.
    pub to: Option<String>,
}

/// What kind of procedure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ProcKind {
    /// `Sub`.
    Sub,
    /// `Function`.
    Function,
    /// `Property Get`.
    PropertyGet,
    /// `Property Let`.
    PropertyLet,
    /// `Property Set`.
    PropertySet,
}

/// A `Sub`, `Function` or `Property` procedure (§5.3).
///
/// A procedure whose body didn't parse is kept: its `body` holds the statements that parsed before
/// the error and ends with one [`Stmt::Unknown`] covering the skipped source
/// ([`Proc::is_recovered`]).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Proc {
    /// What kind of procedure.
    pub kind: ProcKind,
    /// The name as written.
    pub name: String,
    /// A type suffix after the name (`Function F$()`).
    pub suffix: Option<char>,
    /// The access modifier, when written.
    pub visibility: Option<Visibility>,
    /// `Static`.
    pub is_static: bool,
    /// The parameters.
    pub params: Vec<Param>,
    /// `As …` of a `Function` or `Property Get`.
    pub return_type: Option<TypeRef>,
    /// The body statements.
    pub body: Vec<Stmt>,
    /// Header to `End …`.
    pub span: Span,
}

impl Proc {
    /// `true` when the body failed to parse and ends with a [`Stmt::Unknown`].
    #[must_use]
    pub fn is_recovered(&self) -> bool {
        matches!(self.body.last(), Some(Stmt::Unknown { .. }))
    }
}

/// How an argument is passed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Passing {
    /// `ByVal`.
    ByVal,
    /// `ByRef`.
    ByRef,
}

/// One procedure parameter (§5.3.1.5).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Param {
    /// The name as written.
    pub name: String,
    /// A type suffix after the name.
    pub suffix: Option<char>,
    /// `Optional`.
    pub optional: bool,
    /// `ByVal` or `ByRef`, when written.
    pub passing: Option<Passing>,
    /// `ParamArray`.
    pub param_array: bool,
    /// `name()`, an array parameter.
    pub is_array: bool,
    /// `As …`.
    pub ty: Option<TypeRef>,
    /// `= default` of an optional parameter.
    pub default: Option<Expr>,
    /// The parameter.
    pub span: Span,
}

/// A statement inside a procedure (§5.4).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Stmt {
    /// `target = value`, `Let`, `Set`, `LSet` or `RSet`.
    Assign {
        /// The assignment keyword.
        kind: AssignKind,
        /// The assigned l-expression.
        target: Expr,
        /// The value.
        value: Expr,
        /// The statement.
        span: Span,
    },
    /// `Call f(x)`, `f x, y`, `obj.Method a`, or `f(x)`. Without `Call`, a head written with
    /// parentheses (`f(x)`) stays an [`Expr::CallOrIndex`] target with no `args`.
    Call {
        /// Written with `Call`.
        explicit: bool,
        /// The called expression.
        target: Expr,
        /// The arguments outside the target.
        args: Vec<Arg>,
        /// The statement.
        span: Span,
    },
    /// `Dim` or `Static` local variables.
    Var(VarStmt),
    /// A local `Const`.
    Const(ConstStmt),
    /// Block or single-line `If`.
    If(IfStmt),
    /// `Select Case`.
    Select(SelectStmt),
    /// `For … To … [Step …] … Next`.
    For(ForStmt),
    /// `For Each … In … Next`.
    ForEach(ForEachStmt),
    /// `Do … Loop` with an optional `While`/`Until` at either end.
    Do(DoStmt),
    /// `While … Wend`.
    While {
        /// The loop condition.
        cond: Expr,
        /// The body.
        body: Vec<Stmt>,
        /// `While` to `Wend`.
        span: Span,
    },
    /// `With … End With`.
    With {
        /// The `With` object.
        object: Expr,
        /// The body.
        body: Vec<Stmt>,
        /// `With` to `End With`.
        span: Span,
    },
    /// `Exit Sub|Function|Property|For|Do`.
    Exit {
        /// What is exited.
        kind: ExitKind,
        /// The statement.
        span: Span,
    },
    /// `On Error GoTo …` or `On Error Resume Next`.
    OnError {
        /// The handler.
        action: OnErrorAction,
        /// The statement.
        span: Span,
    },
    /// `On expr GoTo|GoSub label, label…`.
    OnJump {
        /// The selector.
        expr: Expr,
        /// `GoSub` rather than `GoTo`.
        gosub: bool,
        /// The targets.
        labels: Vec<String>,
        /// The statement.
        span: Span,
    },
    /// `Resume`, `Resume Next` or `Resume label`.
    Resume {
        /// Where execution resumes.
        target: ResumeTarget,
        /// The statement.
        span: Span,
    },
    /// `GoTo label`.
    GoTo {
        /// The label or line number.
        label: String,
        /// The statement.
        span: Span,
    },
    /// `GoSub label`.
    GoSub {
        /// The label or line number.
        label: String,
        /// The statement.
        span: Span,
    },
    /// `Return`.
    Return {
        /// The statement.
        span: Span,
    },
    /// `name:` or a line number at the start of a line.
    Label {
        /// The label or line number as written.
        name: String,
        /// A line number rather than a name.
        is_line_number: bool,
        /// The label (and its colon).
        span: Span,
    },
    /// `ReDim [Preserve] a(…) [As T], …`.
    ReDim {
        /// `Preserve`.
        preserve: bool,
        /// The resized arrays.
        vars: Vec<ReDimVar>,
        /// The statement.
        span: Span,
    },
    /// `Erase a, b`.
    Erase {
        /// The erased arrays.
        targets: Vec<Expr>,
        /// The statement.
        span: Span,
    },
    /// `RaiseEvent name(args)`.
    RaiseEvent {
        /// The event name.
        name: String,
        /// The arguments.
        args: Vec<Arg>,
        /// The statement.
        span: Span,
    },
    /// `End` (ends the program).
    End {
        /// The statement.
        span: Span,
    },
    /// `Stop`.
    Stop {
        /// The statement.
        span: Span,
    },
    /// `Debug.Print` with its print list.
    DebugPrint {
        /// The items, in order.
        items: Vec<PrintItem>,
        /// The statement.
        span: Span,
    },
    /// A file statement (`Open`, `Close`, `Print #`, …).
    File(FileStmt),
    /// An `Attribute` line inside a procedure.
    Attribute(Attribute),
    /// A conditional-compilation line (`#If`, `#Else`, …), kept unparsed.
    Directive {
        /// The line's source text from `#` to its last token (trailing comment excluded).
        text: String,
        /// The whole line.
        span: Span,
    },
    /// Source the parser skipped while recovering from an error.
    Unknown {
        /// The skipped source.
        span: Span,
    },
}

/// The keyword of an assignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AssignKind {
    /// No keyword.
    Implicit,
    /// `Let`.
    Let,
    /// `Set`.
    Set,
    /// `LSet`.
    LSet,
    /// `RSet`.
    RSet,
}

/// Block or single-line `If`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IfStmt {
    /// The `If` branch followed by any `ElseIf` branches.
    pub branches: Vec<IfBranch>,
    /// The `Else` body.
    pub else_body: Option<Vec<Stmt>>,
    /// Written on one line (`If c Then a Else b`).
    pub single_line: bool,
    /// `If` to `End If` (or the end of the line).
    pub span: Span,
}

/// One `If`/`ElseIf` condition and its body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IfBranch {
    /// The condition.
    pub cond: Expr,
    /// The body.
    pub body: Vec<Stmt>,
    /// From `If`/`ElseIf` to the end of the body.
    pub span: Span,
}

/// `Select Case expr … End Select`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SelectStmt {
    /// The tested expression.
    pub expr: Expr,
    /// The `Case` clauses before `Case Else`.
    pub cases: Vec<CaseClause>,
    /// The `Case Else` body.
    pub else_body: Option<Vec<Stmt>>,
    /// `Select` to `End Select`.
    pub span: Span,
}

/// One `Case a, b To c, Is > d` clause.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CaseClause {
    /// The tests, in order.
    pub tests: Vec<CaseTest>,
    /// The body.
    pub body: Vec<Stmt>,
    /// From `Case` to the end of the body.
    pub span: Span,
}

/// One test in a `Case` clause.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum CaseTest {
    /// `value`.
    Value(Expr),
    /// `from To to`.
    Range {
        /// The lower end.
        from: Expr,
        /// The upper end.
        to: Expr,
    },
    /// `Is op value` (or `op value`); `op` is a comparison.
    Is {
        /// The comparison.
        op: BinaryOp,
        /// The compared value.
        value: Expr,
    },
}

/// `For var = from To to [Step step] … Next [var]`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ForStmt {
    /// The loop variable.
    pub var: Expr,
    /// The start value.
    pub from: Expr,
    /// The end value.
    pub to: Expr,
    /// The step.
    pub step: Option<Expr>,
    /// The body.
    pub body: Vec<Stmt>,
    /// The variable after `Next`, when written.
    pub next_var: Option<Expr>,
    /// `For` to `Next`.
    pub span: Span,
}

/// `For Each var In collection … Next [var]`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ForEachStmt {
    /// The loop variable.
    pub var: Expr,
    /// The iterated collection.
    pub collection: Expr,
    /// The body.
    pub body: Vec<Stmt>,
    /// The variable after `Next`, when written.
    pub next_var: Option<Expr>,
    /// `For` to `Next`.
    pub span: Span,
}

/// `Do [While|Until c] … Loop [While|Until c]`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DoStmt {
    /// The condition after `Do`.
    pub pre: Option<LoopCond>,
    /// The condition after `Loop`.
    pub post: Option<LoopCond>,
    /// The body.
    pub body: Vec<Stmt>,
    /// `Do` to `Loop`.
    pub span: Span,
}

/// A `While c` or `Until c` loop condition.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LoopCond {
    /// `Until` rather than `While`.
    pub until: bool,
    /// The condition.
    pub cond: Expr,
}

/// What an `Exit` statement leaves.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ExitKind {
    /// `Exit Sub`.
    Sub,
    /// `Exit Function`.
    Function,
    /// `Exit Property`.
    Property,
    /// `Exit For`.
    For,
    /// `Exit Do`.
    Do,
}

/// The handler an `On Error` statement installs.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OnErrorAction {
    /// `On Error GoTo label` (a name or a line number other than 0).
    GoTo(String),
    /// `On Error GoTo 0`.
    GoToZero,
    /// `On Error GoTo -1`.
    GoToMinusOne,
    /// `On Error Resume Next`.
    ResumeNext,
}

/// Where `Resume` continues.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ResumeTarget {
    /// `Resume`: the failing statement.
    Same,
    /// `Resume Next`.
    Next,
    /// `Resume label`.
    Label(String),
}

/// One array in a `ReDim`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReDimVar {
    /// The array (a name or member access).
    pub target: Expr,
    /// The new bounds.
    pub bounds: Vec<Bound>,
    /// `As …`.
    pub ty: Option<TypeRef>,
    /// The declarator.
    pub span: Span,
}

/// The separator after a print item.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum PrintSep {
    /// `;`.
    Semicolon,
    /// `,`.
    Comma,
}

/// One item of a `Print` or `Write` list: an optional expression and an optional separator.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PrintItem {
    /// The printed expression (`Spc(n)` and `Tab(n)` are calls).
    pub expr: Option<Expr>,
    /// The separator after it.
    pub sep: Option<PrintSep>,
    /// The item and its separator.
    pub span: Span,
}

/// The file statements (§5.4.5).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FileOp {
    /// `Open path For mode … As #n [Len = r]`.
    Open,
    /// `Close [#n, …]`.
    Close,
    /// `Print #n, items`.
    Print,
    /// `Write #n, items`.
    Write,
    /// `Input #n, vars`.
    Input,
    /// `Line Input #n, var`.
    LineInput,
    /// `Get #n, [rec], var`.
    Get,
    /// `Put #n, [rec], var`.
    Put,
    /// `Seek #n, pos`.
    Seek,
    /// `Lock #n[, rec [To rec]]`.
    Lock,
    /// `Unlock #n[, rec [To rec]]`.
    Unlock,
    /// `Width #n, w`.
    Width,
}

/// A file statement with its expressions, so the capability analyzer can see them.
///
/// - `Open`: `args` is `[path, Len value?]`, `number` the file number, `mode` the words between
///   the path and `As` without the leading `For` (`["Binary", "Access", "Read"]`).
/// - `Close`: `args` holds the file numbers.
/// - `Print`, `Write`: `number` and the print list in `items`.
/// - The others: `number` and the arguments after it; a missing argument (`Put #1, , x`) is
///   `None`, and `Lock #1, a To b` gives `[a, b]`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileStmt {
    /// Which statement.
    pub op: FileOp,
    /// The file number (without its `#`).
    pub number: Option<Expr>,
    /// The other expressions (see above).
    pub args: Vec<Option<Expr>>,
    /// The print list of `Print #` and `Write #`.
    pub items: Vec<PrintItem>,
    /// The mode words of `Open`.
    pub mode: Vec<String>,
    /// The statement.
    pub span: Span,
}

impl FileStmt {
    /// Every expression the statement holds, in source order.
    pub fn exprs(&self) -> impl Iterator<Item = &Expr> {
        let (head, tail) = match self.op {
            FileOp::Open => (self.args.first(), self.args.get(1..)),
            _ => (None, Some(self.args.as_slice())),
        };
        let head = head.and_then(Option::as_ref);
        let tail = tail.unwrap_or(&[]).iter().filter_map(Option::as_ref);
        head.into_iter().chain(self.number.iter()).chain(tail).chain(self.items.iter().filter_map(|item| item.expr.as_ref()))
    }
}

/// An expression (§5.6).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Expr {
    /// A literal.
    Literal {
        /// The value.
        value: Literal,
        /// The literal.
        span: Span,
    },
    /// A simple name.
    Ident {
        /// The name as written.
        name: String,
        /// A type suffix after it.
        suffix: Option<char>,
        /// The name and suffix.
        span: Span,
    },
    /// `Me`.
    Me {
        /// The keyword.
        span: Span,
    },
    /// `object.name` or `object!name`; `object` is `None` inside `With` (`.name`, `!name`).
    Member {
        /// The object, or `None` for a `With`-relative access.
        object: Option<Box<Expr>>,
        /// `!` (dictionary access) rather than `.`.
        bang: bool,
        /// The member name as written.
        name: String,
        /// A type suffix after it.
        suffix: Option<char>,
        /// The whole access.
        span: Span,
    },
    /// `target(args)`: a call or an index, which VBA can't tell apart while parsing.
    CallOrIndex {
        /// The called or indexed expression.
        target: Box<Expr>,
        /// The arguments.
        args: Vec<Arg>,
        /// The whole expression.
        span: Span,
    },
    /// `(inner)`.
    Paren {
        /// The parenthesised expression.
        inner: Box<Expr>,
        /// From `(` to `)`.
        span: Span,
    },
    /// `-x` or `Not x`.
    Unary {
        /// The operator.
        op: UnaryOp,
        /// The operand.
        operand: Box<Expr>,
        /// The whole expression.
        span: Span,
    },
    /// `left op right`.
    Binary {
        /// The operator.
        op: BinaryOp,
        /// The left operand.
        left: Box<Expr>,
        /// The right operand.
        right: Box<Expr>,
        /// The whole expression.
        span: Span,
    },
    /// `New Class`.
    New {
        /// The class name, dots included.
        class: String,
        /// The whole expression.
        span: Span,
    },
    /// `TypeOf object Is Class`.
    TypeOfIs {
        /// The tested object.
        object: Box<Expr>,
        /// The class name, dots included.
        class: String,
        /// The whole expression.
        span: Span,
    },
    /// `AddressOf procedure`.
    AddressOf {
        /// The procedure (a name or member access).
        target: Box<Expr>,
        /// The whole expression.
        span: Span,
    },
}

impl Expr {
    /// Where the expression is.
    #[must_use]
    pub const fn span(&self) -> Span {
        match self {
            Expr::Literal { span, .. }
            | Expr::Ident { span, .. }
            | Expr::Me { span }
            | Expr::Member { span, .. }
            | Expr::CallOrIndex { span, .. }
            | Expr::Paren { span, .. }
            | Expr::Unary { span, .. }
            | Expr::Binary { span, .. }
            | Expr::New { span, .. }
            | Expr::TypeOfIs { span, .. }
            | Expr::AddressOf { span, .. } => *span,
        }
    }

    /// Moves the sub-expressions of `self` onto `stack`, leaving leaves in their place.
    fn take_children(&mut self, stack: &mut Vec<Expr>) {
        const LEAF: Expr = Expr::Me { span: Span { line: 0, col: 0, start: 0, end: 0 } };
        match self {
            Expr::Binary { left, right, .. } => {
                stack.push(std::mem::replace(&mut **left, LEAF));
                stack.push(std::mem::replace(&mut **right, LEAF));
            }
            Expr::Unary { operand: child, .. }
            | Expr::Paren { inner: child, .. }
            | Expr::AddressOf { target: child, .. }
            | Expr::TypeOfIs { object: child, .. } => stack.push(std::mem::replace(&mut **child, LEAF)),
            Expr::Member { object, .. } => stack.extend(object.take().map(|object| *object)),
            Expr::CallOrIndex { target, args, .. } => {
                stack.push(std::mem::replace(&mut **target, LEAF));
                stack.extend(args.drain(..).filter_map(|arg| arg.value));
            }
            Expr::Literal { .. } | Expr::Ident { .. } | Expr::Me { .. } | Expr::New { .. } => {}
        }
    }
}

/// Drops the tree with a heap stack instead of recursion. The parser builds flat operator chains
/// (`a + b + …`) and member or call chains (`a.b.c…`) in loops without a nesting level per link,
/// so their trees can be as deep as the source is long; the derived drop would recurse once per
/// link and overflow the stack (Req 12.9).
impl Drop for Expr {
    fn drop(&mut self) {
        let mut stack = Vec::new();
        self.take_children(&mut stack);
        while let Some(mut child) = stack.pop() {
            child.take_children(&mut stack);
        }
    }
}

/// A literal value. Numbers and dates keep their source text (suffix included); strings are
/// unquoted with `""` turned into `"`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Literal {
    /// A decimal integer.
    Integer(String),
    /// `&H…`.
    Hex(String),
    /// `&O…` or `&…`.
    Octal(String),
    /// A floating-point number.
    Float(String),
    /// A string, unquoted.
    String(String),
    /// `#…#`, as written.
    Date(String),
    /// `True` or `False`.
    Boolean(bool),
    /// `Nothing`.
    Nothing,
    /// `Empty`.
    Empty,
    /// `Null`.
    Null,
}

/// A prefix operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum UnaryOp {
    /// `-`.
    Neg,
    /// `Not`.
    Not,
}

/// A binary operator (§5.6.9).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BinaryOp {
    /// `^`.
    Pow,
    /// `*`.
    Mul,
    /// `/`.
    Div,
    /// `\`.
    IntDiv,
    /// `Mod`.
    Mod,
    /// `+`.
    Add,
    /// `-`.
    Sub,
    /// `&`.
    Concat,
    /// `=`.
    Eq,
    /// `<>`.
    Ne,
    /// `<`.
    Lt,
    /// `>`.
    Gt,
    /// `<=` (or `=<`).
    Le,
    /// `>=` (or `=>`).
    Ge,
    /// `Like`.
    Like,
    /// `Is`.
    Is,
    /// `And`.
    And,
    /// `Or`.
    Or,
    /// `Xor`.
    Xor,
    /// `Eqv`.
    Eqv,
    /// `Imp`.
    Imp,
}

impl BinaryOp {
    /// The binding power: higher binds tighter. All binary operators are left-associative.
    /// From loosest to tightest: `Imp`, `Eqv`, `Xor`, `Or`, `And`, (prefix `Not`), comparisons
    /// (`= <> < > <= >= Like Is`), `&`, `+ -`, `Mod`, `\`, `* /`, (prefix `-`), `^`.
    #[must_use]
    pub const fn precedence(self) -> u8 {
        match self {
            BinaryOp::Imp => 1,
            BinaryOp::Eqv => 2,
            BinaryOp::Xor => 3,
            BinaryOp::Or => 4,
            BinaryOp::And => 5,
            BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Gt | BinaryOp::Le | BinaryOp::Ge | BinaryOp::Like | BinaryOp::Is => 7,
            BinaryOp::Concat => 8,
            BinaryOp::Add | BinaryOp::Sub => 9,
            BinaryOp::Mod => 10,
            BinaryOp::IntDiv => 11,
            BinaryOp::Mul | BinaryOp::Div => 12,
            BinaryOp::Pow => 14,
        }
    }

    /// `true` for `= <> < > <= >=` (the operators `Case Is` accepts).
    #[must_use]
    pub const fn is_relational(self) -> bool {
        matches!(self, BinaryOp::Eq | BinaryOp::Ne | BinaryOp::Lt | BinaryOp::Gt | BinaryOp::Le | BinaryOp::Ge)
    }
}

/// `Not`'s operand binds at least as tightly as a comparison.
pub(crate) const NOT_OPERAND: u8 = 7;
/// `TypeOf`'s operand stops before `Is`.
pub(crate) const TYPEOF_OPERAND: u8 = 8;
/// Unary minus's operand: only `^` binds tighter.
pub(crate) const NEG_OPERAND: u8 = 14;

/// One argument: positional, named (`name:=value`) or missing (`f(, x)`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Arg {
    /// The parameter name of a named argument.
    pub name: Option<String>,
    /// The value, or `None` for a missing argument.
    pub value: Option<Expr>,
    /// The argument (empty, at the following comma, for a missing one).
    pub span: Span,
}

impl Stmt {
    /// Where the statement is.
    #[must_use]
    pub const fn span(&self) -> Span {
        match self {
            Stmt::Var(VarStmt { span, .. })
            | Stmt::Const(ConstStmt { span, .. })
            | Stmt::If(IfStmt { span, .. })
            | Stmt::Select(SelectStmt { span, .. })
            | Stmt::For(ForStmt { span, .. })
            | Stmt::ForEach(ForEachStmt { span, .. })
            | Stmt::Do(DoStmt { span, .. })
            | Stmt::File(FileStmt { span, .. })
            | Stmt::Attribute(Attribute { span, .. })
            | Stmt::Assign { span, .. }
            | Stmt::Call { span, .. }
            | Stmt::While { span, .. }
            | Stmt::With { span, .. }
            | Stmt::Exit { span, .. }
            | Stmt::OnError { span, .. }
            | Stmt::OnJump { span, .. }
            | Stmt::Resume { span, .. }
            | Stmt::GoTo { span, .. }
            | Stmt::GoSub { span, .. }
            | Stmt::Return { span }
            | Stmt::Label { span, .. }
            | Stmt::ReDim { span, .. }
            | Stmt::Erase { span, .. }
            | Stmt::RaiseEvent { span, .. }
            | Stmt::End { span }
            | Stmt::Stop { span }
            | Stmt::DebugPrint { span, .. }
            | Stmt::Directive { span, .. }
            | Stmt::Unknown { span } => *span,
        }
    }
}

#[cfg(test)]
pub(crate) mod testeq;
