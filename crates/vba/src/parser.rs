//! A recursive-descent VBA parser following the module, procedure, statement and expression
//! grammar of [MS-VBAL] §5. [`parse`] lexes the source with [`crate::lexer::lex`] and builds a
//! [`Module`] tree plus [`ParseDiag`]s.
//!
//! # Bounded depth (Req 12.4)
//!
//! Every block statement (`If`, `Select Case`, `For`, `For Each`, `Do`, `While`, `With`, a
//! single-line `If`) and every nested sub-expression (a parenthesised expression, an argument
//! list, the operand of a prefix operator and the right operand of a binary operator) is one
//! nesting level. The top-level expression of a statement and the procedure itself are not. The
//! parser counts levels in `depth`; when a token would open level 257 ([`MAX_DEPTH`] + 1) it
//! records [`ParseDiagKind::NestingTooDeep`] at that token, stops descending, and recovers at
//! procedure level. Binary operators use precedence climbing, so a long flat chain (`a + b + …`)
//! loops instead of recursing, and the host stack used is bounded by the depth limit.
//!
//! # Recovery (Req 12.5)
//!
//! A syntax error inside a procedure records one [`ParseDiagKind::Syntax`] diagnostic at the
//! first unexpected token, then skips to the next `End Sub|Function|Property` (consumed) or the
//! next procedure header at the start of a line (not consumed), and parsing carries on.
//!
//! **A failed procedure is kept**, not dropped: its body holds the top-level statements that
//! parsed before the error, then one [`Stmt::Unknown`] covering the skipped source
//! ([`crate::ast::Proc::is_recovered`]). The statement that failed (including any blocks it
//! opened) is part of the unknown tail. A procedure whose *header* fails to parse has no name or
//! kind to report, so it is dropped (its diagnostic remains). Keeping recovered procedures lets
//! the capability analyzer still see what parsed (Req 13.11); procedures that parsed without
//! error are the ones with `!is_recovered()`.
//!
//! At module level (outside procedures) a failed line records its diagnostic and is skipped to
//! the end of its line, unless the error sits on a procedure header, which is then parsed. A
//! failed `Type` or `Enum` member line is skipped the same way and the block carries on.
//!
//! # Tokens
//!
//! Trivia are dropped after lexing, but each token remembers whether whitespace, a comment or a
//! line continuation came before it on its logical line, and whether it starts a physical line.
//! That decides:
//!
//! - two-character operators: `<=`, `<>`, `>=`, `=<`, `=>` and `:=` are joined only when the two
//!   characters are adjacent;
//! - statement heads: `f (x), y` is a call of `f` with the arguments `(x)` and `y`, while `f(x)`
//!   is a call or index target; `f .x` passes the `With` member `.x` to `f`, while `f.x` and a
//!   line-continued `f _ .x` are member accesses (in an expression, a spaced `.` or `(` still
//!   continues the term);
//! - labels: `name:` and line numbers count only at the start of a line.
//!
//! # Coverage and limits
//!
//! Conditional-compilation lines (a line starting with `#`) are kept as
//! [`Decl::Directive`]/[`Stmt::Directive`] without parsing their condition, so both branches of
//! `#If … #Else … #End If` are parsed. `Name a As b` and the `Line`/`Circle`/`PSet` drawing
//! methods are not recognised. Sources over [`crate::lexer::MAX_SOURCE`] bytes give no tree and
//! one [`ParseDiagKind::SourceTooLarge`] diagnostic (Req 12.6).
//!
//! [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/
mod expr;
mod stmt;

use crate::ast::{
    Attribute, CompareMode, ConstDecl, ConstStmt, Decl, DeclareDecl, DeclareKind, DefDecl, DefRange, EnumDecl, EnumMember, EventDecl, Module,
    OptionKind, OptionStmt, Param, Passing, Proc, ProcKind, Stmt, TypeDecl, TypeRef, VarDecl, VarKeyword, VarStmt, Visibility,
};
use crate::lexer::{self, Keyword, Span, Token, TokenKind};

/// The deepest nesting [`parse`] accepts (Req 12.4).
pub const MAX_DEPTH: u16 = 256;

/// What went wrong.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ParseDiagKind {
    /// An unexpected token.
    Syntax,
    /// Nesting deeper than [`MAX_DEPTH`] levels.
    NestingTooDeep,
    /// The source is longer than [`crate::lexer::MAX_SOURCE`] bytes.
    SourceTooLarge,
}

/// One parse diagnostic. `span` is the offending token (line and column are 1-based).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParseDiag {
    /// What went wrong.
    pub kind: ParseDiagKind,
    /// Where.
    pub span: Span,
    /// A short English description.
    pub message: String,
}

/// Parses VBA module source.
///
/// Returns the tree (always, unless the source is too large) and the diagnostics. A module
/// that is empty or holds only declarations gives a tree with no procedures and no diagnostics.
#[must_use]
pub fn parse(source: &str) -> (Option<Module>, Vec<ParseDiag>) {
    let tokens = match lexer::lex(source) {
        Ok(tokens) => tokens,
        Err(err) => {
            let span = Span { line: 1, col: 1, start: 0, end: 0 };
            return (None, vec![ParseDiag { kind: ParseDiagKind::SourceTooLarge, span, message: err.to_string() }]);
        }
    };
    let mut parser = Parser::new(source, &tokens);
    let module = parser.module();
    (Some(module), parser.diags)
}

/// A significant token with the trivia facts the grammar needs.
#[derive(Debug, Clone, Copy)]
struct Tok {
    kind: TokenKind,
    span: Span,
    /// Whitespace, a comment or a line continuation came before it on its logical line.
    spaced: bool,
    /// A line continuation came right before it.
    continued: bool,
    /// It is the first significant token of a physical line.
    line_start: bool,
}

/// A failure whose diagnostic is already recorded.
#[derive(Debug)]
struct Failed;

type PResult<T> = Result<T, Failed>;

/// What can end a statement block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Term {
    EndIf,
    Else,
    ElseIf,
    Case,
    EndSelect,
    Loop,
    Next,
    Wend,
    EndWith,
}

/// Reserved words that name built-in functions or objects, usable where a name is.
pub(crate) const BUILTIN_NAMES: &[Keyword] = &[
    Keyword::Abs,
    Keyword::Array,
    Keyword::CBool,
    Keyword::CByte,
    Keyword::CCur,
    Keyword::CDate,
    Keyword::CDbl,
    Keyword::CDec,
    Keyword::CInt,
    Keyword::CLng,
    Keyword::CLngLng,
    Keyword::CLngPtr,
    Keyword::CSng,
    Keyword::CStr,
    Keyword::CVar,
    Keyword::CVErr,
    Keyword::Date,
    Keyword::Debug,
    Keyword::DoEvents,
    Keyword::Fix,
    Keyword::Input,
    Keyword::InputB,
    Keyword::Int,
    Keyword::LBound,
    Keyword::Len,
    Keyword::LenB,
    Keyword::Seek,
    Keyword::Sgn,
    Keyword::Spc,
    Keyword::String,
    Keyword::Tab,
    Keyword::UBound,
];

/// Reserved words that name built-in types after `As`.
const TYPE_NAMES: &[Keyword] = &[
    Keyword::Any,
    Keyword::Boolean,
    Keyword::Byte,
    Keyword::Currency,
    Keyword::Date,
    Keyword::Decimal,
    Keyword::Double,
    Keyword::Integer,
    Keyword::Long,
    Keyword::LongLong,
    Keyword::LongPtr,
    Keyword::Single,
    Keyword::String,
    Keyword::Variant,
];

const DEF_KEYWORDS: &[Keyword] = &[
    Keyword::DefBool,
    Keyword::DefByte,
    Keyword::DefCur,
    Keyword::DefDate,
    Keyword::DefDbl,
    Keyword::DefDec,
    Keyword::DefInt,
    Keyword::DefLng,
    Keyword::DefLngLng,
    Keyword::DefLngPtr,
    Keyword::DefObj,
    Keyword::DefSng,
    Keyword::DefStr,
    Keyword::DefVar,
];

struct Parser<'a> {
    src: &'a str,
    toks: Vec<Tok>,
    pos: usize,
    /// Current nesting level (see the module docs).
    depth: u16,
    diags: Vec<ParseDiag>,
    /// End of the last consumed token.
    last_end: usize,
    /// Where the end of the source is.
    eof: Span,
    /// `Next i, j` closed the inner loop; the outer `For` closes at the current token.
    pending_next: bool,
    /// Inside a single-line `If`, where `Else` ends a statement.
    inline_if: u32,
}

impl<'a> Parser<'a> {
    fn new(src: &'a str, tokens: &[Token]) -> Self {
        let mut toks = Vec::with_capacity(tokens.len());
        let (mut spaced, mut continued, mut line_start) = (false, false, true);
        for token in tokens {
            match token.kind {
                TokenKind::Whitespace | TokenKind::Comment => spaced = true,
                TokenKind::LineContinuation => (spaced, continued) = (true, true),
                kind => {
                    toks.push(Tok { kind, span: token.span, spaced, continued, line_start });
                    (spaced, continued, line_start) = (false, false, kind == TokenKind::Newline);
                }
            }
        }
        let eof = eof_span(src, tokens.last());
        Self { src, toks, pos: 0, depth: 0, diags: Vec::new(), last_end: 0, eof, pending_next: false, inline_if: 0 }
    }

    // ----- token access -------------------------------------------------------------------

    fn tok(&self, n: usize) -> Option<&Tok> {
        self.toks.get(self.pos.saturating_add(n))
    }

    fn at_eof(&self) -> bool {
        self.pos >= self.toks.len()
    }

    fn kind_at(&self, n: usize) -> Option<TokenKind> {
        self.tok(n).map(|t| t.kind)
    }

    fn kw_at(&self, n: usize, keyword: Keyword) -> bool {
        self.kind_at(n) == Some(TokenKind::Keyword(keyword))
    }

    fn is_kw(&self, keyword: Keyword) -> bool {
        self.kw_at(0, keyword)
    }

    /// The keyword at `n`, if any.
    fn keyword_at(&self, n: usize) -> Option<Keyword> {
        match self.kind_at(n) {
            Some(TokenKind::Keyword(keyword)) => Some(keyword),
            _ => None,
        }
    }

    /// The character of the punctuation or operator token at `n`.
    fn char_at(&self, n: usize) -> Option<u8> {
        let tok = self.tok(n)?;
        if matches!(tok.kind, TokenKind::Punct | TokenKind::Operator) { self.src.as_bytes().get(tok.span.start).copied() } else { None }
    }

    fn is_char(&self, n: usize, c: u8) -> bool {
        self.char_at(n) == Some(c)
    }

    /// The token at `n` follows the previous one directly (no trivia, same line).
    fn adjacent(&self, n: usize) -> bool {
        self.tok(n).is_some_and(|t| !t.spaced && !t.line_start)
    }

    fn is_newline(&self) -> bool {
        self.kind_at(0) == Some(TokenKind::Newline)
    }

    fn text_at(&self, n: usize) -> &'a str {
        self.tok(n).and_then(|t| self.src.get(t.span.start..t.span.end)).unwrap_or("")
    }

    /// The source text of `span`.
    fn span_text(&self, span: Span) -> String {
        self.src.get(span.start..span.end).unwrap_or("").to_owned()
    }

    fn bump(&mut self) {
        if let Some(tok) = self.toks.get(self.pos) {
            self.last_end = tok.span.end;
            self.pos = self.pos.saturating_add(1);
        }
    }

    fn bump_n(&mut self, n: usize) {
        for _ in 0..n {
            self.bump();
        }
    }

    /// The current token's span, or the end of the source.
    fn here(&self) -> Span {
        self.tok(0).map_or(self.eof, |t| t.span)
    }

    /// `start` extended to the end of the last consumed token.
    fn span_from(&self, start: Span) -> Span {
        Span { end: self.last_end.max(start.start), ..start }
    }

    fn eat_kw(&mut self, keyword: Keyword) -> bool {
        let found = self.is_kw(keyword);
        if found {
            self.bump();
        }
        found
    }

    fn eat_char(&mut self, c: u8) -> bool {
        let found = self.is_char(0, c);
        if found {
            self.bump();
        }
        found
    }

    fn expect_kw(&mut self, keyword: Keyword) -> PResult<()> {
        if self.eat_kw(keyword) { Ok(()) } else { self.fail(keyword.as_str()) }
    }

    fn expect_char(&mut self, c: u8) -> PResult<()> {
        if self.eat_char(c) { Ok(()) } else { self.fail(&format!("'{}'", char::from(c))) }
    }

    /// Records a syntax error at the current token.
    fn fail<T>(&mut self, expected: &str) -> PResult<T> {
        let found = match self.tok(0) {
            None => "the end of the source".to_owned(),
            Some(t) if t.kind == TokenKind::Newline => "the end of the line".to_owned(),
            Some(_) => format!("'{}'", self.text_at(0).chars().take(40).collect::<String>()),
        };
        let span = self.here();
        self.diags.push(ParseDiag { kind: ParseDiagKind::Syntax, span, message: format!("expected {expected}, found {found}") });
        Err(Failed)
    }

    /// Runs `f` one nesting level deeper, or records [`ParseDiagKind::NestingTooDeep`] at the
    /// current token when that level would pass [`MAX_DEPTH`].
    fn nested<T>(&mut self, f: impl FnOnce(&mut Self) -> PResult<T>) -> PResult<T> {
        if self.depth >= MAX_DEPTH {
            let span = self.here();
            let message = format!("nesting is deeper than {MAX_DEPTH} levels");
            self.diags.push(ParseDiag { kind: ParseDiagKind::NestingTooDeep, span, message });
            return Err(Failed);
        }
        self.depth = self.depth.saturating_add(1);
        let result = f(self);
        self.depth = self.depth.saturating_sub(1);
        result
    }

    /// Clears the per-statement state after a failure.
    fn reset(&mut self) {
        self.depth = 0;
        self.pending_next = false;
        self.inline_if = 0;
    }

    // ----- statement boundaries ----------------------------------------------------------

    fn at_stmt_end(&self) -> bool {
        match self.tok(0) {
            None => true,
            Some(t) => t.kind == TokenKind::Newline || self.is_char(0, b':') || (self.inline_if > 0 && self.is_kw(Keyword::Else)),
        }
    }

    fn expect_stmt_end(&mut self) -> PResult<()> {
        if self.at_stmt_end() { Ok(()) } else { self.fail("the end of the statement") }
    }

    fn skip_separators(&mut self) {
        while self.is_newline() || self.is_char(0, b':') {
            self.bump();
        }
    }

    /// Skips to the end of the current line, unless it sits on a procedure header.
    fn skip_line(&mut self) {
        if self.at_header_line() {
            return;
        }
        while !self.at_eof() && !self.is_newline() {
            self.bump();
        }
    }

    /// `[modifiers] Sub|Function|Property Get|Let|Set` at the current token.
    fn at_proc_header(&self) -> bool {
        let mut n = 0;
        while matches!(self.keyword_at(n), Some(Keyword::Public | Keyword::Private | Keyword::Friend | Keyword::Static | Keyword::Global)) {
            n = n.saturating_add(1);
        }
        match self.keyword_at(n) {
            Some(Keyword::Sub | Keyword::Function) => true,
            Some(Keyword::Property) => matches!(self.keyword_at(n.saturating_add(1)), Some(Keyword::Get | Keyword::Let | Keyword::Set)),
            _ => false,
        }
    }

    fn at_header_line(&self) -> bool {
        self.tok(0).is_some_and(|t| t.line_start) && self.at_proc_header()
    }

    /// `End Sub`, `End Function` or `End Property`.
    fn at_end_proc(&self) -> bool {
        self.is_kw(Keyword::End) && matches!(self.keyword_at(1), Some(Keyword::Sub | Keyword::Function | Keyword::Property))
    }

    /// Skips to just after the next `End Sub|Function|Property`, or to the next procedure header
    /// at the start of a line.
    fn skip_to_proc_end(&mut self) {
        while !self.at_eof() {
            if self.at_header_line() {
                return;
            }
            if self.at_end_proc() {
                self.bump_n(2);
                return;
            }
            self.bump();
        }
    }

    // ----- names --------------------------------------------------------------------------

    /// The name at `n`: an identifier, a `[foreign]` name, a contextual keyword, and with
    /// `builtins` a reserved built-in function name.
    fn name_at(&self, n: usize, builtins: bool) -> Option<&'a str> {
        match self.kind_at(n)? {
            TokenKind::Identifier | TokenKind::ForeignIdentifier => Some(self.text_at(n)),
            TokenKind::Keyword(keyword) if !keyword.is_reserved() || (builtins && BUILTIN_NAMES.contains(&keyword)) => Some(self.text_at(n)),
            _ => None,
        }
    }

    /// Any word at `n` (member names after `.` may be keywords).
    fn word_at(&self, n: usize) -> Option<&'a str> {
        matches!(self.kind_at(n)?, TokenKind::Identifier | TokenKind::ForeignIdentifier | TokenKind::Keyword(_)).then(|| self.text_at(n))
    }

    /// A type suffix right after the previous word.
    fn suffix(&mut self) -> Option<char> {
        if self.kind_at(0) != Some(TokenKind::TypeSuffix) {
            return None;
        }
        let c = self.text_at(0).chars().next();
        self.bump();
        c
    }

    /// A declared name with its optional suffix.
    fn decl_name(&mut self, what: &str) -> PResult<(String, Option<char>)> {
        let Some(name) = self.name_at(0, false) else { return self.fail(what) };
        self.bump();
        Ok((name.to_owned(), self.suffix()))
    }

    /// A type or class name, dots included.
    fn type_name(&mut self) -> PResult<String> {
        let first = match self.name_at(0, false) {
            Some(name) => name,
            None if self.keyword_at(0).is_some_and(|k| TYPE_NAMES.contains(&k)) => self.text_at(0),
            None => return self.fail("a type name"),
        };
        let mut name = first.to_owned();
        self.bump();
        while self.is_char(0, b'.') && self.adjacent(1) {
            let Some(part) = self.word_at(1) else { break };
            name.push('.');
            name.push_str(part);
            self.bump_n(2);
        }
        Ok(name)
    }

    /// A string literal's value.
    fn string_lit(&mut self, what: &str) -> PResult<String> {
        if self.kind_at(0) != Some(TokenKind::String) {
            return self.fail(what);
        }
        let value = unquote(self.text_at(0));
        self.bump();
        Ok(value)
    }

    // ----- module level ------------------------------------------------------------------

    fn module(&mut self) -> Module {
        let span = Span { line: 1, col: 1, start: 0, end: self.src.len() };
        let mut module = Module { attrs: Vec::new(), options: Vec::new(), decls: Vec::new(), procs: Vec::new(), span };
        loop {
            self.skip_separators();
            if self.at_eof() {
                break;
            }
            let before = self.pos;
            if self.module_item(&mut module).is_err() {
                self.reset();
                self.skip_line();
            }
            if self.pos == before {
                self.bump();
            }
        }
        module
    }

    fn module_item(&mut self, module: &mut Module) -> PResult<()> {
        let start = self.here();
        if self.is_char(0, b'#') {
            let span = self.directive();
            module.decls.push(Decl::Directive { text: self.span_text(span), span });
            return Ok(());
        }
        match self.keyword_at(0) {
            Some(Keyword::Attribute) => module.attrs.push(self.attribute()?),
            Some(Keyword::Option) => module.options.push(self.option()?),
            Some(Keyword::Implements) => {
                self.bump();
                let name = self.type_name()?;
                module.decls.push(Decl::Implements { name, span: self.span_from(start) });
            }
            Some(keyword) if DEF_KEYWORDS.contains(&keyword) => module.decls.push(Decl::Def(self.def(keyword)?)),
            _ => return self.declaration_or_proc(module),
        }
        self.expect_stmt_end()
    }

    /// A `#…` conditional-compilation line, skipped to its end.
    fn directive(&mut self) -> Span {
        let start = self.here();
        while !self.at_eof() && !self.is_newline() {
            self.bump();
        }
        self.span_from(start)
    }

    /// `Attribute name = value[, value…]`.
    fn attribute(&mut self) -> PResult<Attribute> {
        let start = self.here();
        self.bump();
        let Some(first) = self.word_at(0) else { return self.fail("an attribute name") };
        let mut name = first.to_owned();
        self.bump();
        while self.is_char(0, b'.') {
            let Some(part) = self.word_at(1) else { break };
            name.push('.');
            name.push_str(part);
            self.bump_n(2);
        }
        self.expect_char(b'=')?;
        let mut values = vec![self.expr()?];
        while self.eat_char(b',') {
            values.push(self.expr()?);
        }
        Ok(Attribute { name, values, span: self.span_from(start) })
    }

    /// `Option Explicit|Base n|Compare mode|Private Module`.
    fn option(&mut self) -> PResult<OptionStmt> {
        let start = self.here();
        self.bump();
        let kind = if self.eat_kw(Keyword::Explicit) {
            OptionKind::Explicit
        } else if self.eat_kw(Keyword::Base) {
            if self.kind_at(0) != Some(TokenKind::Integer) {
                return self.fail("0 or 1");
            }
            let base = self.text_at(0).to_owned();
            self.bump();
            OptionKind::Base(base)
        } else if self.eat_kw(Keyword::Compare) {
            OptionKind::Compare(self.compare_mode()?)
        } else if self.eat_kw(Keyword::Private) {
            self.expect_kw(Keyword::Module)?;
            OptionKind::PrivateModule
        } else {
            return self.fail("Explicit, Base, Compare or Private Module");
        };
        Ok(OptionStmt { kind, span: self.span_from(start) })
    }

    fn compare_mode(&mut self) -> PResult<CompareMode> {
        let mode = if self.is_kw(Keyword::Binary) {
            CompareMode::Binary
        } else if self.is_kw(Keyword::Text) {
            CompareMode::Text
        } else if self.kind_at(0) == Some(TokenKind::Identifier) && self.text_at(0).eq_ignore_ascii_case("Database") {
            CompareMode::Database
        } else {
            return self.fail("Binary, Text or Database");
        };
        self.bump();
        Ok(mode)
    }

    /// `DefInt A-Z, I` and friends.
    fn def(&mut self, keyword: Keyword) -> PResult<DefDecl> {
        let start = self.here();
        self.bump();
        let mut ranges = Vec::new();
        loop {
            let (from, _) = self.decl_name("a letter")?;
            let to = if self.eat_char(b'-') { Some(self.decl_name("a letter")?.0) } else { None };
            ranges.push(DefRange { from, to });
            if !self.eat_char(b',') {
                break;
            }
        }
        Ok(DefDecl { keyword, ranges, span: self.span_from(start) })
    }

    /// `Public`/`Private`/`Friend`/`Global` and `Static`, in any order.
    fn modifiers(&mut self) -> (Option<Visibility>, bool) {
        let (mut visibility, mut is_static) = (None, false);
        loop {
            match self.keyword_at(0) {
                Some(Keyword::Public) => visibility = Some(Visibility::Public),
                Some(Keyword::Private) => visibility = Some(Visibility::Private),
                Some(Keyword::Friend) => visibility = Some(Visibility::Friend),
                Some(Keyword::Global) => visibility = Some(Visibility::Global),
                Some(Keyword::Static) => is_static = true,
                _ => return (visibility, is_static),
            }
            self.bump();
        }
    }

    fn declaration_or_proc(&mut self, module: &mut Module) -> PResult<()> {
        let start = self.here();
        let (visibility, is_static) = self.modifiers();
        let decl = match self.keyword_at(0) {
            Some(Keyword::Sub | Keyword::Function) => {
                self.procedure(start, visibility, is_static, module);
                return Ok(());
            }
            Some(Keyword::Property) if matches!(self.keyword_at(1), Some(Keyword::Get | Keyword::Let | Keyword::Set)) => {
                self.procedure(start, visibility, is_static, module);
                return Ok(());
            }
            Some(Keyword::Const) => Decl::Const(self.const_stmt(start, visibility)?),
            Some(Keyword::Type) => Decl::Type(self.type_decl(start, visibility)?),
            Some(Keyword::Enum) => Decl::Enum(self.enum_decl(start, visibility)?),
            Some(Keyword::Declare) => Decl::Declare(self.declare(start, visibility)?),
            Some(Keyword::Event) => Decl::Event(self.event(start, visibility)?),
            Some(Keyword::Dim) => {
                self.bump();
                Decl::Var(self.var_stmt(start, VarKeyword::Dim)?)
            }
            _ => {
                let keyword = match (visibility, is_static) {
                    (Some(Visibility::Public), _) => VarKeyword::Public,
                    (Some(Visibility::Private), _) => VarKeyword::Private,
                    (Some(Visibility::Global), _) => VarKeyword::Global,
                    (None, true) => VarKeyword::Static,
                    _ => return self.fail("a declaration or procedure"),
                };
                Decl::Var(self.var_stmt(start, keyword)?)
            }
        };
        module.decls.push(decl);
        self.expect_stmt_end()
    }

    // ----- declarations ------------------------------------------------------------------

    /// The declarators after `Dim`, `Static`, `Public`, ….
    fn var_stmt(&mut self, start: Span, keyword: VarKeyword) -> PResult<VarStmt> {
        let mut vars = vec![self.var_decl()?];
        while self.eat_char(b',') {
            vars.push(self.var_decl()?);
        }
        Ok(VarStmt { keyword, vars, span: self.span_from(start) })
    }

    /// `[WithEvents] name[suffix][(bounds)] [As [New] T]`.
    fn var_decl(&mut self) -> PResult<VarDecl> {
        let start = self.here();
        let with_events = self.eat_kw(Keyword::WithEvents);
        let (name, suffix) = self.decl_name("a variable name")?;
        let bounds = if self.is_char(0, b'(') { Some(self.bounds()?) } else { None };
        let ty = if self.eat_kw(Keyword::As) { Some(self.type_ref(false)?) } else { None };
        Ok(VarDecl { name, suffix, with_events, bounds, ty, span: self.span_from(start) })
    }

    /// `([lower To] upper, …)`, possibly empty.
    fn bounds(&mut self) -> PResult<Vec<crate::ast::Bound>> {
        self.expect_char(b'(')?;
        let mut bounds = Vec::new();
        if self.eat_char(b')') {
            return Ok(bounds);
        }
        loop {
            let start = self.here();
            let first = self.expr()?;
            let (lower, upper) = if self.eat_kw(Keyword::To) { (Some(first), self.expr()?) } else { (None, first) };
            bounds.push(crate::ast::Bound { lower, upper, span: self.span_from(start) });
            if !self.eat_char(b',') {
                break;
            }
        }
        self.expect_char(b')')?;
        Ok(bounds)
    }

    /// The type after `As`: `[New] name[.name…] [* len]`, and with `allow_array` a trailing `()`.
    fn type_ref(&mut self, allow_array: bool) -> PResult<TypeRef> {
        let start = self.here();
        let is_new = self.eat_kw(Keyword::New);
        let name = self.type_name()?;
        let fixed_len = if !is_new && self.eat_char(b'*') { Some(Box::new(self.primary()?)) } else { None };
        let is_array = allow_array && self.is_char(0, b'(') && self.is_char(1, b')');
        if is_array {
            self.bump_n(2);
        }
        Ok(TypeRef { name, is_new, fixed_len, is_array, span: self.span_from(start) })
    }

    /// `Const a [As T] = v, …`.
    fn const_stmt(&mut self, start: Span, visibility: Option<Visibility>) -> PResult<ConstStmt> {
        self.expect_kw(Keyword::Const)?;
        let mut consts = Vec::new();
        loop {
            let item = self.here();
            let (name, suffix) = self.decl_name("a constant name")?;
            let ty = if self.eat_kw(Keyword::As) { Some(self.type_ref(false)?) } else { None };
            self.expect_char(b'=')?;
            let value = self.expr()?;
            consts.push(ConstDecl { name, suffix, ty, value, span: self.span_from(item) });
            if !self.eat_char(b',') {
                break;
            }
        }
        Ok(ConstStmt { visibility, consts, span: self.span_from(start) })
    }

    /// The member lines of a `Type` or `Enum` block up to `End <closer>`; a failed line is
    /// skipped and the block carries on.
    fn block_members<T>(&mut self, closer: Keyword, mut member: impl FnMut(&mut Self) -> PResult<T>) -> PResult<Vec<T>> {
        let mut members = Vec::new();
        loop {
            self.skip_separators();
            if self.at_eof() || self.at_header_line() {
                return self.fail(&format!("End {}", closer.as_str()));
            }
            if self.is_kw(Keyword::End) && self.kw_at(1, closer) {
                self.bump_n(2);
                return Ok(members);
            }
            if self.is_char(0, b'#') {
                self.directive();
                continue;
            }
            match member(self).and_then(|m| self.expect_stmt_end().map(|()| m)) {
                Ok(m) => members.push(m),
                Err(Failed) => {
                    self.reset();
                    self.skip_line();
                }
            }
        }
    }

    fn type_decl(&mut self, start: Span, visibility: Option<Visibility>) -> PResult<TypeDecl> {
        self.bump();
        let (name, _) = self.decl_name("a type name")?;
        self.expect_stmt_end()?;
        let members = self.block_members(Keyword::Type, Self::var_decl)?;
        Ok(TypeDecl { visibility, name, members, span: self.span_from(start) })
    }

    fn enum_decl(&mut self, start: Span, visibility: Option<Visibility>) -> PResult<EnumDecl> {
        self.bump();
        let (name, _) = self.decl_name("an enum name")?;
        self.expect_stmt_end()?;
        let members = self.block_members(Keyword::Enum, |p| {
            let item = p.here();
            let (name, _) = p.decl_name("an enum member")?;
            let value = if p.eat_char(b'=') { Some(p.expr()?) } else { None };
            Ok(EnumMember { name, value, span: p.span_from(item) })
        })?;
        Ok(EnumDecl { visibility, name, members, span: self.span_from(start) })
    }

    /// `Declare [PtrSafe] Sub|Function name Lib "x" [Alias "y"] [(params)] [As T]`.
    fn declare(&mut self, start: Span, visibility: Option<Visibility>) -> PResult<DeclareDecl> {
        self.bump();
        let ptr_safe = self.eat_kw(Keyword::PtrSafe);
        let kind = if self.eat_kw(Keyword::Sub) {
            DeclareKind::Sub
        } else if self.eat_kw(Keyword::Function) {
            DeclareKind::Function
        } else {
            return self.fail("Sub or Function");
        };
        let (name, suffix) = self.decl_name("a procedure name")?;
        self.expect_kw(Keyword::Lib)?;
        let lib = self.string_lit("a library name")?;
        let alias = if self.eat_kw(Keyword::Alias) { Some(self.string_lit("an alias")?) } else { None };
        let params = if self.is_char(0, b'(') { self.params()? } else { Vec::new() };
        let return_type = if self.eat_kw(Keyword::As) { Some(self.type_ref(true)?) } else { None };
        Ok(DeclareDecl { visibility, ptr_safe, kind, name, suffix, lib, alias, params, return_type, span: self.span_from(start) })
    }

    fn event(&mut self, start: Span, visibility: Option<Visibility>) -> PResult<EventDecl> {
        self.bump();
        let (name, _) = self.decl_name("an event name")?;
        let params = if self.is_char(0, b'(') { self.params()? } else { Vec::new() };
        Ok(EventDecl { visibility, name, params, span: self.span_from(start) })
    }

    /// `(param, …)`, possibly empty.
    fn params(&mut self) -> PResult<Vec<Param>> {
        self.expect_char(b'(')?;
        let mut params = Vec::new();
        if self.eat_char(b')') {
            return Ok(params);
        }
        loop {
            params.push(self.param()?);
            if !self.eat_char(b',') {
                break;
            }
        }
        self.expect_char(b')')?;
        Ok(params)
    }

    /// `[Optional] [ByVal|ByRef] [ParamArray] name[suffix][()] [As T] [= default]`.
    fn param(&mut self) -> PResult<Param> {
        let start = self.here();
        let optional = self.eat_kw(Keyword::Optional);
        let passing = if self.eat_kw(Keyword::ByVal) {
            Some(Passing::ByVal)
        } else if self.eat_kw(Keyword::ByRef) {
            Some(Passing::ByRef)
        } else {
            None
        };
        let param_array = self.eat_kw(Keyword::ParamArray);
        let (name, suffix) = self.decl_name("a parameter name")?;
        let is_array = self.is_char(0, b'(');
        if is_array {
            self.bump();
            self.expect_char(b')')?;
        }
        let ty = if self.eat_kw(Keyword::As) { Some(self.type_ref(false)?) } else { None };
        let default = if self.eat_char(b'=') { Some(self.expr()?) } else { None };
        Ok(Param { name, suffix, optional, passing, param_array, is_array, ty, default, span: self.span_from(start) })
    }

    // ----- procedures --------------------------------------------------------------------

    /// A procedure from its header to `End …`, recovering inside it (see the module docs).
    fn procedure(&mut self, start: Span, visibility: Option<Visibility>, is_static: bool, module: &mut Module) {
        let mut proc = match self.proc_header(start, visibility, is_static) {
            Ok(proc) => proc,
            Err(Failed) => {
                self.reset();
                self.skip_to_proc_end();
                return;
            }
        };
        self.proc_body(&mut proc);
        proc.span = self.span_from(start);
        module.procs.push(proc);
    }

    fn proc_header(&mut self, start: Span, visibility: Option<Visibility>, is_static: bool) -> PResult<Proc> {
        let kind = match self.keyword_at(0) {
            Some(Keyword::Sub) => ProcKind::Sub,
            Some(Keyword::Function) => ProcKind::Function,
            _ => {
                self.bump();
                match self.keyword_at(0) {
                    Some(Keyword::Get) => ProcKind::PropertyGet,
                    Some(Keyword::Let) => ProcKind::PropertyLet,
                    _ => ProcKind::PropertySet,
                }
            }
        };
        self.bump();
        let (name, suffix) = self.decl_name("a procedure name")?;
        let params = if self.is_char(0, b'(') { self.params()? } else { Vec::new() };
        let return_type = if self.eat_kw(Keyword::As) { Some(self.type_ref(true)?) } else { None };
        self.expect_stmt_end()?;
        Ok(Proc { kind, name, suffix, visibility, is_static, params, return_type, body: Vec::new(), span: start })
    }

    /// The body statements up to the matching `End …`. On a failure, the rest up to the
    /// procedure's end becomes one [`Stmt::Unknown`].
    fn proc_body(&mut self, proc: &mut Proc) {
        let closer = match proc.kind {
            ProcKind::Sub => Keyword::Sub,
            ProcKind::Function => Keyword::Function,
            ProcKind::PropertyGet | ProcKind::PropertyLet | ProcKind::PropertySet => Keyword::Property,
        };
        loop {
            self.skip_separators();
            let start = self.here();
            let result = if self.at_eof() {
                self.fail(&format!("End {}", closer.as_str()))
            } else if self.at_end_proc() && !self.pending_next {
                if self.kw_at(1, closer) {
                    self.bump_n(2);
                    if self.expect_stmt_end().is_err() {
                        self.skip_line();
                    }
                    return;
                }
                self.fail(&format!("End {}", closer.as_str()))
            } else if self.pending_next {
                self.fail("a statement ('Next' has more variables than open 'For' loops)")
            } else {
                self.statement()
            };
            match result {
                Ok(stmt) => proc.body.push(stmt),
                Err(Failed) => {
                    self.reset();
                    self.skip_to_proc_end();
                    proc.body.push(Stmt::Unknown { span: self.span_from(start) });
                    return;
                }
            }
        }
    }
}

/// The end of the source as a span, for diagnostics at the end of input.
fn eof_span(src: &str, last: Option<&Token>) -> Span {
    let len = src.len();
    let Some(last) = last else { return Span { line: 1, col: 1, start: len, end: len } };
    if last.kind == TokenKind::Newline {
        return Span { line: last.span.line.saturating_add(1), col: 1, start: len, end: len };
    }
    let chars = u32::try_from(lexer::text(src, last).chars().count()).unwrap_or(u32::MAX);
    Span { line: last.span.line, col: last.span.col.saturating_add(chars), start: len, end: len }
}

/// The value of a `"…"` string token: quotes removed, `""` turned into `"`.
fn unquote(text: &str) -> String {
    let inner = text.strip_prefix('"').unwrap_or(text);
    let inner = inner.strip_suffix('"').unwrap_or(inner);
    inner.replace("\"\"", "\"")
}

#[cfg(test)]
mod props;
#[cfg(test)]
mod tests;
