//! A pretty printer that turns a [`Module`] tree back into VBA source (Req 12.7).
//!
//! The output is canonical, not a copy of the original layout: trivia (comments, blank lines,
//! line continuations, spacing) isn't in the tree, so it can't come back. What the printer
//! guarantees is that [`crate::parser::parse`] reads its output with zero diagnostics and, for a
//! tree the parser built without diagnostics, gives back an equivalent tree (Req 12.8): the
//! same node kinds, identifiers, literal values and child order, ignoring spans, trivia and
//! [`IfStmt::single_line`].
//!
//! # Formatting choices
//!
//! - Lines end in LF (`\n`), including the last one. The parser accepts every VBA line
//!   terminator, and LF keeps the output the same on every platform; a writer that needs CRLF
//!   can convert.
//! - Blocks indent by four spaces per level. `Case` clauses sit one level inside `Select Case`
//!   and their bodies two. Labels and `#` directive lines start at column 1.
//! - Keywords use VBA's canonical capitalisation (`End If`, `ByVal`, `Line Input`). Identifiers,
//!   type names, member names, labels and number and date literals keep the spelling stored in
//!   the tree. Strings are quoted with `"` doubled.
//! - One statement per line. A single-line `If` prints as a block `If … End If`, which the parser
//!   reads back as the same branches with `single_line` false; the round-trip equivalence ignores
//!   that flag. `Next j, i` prints as `Next j` and `Next i` on two lines, which parses to the same
//!   tree.
//! - Module items (attributes, options, declarations, procedures) print in source order, merged
//!   by span start, with a blank line around each procedure. Each kind keeps its own order.
//! - Expressions get the fewest parentheses that keep the tree: [`Expr::Paren`] nodes print as
//!   the parentheses they came from, and the printer adds parentheses of its own only where the
//!   parser would otherwise read a different tree, using [`BinaryOp::precedence`],
//!   left-associativity of every binary operator (`^` included, as in VBA: `a ^ b ^ c` is
//!   `(a ^ b) ^ c`), and the reach of prefix operators: the operand of `-` takes in `^`
//!   (`-a ^ b` is `-(a ^ b)`) and the operand of `Not` takes in comparisons and everything
//!   tighter (`Not a = b` is `Not (a = b)`). A tree the parser built never needs added
//!   parentheses, so they only appear for hand-built trees, where they become `Paren` nodes on
//!   the next parse.
//! - Call statements keep their shape: `Call f(a, b)` when written with `Call` (`Call f` with no
//!   arguments, `Call f(1)()` when the target itself ends in an argument list); otherwise
//!   `f a, b` with a space after the target, so a first argument in parentheses prints as
//!   `f (x), y` and a `With` member as `f .x`, while a parenthesised target prints as `f(x)`.
//!   A missing argument prints as nothing between commas (`f(a, , c)`).
//! - A name that the lexer wouldn't read back as the same identifier (a reserved word, or text
//!   with spaces) prints as a `[foreign]` name. The parser keeps the brackets of a foreign name
//!   in the tree, so this only changes hand-built trees.
//! - A statement head written `Debug` + line continuation + `.Print` (a call of the member, not a
//!   `Debug.Print` statement) prints with the continuation, because `Debug.Print` on one line
//!   would parse as the `Debug.Print` statement.
//!
//! # What can't round-trip
//!
//! - [`Stmt::Unknown`] (source the parser skipped while recovering) prints as the comment
//!   `' (unparsed)`, because the tree doesn't hold the skipped text. Trees with `Unknown` only
//!   come from sources with diagnostics, which the round-trip property excludes (Req 12.8).
//! - `#` directive lines print their stored text verbatim, so they re-parse to the same
//!   directive, but always at the position the tree gives them (a directive between two
//!   procedures stays between them; one between `Type` members was never kept by the parser).
//! - The printer assumes shapes the parser can build. A hand-built tree with, say, a member
//!   access on a literal or an empty `If` has no VBA spelling; the printer writes its best
//!   effort (an `If` with no branches prints only its `Else` body), which may not parse.
//!
//! The printer never panics. It walks flat operator chains (`a + b + …`) and member or call
//! chains (`a.b(1).c…`) in loops, so its recursion depth follows the tree's nesting, which the
//! parser bounds at [`crate::parser::MAX_DEPTH`].
use crate::ast::{
    Arg, AssignKind, Attribute, BinaryOp, Bound, CaseTest, CompareMode, ConstStmt, Decl, DeclareDecl, DeclareKind, ExitKind, Expr, FileOp, FileStmt,
    IfStmt, Literal, LoopCond, Module, NEG_OPERAND, NOT_OPERAND, OnErrorAction, OptionKind, OptionStmt, Param, Passing, PrintItem, PrintSep, Proc,
    ProcKind, ResumeTarget, SelectStmt, Stmt, TYPEOF_OPERAND, TypeRef, UnaryOp, VarDecl, VarKeyword, VarStmt, Visibility,
};
use crate::lexer::{Keyword, Span};
use crate::parser::BUILTIN_NAMES;

/// One indentation level.
const INDENT: &str = "    ";

/// The "tail" of an expression that takes in no following operator (see [`expr`]).
const CLOSED: u8 = u8::MAX;

/// Prints a module as canonical VBA source (see the module docs).
#[must_use]
pub fn print_module(module: &Module) -> String {
    let mut printer = Printer::default();
    printer.module(module);
    printer.out
}

/// Prints one statement (and any block it opens) at indentation level 0.
#[must_use]
pub fn print_stmt(stmt: &Stmt) -> String {
    let mut printer = Printer::default();
    printer.stmt(stmt);
    printer.out
}

/// Prints one expression.
#[must_use]
pub fn print_expr(e: &Expr) -> String {
    let mut out = String::new();
    expr(&mut out, e, 0);
    out
}

#[derive(Default)]
struct Printer {
    out: String,
    level: usize,
}

/// Which module item comes next in [`Printer::module`].
#[derive(Clone, Copy, PartialEq, Eq)]
enum Item {
    Attr,
    Option,
    Decl,
    Proc,
}

impl Printer {
    /// One indented line.
    fn line(&mut self, text: &str) {
        for _ in 0..self.level {
            self.out.push_str(INDENT);
        }
        self.out.push_str(text);
        self.out.push('\n');
    }

    /// One line at column 1.
    fn flush_line(&mut self, text: &str) {
        self.out.push_str(text);
        self.out.push('\n');
    }

    fn block(&mut self, stmts: &[Stmt]) {
        self.level = self.level.saturating_add(1);
        for s in stmts {
            self.stmt(s);
        }
        self.level = self.level.saturating_sub(1);
    }

    // ----- module level ------------------------------------------------------------------

    /// Merges the four item lists by span start (ties keep attributes, options, declarations,
    /// procedures in that order) without reordering any one list.
    fn module(&mut self, m: &Module) {
        let (mut a, mut o, mut d, mut p) = (0_usize, 0_usize, 0_usize, 0_usize);
        let mut after_proc = false;
        loop {
            let heads = [
                (Item::Attr, m.attrs.get(a).map(|x| x.span.start)),
                (Item::Option, m.options.get(o).map(|x| x.span.start)),
                (Item::Decl, m.decls.get(d).map(|x| decl_span(x).start)),
                (Item::Proc, m.procs.get(p).map(|x| x.span.start)),
            ];
            let mut next: Option<(Item, usize)> = None;
            for (item, start) in heads {
                if let Some(start) = start
                    && next.is_none_or(|(_, best)| start < best)
                {
                    next = Some((item, start));
                }
            }
            let Some((item, _)) = next else { break };
            if (item == Item::Proc || after_proc) && !self.out.is_empty() {
                self.out.push('\n');
            }
            after_proc = item == Item::Proc;
            match item {
                Item::Attr => {
                    if let Some(x) = m.attrs.get(a) {
                        self.line(&attribute(x));
                    }
                    a = a.saturating_add(1);
                }
                Item::Option => {
                    if let Some(x) = m.options.get(o) {
                        self.line(&option(x));
                    }
                    o = o.saturating_add(1);
                }
                Item::Decl => {
                    if let Some(x) = m.decls.get(d) {
                        self.decl(x);
                    }
                    d = d.saturating_add(1);
                }
                Item::Proc => {
                    if let Some(x) = m.procs.get(p) {
                        self.proc(x);
                    }
                    p = p.saturating_add(1);
                }
            }
        }
    }

    fn decl(&mut self, d: &Decl) {
        match d {
            Decl::Var(v) => self.line(&var_stmt(v, false)),
            Decl::Const(c) => self.line(&const_stmt(c)),
            Decl::Type(t) => {
                self.line(&format!("{}Type {}", visibility(t.visibility), name(&t.name, false)));
                self.level = self.level.saturating_add(1);
                for member in &t.members {
                    self.line(&var_decl(member));
                }
                self.level = self.level.saturating_sub(1);
                self.line("End Type");
            }
            Decl::Enum(e) => {
                self.line(&format!("{}Enum {}", visibility(e.visibility), name(&e.name, false)));
                self.level = self.level.saturating_add(1);
                for member in &e.members {
                    let mut text = name(&member.name, false);
                    if let Some(value) = &member.value {
                        text.push_str(" = ");
                        expr(&mut text, value, 0);
                    }
                    self.line(&text);
                }
                self.level = self.level.saturating_sub(1);
                self.line("End Enum");
            }
            Decl::Declare(x) => self.line(&declare(x)),
            Decl::Event(e) => self.line(&format!("{}Event {}({})", visibility(e.visibility), name(&e.name, false), params(&e.params))),
            Decl::Implements { name: interface, .. } => self.line(&format!("Implements {interface}")),
            Decl::Def(def) => {
                let ranges: Vec<String> = def
                    .ranges
                    .iter()
                    .map(|r| match &r.to {
                        Some(to) => format!("{}-{}", name(&r.from, false), name(to, false)),
                        None => name(&r.from, false),
                    })
                    .collect();
                self.line(&format!("{} {}", def.keyword.as_str(), ranges.join(", ")));
            }
            Decl::Directive { text, .. } => self.flush_line(text),
        }
    }

    fn proc(&mut self, p: &Proc) {
        let (kind, closer) = match p.kind {
            ProcKind::Sub => ("Sub", "Sub"),
            ProcKind::Function => ("Function", "Function"),
            ProcKind::PropertyGet => ("Property Get", "Property"),
            ProcKind::PropertyLet => ("Property Let", "Property"),
            ProcKind::PropertySet => ("Property Set", "Property"),
        };
        let mut header = format!("{}{}{kind} {}", visibility(p.visibility), if p.is_static { "Static " } else { "" }, name(&p.name, false));
        push_suffix(&mut header, p.suffix);
        header.push('(');
        header.push_str(&params(&p.params));
        header.push(')');
        if let Some(t) = &p.return_type {
            header.push_str(" As ");
            header.push_str(&type_ref(t));
        }
        self.line(&header);
        self.block(&p.body);
        self.line(&format!("End {closer}"));
    }

    // ----- statements --------------------------------------------------------------------

    fn stmt(&mut self, s: &Stmt) {
        match s {
            Stmt::If(i) => self.if_stmt(i),
            Stmt::Select(sel) => self.select(sel),
            Stmt::For(f) => {
                let mut head = format!("For {} = {} To {}", postfix_text(&f.var), expr_text(&f.from), expr_text(&f.to));
                if let Some(step) = &f.step {
                    head.push_str(" Step ");
                    expr(&mut head, step, 0);
                }
                self.line(&head);
                self.block(&f.body);
                self.line(&next(f.next_var.as_ref()));
            }
            Stmt::ForEach(f) => {
                self.line(&format!("For Each {} In {}", postfix_text(&f.var), expr_text(&f.collection)));
                self.block(&f.body);
                self.line(&next(f.next_var.as_ref()));
            }
            Stmt::Do(d) => {
                self.line(&format!("Do{}", loop_cond(d.pre.as_ref())));
                self.block(&d.body);
                self.line(&format!("Loop{}", loop_cond(d.post.as_ref())));
            }
            Stmt::While { cond, body, .. } => {
                self.line(&format!("While {}", expr_text(cond)));
                self.block(body);
                self.line("Wend");
            }
            Stmt::With { object, body, .. } => {
                self.line(&format!("With {}", expr_text(object)));
                self.block(body);
                self.line("End With");
            }
            Stmt::Label { name: label, is_line_number, .. } => {
                if *is_line_number {
                    self.flush_line(label);
                } else {
                    self.flush_line(&format!("{label}:"));
                }
            }
            Stmt::Directive { text, .. } => self.flush_line(text),
            Stmt::Unknown { .. } => self.line("' (unparsed)"),
            Stmt::Assign { kind: AssignKind::Implicit, target, value, .. } => {
                let text = format!("{} = {}", self.head(target), expr_text(value));
                self.line(&text);
            }
            Stmt::Call { explicit: false, target, args, .. } => {
                let mut text = self.head(target);
                if !args.is_empty() {
                    text.push(' ');
                    arg_list(&mut text, args);
                }
                self.line(&text);
            }
            _ => self.line(&simple_stmt(s)),
        }
    }

    /// The target of an assignment or call statement without a keyword. `Debug.Print` on one
    /// line would start a `Debug.Print` statement, so a member `Print` of `Debug` there is
    /// written across a line continuation, as the parser read it.
    fn head(&self, target: &Expr) -> String {
        let text = postfix_text(target);
        let is_debug_print = text.get(..11).is_some_and(|start| start.eq_ignore_ascii_case("Debug.Print"))
            && !text.get(11..).and_then(|rest| rest.chars().next()).is_some_and(|c| c.is_alphanumeric() || c == '_');
        match (is_debug_print, text.get(5..)) {
            (true, Some(rest)) => format!("Debug _\n{}{rest}", INDENT.repeat(self.level.saturating_add(1))),
            _ => text,
        }
    }

    fn if_stmt(&mut self, i: &IfStmt) {
        let mut branches = i.branches.iter();
        let Some(first) = branches.next() else {
            for s in i.else_body.iter().flatten() {
                self.stmt(s);
            }
            return;
        };
        self.line(&format!("If {} Then", expr_text(&first.cond)));
        self.block(&first.body);
        for branch in branches {
            self.line(&format!("ElseIf {} Then", expr_text(&branch.cond)));
            self.block(&branch.body);
        }
        if let Some(body) = &i.else_body {
            self.line("Else");
            self.block(body);
        }
        self.line("End If");
    }

    fn select(&mut self, sel: &SelectStmt) {
        self.line(&format!("Select Case {}", expr_text(&sel.expr)));
        self.level = self.level.saturating_add(1);
        for clause in &sel.cases {
            let tests: Vec<String> = clause.tests.iter().map(case_test).collect();
            self.line(&format!("Case {}", tests.join(", ")));
            self.block(&clause.body);
        }
        if let Some(body) = &sel.else_body {
            self.line("Case Else");
            self.block(body);
        }
        self.level = self.level.saturating_sub(1);
        self.line("End Select");
    }
}

/// The one-line statements.
fn simple_stmt(s: &Stmt) -> String {
    match s {
        Stmt::Assign { kind, target, value, .. } => {
            let keyword = match kind {
                AssignKind::Implicit => "",
                AssignKind::Let => "Let ",
                AssignKind::Set => "Set ",
                AssignKind::LSet => "LSet ",
                AssignKind::RSet => "RSet ",
            };
            format!("{keyword}{} = {}", postfix_text(target), expr_text(value))
        }
        Stmt::Call { target, args, .. } => {
            let mut text = format!("Call {}", postfix_text(target));
            if !args.is_empty() || matches!(target, Expr::CallOrIndex { .. }) {
                text.push('(');
                arg_list(&mut text, args);
                text.push(')');
            }
            text
        }
        Stmt::Var(v) => var_stmt(v, true),
        Stmt::Const(c) => const_stmt(c),
        Stmt::Exit { kind, .. } => {
            let what = match kind {
                ExitKind::Sub => "Sub",
                ExitKind::Function => "Function",
                ExitKind::Property => "Property",
                ExitKind::For => "For",
                ExitKind::Do => "Do",
            };
            format!("Exit {what}")
        }
        Stmt::OnError { action, .. } => match action {
            OnErrorAction::GoTo(label) => format!("On Error GoTo {label}"),
            OnErrorAction::GoToZero => "On Error GoTo 0".to_owned(),
            OnErrorAction::GoToMinusOne => "On Error GoTo -1".to_owned(),
            OnErrorAction::ResumeNext => "On Error Resume Next".to_owned(),
        },
        Stmt::OnJump { expr: selector, gosub, labels, .. } => {
            format!("On {} {} {}", expr_text(selector), if *gosub { "GoSub" } else { "GoTo" }, labels.join(", "))
        }
        Stmt::Resume { target, .. } => match target {
            ResumeTarget::Same => "Resume".to_owned(),
            ResumeTarget::Next => "Resume Next".to_owned(),
            ResumeTarget::Label(label) => format!("Resume {label}"),
        },
        Stmt::GoTo { label, .. } => format!("GoTo {label}"),
        Stmt::GoSub { label, .. } => format!("GoSub {label}"),
        _ => other_stmt(s),
    }
}

/// The rest of the one-line statements.
fn other_stmt(s: &Stmt) -> String {
    match s {
        Stmt::Return { .. } => "Return".to_owned(),
        Stmt::End { .. } => "End".to_owned(),
        Stmt::Stop { .. } => "Stop".to_owned(),
        Stmt::ReDim { preserve, vars, .. } => {
            let vars: Vec<String> = vars
                .iter()
                .map(|v| {
                    let mut text = postfix_text(&v.target);
                    text.push_str(&bounds(&v.bounds));
                    if let Some(t) = &v.ty {
                        text.push_str(" As ");
                        text.push_str(&type_ref(t));
                    }
                    text
                })
                .collect();
            format!("ReDim {}{}", if *preserve { "Preserve " } else { "" }, vars.join(", "))
        }
        Stmt::Erase { targets, .. } => {
            let targets: Vec<String> = targets.iter().map(postfix_text).collect();
            format!("Erase {}", targets.join(", "))
        }
        Stmt::RaiseEvent { name: event, args, .. } => {
            let mut text = format!("RaiseEvent {}", name(event, false));
            if !args.is_empty() {
                text.push('(');
                arg_list(&mut text, args);
                text.push(')');
            }
            text
        }
        Stmt::DebugPrint { items, .. } => {
            if items.is_empty() {
                "Debug.Print".to_owned()
            } else {
                format!("Debug.Print {}", print_items(items))
            }
        }
        Stmt::File(f) => file_stmt(f),
        Stmt::Attribute(a) => attribute(a),
        // Block statements, labels, directives and the rest are printed by `Printer::stmt`.
        _ => String::new(),
    }
}

fn next(var: Option<&Expr>) -> String {
    match var {
        Some(var) => format!("Next {}", postfix_text(var)),
        None => "Next".to_owned(),
    }
}

fn loop_cond(cond: Option<&LoopCond>) -> String {
    match cond {
        Some(c) => format!(" {} {}", if c.until { "Until" } else { "While" }, expr_text(&c.cond)),
        None => String::new(),
    }
}

fn case_test(test: &CaseTest) -> String {
    match test {
        CaseTest::Value(v) => expr_text(v),
        CaseTest::Range { from, to } => format!("{} To {}", expr_text(from), expr_text(to)),
        CaseTest::Is { op, value } => format!("Is {} {}", op_text(*op), expr_text(value)),
    }
}

fn print_items(items: &[PrintItem]) -> String {
    let parts: Vec<String> = items
        .iter()
        .map(|item| {
            let mut text = item.expr.as_ref().map(expr_text).unwrap_or_default();
            match item.sep {
                Some(PrintSep::Semicolon) => text.push(';'),
                Some(PrintSep::Comma) => text.push(','),
                None => {}
            }
            text
        })
        .collect();
    parts.join(" ")
}

/// `#n`, or nothing when the tree has no number.
fn file_number(number: Option<&Expr>) -> String {
    number.map(|n| format!("#{}", expr_text(n))).unwrap_or_default()
}

fn file_stmt(f: &FileStmt) -> String {
    let number = file_number(f.number.as_ref());
    match f.op {
        FileOp::Open => {
            let mut text = format!("Open {}", f.args.first().and_then(Option::as_ref).map(expr_text).unwrap_or_default());
            if !f.mode.is_empty() {
                text.push_str(" For ");
                text.push_str(&f.mode.join(" "));
            }
            text.push_str(" As ");
            text.push_str(&number);
            if let Some(Some(len)) = f.args.get(1) {
                text.push_str(" Len = ");
                expr(&mut text, len, 0);
            }
            text
        }
        FileOp::Close => {
            let numbers: Vec<String> = f.args.iter().flatten().map(|n| file_number(Some(n))).collect();
            if numbers.is_empty() { "Close".to_owned() } else { format!("Close {}", numbers.join(", ")) }
        }
        FileOp::Print | FileOp::Write => {
            let keyword = if f.op == FileOp::Print { "Print" } else { "Write" };
            match (f.number.is_some(), f.items.is_empty()) {
                (true, true) => format!("{keyword} {number}"),
                (true, false) => format!("{keyword} {number}, {}", print_items(&f.items)),
                (false, true) => keyword.to_owned(),
                (false, false) => format!("{keyword} {}", print_items(&f.items)),
            }
        }
        op => {
            let keyword = match op {
                FileOp::Input => "Input",
                FileOp::LineInput => "Line Input",
                FileOp::Get => "Get",
                FileOp::Put => "Put",
                FileOp::Seek => "Seek",
                FileOp::Lock => "Lock",
                FileOp::Unlock => "Unlock",
                _ => "Width",
            };
            let mut text = format!("{keyword} {number}");
            for arg in &f.args {
                text.push_str(", ");
                if let Some(arg) = arg {
                    expr(&mut text, arg, 0);
                }
            }
            text
        }
    }
}

// ----- declarations ----------------------------------------------------------------------

fn decl_span(d: &Decl) -> Span {
    match d {
        Decl::Var(v) => v.span,
        Decl::Const(c) => c.span,
        Decl::Type(t) => t.span,
        Decl::Enum(e) => e.span,
        Decl::Declare(x) => x.span,
        Decl::Event(e) => e.span,
        Decl::Def(def) => def.span,
        Decl::Implements { span, .. } | Decl::Directive { span, .. } => *span,
    }
}

fn attribute(a: &Attribute) -> String {
    let values: Vec<String> = a.values.iter().map(expr_text).collect();
    format!("Attribute {} = {}", a.name, values.join(", "))
}

fn option(o: &OptionStmt) -> String {
    match &o.kind {
        OptionKind::Explicit => "Option Explicit".to_owned(),
        OptionKind::Base(base) => format!("Option Base {base}"),
        OptionKind::Compare(mode) => {
            let mode = match mode {
                CompareMode::Binary => "Binary",
                CompareMode::Text => "Text",
                CompareMode::Database => "Database",
            };
            format!("Option Compare {mode}")
        }
        OptionKind::PrivateModule => "Option Private Module".to_owned(),
    }
}

const fn visibility(v: Option<Visibility>) -> &'static str {
    match v {
        Some(Visibility::Public) => "Public ",
        Some(Visibility::Private) => "Private ",
        Some(Visibility::Friend) => "Friend ",
        Some(Visibility::Global) => "Global ",
        None => "",
    }
}

/// `Dim …` and friends. Inside a procedure only `Dim` and `Static` exist, so the module-level
/// keywords print as `Dim` there.
fn var_stmt(v: &VarStmt, in_proc: bool) -> String {
    let keyword = match v.keyword {
        VarKeyword::Static => "Static",
        VarKeyword::Public if !in_proc => "Public",
        VarKeyword::Private if !in_proc => "Private",
        VarKeyword::Global if !in_proc => "Global",
        _ => "Dim",
    };
    let vars: Vec<String> = v.vars.iter().map(var_decl).collect();
    format!("{keyword} {}", vars.join(", "))
}

fn var_decl(v: &VarDecl) -> String {
    let mut text = if v.with_events { "WithEvents ".to_owned() } else { String::new() };
    text.push_str(&name(&v.name, false));
    push_suffix(&mut text, v.suffix);
    if let Some(b) = &v.bounds {
        text.push_str(&bounds(b));
    }
    if let Some(t) = &v.ty {
        text.push_str(" As ");
        text.push_str(&type_ref(t));
    }
    text
}

fn bounds(bounds: &[Bound]) -> String {
    let dims: Vec<String> = bounds
        .iter()
        .map(|b| match &b.lower {
            Some(lower) => format!("{} To {}", expr_text(lower), expr_text(&b.upper)),
            None => expr_text(&b.upper),
        })
        .collect();
    format!("({})", dims.join(", "))
}

fn type_ref(t: &TypeRef) -> String {
    let mut text = if t.is_new { "New ".to_owned() } else { String::new() };
    text.push_str(&t.name);
    if let Some(len) = &t.fixed_len {
        text.push_str(" * ");
        // The parser reads the length as a single term.
        if matches!(
            **len,
            Expr::Literal { .. } | Expr::Ident { .. } | Expr::Me { .. } | Expr::Member { .. } | Expr::CallOrIndex { .. } | Expr::Paren { .. }
        ) {
            expr(&mut text, len, 0);
        } else {
            text.push('(');
            expr(&mut text, len, 0);
            text.push(')');
        }
    }
    if t.is_array {
        text.push_str("()");
    }
    text
}

fn const_stmt(c: &ConstStmt) -> String {
    let consts: Vec<String> = c
        .consts
        .iter()
        .map(|item| {
            let mut text = name(&item.name, false);
            push_suffix(&mut text, item.suffix);
            if let Some(t) = &item.ty {
                text.push_str(" As ");
                text.push_str(&type_ref(t));
            }
            text.push_str(" = ");
            expr(&mut text, &item.value, 0);
            text
        })
        .collect();
    format!("{}Const {}", visibility(c.visibility), consts.join(", "))
}

fn params(params: &[Param]) -> String {
    let params: Vec<String> = params
        .iter()
        .map(|p| {
            let mut text = String::new();
            if p.optional {
                text.push_str("Optional ");
            }
            match p.passing {
                Some(Passing::ByVal) => text.push_str("ByVal "),
                Some(Passing::ByRef) => text.push_str("ByRef "),
                None => {}
            }
            if p.param_array {
                text.push_str("ParamArray ");
            }
            text.push_str(&name(&p.name, false));
            push_suffix(&mut text, p.suffix);
            if p.is_array {
                text.push_str("()");
            }
            if let Some(t) = &p.ty {
                text.push_str(" As ");
                text.push_str(&type_ref(t));
            }
            if let Some(default) = &p.default {
                text.push_str(" = ");
                expr(&mut text, default, 0);
            }
            text
        })
        .collect();
    params.join(", ")
}

fn declare(d: &DeclareDecl) -> String {
    let mut text = format!(
        "{}Declare {}{} {}",
        visibility(d.visibility),
        if d.ptr_safe { "PtrSafe " } else { "" },
        if d.kind == DeclareKind::Sub { "Sub" } else { "Function" },
        name(&d.name, false)
    );
    push_suffix(&mut text, d.suffix);
    text.push_str(" Lib ");
    push_string(&mut text, &d.lib);
    if let Some(alias) = &d.alias {
        text.push_str(" Alias ");
        push_string(&mut text, alias);
    }
    text.push_str(" (");
    text.push_str(&params(&d.params));
    text.push(')');
    if let Some(t) = &d.return_type {
        text.push_str(" As ");
        text.push_str(&type_ref(t));
    }
    text
}

// ----- names and literals ----------------------------------------------------------------

/// `n` as written, or as a `[foreign]` name when the lexer and parser wouldn't read it back as
/// this identifier. `builtins` allows the reserved built-in names (`Len`, `CStr`, …) that the
/// parser accepts where an expression names something.
fn name(n: &str, builtins: bool) -> String {
    if needs_brackets(n, builtins) { format!("[{n}]") } else { n.to_owned() }
}

fn needs_brackets(n: &str, builtins: bool) -> bool {
    if let Some(inner) = n.strip_prefix('[').and_then(|rest| rest.strip_suffix(']')) {
        return inner.is_empty() || inner.contains(']');
    }
    let mut chars = n.chars();
    let plain = chars.next().is_some_and(char::is_alphabetic) && chars.all(|c| c.is_alphanumeric() || c == '_');
    if !plain {
        return true;
    }
    Keyword::lookup(n).is_some_and(|k| k.is_reserved() && !(builtins && BUILTIN_NAMES.contains(&k)))
}

fn push_suffix(out: &mut String, suffix: Option<char>) {
    if let Some(c) = suffix {
        out.push(c);
    }
}

/// A `"…"` literal with `"` doubled.
fn push_string(out: &mut String, s: &str) {
    out.push('"');
    out.push_str(&s.replace('"', "\"\""));
    out.push('"');
}

fn literal(out: &mut String, value: &Literal) {
    match value {
        Literal::Integer(text) | Literal::Hex(text) | Literal::Octal(text) | Literal::Float(text) | Literal::Date(text) => out.push_str(text),
        Literal::String(s) => push_string(out, s),
        Literal::Boolean(true) => out.push_str("True"),
        Literal::Boolean(false) => out.push_str("False"),
        Literal::Nothing => out.push_str("Nothing"),
        Literal::Empty => out.push_str("Empty"),
        Literal::Null => out.push_str("Null"),
    }
}

// ----- expressions -----------------------------------------------------------------------

const fn op_text(op: BinaryOp) -> &'static str {
    match op {
        BinaryOp::Pow => "^",
        BinaryOp::Mul => "*",
        BinaryOp::Div => "/",
        BinaryOp::IntDiv => "\\",
        BinaryOp::Mod => "Mod",
        BinaryOp::Add => "+",
        BinaryOp::Sub => "-",
        BinaryOp::Concat => "&",
        BinaryOp::Eq => "=",
        BinaryOp::Ne => "<>",
        BinaryOp::Lt => "<",
        BinaryOp::Gt => ">",
        BinaryOp::Le => "<=",
        BinaryOp::Ge => ">=",
        BinaryOp::Like => "Like",
        BinaryOp::Is => "Is",
        BinaryOp::And => "And",
        BinaryOp::Or => "Or",
        BinaryOp::Xor => "Xor",
        BinaryOp::Eqv => "Eqv",
        BinaryOp::Imp => "Imp",
    }
}

fn expr_text(e: &Expr) -> String {
    let mut out = String::new();
    expr(&mut out, e, 0);
    out
}

fn postfix_text(e: &Expr) -> String {
    let mut out = String::new();
    postfix(&mut out, e);
    out
}

/// Writes `e` where the parser reads operators binding at least `min` tightly, and returns its
/// tail: the lowest binding power of a following operator that a prefix operator at the end of
/// `e` would take in (`-x` takes in `^`, `Not x` comparisons and tighter), or [`CLOSED`].
fn expr(out: &mut String, e: &Expr, min: u8) -> u8 {
    match e {
        Expr::Binary { op, .. } if op.precedence() < min => {
            out.push('(');
            binary(out, e);
            out.push(')');
            CLOSED
        }
        Expr::Binary { .. } => binary(out, e),
        Expr::Unary { op, operand, .. } => {
            let operand_min = match op {
                UnaryOp::Neg => {
                    out.push('-');
                    NEG_OPERAND
                }
                UnaryOp::Not => {
                    out.push_str("Not ");
                    NOT_OPERAND
                }
            };
            expr(out, operand, operand_min).min(operand_min)
        }
        Expr::TypeOfIs { object, class, .. } => {
            out.push_str("TypeOf ");
            // `Is` must not be taken in by a `Not` at the end of the object.
            let start = out.len();
            if expr(out, object, TYPEOF_OPERAND) <= BinaryOp::Is.precedence() {
                wrap_from(out, start);
            }
            out.push_str(" Is ");
            out.push_str(class);
            CLOSED
        }
        Expr::AddressOf { target, .. } => {
            out.push_str("AddressOf ");
            postfix(out, target);
            CLOSED
        }
        Expr::New { class, .. } => {
            out.push_str("New ");
            out.push_str(class);
            CLOSED
        }
        Expr::Literal { value, .. } => {
            literal(out, value);
            CLOSED
        }
        Expr::Paren { inner, .. } => {
            out.push('(');
            expr(out, inner, 0);
            out.push(')');
            CLOSED
        }
        Expr::Ident { .. } | Expr::Me { .. } | Expr::Member { .. } | Expr::CallOrIndex { .. } => {
            postfix(out, e);
            CLOSED
        }
    }
}

/// Puts parentheses around `out[start..]`.
fn wrap_from(out: &mut String, start: usize) {
    let inner = out.split_off(start.min(out.len()));
    out.push('(');
    out.push_str(&inner);
    out.push(')');
}

/// A binary expression without parentheses of its own. The left spine of operators that bind
/// at least as tightly as their parent (`a + b - c …`) is walked in a loop, as the parser
/// builds it, so long chains don't recurse.
fn binary(out: &mut String, e: &Expr) -> u8 {
    let mut spine: Vec<(BinaryOp, &Expr)> = Vec::new();
    let mut cur = e;
    let base = loop {
        match cur {
            Expr::Binary { op, left, right, .. } => {
                spine.push((*op, right));
                if matches!(&**left, Expr::Binary { op: inner, .. } if inner.precedence() >= op.precedence()) {
                    cur = left;
                } else {
                    break &**left;
                }
            }
            other => break other,
        }
    };
    let start = out.len();
    let mut tail = expr(out, base, spine.last().map_or(0, |(op, _)| op.precedence()));
    for (op, right) in spine.iter().rev() {
        let power = op.precedence();
        if tail <= power {
            // A prefix operator at the end of the left side would take in `op`.
            wrap_from(out, start);
        }
        out.push(' ');
        out.push_str(op_text(*op));
        out.push(' ');
        tail = expr(out, right, power.saturating_add(1));
    }
    tail
}

/// A name, `Me` or `With` member followed by member accesses and argument lists, walked in a
/// loop so long chains don't recurse. Nothing may sit between the links: a space before `(` or
/// `.` means something else at the start of a statement.
fn postfix(out: &mut String, e: &Expr) {
    let mut links: Vec<&Expr> = Vec::new();
    let mut cur = e;
    loop {
        match cur {
            Expr::Member { object: Some(object), .. } => {
                links.push(cur);
                cur = object;
            }
            Expr::CallOrIndex { target, .. } => {
                links.push(cur);
                cur = target;
            }
            _ => break,
        }
    }
    match cur {
        Expr::Ident { name: n, suffix, .. } => {
            out.push_str(&name(n, true));
            push_suffix(out, *suffix);
        }
        Expr::Me { .. } => out.push_str("Me"),
        Expr::Member { object: None, bang, name: n, suffix, .. } => member(out, *bang, n, *suffix),
        Expr::Paren { .. } => {
            expr(out, cur, 0);
        }
        other => {
            // No VBA spelling extends this with members or arguments; best effort.
            out.push('(');
            expr(out, other, 0);
            out.push(')');
        }
    }
    for link in links.iter().rev() {
        match link {
            Expr::Member { bang, name: n, suffix, .. } => member(out, *bang, n, *suffix),
            Expr::CallOrIndex { args, .. } => {
                out.push('(');
                arg_list(out, args);
                out.push(')');
            }
            _ => {}
        }
    }
}

fn member(out: &mut String, bang: bool, n: &str, suffix: Option<char>) {
    out.push(if bang { '!' } else { '.' });
    out.push_str(n);
    push_suffix(out, suffix);
}

/// `a, name:=b, , c`: a missing argument is nothing between commas.
fn arg_list(out: &mut String, args: &[Arg]) {
    for (i, arg) in args.iter().enumerate() {
        if i > 0 {
            out.push_str(", ");
        }
        if let Some(n) = &arg.name {
            out.push_str(n);
            out.push_str(":=");
        }
        if let Some(value) = &arg.value {
            expr(out, value, 0);
        }
    }
}

#[cfg(test)]
mod tests;
