//! Statements inside procedures (§5.4).
use super::{PResult, Parser, Term};
use crate::ast::{
    AssignKind, Bound, CaseClause, CaseTest, DoStmt, ExitKind, Expr, FileOp, FileStmt, ForEachStmt, ForStmt, IfBranch, IfStmt, LoopCond,
    OnErrorAction, PrintItem, PrintSep, ReDimVar, ResumeTarget, SelectStmt, Stmt, VarKeyword,
};
use crate::lexer::{Keyword, Span, TokenKind};

impl Parser<'_> {
    /// One statement, including its end check (a label needs none).
    pub(super) fn statement(&mut self) -> PResult<Stmt> {
        if let Some(label) = self.label() {
            return Ok(label);
        }
        let stmt = self.statement_body()?;
        if !self.pending_next {
            self.expect_stmt_end()?;
        }
        Ok(stmt)
    }

    /// `name:` or a line number at the start of a line.
    fn label(&mut self) -> Option<Stmt> {
        let tok = *self.tok(0)?;
        if !tok.line_start {
            return None;
        }
        let start = tok.span;
        if tok.kind == TokenKind::Integer {
            let name = self.text_at(0).to_owned();
            self.bump();
            if self.is_char(0, b':') && !(self.is_char(1, b'=') && self.adjacent(1)) {
                self.bump();
            }
            return Some(Stmt::Label { name, is_line_number: true, span: self.span_from(start) });
        }
        let name = self.name_at(0, false)?;
        if !self.is_char(1, b':') || (self.is_char(2, b'=') && self.adjacent(2)) {
            return None;
        }
        self.bump_n(2);
        Some(Stmt::Label { name: name.to_owned(), is_line_number: false, span: self.span_from(start) })
    }

    fn statement_body(&mut self) -> PResult<Stmt> {
        let start = self.here();
        if self.is_char(0, b'#') {
            let span = self.directive();
            return Ok(Stmt::Directive { text: self.span_text(span), span });
        }
        let Some(keyword) = self.keyword_at(0) else { return self.simple(start) };
        match keyword {
            Keyword::Dim | Keyword::Static => {
                self.bump();
                let kind = if keyword == Keyword::Dim { VarKeyword::Dim } else { VarKeyword::Static };
                Ok(Stmt::Var(self.var_stmt(start, kind)?))
            }
            Keyword::Const => Ok(Stmt::Const(self.const_stmt(start, None)?)),
            Keyword::If => self.nested(|p| p.if_stmt(start)),
            Keyword::Select => self.nested(|p| p.select_stmt(start)),
            Keyword::For => self.nested(|p| p.for_stmt(start)),
            Keyword::Do => self.nested(|p| p.do_stmt(start)),
            Keyword::While => self.nested(|p| p.while_stmt(start)),
            Keyword::With => self.nested(|p| p.with_stmt(start)),
            Keyword::Call => self.call_stmt(start),
            Keyword::Let | Keyword::Set | Keyword::LSet | Keyword::RSet => self.assign_kw(start, keyword),
            Keyword::Attribute => Ok(Stmt::Attribute(self.attribute()?)),
            _ => self.flow_or_file(start, keyword),
        }
    }

    fn flow_or_file(&mut self, start: Span, keyword: Keyword) -> PResult<Stmt> {
        match keyword {
            Keyword::Exit => self.exit_stmt(start),
            Keyword::On => self.on_stmt(start),
            Keyword::Resume => self.resume_stmt(start),
            Keyword::GoTo | Keyword::GoSub => {
                self.bump();
                let label = self.label_name()?;
                let span = self.span_from(start);
                Ok(if keyword == Keyword::GoTo { Stmt::GoTo { label, span } } else { Stmt::GoSub { label, span } })
            }
            Keyword::Return | Keyword::Stop => {
                self.bump();
                let span = self.span_from(start);
                Ok(if keyword == Keyword::Return { Stmt::Return { span } } else { Stmt::Stop { span } })
            }
            Keyword::End => self.end_stmt(start),
            Keyword::ReDim => self.redim_stmt(start),
            Keyword::Erase => self.erase_stmt(start),
            Keyword::RaiseEvent => self.raise_event(start),
            _ => self.file_or_simple(start, keyword),
        }
    }

    fn file_or_simple(&mut self, start: Span, keyword: Keyword) -> PResult<Stmt> {
        let op = match keyword {
            Keyword::Open => return self.open_stmt(start),
            Keyword::Close => return self.close_stmt(start),
            Keyword::Print => FileOp::Print,
            Keyword::Write => FileOp::Write,
            Keyword::Input => FileOp::Input,
            Keyword::LineInput => FileOp::LineInput,
            Keyword::Get => FileOp::Get,
            Keyword::Put => FileOp::Put,
            Keyword::Seek => FileOp::Seek,
            Keyword::Lock => FileOp::Lock,
            Keyword::Unlock => FileOp::Unlock,
            Keyword::Line if self.kw_at(1, Keyword::Input) => {
                self.bump();
                FileOp::LineInput
            }
            Keyword::Width if self.is_char(1, b'#') => FileOp::Width,
            Keyword::Debug if self.is_char(1, b'.') && self.adjacent(1) && self.kw_at(2, Keyword::Print) => {
                self.bump_n(3);
                let items = self.print_items()?;
                return Ok(Stmt::DebugPrint { items, span: self.span_from(start) });
            }
            _ => return self.simple(start),
        };
        self.bump();
        Ok(Stmt::File(self.file_stmt(start, op)?))
    }

    /// An assignment or a call statement without `Call`.
    fn simple(&mut self, start: Span) -> PResult<Stmt> {
        let target = self.l_expr()?;
        if self.eat_char(b'=') {
            let value = self.expr()?;
            return Ok(Stmt::Assign { kind: AssignKind::Implicit, target, value, span: self.span_from(start) });
        }
        let args = if self.at_stmt_end() { Vec::new() } else { self.bare_args()? };
        Ok(Stmt::Call { explicit: false, target, args, span: self.span_from(start) })
    }

    fn assign_kw(&mut self, start: Span, keyword: Keyword) -> PResult<Stmt> {
        self.bump();
        let kind = match keyword {
            Keyword::Let => AssignKind::Let,
            Keyword::Set => AssignKind::Set,
            Keyword::LSet => AssignKind::LSet,
            _ => AssignKind::RSet,
        };
        let target = self.l_expr()?;
        self.expect_char(b'=')?;
        let value = self.expr()?;
        Ok(Stmt::Assign { kind, target, value, span: self.span_from(start) })
    }

    /// `Call target[(args)]`.
    fn call_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        // `Expr` has a `Drop` impl, so its fields are taken rather than moved out.
        let mut called = self.call_target()?;
        let split = match &mut called {
            Expr::CallOrIndex { target, args, span } => Some((std::mem::replace(&mut **target, Expr::Me { span: *span }), std::mem::take(args))),
            _ => None,
        };
        let (target, args) = match split {
            Some(split) => split,
            None => (called, Vec::new()),
        };
        Ok(Stmt::Call { explicit: true, target, args, span: self.span_from(start) })
    }

    /// A statement block up to (not including) one of `terms`.
    pub(super) fn block(&mut self, terms: &[Term]) -> PResult<Vec<Stmt>> {
        let saved = std::mem::take(&mut self.inline_if);
        let result = self.block_inner(terms);
        self.inline_if = saved;
        result
    }

    fn block_inner(&mut self, terms: &[Term]) -> PResult<Vec<Stmt>> {
        let mut body = Vec::new();
        loop {
            self.skip_separators();
            if self.pending_next {
                if terms.contains(&Term::Next) {
                    return Ok(body);
                }
                return self.fail("a statement ('Next' has more variables than open 'For' loops)");
            }
            if self.at_eof() {
                return self.fail(terms.last().map_or("a statement", |&term| term_text(term)));
            }
            if terms.iter().any(|&term| self.at_term(term)) {
                return Ok(body);
            }
            body.push(self.statement()?);
        }
    }

    fn at_term(&self, term: Term) -> bool {
        let end_with = |keyword| self.is_kw(Keyword::End) && self.kw_at(1, keyword);
        match term {
            Term::EndIf => self.is_kw(Keyword::EndIf) || end_with(Keyword::If),
            Term::Else => self.is_kw(Keyword::Else),
            Term::ElseIf => self.is_kw(Keyword::ElseIf),
            Term::Case => self.is_kw(Keyword::Case),
            Term::EndSelect => end_with(Keyword::Select),
            Term::Loop => self.is_kw(Keyword::Loop),
            Term::Next => self.is_kw(Keyword::Next),
            Term::Wend => self.is_kw(Keyword::Wend),
            Term::EndWith => end_with(Keyword::With),
        }
    }

    // ----- If ----------------------------------------------------------------------------

    fn if_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let cond = self.expr()?;
        self.expect_kw(Keyword::Then)?;
        if self.at_eof() || self.is_newline() {
            return self.block_if(start, cond);
        }
        self.inline_if = self.inline_if.saturating_add(1);
        let parts = self.inline_if_parts();
        self.inline_if = self.inline_if.saturating_sub(1);
        let (body, else_body) = parts?;
        let span = self.span_from(start);
        Ok(Stmt::If(IfStmt { branches: vec![IfBranch { cond, body, span }], else_body, single_line: true, span }))
    }

    fn inline_if_parts(&mut self) -> PResult<(Vec<Stmt>, Option<Vec<Stmt>>)> {
        let body = self.inline_stmts()?;
        let else_body = if self.eat_kw(Keyword::Else) { Some(self.inline_stmts()?) } else { None };
        Ok((body, else_body))
    }

    /// `:`-separated statements up to the end of the line or `Else`. A line number right after
    /// `Then` or `Else` is a `GoTo`.
    fn inline_stmts(&mut self) -> PResult<Vec<Stmt>> {
        let mut stmts = Vec::new();
        if self.kind_at(0) == Some(TokenKind::Integer) {
            let start = self.here();
            let label = self.text_at(0).to_owned();
            self.bump();
            stmts.push(Stmt::GoTo { label, span: self.span_from(start) });
        }
        loop {
            while self.eat_char(b':') {}
            if self.at_eof() || self.is_newline() || self.is_kw(Keyword::Else) {
                return Ok(stmts);
            }
            stmts.push(self.statement()?);
        }
    }

    fn block_if(&mut self, start: Span, cond: Expr) -> PResult<Stmt> {
        const TERMS: &[Term] = &[Term::ElseIf, Term::Else, Term::EndIf];
        let body = self.block(TERMS)?;
        let mut branches = vec![IfBranch { cond, body, span: self.span_from(start) }];
        while self.is_kw(Keyword::ElseIf) {
            let branch = self.here();
            self.bump();
            let cond = self.expr()?;
            self.expect_kw(Keyword::Then)?;
            let body = self.block(TERMS)?;
            branches.push(IfBranch { cond, body, span: self.span_from(branch) });
        }
        let else_body = if self.eat_kw(Keyword::Else) { Some(self.block(&[Term::EndIf])?) } else { None };
        if !self.eat_kw(Keyword::EndIf) {
            self.expect_kw(Keyword::End)?;
            self.expect_kw(Keyword::If)?;
        }
        Ok(Stmt::If(IfStmt { branches, else_body, single_line: false, span: self.span_from(start) }))
    }

    // ----- Select Case -------------------------------------------------------------------

    fn select_stmt(&mut self, start: Span) -> PResult<Stmt> {
        const TERMS: &[Term] = &[Term::Case, Term::EndSelect];
        self.bump();
        self.expect_kw(Keyword::Case)?;
        let expr = self.expr()?;
        self.expect_stmt_end()?;
        let (mut cases, mut else_body) = (Vec::new(), None);
        loop {
            self.skip_separators();
            if self.at_term(Term::EndSelect) {
                self.bump_n(2);
                break;
            }
            let clause = self.here();
            if !self.eat_kw(Keyword::Case) {
                return self.fail("Case or End Select");
            }
            if self.eat_kw(Keyword::Else) {
                else_body = Some(self.block(TERMS)?);
                continue;
            }
            let mut tests = vec![self.case_test()?];
            while self.eat_char(b',') {
                tests.push(self.case_test()?);
            }
            self.expect_stmt_end()?;
            let body = self.block(TERMS)?;
            cases.push(CaseClause { tests, body, span: self.span_from(clause) });
        }
        Ok(Stmt::Select(SelectStmt { expr, cases, else_body, span: self.span_from(start) }))
    }

    /// `Is op v`, `op v`, `a To b` or `v`.
    fn case_test(&mut self) -> PResult<CaseTest> {
        let is = self.eat_kw(Keyword::Is);
        match self.binary_op() {
            Some((op, _, len)) if op.is_relational() => {
                self.bump_n(len);
                return Ok(CaseTest::Is { op, value: self.expr()? });
            }
            _ if is => return self.fail("a comparison operator"),
            _ => {}
        }
        let first = self.expr()?;
        if self.eat_kw(Keyword::To) { Ok(CaseTest::Range { from: first, to: self.expr()? }) } else { Ok(CaseTest::Value(first)) }
    }

    // ----- loops -------------------------------------------------------------------------

    fn for_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        if self.eat_kw(Keyword::Each) {
            let var = self.l_expr()?;
            self.expect_kw(Keyword::In)?;
            let collection = self.expr()?;
            self.expect_stmt_end()?;
            let body = self.block(&[Term::Next])?;
            let next_var = self.close_next()?;
            return Ok(Stmt::ForEach(ForEachStmt { var, collection, body, next_var, span: self.span_from(start) }));
        }
        let var = self.l_expr()?;
        self.expect_char(b'=')?;
        let from = self.expr()?;
        self.expect_kw(Keyword::To)?;
        let to = self.expr()?;
        let step = if self.eat_kw(Keyword::Step) { Some(self.expr()?) } else { None };
        self.expect_stmt_end()?;
        let body = self.block(&[Term::Next])?;
        let next_var = self.close_next()?;
        Ok(Stmt::For(ForStmt { var, from, to, step, body, next_var, span: self.span_from(start) }))
    }

    /// `Next [var]`, or the next variable of a `Next a, b` that closed an inner loop. A trailing
    /// `, var` leaves the next outer loop to close at `var`.
    fn close_next(&mut self) -> PResult<Option<Expr>> {
        if self.pending_next {
            self.pending_next = false;
        } else {
            self.expect_kw(Keyword::Next)?;
            if self.at_stmt_end() {
                return Ok(None);
            }
        }
        let var = self.l_expr()?;
        if self.eat_char(b',') {
            self.pending_next = true;
        }
        Ok(Some(var))
    }

    fn do_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let pre = self.loop_cond()?;
        self.expect_stmt_end()?;
        let body = self.block(&[Term::Loop])?;
        self.expect_kw(Keyword::Loop)?;
        let post = self.loop_cond()?;
        Ok(Stmt::Do(DoStmt { pre, post, body, span: self.span_from(start) }))
    }

    fn loop_cond(&mut self) -> PResult<Option<LoopCond>> {
        let until = if self.eat_kw(Keyword::While) {
            false
        } else if self.eat_kw(Keyword::Until) {
            true
        } else {
            return Ok(None);
        };
        Ok(Some(LoopCond { until, cond: self.expr()? }))
    }

    fn while_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let cond = self.expr()?;
        self.expect_stmt_end()?;
        let body = self.block(&[Term::Wend])?;
        self.expect_kw(Keyword::Wend)?;
        Ok(Stmt::While { cond, body, span: self.span_from(start) })
    }

    fn with_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let object = self.expr()?;
        self.expect_stmt_end()?;
        let body = self.block(&[Term::EndWith])?;
        self.bump_n(2);
        Ok(Stmt::With { object, body, span: self.span_from(start) })
    }

    // ----- jumps and errors --------------------------------------------------------------

    fn exit_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let kind = match self.keyword_at(0) {
            Some(Keyword::Sub) => ExitKind::Sub,
            Some(Keyword::Function) => ExitKind::Function,
            Some(Keyword::Property) => ExitKind::Property,
            Some(Keyword::For) => ExitKind::For,
            Some(Keyword::Do) => ExitKind::Do,
            _ => return self.fail("Sub, Function, Property, For or Do"),
        };
        self.bump();
        Ok(Stmt::Exit { kind, span: self.span_from(start) })
    }

    /// A label name or line number.
    fn label_name(&mut self) -> PResult<String> {
        let name = match self.name_at(0, false) {
            Some(name) => name,
            None if self.kind_at(0) == Some(TokenKind::Integer) => self.text_at(0),
            None => return self.fail("a label or line number"),
        };
        self.bump();
        Ok(name.to_owned())
    }

    /// `On Error GoTo …|Resume Next`, or `On expr GoTo|GoSub labels`.
    fn on_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        if !self.eat_kw(Keyword::Error) {
            let expr = self.expr()?;
            let gosub = if self.eat_kw(Keyword::GoSub) {
                true
            } else {
                self.expect_kw(Keyword::GoTo)?;
                false
            };
            let mut labels = vec![self.label_name()?];
            while self.eat_char(b',') {
                labels.push(self.label_name()?);
            }
            return Ok(Stmt::OnJump { expr, gosub, labels, span: self.span_from(start) });
        }
        let action = if self.eat_kw(Keyword::Resume) {
            self.expect_kw(Keyword::Next)?;
            OnErrorAction::ResumeNext
        } else {
            self.expect_kw(Keyword::GoTo)?;
            if self.is_char(0, b'-') && self.kind_at(1) == Some(TokenKind::Integer) && self.text_at(1) == "1" {
                self.bump_n(2);
                OnErrorAction::GoToMinusOne
            } else {
                match self.label_name()? {
                    label if label == "0" => OnErrorAction::GoToZero,
                    label => OnErrorAction::GoTo(label),
                }
            }
        };
        Ok(Stmt::OnError { action, span: self.span_from(start) })
    }

    fn resume_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let target = if self.at_stmt_end() {
            ResumeTarget::Same
        } else if self.eat_kw(Keyword::Next) {
            ResumeTarget::Next
        } else {
            ResumeTarget::Label(self.label_name()?)
        };
        Ok(Stmt::Resume { target, span: self.span_from(start) })
    }

    /// `End` alone; `End If` and friends out of place are errors.
    fn end_stmt(&mut self, start: Span) -> PResult<Stmt> {
        if matches!(
            self.keyword_at(1),
            Some(
                Keyword::If | Keyword::Select | Keyword::With | Keyword::Sub | Keyword::Function | Keyword::Property | Keyword::Type | Keyword::Enum
            )
        ) {
            return self.fail("a statement");
        }
        self.bump();
        Ok(Stmt::End { span: self.span_from(start) })
    }

    // ----- arrays and events -------------------------------------------------------------

    fn redim_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let preserve = self.eat_kw(Keyword::Preserve);
        let mut vars = Vec::new();
        loop {
            let item = self.here();
            let target = self.member_chain()?;
            let bounds: Vec<Bound> = self.bounds()?;
            let ty = if self.eat_kw(Keyword::As) { Some(self.type_ref(false)?) } else { None };
            vars.push(ReDimVar { target, bounds, ty, span: self.span_from(item) });
            if !self.eat_char(b',') {
                break;
            }
        }
        Ok(Stmt::ReDim { preserve, vars, span: self.span_from(start) })
    }

    fn erase_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let mut targets = vec![self.l_expr()?];
        while self.eat_char(b',') {
            targets.push(self.l_expr()?);
        }
        Ok(Stmt::Erase { targets, span: self.span_from(start) })
    }

    fn raise_event(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let (name, _) = self.decl_name("an event name")?;
        let args = if self.is_char(0, b'(') { self.paren_args()? } else { Vec::new() };
        Ok(Stmt::RaiseEvent { name, args, span: self.span_from(start) })
    }

    // ----- printing and files ------------------------------------------------------------

    /// A print list: items separated by `;` or `,`, each part optional.
    fn print_items(&mut self) -> PResult<Vec<PrintItem>> {
        let mut items = Vec::new();
        while !self.at_stmt_end() {
            let start = self.here();
            let expr = if self.is_char(0, b';') || self.is_char(0, b',') { None } else { Some(self.expr()?) };
            let sep = if self.eat_char(b';') {
                Some(PrintSep::Semicolon)
            } else if self.eat_char(b',') {
                Some(PrintSep::Comma)
            } else {
                None
            };
            items.push(PrintItem { expr, sep, span: self.span_from(start) });
            if sep.is_none() {
                break;
            }
        }
        Ok(items)
    }

    /// `[#]expr`.
    fn file_number(&mut self) -> PResult<Expr> {
        self.eat_char(b'#');
        self.expr()
    }

    /// The file statements after their keyword, except `Open` and `Close`.
    fn file_stmt(&mut self, start: Span, op: FileOp) -> PResult<FileStmt> {
        let mut stmt = FileStmt { op, number: None, args: Vec::new(), items: Vec::new(), mode: Vec::new(), span: start };
        if matches!(op, FileOp::Print | FileOp::Write) {
            if self.is_char(0, b'#') {
                stmt.number = Some(self.file_number()?);
                if !self.at_stmt_end() {
                    self.expect_char(b',')?;
                }
            }
            stmt.items = self.print_items()?;
        } else {
            stmt.number = Some(self.file_number()?);
            while self.eat_char(b',') {
                if self.is_char(0, b',') || self.at_stmt_end() {
                    stmt.args.push(None);
                    continue;
                }
                stmt.args.push(Some(self.expr()?));
                if self.eat_kw(Keyword::To) {
                    stmt.args.push(Some(self.expr()?));
                }
            }
        }
        stmt.span = self.span_from(start);
        Ok(stmt)
    }

    /// `Open path [For mode] [Access …] [lock] As [#]n [Len = r]`.
    fn open_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let path = self.expr()?;
        self.eat_kw(Keyword::For);
        let mut mode = Vec::new();
        while !self.is_kw(Keyword::As) {
            let Some(word) = self.word_at(0) else { return self.fail("As") };
            mode.push(word.to_owned());
            self.bump();
        }
        self.bump();
        let number = Some(self.file_number()?);
        let mut args = vec![Some(path)];
        if self.eat_kw(Keyword::Len) {
            self.expect_char(b'=')?;
            args.push(Some(self.expr()?));
        }
        Ok(Stmt::File(FileStmt { op: FileOp::Open, number, args, items: Vec::new(), mode, span: self.span_from(start) }))
    }

    /// `Close [[#]n, …]`.
    fn close_stmt(&mut self, start: Span) -> PResult<Stmt> {
        self.bump();
        let mut args = Vec::new();
        if !self.at_stmt_end() {
            loop {
                args.push(Some(self.file_number()?));
                if !self.eat_char(b',') {
                    break;
                }
            }
        }
        Ok(Stmt::File(FileStmt { op: FileOp::Close, number: None, args, items: Vec::new(), mode: Vec::new(), span: self.span_from(start) }))
    }
}

/// What a block was waiting for, for an end-of-source diagnostic.
const fn term_text(term: Term) -> &'static str {
    match term {
        Term::EndIf | Term::Else | Term::ElseIf => "End If",
        Term::Case | Term::EndSelect => "End Select",
        Term::Loop => "Loop",
        Term::Next => "Next",
        Term::Wend => "Wend",
        Term::EndWith => "End With",
    }
}
