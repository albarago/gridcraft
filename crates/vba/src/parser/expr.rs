//! Expressions (§5.6): precedence climbing over the binary operators, with every nested
//! sub-expression one level of `depth`.
use super::{PResult, Parser, unquote};
use crate::ast::{Arg, BinaryOp, Expr, Literal, NEG_OPERAND, NOT_OPERAND, TYPEOF_OPERAND, UnaryOp};
use crate::lexer::{Keyword, Span, TokenKind};

/// Whether a term may continue with a spaced `.`, `!` or `(`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    /// Inside an expression: `a (1)` indexes `a`.
    Expr,
    /// A statement head: `f (1)` passes `(1)` to `f`, `f .x` passes `.x`.
    Head,
}

impl Parser<'_> {
    /// A full expression (the top level of a statement is not a nesting level).
    pub(super) fn expr(&mut self) -> PResult<Expr> {
        self.bp_expr(0)
    }

    /// Precedence climbing: operators binding at least `min` tightly.
    fn bp_expr(&mut self, min: u8) -> PResult<Expr> {
        let mut left = self.prefix()?;
        while let Some((op, power, len)) = self.binary_op() {
            if power < min {
                break;
            }
            self.bump_n(len);
            let right = self.nested(|p| p.bp_expr(power.saturating_add(1)))?;
            let span = self.span_from(left.span());
            left = Expr::Binary { op, left: Box::new(left), right: Box::new(right), span };
        }
        Ok(left)
    }

    /// The binary operator at the current token: the operator, its binding power and how many
    /// tokens it takes (`<=` is two).
    pub(super) fn binary_op(&self) -> Option<(BinaryOp, u8, usize)> {
        let op = match self.kind_at(0)? {
            TokenKind::Keyword(keyword) => match keyword {
                Keyword::Mod => BinaryOp::Mod,
                Keyword::Like => BinaryOp::Like,
                Keyword::Is => BinaryOp::Is,
                Keyword::And => BinaryOp::And,
                Keyword::Or => BinaryOp::Or,
                Keyword::Xor => BinaryOp::Xor,
                Keyword::Eqv => BinaryOp::Eqv,
                Keyword::Imp => BinaryOp::Imp,
                _ => return None,
            },
            TokenKind::Operator => return self.symbol_op(),
            _ => return None,
        };
        Some((op, op.precedence(), 1))
    }

    fn symbol_op(&self) -> Option<(BinaryOp, u8, usize)> {
        let then = |c| self.is_char(1, c) && self.adjacent(1);
        let (op, len) = match self.char_at(0)? {
            b'^' => (BinaryOp::Pow, 1),
            b'*' => (BinaryOp::Mul, 1),
            b'/' => (BinaryOp::Div, 1),
            b'\\' => (BinaryOp::IntDiv, 1),
            b'+' => (BinaryOp::Add, 1),
            b'-' => (BinaryOp::Sub, 1),
            b'&' => (BinaryOp::Concat, 1),
            b'=' if then(b'<') => (BinaryOp::Le, 2),
            b'=' if then(b'>') => (BinaryOp::Ge, 2),
            b'=' => (BinaryOp::Eq, 1),
            b'<' if then(b'>') => (BinaryOp::Ne, 2),
            b'<' if then(b'=') => (BinaryOp::Le, 2),
            b'<' => (BinaryOp::Lt, 1),
            b'>' if then(b'=') => (BinaryOp::Ge, 2),
            b'>' => (BinaryOp::Gt, 1),
            _ => return None,
        };
        Some((op, op.precedence(), len))
    }

    /// Prefix operators, `New`, `TypeOf … Is`, `AddressOf`, or a term.
    fn prefix(&mut self) -> PResult<Expr> {
        let start = self.here();
        let unary = |p: &mut Self, op, min| -> PResult<Expr> {
            p.bump();
            let operand = p.nested(|p| p.bp_expr(min))?;
            Ok(Expr::Unary { op, operand: Box::new(operand), span: p.span_from(start) })
        };
        if self.is_char(0, b'-') {
            return unary(self, UnaryOp::Neg, NEG_OPERAND);
        }
        match self.keyword_at(0) {
            Some(Keyword::Not) => unary(self, UnaryOp::Not, NOT_OPERAND),
            Some(Keyword::New) => {
                self.bump();
                let class = self.type_name()?;
                Ok(Expr::New { class, span: self.span_from(start) })
            }
            Some(Keyword::TypeOf) => {
                self.bump();
                let object = self.nested(|p| p.bp_expr(TYPEOF_OPERAND))?;
                self.expect_kw(Keyword::Is)?;
                let class = self.type_name()?;
                Ok(Expr::TypeOfIs { object: Box::new(object), class, span: self.span_from(start) })
            }
            Some(Keyword::AddressOf) => {
                self.bump();
                let target = self.nested(Self::member_chain)?;
                Ok(Expr::AddressOf { target: Box::new(target), span: self.span_from(start) })
            }
            _ => self.primary(),
        }
    }

    /// A literal, a parenthesised expression, or a name, `Me` or `With` member followed by
    /// member accesses and argument lists.
    pub(super) fn primary(&mut self) -> PResult<Expr> {
        let start = self.here();
        let literal = match self.kind_at(0) {
            Some(TokenKind::Integer) => Some(Literal::Integer(self.text_at(0).to_owned())),
            Some(TokenKind::Hex) => Some(Literal::Hex(self.text_at(0).to_owned())),
            Some(TokenKind::Octal) => Some(Literal::Octal(self.text_at(0).to_owned())),
            Some(TokenKind::Float) => Some(Literal::Float(self.text_at(0).to_owned())),
            Some(TokenKind::Date) => Some(Literal::Date(self.text_at(0).to_owned())),
            Some(TokenKind::String) => Some(Literal::String(unquote(self.text_at(0)))),
            Some(TokenKind::Keyword(Keyword::True)) => Some(Literal::Boolean(true)),
            Some(TokenKind::Keyword(Keyword::False)) => Some(Literal::Boolean(false)),
            Some(TokenKind::Keyword(Keyword::Nothing)) => Some(Literal::Nothing),
            Some(TokenKind::Keyword(Keyword::Empty)) => Some(Literal::Empty),
            Some(TokenKind::Keyword(Keyword::Null)) => Some(Literal::Null),
            _ => None,
        };
        if let Some(value) = literal {
            self.bump();
            return Ok(Expr::Literal { value, span: start });
        }
        if self.is_char(0, b'(') {
            return self.nested(|p| {
                p.bump();
                let inner = p.bp_expr(0)?;
                p.expect_char(b')')?;
                Ok(Expr::Paren { inner: Box::new(inner), span: p.span_from(start) })
            });
        }
        let head = self.head("an expression")?;
        self.postfix(head, start, Context::Expr)
    }

    /// A name, `Me`, or a `With`-relative `.name`/`!name`.
    fn head(&mut self, expected: &str) -> PResult<Expr> {
        let start = self.here();
        if self.is_char(0, b'.') || self.is_char(0, b'!') {
            let bang = self.is_char(0, b'!');
            self.bump();
            let (name, suffix) = self.member_name()?;
            return Ok(Expr::Member { object: None, bang, name, suffix, span: self.span_from(start) });
        }
        if self.eat_kw(Keyword::Me) {
            return Ok(Expr::Me { span: start });
        }
        let Some(name) = self.name_at(0, true) else { return self.fail(expected) };
        self.bump();
        let suffix = self.suffix();
        Ok(Expr::Ident { name: name.to_owned(), suffix, span: self.span_from(start) })
    }

    /// The word after `.` or `!` (keywords allowed) and its suffix.
    fn member_name(&mut self) -> PResult<(String, Option<char>)> {
        let Some(name) = self.word_at(0) else { return self.fail("a member name") };
        self.bump();
        Ok((name.to_owned(), self.suffix()))
    }

    /// `.name`, `!name` and `(args)` after a term.
    fn postfix(&mut self, mut left: Expr, start: Span, context: Context) -> PResult<Expr> {
        loop {
            let joined = self.tok(0).is_some_and(|t| !t.line_start && (!t.spaced || t.continued || context == Context::Expr));
            if joined && (self.is_char(0, b'.') || self.is_char(0, b'!')) {
                let bang = self.is_char(0, b'!');
                self.bump();
                let (name, suffix) = self.member_name()?;
                left = Expr::Member { object: Some(Box::new(left)), bang, name, suffix, span: self.span_from(start) };
            } else if self.is_char(0, b'(') && (self.adjacent(0) || context == Context::Expr) {
                let args = self.paren_args()?;
                left = Expr::CallOrIndex { target: Box::new(left), args, span: self.span_from(start) };
            } else {
                return Ok(left);
            }
        }
    }

    /// A statement head: an assignment target or a called expression.
    pub(super) fn l_expr(&mut self) -> PResult<Expr> {
        let start = self.here();
        let head = self.head("a statement")?;
        self.postfix(head, start, Context::Head)
    }

    /// The target after `Call`, where `Call f (x)` still passes `x`.
    pub(super) fn call_target(&mut self) -> PResult<Expr> {
        let start = self.here();
        let head = self.head("a procedure name")?;
        self.postfix(head, start, Context::Expr)
    }

    /// A name followed by member accesses only (`ReDim` targets, `AddressOf`).
    pub(super) fn member_chain(&mut self) -> PResult<Expr> {
        let start = self.here();
        let mut left = self.head("a name")?;
        while (self.is_char(0, b'.') || self.is_char(0, b'!')) && self.adjacent(0) {
            let bang = self.is_char(0, b'!');
            self.bump();
            let (name, suffix) = self.member_name()?;
            left = Expr::Member { object: Some(Box::new(left)), bang, name, suffix, span: self.span_from(start) };
        }
        Ok(left)
    }

    /// `(arg, …)`: one nesting level for the whole list.
    pub(super) fn paren_args(&mut self) -> PResult<Vec<Arg>> {
        self.nested(|p| {
            p.expect_char(b'(')?;
            let mut args = Vec::new();
            if p.eat_char(b')') {
                return Ok(args);
            }
            loop {
                args.push(p.arg()?);
                if !p.eat_char(b',') {
                    break;
                }
            }
            p.expect_char(b')')?;
            Ok(args)
        })
    }

    /// The arguments of a call statement without parentheses, up to the end of the statement.
    pub(super) fn bare_args(&mut self) -> PResult<Vec<Arg>> {
        let mut args = vec![self.arg()?];
        while self.eat_char(b',') {
            args.push(self.arg()?);
        }
        Ok(args)
    }

    /// A positional, named (`name:=v`) or missing argument. A `#` before a file number is
    /// dropped.
    fn arg(&mut self) -> PResult<Arg> {
        let start = self.here();
        if self.is_char(0, b',') || self.is_char(0, b')') || self.at_stmt_end() {
            return Ok(Arg { name: None, value: None, span: Span { end: start.start, ..start } });
        }
        let named = self.word_at(0).filter(|_| self.is_char(1, b':') && self.is_char(2, b'=') && self.adjacent(2));
        let name = named.map(str::to_owned);
        if name.is_some() {
            self.bump_n(3);
        } else {
            self.eat_char(b'#');
        }
        let value = self.expr()?;
        Ok(Arg { name, value: Some(value), span: self.span_from(start) })
    }
}
