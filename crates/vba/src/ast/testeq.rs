//! Test-only tree equivalence for the print/parse round trip (Req 12.8, design property P6).
//!
//! Two trees are equivalent when they have the same node kinds, the same identifiers compared
//! without regard to case, the same literal values and the same child order. Spans are ignored,
//! and so is [`IfStmt::single_line`], because the printer writes every `If` as a block `If`.
//! [`normalized`] makes that concrete: it zeroes every span, lowercases every identifier and
//! clears `single_line`, so plain `==` on two normalized trees is the equivalence. String
//! literals, `Lib`/`Alias` strings, number and date literal text and directive text are kept
//! as they are.
use super::{
    Arg, Attribute, Bound, CaseTest, ConstStmt, Decl, Expr, FileStmt, IfStmt, Module, OnErrorAction, Param, PrintItem, Proc, ResumeTarget, Stmt,
    TypeRef, VarDecl, VarStmt,
};
use crate::lexer::Span;

const ZERO: Span = Span { line: 0, col: 0, start: 0, end: 0 };

/// `m` with every span zeroed, every identifier lowercased and every `If` marked as a block.
pub(crate) fn normalized(m: &Module) -> Module {
    let mut m = m.clone();
    m.span = ZERO;
    for attr in &mut m.attrs {
        attribute(attr);
    }
    for option in &mut m.options {
        option.span = ZERO;
    }
    for d in &mut m.decls {
        decl(d);
    }
    for p in &mut m.procs {
        proc(p);
    }
    m
}

/// `true` when `a` and `b` are equivalent (see the module docs).
pub(crate) fn equivalent(a: &Module, b: &Module) -> bool {
    normalized(a) == normalized(b)
}

fn name(s: &mut String) {
    *s = s.to_lowercase();
}

fn attribute(a: &mut Attribute) {
    name(&mut a.name);
    a.values.iter_mut().for_each(expr);
    a.span = ZERO;
}

fn type_ref(t: &mut TypeRef) {
    name(&mut t.name);
    if let Some(len) = &mut t.fixed_len {
        expr(len);
    }
    t.span = ZERO;
}

fn bounds(bs: &mut [Bound]) {
    for b in bs {
        if let Some(lower) = &mut b.lower {
            expr(lower);
        }
        expr(&mut b.upper);
        b.span = ZERO;
    }
}

fn var_decl(v: &mut VarDecl) {
    name(&mut v.name);
    if let Some(bs) = &mut v.bounds {
        bounds(bs);
    }
    if let Some(t) = &mut v.ty {
        type_ref(t);
    }
    v.span = ZERO;
}

fn var_stmt(v: &mut VarStmt) {
    v.vars.iter_mut().for_each(var_decl);
    v.span = ZERO;
}

fn const_stmt(c: &mut ConstStmt) {
    for item in &mut c.consts {
        name(&mut item.name);
        if let Some(t) = &mut item.ty {
            type_ref(t);
        }
        expr(&mut item.value);
        item.span = ZERO;
    }
    c.span = ZERO;
}

fn params(ps: &mut [Param]) {
    for p in ps {
        name(&mut p.name);
        if let Some(t) = &mut p.ty {
            type_ref(t);
        }
        if let Some(d) = &mut p.default {
            expr(d);
        }
        p.span = ZERO;
    }
}

fn decl(d: &mut Decl) {
    match d {
        Decl::Var(v) => var_stmt(v),
        Decl::Const(c) => const_stmt(c),
        Decl::Type(t) => {
            name(&mut t.name);
            t.members.iter_mut().for_each(var_decl);
            t.span = ZERO;
        }
        Decl::Enum(e) => {
            name(&mut e.name);
            for m in &mut e.members {
                name(&mut m.name);
                if let Some(v) = &mut m.value {
                    expr(v);
                }
                m.span = ZERO;
            }
            e.span = ZERO;
        }
        Decl::Declare(d) => {
            name(&mut d.name);
            params(&mut d.params);
            if let Some(t) = &mut d.return_type {
                type_ref(t);
            }
            d.span = ZERO;
        }
        Decl::Event(e) => {
            name(&mut e.name);
            params(&mut e.params);
            e.span = ZERO;
        }
        Decl::Implements { name: n, span } => {
            name(n);
            *span = ZERO;
        }
        Decl::Def(def) => {
            for r in &mut def.ranges {
                name(&mut r.from);
                if let Some(to) = &mut r.to {
                    name(to);
                }
            }
            def.span = ZERO;
        }
        Decl::Directive { span, .. } => *span = ZERO,
    }
}

fn proc(p: &mut Proc) {
    name(&mut p.name);
    params(&mut p.params);
    if let Some(t) = &mut p.return_type {
        type_ref(t);
    }
    body(&mut p.body);
    p.span = ZERO;
}

fn body(stmts: &mut [Stmt]) {
    stmts.iter_mut().for_each(stmt);
}

fn args(args: &mut [Arg]) {
    for a in args {
        if let Some(n) = &mut a.name {
            name(n);
        }
        if let Some(v) = &mut a.value {
            expr(v);
        }
        a.span = ZERO;
    }
}

fn print_items(items: &mut [PrintItem]) {
    for item in items {
        if let Some(e) = &mut item.expr {
            expr(e);
        }
        item.span = ZERO;
    }
}

fn if_stmt(i: &mut IfStmt) {
    for b in &mut i.branches {
        expr(&mut b.cond);
        body(&mut b.body);
        b.span = ZERO;
    }
    if let Some(e) = &mut i.else_body {
        body(e);
    }
    i.single_line = false;
    i.span = ZERO;
}

fn file(f: &mut FileStmt) {
    if let Some(n) = &mut f.number {
        expr(n);
    }
    f.args.iter_mut().flatten().for_each(expr);
    print_items(&mut f.items);
    f.mode.iter_mut().for_each(name);
    f.span = ZERO;
}

fn stmt(s: &mut Stmt) {
    match s {
        Stmt::Assign { target, value, span, .. } => {
            expr(target);
            expr(value);
            *span = ZERO;
        }
        Stmt::Call { target, args: a, span, .. } => {
            expr(target);
            args(a);
            *span = ZERO;
        }
        Stmt::Var(v) => var_stmt(v),
        Stmt::Const(c) => const_stmt(c),
        Stmt::If(i) => if_stmt(i),
        Stmt::Select(sel) => {
            expr(&mut sel.expr);
            for c in &mut sel.cases {
                for t in &mut c.tests {
                    match t {
                        CaseTest::Value(v) | CaseTest::Is { value: v, .. } => expr(v),
                        CaseTest::Range { from, to } => {
                            expr(from);
                            expr(to);
                        }
                    }
                }
                body(&mut c.body);
                c.span = ZERO;
            }
            if let Some(e) = &mut sel.else_body {
                body(e);
            }
            sel.span = ZERO;
        }
        _ => loop_or_simple(s),
    }
}

fn loop_or_simple(s: &mut Stmt) {
    match s {
        Stmt::For(f) => {
            for e in [&mut f.var, &mut f.from, &mut f.to] {
                expr(e);
            }
            for e in [&mut f.step, &mut f.next_var].into_iter().flatten() {
                expr(e);
            }
            body(&mut f.body);
            f.span = ZERO;
        }
        Stmt::ForEach(f) => {
            expr(&mut f.var);
            expr(&mut f.collection);
            if let Some(v) = &mut f.next_var {
                expr(v);
            }
            body(&mut f.body);
            f.span = ZERO;
        }
        Stmt::Do(d) => {
            for c in [&mut d.pre, &mut d.post].into_iter().flatten() {
                expr(&mut c.cond);
            }
            body(&mut d.body);
            d.span = ZERO;
        }
        Stmt::While { cond: e, body: b, span } | Stmt::With { object: e, body: b, span } => {
            expr(e);
            body(b);
            *span = ZERO;
        }
        _ => simple(s),
    }
}

fn simple(s: &mut Stmt) {
    match s {
        Stmt::OnError { action, span } => {
            if let OnErrorAction::GoTo(l) = action {
                name(l);
            }
            *span = ZERO;
        }
        Stmt::OnJump { expr: e, labels, span, .. } => {
            expr(e);
            labels.iter_mut().for_each(name);
            *span = ZERO;
        }
        Stmt::Resume { target, span } => {
            if let ResumeTarget::Label(l) = target {
                name(l);
            }
            *span = ZERO;
        }
        Stmt::GoTo { label, span } | Stmt::GoSub { label, span } | Stmt::Label { name: label, span, .. } => {
            name(label);
            *span = ZERO;
        }
        Stmt::ReDim { vars, span, .. } => {
            for v in vars {
                expr(&mut v.target);
                bounds(&mut v.bounds);
                if let Some(t) = &mut v.ty {
                    type_ref(t);
                }
                v.span = ZERO;
            }
            *span = ZERO;
        }
        Stmt::Erase { targets, span } => {
            targets.iter_mut().for_each(expr);
            *span = ZERO;
        }
        Stmt::RaiseEvent { name: n, args: a, span } => {
            name(n);
            args(a);
            *span = ZERO;
        }
        Stmt::DebugPrint { items, span } => {
            print_items(items);
            *span = ZERO;
        }
        Stmt::File(f) => file(f),
        Stmt::Attribute(a) => attribute(a),
        Stmt::Exit { span, .. }
        | Stmt::Return { span }
        | Stmt::End { span }
        | Stmt::Stop { span }
        | Stmt::Directive { span, .. }
        | Stmt::Unknown { span } => *span = ZERO,
        _ => {}
    }
}

fn expr(e: &mut Expr) {
    match e {
        Expr::Literal { span, .. } | Expr::Me { span } => *span = ZERO,
        Expr::Ident { name: n, span, .. } => {
            name(n);
            *span = ZERO;
        }
        Expr::Member { object, name: n, span, .. } => {
            if let Some(o) = object {
                expr(o);
            }
            name(n);
            *span = ZERO;
        }
        Expr::CallOrIndex { target, args: a, span } => {
            expr(target);
            args(a);
            *span = ZERO;
        }
        Expr::Paren { inner: x, span } | Expr::Unary { operand: x, span, .. } | Expr::AddressOf { target: x, span } => {
            expr(x);
            *span = ZERO;
        }
        Expr::Binary { left, right, span, .. } => {
            expr(left);
            expr(right);
            *span = ZERO;
        }
        Expr::New { class, span } => {
            name(class);
            *span = ZERO;
        }
        Expr::TypeOfIs { object, class, span } => {
            expr(object);
            name(class);
            *span = ZERO;
        }
    }
}
