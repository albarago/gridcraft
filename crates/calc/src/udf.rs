//! Which function calls in a formula are calls to user-defined (VBA) functions (Req 17.1, 17.2).
//!
//! A call is a **UDF call** when its name is not a GridCraft built-in (the functions
//! [`is_known_function`] knows, the same lookup the evaluator dispatches on) and either carries
//! the `_xludf.` prefix or is in the workbook's set of UDF names. The engine gives calc that set
//! as data ([`crate::Calc::set_udf_names`]): the public `Function`s of the VBA project's standard
//! modules, and for an `.xls` the macro names its Lbl records flag `fFunc`/`fProc`. Calc never
//! reads VBA itself and never runs it.
//!
//! A built-in always wins: `SUM` stays `SUM` even when a VBA module declares `Function Sum`.
//! `_xludf.SUM` is not the built-in (the formula parser keeps the prefix), so it is a UDF call.
//! A workbook defined name (a LAMBDA) wins too, and so does a name the formula binds itself
//! with `LET` or as a `LAMBDA` parameter ([`UdfNames::udf_calls`]).
//!
//! A formula cell with at least one UDF call is a **foreign-result cell** (Req 17.3–17.10): its
//! value came from the application that saved the file, and calc never computes it. While
//! nothing it reads has changed, it keeps that value, through full recalculation too; once a
//! precedent changes, or its formula is edited, it shows `#BLOCKED!`, as a Python cell with new
//! code does (`Calc::run`). Undo puts the cell's value back with the rest of the snapshot.
//!
//! A UDF-calling formula still reads its argument cells, so it depends on them like any other
//! formula, but it is not *dynamic*: nothing about the call can reach cells the graph doesn't see
//! (`precedents`).

use std::collections::BTreeSet;

use gridcraft_formula::Expr;
use gridcraft_model::Workbook;

use crate::eval::is_known_function;

/// The prefix files write before a user-defined function's name, as the formula parser keeps it.
const XLUDF: &str = "_xludf.";

/// The workbook's UDF names, ASCII upper-cased (the formula parser upper-cases call names the
/// same way). Empty for a workbook with no VBA project and no `.xls` macro names: then only
/// `_xludf.` calls are UDF calls.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct UdfNames(BTreeSet<String>);

impl UdfNames {
    /// The set of `names`, compared case-insensitively (ASCII).
    pub fn new<I: IntoIterator<Item = S>, S: AsRef<str>>(names: I) -> UdfNames {
        UdfNames(names.into_iter().map(|n| n.as_ref().to_ascii_uppercase()).filter(|n| !n.is_empty()).collect())
    }

    /// The names, upper-cased, in order.
    pub fn names(&self) -> impl Iterator<Item = &str> {
        self.0.iter().map(String::as_str)
    }

    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// Whether a call to `name` (as the formula parser stores it) is a UDF call.
    pub fn is_udf_call(&self, name: &str) -> bool {
        if is_known_function(name) {
            return false;
        }
        has_xludf_prefix(name) || self.0.contains(&name.to_ascii_uppercase())
    }

    /// Whether a call to `name` in a formula on `sheet` of `wb` is a UDF call: as
    /// [`UdfNames::is_udf_call`], except that a defined name visible from the sheet (a LAMBDA, or
    /// any other name the evaluator would call) wins over a VBA `Function` (Req 17.2).
    pub fn is_udf_call_in(&self, wb: &Workbook, sheet: usize, name: &str) -> bool {
        self.is_udf_call(name) && wb.name(name, sheet).is_none()
    }

    /// The UDF calls in `e`, a formula on `sheet` of `wb`, by name as written in the formula
    /// (`DIAGONALWERT`, `_xludf.BAR`), each once, in the order they first appear. A call to a
    /// defined name ([`UdfNames::is_udf_call_in`]) or to a name the formula itself binds with
    /// `LET` or as a `LAMBDA` parameter is not a UDF call (Req 17.2). The walk is iterative
    /// ([`Expr::walk`]).
    pub fn udf_calls(&self, wb: &Workbook, sheet: usize, e: &Expr) -> Vec<String> {
        let mut out: Vec<String> = Vec::new();
        let mut bound: Option<Vec<&str>> = None;
        e.walk(&mut |x| {
            if let Expr::Call(n, _) = x
                && self.is_udf_call_in(wb, sheet, n)
                && !out.iter().any(|o| o.eq_ignore_ascii_case(n))
                && !bound.get_or_insert_with(|| bound_names(e)).iter().any(|b| b.eq_ignore_ascii_case(n))
            {
                out.push(n.clone());
            }
        });
        out
    }
}

/// The names `e` binds anywhere in it: `LET` names and `LAMBDA` parameters. Scopes are not
/// tracked: a name bound anywhere in the formula is not a UDF call anywhere in it, which only
/// errs towards evaluating a call (to `#NAME?`) instead of keeping a saved result.
fn bound_names(e: &Expr) -> Vec<&str> {
    let mut out = Vec::new();
    e.walk(&mut |x| match x {
        Expr::Call(n, args) if n == "LET" => {
            let pairs = args.len().saturating_sub(1);
            out.extend(args.iter().take(pairs).step_by(2).filter_map(|a| if let Expr::Name(b) = a { Some(b.as_str()) } else { None }));
        }
        Expr::Call(n, args) if n == "LAMBDA" => {
            let params = args.len().saturating_sub(1);
            out.extend(args.iter().take(params).filter_map(|a| if let Expr::Name(b) = a { Some(b.as_str()) } else { None }));
        }
        _ => {}
    });
    out
}

/// `name` starts with `_xludf.` (ASCII case-insensitive).
fn has_xludf_prefix(name: &str) -> bool {
    name.get(..XLUDF.len()).is_some_and(|p| p.eq_ignore_ascii_case(XLUDF))
}
