//! Saving formulas that call VBA functions (Req 17.11, 17.12): the formula text is written as
//! it was read, `_xludf.` included, with the saved result in its own type; a formula showing
//! `#BLOCKED!` is written with no cached value, so the next application calculates it. A Python
//! cell (no formula) still caches `#BLOCKED!` as `#VALUE!`.

use std::sync::Arc;

use gridcraft_core::{CellError, CellRef, Value};
use gridcraft_model::{Cell, Formula, PyCode, StyleId, Workbook};

use crate::xml::El;

const BLOCKED: Value = Value::Error(CellError::Blocked);

fn at(a1: &str) -> CellRef {
    CellRef::parse(a1).unwrap()
}

fn workbook(cells: &[(&str, &str, Value)]) -> Workbook {
    let mut wb = Workbook::new();
    let sh = wb.sheet_mut(0).unwrap();
    sh.cells.set(at("A1"), Cell::value(Value::Number(21.0)));
    for (a1, f, v) in cells {
        sh.cells.set(at(a1), Cell { value: v.clone(), ..Cell::formula(Formula::new(f)) });
    }
    wb
}

/// `(t, <f>, <v>)` of each written cell.
fn written(bytes: &[u8], cells: &[&str]) -> Vec<(Option<String>, Option<String>, Option<String>)> {
    let mut pkg = crate::package::Package::open(bytes).unwrap();
    let sheet = pkg.read_xml("xl/worksheets/sheet1.xml").unwrap().unwrap();
    let all: Vec<&El> = sheet.child("sheetData").unwrap().kids("row").flat_map(|r| r.kids("c")).collect();
    cells
        .iter()
        .map(|a1| {
            let c = all.iter().find(|c| c.attr("r") == Some(*a1)).unwrap_or_else(|| panic!("{a1} not written"));
            (c.attr("t").map(str::to_string), c.child("f").map(|f| f.text.clone()), c.child("v").map(|v| v.text.clone()))
        })
        .collect()
}

fn workbook_xml(bytes: &[u8]) -> String {
    let mut pkg = crate::package::Package::open(bytes).unwrap();
    String::from_utf8(pkg.read("xl/workbook.xml").unwrap().unwrap()).unwrap()
}

fn s(x: &str) -> Option<String> {
    Some(x.to_string())
}

/// Req 17.11: every saved result keeps its value and type, and the formula text is unchanged:
/// `_xludf.` stays, a bare call (a function of the workbook's VBA project) stays bare.
#[test]
fn udf_calls_are_written_unchanged_with_their_saved_results() {
    let wb = workbook(&[
        ("B1", "_xludf.Twice(A1)", Value::Number(42.0)),
        ("B2", "_xludf.Name(A1)", Value::text("forty-two")),
        ("B3", "_xludf.Flag(A1)", Value::Bool(true)),
        ("B4", "_xludf.Fails(A1)", Value::Error(CellError::Div0)),
        ("B5", "TWICE(A1)", Value::Number(0.1)),
        ("B6", "DIAGONALWERT(A1)", Value::text("Diagonale")),
    ]);
    let bytes = crate::write_xlsx(&wb).unwrap();
    assert_eq!(
        written(&bytes, &["B1", "B2", "B3", "B4", "B5", "B6"]),
        [
            (None, s("_xludf.TWICE(A1)"), s("42")),
            (s("str"), s("_xludf.NAME(A1)"), s("forty-two")),
            (s("b"), s("_xludf.FLAG(A1)"), s("1")),
            (s("e"), s("_xludf.FAILS(A1)"), s("#DIV/0!")),
            (None, s("TWICE(A1)"), s("0.1")),
            (s("str"), s("DIAGONALWERT(A1)"), s("Diagonale")),
        ]
    );
    assert!(!workbook_xml(&bytes).contains("fullCalcOnLoad"), "nothing to calculate on load");
    // Read back: the same formulas and values.
    let (back, _) = crate::read_xlsx(&bytes).unwrap();
    for a1 in ["B1", "B2", "B3", "B4", "B5", "B6"] {
        let (a, b) = (wb.sheets[0].cell(at(a1)).unwrap(), back.sheets[0].cell(at(a1)).unwrap());
        assert_eq!((&a.value, a.formula.as_ref().map(|f| &f.text)), (&b.value, b.formula.as_ref().map(|f| &f.text)), "{a1}");
    }
}

/// Req 17.12: a formula showing `#BLOCKED!`, a VBA call or a cell reading one, is written as
/// `<f>` with no `<v>` and no `t`, and the workbook asks for a full calculation on load; a
/// Python cell showing `#BLOCKED!` keeps today's cached `#VALUE!`.
#[test]
fn a_blocked_formula_is_written_with_no_cached_value() {
    let mut wb = workbook(&[
        ("B1", "_xludf.Twice(A1)", BLOCKED),
        ("B2", "TWICE(A1)", BLOCKED),
        ("C1", "B1+1", BLOCKED),
        ("C2", "SUM(A1)", Value::Number(21.0)),
    ]);
    wb.sheet_mut(0)
        .unwrap()
        .cells
        .set(at("D1"), Cell { value: BLOCKED, formula: None, style: StyleId::DEFAULT, py: Some(Arc::new(PyCode::new("1 + 1"))) });
    let bytes = crate::write_xlsx(&wb).unwrap();
    assert_eq!(
        written(&bytes, &["B1", "B2", "C1", "C2", "D1"]),
        [
            (None, s("_xludf.TWICE(A1)"), None),
            (None, s("TWICE(A1)"), None),
            (None, s("B1+1"), None),
            (None, s("SUM(A1)"), s("21")),
            (s("e"), s("_xlfn._xlws.PY(0,0)"), s("#VALUE!")),
        ]
    );
    assert!(workbook_xml(&bytes).contains("fullCalcOnLoad=\"1\""));
    let (back, _) = crate::read_xlsx(&bytes).unwrap();
    let sh = &back.sheets[0];
    assert_eq!(sh.cell(at("B1")).unwrap().formula.as_ref().unwrap().text, "_xludf.TWICE(A1)");
    assert_eq!(sh.value(at("B1")), Value::Empty, "no saved result");
}
