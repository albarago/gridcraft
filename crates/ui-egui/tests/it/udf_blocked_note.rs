//! A `#BLOCKED!` cell whose formula calls VBA functions shows a note naming them and saying VBA
//! functions are not executed (Req 17.8).

use egui_kittest::kittest::Queryable as _;
use gridcraft_engine::Session;
use gridcraft_engine::core::{CellRef, Value};
use gridcraft_engine::model::{Cell, Formula};
use gridcraft_ui_egui::SheetApp;
use serde_json::json;

type Harness = egui_kittest::Harness<'static, SheetApp>;

/// A1 = 3, and a UDF-calling cell `=_xludf.Twice(A1)` that is `#BLOCKED!` because A1 was edited
/// after load (no saved result here). `active` is selected.
fn harness(active: &str) -> Harness {
    let mut s = Session::new();
    s.new_workbook();
    {
        let d = s.doc_mut().unwrap();
        let wb = std::sync::Arc::make_mut(&mut d.wb);
        let sh = wb.sheet_mut(0).unwrap();
        sh.cells.set(CellRef::parse("A1").unwrap(), Cell::value(Value::Number(3.0)));
        sh.cells.set(CellRef::parse("B1").unwrap(), Cell::formula(Formula::new("=_xludf.Twice(A1)")));
        sh.cells.set(CellRef::parse("B2").unwrap(), Cell::value(Value::Number(1.0)));
        let mut wb = (*d.wb).clone();
        d.calc.recalc_all(&mut wb);
        d.wb = std::sync::Arc::new(wb);
    }
    s.execute("selection.set", json!({"cell": active})).unwrap();
    let mut h = crate::common::builder().with_size(egui::vec2(1200.0, 800.0)).build_ui_state(
        |ui, app: &mut SheetApp| {
            app.logic(ui.ctx());
            app.ui(ui);
        },
        crate::common::app(s),
    );
    h.run_steps(4);
    h
}

const NOT_RUN: &str = "VBA functions are not executed, so this cell shows #BLOCKED! instead of a result.";

#[test]
fn the_note_names_the_vba_function_on_the_selected_cell() {
    let mut h = harness("B1");
    assert_eq!(h.state_mut().session.execute("cell.get", json!({"cell": "B1"})).unwrap()["value"]["error"], "#BLOCKED!");
    assert!(h.query_by_label("Calls VBA functions: TWICE").is_some(), "the note names the function as the formula does, without the _xludf. prefix");
    assert!(h.query_by_label(NOT_RUN).is_some());
}

#[test]
fn an_ordinary_cell_shows_no_note() {
    let h = harness("B2");
    assert!(h.query_by_label(NOT_RUN).is_none());
}
