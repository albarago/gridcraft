//! View Code from the keyboard and to a screen reader (Req 8.9, 8.10): Up, Down, Home and End
//! walk the modules across their groups, Tab and Shift+Tab move between the list and the source,
//! and every module, heading and the source carry a label. The project is built in code by the
//! Fixture_Builder and opened from an .xlsm the engine wrote.
use egui::accesskit::Role;
use egui::{Key, Modifiers};
use egui_kittest::kittest::NodeT as _;
use egui_kittest::kittest::Queryable as _;
use gridcraft_engine::core::CellRef;
use gridcraft_engine::model::Workbook;
use gridcraft_engine::{Session, io};
use gridcraft_ui_egui::SheetApp;
use gridcraft_vba::ModuleKind;
use gridcraft_vba::fixture::{Fixture, FixtureModule};
use serde_json::json;
use std::sync::Arc;
type Harness = egui_kittest::Harness<'static, SheetApp>;
/// A window on a workbook whose project holds two standard modules, a class and a document
/// module, with View Code open; the .xlsm is removed once it is read.
fn harness(name: &str) -> Harness {
    let fixture = Fixture::new(vec![
        FixtureModule::new("Module1", ModuleKind::Standard, "Attribute VB_Name = \"Module1\"\r\nSub One()\r\nEnd Sub\r\n"),
        FixtureModule::new("alpha", ModuleKind::Standard, "Attribute VB_Name = \"alpha\"\r\nSub Two()\r\nEnd Sub\r\n"),
        FixtureModule::new("Klasse", ModuleKind::Class, "Attribute VB_Name = \"Klasse\"\r\nPublic X As Long\r\n"),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, "Attribute VB_Name = \"ThisWorkbook\"\r\n"),
    ]);
    let path = std::env::temp_dir().join(format!("gridcraft-vba-view-{}-{name}.xlsm", std::process::id()));
    let path = path.to_str().unwrap().to_string();
    let mut wb = Workbook::new();
    wb.vba_project = Some(Arc::new(fixture.to_vba_project_bin().unwrap()));
    io::write_file(&path, &io::save_bytes(&wb, &path).unwrap()).unwrap();
    let mut session = Session::new();
    session.execute("file.open", json!({"path": path})).unwrap();
    let _ = std::fs::remove_file(&path);
    let app = SheetApp::new(session, Default::default());
    let mut h = egui_kittest::Harness::builder().with_size(egui::vec2(1280.0, 800.0)).build_ui_state(
        |ui, app: &mut SheetApp| {
            let ctx = ui.ctx().clone();
            app.logic(&ctx);
            app.ui(ui);
        },
        app,
    );
    h.run_steps(4);
    h.state_mut().open_dialog("vbaCode", json!({}));
    h.run_steps(4);
    // The workbook's open said its macros were kept but not run. That box is modal: until it is
    // dismissed nothing under it takes the keyboard.
    assert!(h.state().message.as_ref().is_some_and(|(_, m)| m.contains("not executed")), "{:?}", h.state().message);
    h.get_by_label("  OK  ").click();
    h.run_steps(4);
    assert!(h.state().message.is_none());
    assert_eq!(h.state().dialog.as_ref().map(|d| d.name().to_string()).as_deref(), Some("vbaCode"));
    h
}
fn press(h: &mut Harness, modifiers: Modifiers, key: Key) {
    h.key_press_modifiers(modifiers, key);
    h.run_steps(3);
}
/// The module whose source shows, read from the source view's label.
fn shown(h: &Harness) -> String {
    let source = h.get_by_label_contains("Source of ");
    source.accesskit_node().label().unwrap_or_default().trim_start_matches("Source of ").to_string()
}
fn active_cell(h: &Harness) -> CellRef {
    h.state().session.active().map(|d| d.selection.active).unwrap_or_default()
}
#[test]
fn arrows_walk_the_modules_across_groups_and_stop_at_the_ends() {
    let mut h = harness("arrows");
    let cell = active_cell(&h);
    assert_eq!(shown(&h), "alpha", "the first module of the first group is selected");
    assert!(h.get_by_label("alpha, module").is_focused(), "the selected module has the focus");
    press(&mut h, Modifiers::NONE, Key::ArrowUp);
    assert_eq!(shown(&h), "alpha", "Up stops at the first module");
    press(&mut h, Modifiers::NONE, Key::ArrowDown);
    assert_eq!(shown(&h), "Module1");
    press(&mut h, Modifiers::NONE, Key::ArrowDown);
    assert_eq!(shown(&h), "Klasse", "Down goes on past the Class Modules heading");
    assert!(h.get_by_label("Klasse, class module").is_focused(), "the focus follows the selection");
    press(&mut h, Modifiers::NONE, Key::End);
    assert_eq!(shown(&h), "ThisWorkbook");
    press(&mut h, Modifiers::NONE, Key::ArrowDown);
    assert_eq!(shown(&h), "ThisWorkbook", "Down stops at the last module");
    press(&mut h, Modifiers::NONE, Key::ArrowUp);
    assert_eq!(shown(&h), "Klasse");
    press(&mut h, Modifiers::NONE, Key::Home);
    assert_eq!(shown(&h), "alpha");
    assert_eq!(active_cell(&h), cell, "the arrows never reach the grid behind the window");
}
#[test]
fn tab_moves_between_the_list_and_the_source() {
    let mut h = harness("tab");
    press(&mut h, Modifiers::NONE, Key::ArrowDown);
    press(&mut h, Modifiers::NONE, Key::Tab);
    assert!(h.get_by_label("Source of Module1").is_focused(), "Tab leaves the list for the source");
    // In the source the arrows move the caret, not the selection.
    press(&mut h, Modifiers::NONE, Key::ArrowDown);
    assert_eq!(shown(&h), "Module1");
    press(&mut h, Modifiers::SHIFT, Key::Tab);
    assert!(h.get_by_label("Module1, module").is_focused(), "Shift+Tab comes back to the selected module");
    press(&mut h, Modifiers::NONE, Key::ArrowDown);
    assert_eq!(shown(&h), "Klasse", "the list takes the arrows again");
}
#[test]
fn modules_headings_and_source_have_labels() {
    let h = harness("labels");
    for label in ["alpha, module", "Module1, module", "Klasse, class module", "ThisWorkbook, document module"] {
        assert!(h.query_by_label(label).is_some(), "{label}");
    }
    let headings: Vec<String> = h.get_all_by_role(Role::Heading).filter_map(|n| n.accesskit_node().value()).collect();
    for heading in ["Modules", "Class Modules", "Document Modules"] {
        assert!(headings.iter().any(|x| x == heading), "{heading} in {headings:?}");
    }
    assert!(h.query_by_label("Source of alpha").is_some());
}
