//! Automate › Macros › View Code in the window (Req 8.1 to 8.8): the button shows only with a
//! VBA project, the modules are grouped and sorted with the first one selected, the selected
//! module's source shows read only, and the banner carries what reading the project found. The
//! project is built in code by the Fixture_Builder and opened from an .xlsm the engine wrote; the
//! harness is the suite's strict one (pinned fonts, accessibility checks on).
use crate::common::{Harness, app, builder, harness_with};
use egui::accesskit::Role;
use egui::{Event, Key, Modifiers};
use egui_kittest::kittest::NodeT as _;
use egui_kittest::kittest::Queryable as _;
use gridcraft_engine::model::Workbook;
use gridcraft_engine::{Session, io};
use gridcraft_vba::ModuleKind;
use gridcraft_vba::fixture::{Fixture, FixtureModule, SignatureKind};
use serde_json::json;
use std::sync::Arc;
/// The Macros dialog's sentence about a VBA project, and the View Code header's (Req 8.8).
const DIALOG_KEPT: &str = "GridCraft keeps it when you save but does not run it.";
const VIEW_KEPT: &str = "GridCraft keeps the VBA project when you save but does not run it.";
const ZETA: &str = "Attribute VB_Name = \"zeta\"\r\nSub Zed()\r\n    MsgBox \"zeta\"\r\nEnd Sub\r\n";
const ALPHA: &str = "Attribute VB_Name = \"Alpha\"\r\nSub First()\r\n    Range(\"A1\").Value = 1\r\nEnd Sub\r\n";
/// A signed project with modules of every kind, given out of order: standard `zeta`, `Alpha`
/// and `Empty` (no stored source), class `beta`, document `ThisWorkbook` and `Sheet1` (only its
/// attributes are left beside compiled code: possible stomping), form `UserForm1`.
fn project() -> Fixture {
    let mut f = Fixture::new(vec![
        FixtureModule::new("zeta", ModuleKind::Standard, ZETA),
        FixtureModule::new("UserForm1", ModuleKind::Form, "Attribute VB_Name = \"UserForm1\"\r\nPrivate Sub UserForm_Click()\r\nEnd Sub\r\n"),
        FixtureModule::new(
            "ThisWorkbook",
            ModuleKind::Document,
            "Attribute VB_Name = \"ThisWorkbook\"\r\nPrivate Sub Workbook_Open()\r\nEnd Sub\r\n",
        ),
        FixtureModule::new("beta", ModuleKind::Class, "Attribute VB_Name = \"beta\"\r\nPublic Count As Long\r\n"),
        FixtureModule::new("Sheet1", ModuleKind::Document, "Attribute VB_Name = \"Sheet1\"\r\n").with_pcode(256),
        FixtureModule::new("Alpha", ModuleKind::Standard, ALPHA),
        FixtureModule::new("Empty", ModuleKind::Standard, ""),
    ]);
    f.signatures = vec![SignatureKind::Agile];
    f
}
/// A window on a workbook carrying `project` (or no project), opened from an .xlsm (or .xlsx)
/// the engine wrote; the file is removed once it is read. The box saying the macros were kept but
/// not executed is modal, so it is dismissed first.
fn harness(name: &str, project: Option<Fixture>) -> Harness {
    let ext = if project.is_some() { "xlsm" } else { "xlsx" };
    let path = std::env::temp_dir().join(format!("gridcraft-vba-view-ui-{}-{name}.{ext}", std::process::id()));
    let path = path.to_str().unwrap().to_string();
    let mut wb = Workbook::new();
    let with_project = project.is_some();
    wb.vba_project = project.map(|f| Arc::new(f.to_vba_project_bin().unwrap()));
    io::write_file(&path, &io::save_bytes(&wb, &path).unwrap()).unwrap();
    let mut session = Session::new();
    session.execute("file.open", json!({"path": path})).unwrap();
    let _ = std::fs::remove_file(&path);
    let mut h = harness_with(builder().with_size(egui::vec2(1280.0, 800.0)), app(session), 4);
    if with_project {
        // The open's note waits in the session until the window next runs a command; it comes up
        // then as a modal box, which takes the clicks until it is dismissed.
        h.state_mut().run("view.macros", json!({})).unwrap();
        h.run_steps(3);
        assert!(h.state().message.as_ref().is_some_and(|(_, m)| m.contains("not executed")), "{:?}", h.state().message);
        h.get_by_label("  OK  ").click();
        h.run_steps(4);
    }
    assert!(h.state().message.is_none(), "{:?}", h.state().message);
    h
}
fn dialog_name(h: &Harness) -> Option<String> {
    h.state().dialog.as_ref().map(|d| d.name().to_string())
}
/// Automate › Macros, then View Code.
fn view_code(h: &mut Harness) {
    h.state_mut().open_dialog("macros", json!({}));
    h.run_steps(4);
    h.get_by_label("View Code").click();
    h.run_steps(6);
    assert_eq!(dialog_name(h).as_deref(), Some("vbaCode"));
}
/// The selected module and the source the window shows for it, read from the source view.
fn shown(h: &Harness) -> (String, String) {
    let node = h.get_by_label_contains("Source of ");
    let name = node.accesskit_node().label().unwrap_or_default().trim_start_matches("Source of ").to_string();
    (name, node.accesskit_node().value().unwrap_or_default())
}
fn select(h: &mut Harness, label: &str) {
    h.get_by_label(label).click();
    h.run_steps(4);
}
/// What the workbook is: its revision, whether it has unsaved changes, and its undo depth.
fn state(h: &Harness) -> (u64, bool, usize) {
    let d = h.state().session.active().unwrap();
    (d.revision, d.is_dirty(), d.undo.len())
}
#[test]
fn a_workbook_without_a_project_offers_no_view_code() {
    let mut h = harness("none", None);
    h.state_mut().open_dialog("macros", json!({}));
    h.run_steps(4);
    assert_eq!(dialog_name(&h).as_deref(), Some("macros"));
    assert!(h.query_by_label("View Code").is_none(), "no project, no View Code (Req 8.7)");
    assert!(h.query_by_label_contains(DIALOG_KEPT).is_none());
}
#[test]
fn view_code_groups_sorts_and_selects_the_first_module() {
    let mut h = harness("groups", Some(project()));
    view_code(&mut h);
    let headings: Vec<String> = h.get_all_by_role(Role::Heading).filter_map(|n| n.accesskit_node().value()).collect();
    assert_eq!(headings, ["Modules", "Class Modules", "Document Modules", "Forms"], "groups in kind order (Req 8.1)");
    let expected = [
        "Alpha, module",
        "Empty, module",
        "zeta, module",
        "beta, class module",
        "Sheet1, document module",
        "ThisWorkbook, document module",
        "UserForm1, form",
    ];
    let items: Vec<String> =
        h.get_all_by_role(Role::Button).filter_map(|n| n.accesskit_node().label()).filter(|l| expected.contains(&l.as_str())).collect();
    assert_eq!(items, expected, "modules sorted by name ignoring case within each group (Req 8.1)");
    let (name, source) = shown(&h);
    assert_eq!(name, "Alpha", "the first module of the first group is selected (Req 8.2)");
    assert!(source.contains("Range(\"A1\").Value = 1"), "{source:?}");
}
#[test]
fn selecting_a_module_shows_its_source() {
    let mut h = harness("select", Some(project()));
    view_code(&mut h);
    select(&mut h, "zeta, module");
    let (name, source) = shown(&h);
    assert_eq!(name, "zeta");
    assert!(source.contains("MsgBox \"zeta\""), "{source:?}");
    assert!(!source.contains("Range(\"A1\")"), "the previous module's source is gone (Req 8.3)");
    select(&mut h, "beta, class module");
    assert!(shown(&h).1.contains("Public Count As Long"));
}
#[test]
fn edits_in_the_source_change_neither_the_text_nor_the_workbook() {
    let mut h = harness("readonly", Some(project()));
    view_code(&mut h);
    let before = state(&h);
    let (_, text) = shown(&h);
    h.get_by_label("Source of Alpha").click();
    h.run_steps(3);
    assert!(h.get_by_label("Source of Alpha").is_focused());
    // Each edit with the whole text selected, so a cut or a delete would take everything.
    for edit in ["type", "paste", "cut", "backspace", "delete", "enter"] {
        h.key_press_modifiers(Modifiers::COMMAND, Key::A);
        h.run_steps(2);
        match edit {
            "type" => h.event(Event::Text("typed".into())),
            "paste" => h.event(Event::Paste("pasted".into())),
            "cut" => h.event(Event::Cut),
            "backspace" => h.key_press(Key::Backspace),
            "delete" => h.key_press(Key::Delete),
            _ => h.key_press(Key::Enter),
        }
        h.run_steps(3);
        assert_eq!(shown(&h), ("Alpha".to_string(), text.clone()), "{edit}: the source stays as it was (Req 8.4)");
        assert_eq!(state(&h), before, "{edit}: the workbook stays as it was (Req 8.4)");
    }
    // The text read again from the workbook is the same too.
    assert_eq!(h.state_mut().run("vba.source", json!({"module": "Alpha"})).unwrap()["source"].as_str(), Some(text.as_str()));
}
#[test]
fn the_banner_shows_the_project_notes_and_the_selected_modules_warnings() {
    let mut h = harness("banner", Some(project()));
    let modules = h.state_mut().run("vba.modules", json!({})).unwrap();
    let stomped = modules["modules"]
        .as_array()
        .unwrap()
        .iter()
        .find(|m| m["name"] == "Sheet1")
        .and_then(|m| m["diagnostics"].as_array()?.iter().find(|d| d["severity"] == "warning")?["message"].as_str().map(str::to_string))
        .expect("Sheet1 carries a stomping warning");
    let project_notes: Vec<String> =
        modules["project"]["diagnostics"].as_array().unwrap().iter().filter_map(|d| d["message"].as_str().map(str::to_string)).collect();
    assert!(!project_notes.is_empty(), "a signed project has a note");
    view_code(&mut h);
    assert!(h.query_by_label("signed (not verified)").is_some(), "the project is marked signed, not verified (Req 8.5)");
    for note in &project_notes {
        assert!(h.query_by_label_contains(note).is_some(), "the project's note shows: {note}");
    }
    assert!(h.query_by_label_contains(&stomped).is_none(), "another module's warning stays out of the banner");
    select(&mut h, "Sheet1, document module");
    assert!(h.query_by_label_contains(&format!("Warning: {stomped}")).is_some(), "the selected module's warning shows (Req 8.5)");
    for note in &project_notes {
        assert!(h.query_by_label_contains(note).is_some(), "the project's notes stay: {note}");
    }
}
#[test]
fn a_module_without_stored_source_says_so() {
    let mut h = harness("empty", Some(project()));
    view_code(&mut h);
    assert!(h.query_by_label("This module has no stored source.").is_none());
    select(&mut h, "Empty, module");
    assert!(h.query_by_label("This module has no stored source.").is_some(), "Req 8.6");
    assert!(h.query_by_label_contains("Source of ").is_none(), "no source view for a module without source");
}
#[test]
fn the_kept_sentence_shows_in_the_dialog_and_in_view_code() {
    let mut h = harness("kept", Some(project()));
    h.state_mut().open_dialog("macros", json!({}));
    h.run_steps(4);
    assert!(h.query_by_label_contains(DIALOG_KEPT).is_some(), "the Macros dialog keeps its sentence (Req 8.8)");
    h.get_by_label("View Code").click();
    h.run_steps(6);
    assert!(h.query_by_label(VIEW_KEPT).is_some(), "View Code says it too (Req 8.8)");
}
/// A project whose standard module calls `Shell` and runs on open.
fn shell_project() -> Fixture {
    Fixture::new(vec![
        FixtureModule::new(
            "Runner",
            ModuleKind::Standard,
            "Attribute VB_Name = \"Runner\"\r\nSub Auto_Open()\r\n    Shell \"notepad.exe\"\r\nEnd Sub\r\n",
        ),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, "Attribute VB_Name = \"ThisWorkbook\"\r\n"),
    ])
}
#[test]
fn the_compatibility_tab_shows_the_verdict_and_the_flags() {
    let mut h = harness("compat", Some(shell_project()));
    view_code(&mut h);
    let before = state(&h);
    assert!(h.query_by_label_contains("Verdict:").is_none(), "the Source tab comes first");
    h.get_by_label("Compatibility").click();
    h.run_steps(6);
    assert_eq!(h.state().dialog.as_ref().map(|d| d.tab.clone()).as_deref(), Some("Compatibility"));
    let headings: Vec<String> = h.get_all_by_role(Role::Heading).filter_map(|n| n.accesskit_node().value()).collect();
    assert!(headings.iter().any(|t| t == "Verdict: Uses unsupported features"), "the verdict headline (Req 13.14): {headings:?}");
    assert!(headings.iter().any(|t| t == "Flagged constructs (1)"), "{headings:?}");
    assert!(h.query_by_label_contains("Shell in Runner, line 3").is_some(), "the verdict's reason");
    assert!(h.query_by_label("Shell: Runner, line 3").is_some(), "a row for the Shell flag");
    assert!(h.query_by_label_contains("Processes and native code: present (1)").is_some(), "the process capability");
    assert!(h.query_by_label("Network: absent").is_some(), "an absent capability");
    assert!(h.query_by_label("entry point, auto-exec").is_some(), "Auto_Open's marks");
    assert!(h.query_by_label_contains("Source of ").is_none(), "the source view is on the other tab");
    // The tab shows what vba.analyze returns, and reading it changed nothing.
    let report = h.state_mut().run("vba.analyze", json!({})).unwrap();
    assert_eq!(h.state().dialog.as_ref().and_then(|d| d.values.get("analysis")), Some(&report));
    assert_eq!(state(&h), before);
    h.get_by_label("Source").click();
    h.run_steps(4);
    assert_eq!(shown(&h).0, "Runner");
}
