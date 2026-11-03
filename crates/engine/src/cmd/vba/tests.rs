//! `vba.modules` and `vba.source` (Req 9.1-9.6, 9.9).
use super::{find_module, line_count};
use crate::{EngineError, Session};
use gridcraft_model::Workbook;
use gridcraft_vba::fixture::{Fixture, FixtureModule, SignatureKind};
use gridcraft_vba::{ModuleKind, VbaModule, VbaProject};
use serde_json::{Value as Json, json};
use std::sync::Arc;
const MODULE1: &str = "Attribute VB_Name = \"Module1\"\r\nSub A()\r\n  MsgBox \"Café — 5 €\"\r\nEnd Sub\r\n";
fn fixture() -> Fixture {
    Fixture {
        signatures: vec![SignatureKind::Legacy],
        ..Fixture::new(vec![
            FixtureModule::new("Module1", ModuleKind::Standard, MODULE1),
            FixtureModule::new("Class1", ModuleKind::Class, "Attribute VB_Name = \"Class1\"\r\nPublic X As Long"),
            FixtureModule::new("ThisWorkbook", ModuleKind::Document, ""),
            FixtureModule::new("UserForm1", ModuleKind::Form, "Attribute VB_Name = \"UserForm1\"\n"),
            // Compiled code with only Attribute lines: possible stomping.
            FixtureModule::new("Stomped", ModuleKind::Standard, "Attribute VB_Name = \"Stomped\"\r\n").with_pcode(64),
        ])
    }
}
/// A session whose active workbook is an .xlsm with `fixture`'s project, opened through
/// `file.open`.
fn opened(fixture: &Fixture) -> Session {
    let mut wb = Workbook::new();
    wb.vba_project = Some(Arc::new(fixture.to_vba_project_bin().unwrap()));
    let xlsm = gridcraft_xlsx::write_xlsx_as(&wb, gridcraft_xlsx::XlsxKind::MacroWorkbook).unwrap();
    let mut s = Session::new();
    s.execute("file.open", json!({"name": "m.xlsm", "base64": crate::io::base64_encode(&xlsm)})).unwrap();
    s
}
fn no_project() -> Session {
    let mut s = Session::new();
    s.execute("file.new", json!({})).unwrap();
    s
}
fn module<'a>(r: &'a Json, name: &str) -> &'a Json {
    r["modules"].as_array().unwrap().iter().find(|m| m["name"] == name).unwrap()
}
#[test]
fn vba_modules_lists_kinds_code_names_lines_and_diagnostics() {
    let mut s = opened(&fixture());
    let r = s.execute("vba.modules", json!({})).unwrap();
    let p = &r["project"];
    assert_eq!((&p["name"], &p["codePage"], &p["signed"]), (&json!("VBAProject"), &json!(1252), &json!(true)));
    assert!(p["diagnostics"].as_array().unwrap().iter().any(|d| d["code"] == "signed" && d["severity"] == "info" && d["module"].is_null()), "{p}");
    let names: Vec<_> = r["modules"].as_array().unwrap().iter().map(|m| m["name"].as_str().unwrap()).collect();
    assert_eq!(names, ["Module1", "Class1", "ThisWorkbook", "UserForm1", "Stomped"], "dir order");
    for (name, kind, lines) in
        [("Module1", "standard", 4), ("Class1", "class", 2), ("ThisWorkbook", "document", 0), ("UserForm1", "form", 1), ("Stomped", "standard", 1)]
    {
        let m = module(&r, name);
        assert_eq!((&m["kind"], &m["codeName"], &m["lines"]), (&json!(kind), &json!(name), &json!(lines)), "{m}");
        assert!(m["diagnostics"].is_array(), "{m}");
    }
    let stomped = &module(&r, "Stomped")["diagnostics"];
    let d = stomped.as_array().unwrap().iter().find(|d| d["code"] == "possibleStomping").unwrap();
    assert_eq!((&d["severity"], &d["module"]), (&json!("warning"), &json!("Stomped")));
    assert!(d["message"].as_str().is_some_and(|m| !m.is_empty()));
    assert_eq!(module(&r, "Module1")["diagnostics"], json!([]));
}
#[test]
fn vba_modules_without_a_project_is_empty_not_an_error() {
    let mut s = no_project();
    let r = s.execute("vba.modules", json!({})).unwrap();
    assert_eq!(r, json!({"project": {"name": null, "codePage": null, "signed": false, "diagnostics": []}, "modules": []}));
}
#[test]
fn vba_source_returns_the_source_exactly_matching_names_ignoring_case() {
    let mut s = opened(&fixture());
    for asked in ["Module1", "module1", "MODULE1"] {
        let r = s.execute("vba.source", json!({"module": asked})).unwrap();
        assert_eq!(r, json!({"module": "Module1", "source": MODULE1}), "{asked}");
    }
    let r = s.execute("vba.source", json!({"module": "thisworkbook"})).unwrap();
    assert_eq!(r, json!({"module": "ThisWorkbook", "source": ""}));
}
#[test]
fn vba_source_errors_are_typed() {
    let mut s = opened(&fixture());
    for params in [json!({}), json!({"module": ""}), json!({"module": 5}), json!({"module": null}), json!([1, 2])] {
        let r = s.execute("vba.source", params.clone());
        assert!(matches!(&r, Err(EngineError::BadParams { cmd, .. }) if cmd == "vba.source"), "{params}: {r:?}");
    }
    for mut s in [opened(&fixture()), no_project()] {
        let r = s.execute("vba.source", json!({"module": "Foo"}));
        assert!(matches!(&r, Err(EngineError::Other(m)) if m == "unknown module: Foo"), "{r:?}");
    }
    assert!(s.take_ui_requests().iter().all(|q| !matches!(q, crate::UiRequest::Dialog(..))), "no dialog");
}
#[test]
fn the_commands_open_no_dialog_and_change_nothing() {
    let mut s = opened(&fixture());
    s.take_ui_requests();
    let (dirty, revision) = (s.doc().unwrap().is_dirty(), s.doc().unwrap().revision);
    let wb = s.doc().unwrap().wb.clone();
    for (id, params) in [("vba.modules", json!({})), ("vba.source", json!({})), ("vba.source", json!({"module": "Class1"}))] {
        let _ = s.execute(id, params);
        let d = s.doc().unwrap();
        assert!(Arc::ptr_eq(&d.wb, &wb) && d.is_dirty() == dirty && d.revision == revision && d.undo.is_empty(), "{id}");
    }
    assert!(s.take_ui_requests().iter().all(|q| !matches!(q, crate::UiRequest::Dialog(..))));
}
#[test]
fn line_counts_follow_crlf_lf_and_cr() {
    for (src, n) in [("", 0), ("a", 1), ("a\r\n", 1), ("a\nb", 2), ("a\rb\r\n", 2), ("a\r\n\r\n", 2), ("\r", 1), ("\n\r", 2), ("é", 1)] {
        assert_eq!(line_count(src), n, "{src:?}");
    }
}
fn m(name: &str, code_name: &str) -> VbaModule {
    VbaModule {
        name: name.into(),
        code_name: code_name.into(),
        stream_name: code_name.into(),
        kind: ModuleKind::Standard,
        text_offset: 0,
        source: name.into(),
        diagnostics: Vec::new(),
    }
}
#[test]
fn a_name_match_wins_over_a_code_name_match() {
    let p = VbaProject { modules: vec![m("Alpha", "Beta"), m("Beta", "Gamma"), m("Ärger", "Arger")], ..VbaProject::default() };
    assert_eq!(find_module(&p, "beta").map(|m| m.name.as_str()), Some("Beta"));
    assert_eq!(find_module(&p, "GAMMA").map(|m| m.name.as_str()), Some("Beta"), "code name when no name matches");
    assert_eq!(find_module(&p, "ÄRGER").map(|m| m.name.as_str()), Some("Ärger"), "Unicode case");
    assert!(find_module(&p, "Delta").is_none());
}
