//! `vba.modules` and `vba.source` on a macro workbook that went through a real save and open
//! (Req 9.1-9.6, 9.8; Req 1.3 for the .xlsx case). The project is built in code by the
//! Fixture_Builder, written into an .xlsm with the engine's `io` API, then opened from disk in a
//! fresh session.
//!
//! Surfaces covered here: `Session::execute` (UI and control channel) and
//! `Session::execute_headless` / `execute_headless_in` (what the CLI and the MCP `Headless` backend
//! call). The CLI binary and the MCP `execute_command` tool are checked byte-for-byte against
//! `Session::execute` in `apps/gridcraft-cli/tests/it/vba.rs`.
use gridcraft_engine::model::Workbook;
use gridcraft_engine::{EngineError, Session, Target, io};
use gridcraft_vba::ModuleKind;
use gridcraft_vba::fixture::{Fixture, FixtureModule};
use serde_json::{Value as Json, json};
use std::path::{Path, PathBuf};
use std::sync::Arc;

const MODULE1: &str = "Attribute VB_Name = \"Module1\"\r\nSub Hello()\r\n  MsgBox \"Café — 5 €\"\r\nEnd Sub\r\n";
const CLASS1: &str = "Attribute VB_Name = \"Klasse\"\nPublic Größe As Long";
const THIS_WORKBOOK: &str = "Attribute VB_Name = \"ThisWorkbook\"\rPrivate Sub Workbook_Open()\rEnd Sub";

fn fixture() -> Fixture {
    Fixture::new(vec![
        FixtureModule::new("Module1", ModuleKind::Standard, MODULE1),
        FixtureModule::new("Klasse", ModuleKind::Class, CLASS1),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, THIS_WORKBOOK),
    ])
}

/// A temp dir only this test uses; removed by [`TempDir`]'s drop.
struct TempDir(PathBuf);
impl TempDir {
    fn new(name: &str) -> Self {
        let d = std::env::temp_dir().join(format!("gridcraft-vba-it-{}-{name}", std::process::id()));
        std::fs::create_dir_all(&d).unwrap();
        Self(d)
    }
    fn file(&self, name: &str) -> String {
        self.0.join(name).to_str().unwrap().to_string()
    }
}
impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

/// Writes a new workbook holding `fixture`'s project to `path` (.xlsm) through `io::save_bytes`.
fn write_xlsm(fixture: &Fixture, path: &str) {
    let mut wb = Workbook::new();
    wb.vba_project = Some(Arc::new(fixture.to_vba_project_bin().unwrap()));
    let bytes = io::save_bytes(&wb, path).unwrap();
    io::write_file(path, &bytes).unwrap();
}

/// A fresh session with `path` opened through `file.open`.
fn open(path: &str) -> Session {
    let mut s = Session::new();
    s.execute("file.open", json!({"path": path})).unwrap();
    s
}

fn modules(r: &Json) -> &[Json] {
    r["modules"].as_array().unwrap()
}

fn assert_fixture_modules(s: &mut Session) {
    let r = s.execute("vba.modules", json!({})).unwrap();
    assert_eq!(r["project"], json!({"name": "VBAProject", "codePage": 1252, "signed": false, "diagnostics": []}), "{r}");
    let got: Vec<_> = modules(&r)
        .iter()
        .map(|m| (m["name"].as_str().unwrap(), m["kind"].as_str().unwrap(), m["codeName"].as_str().unwrap(), m["lines"].as_u64().unwrap()))
        .collect();
    assert_eq!(got, [("Module1", "standard", "Module1", 4), ("Klasse", "class", "Klasse", 2), ("ThisWorkbook", "document", "ThisWorkbook", 3)]);
    assert!(modules(&r).iter().all(|m| m["diagnostics"] == json!([])), "{r}");
    for (name, source) in [("Module1", MODULE1), ("Klasse", CLASS1), ("ThisWorkbook", THIS_WORKBOOK)] {
        let r = s.execute("vba.source", json!({"module": name})).unwrap();
        assert_eq!(r, json!({"module": name, "source": source}));
    }
    // Names match ignoring case; the reply carries the module's own spelling.
    assert_eq!(s.execute("vba.source", json!({"module": "KLASSE"})).unwrap()["module"], "Klasse");
}

#[test]
fn vba_commands_read_an_xlsm_after_save_and_open() {
    let dir = TempDir::new("roundtrip");
    let path = dir.file("Macros.xlsm");
    write_xlsm(&fixture(), &path);
    let mut s = open(&path);
    assert_fixture_modules(&mut s);
    // A second trip through the command API (`file.saveAs` .xlsm) keeps the project.
    let again = dir.file("Again.xlsm");
    s.execute("file.saveAs", json!({"path": again})).unwrap();
    assert_fixture_modules(&mut open(&again));
}

#[test]
fn vba_commands_give_identical_json_through_every_session_entry_point() {
    let dir = TempDir::new("surfaces");
    let path = dir.file("Macros.xlsm");
    write_xlsm(&fixture(), &path);
    let mut s = open(&path);
    for (id, params) in [("vba.modules", json!({})), ("vba.source", json!({"module": "Module1"})), ("vba.source", json!({"module": "klasse"}))] {
        let ui = serde_json::to_string(&s.execute(id, params.clone()).unwrap()).unwrap();
        let headless = serde_json::to_string(&s.execute_headless(id, params.clone()).unwrap()).unwrap();
        let targeted = serde_json::to_string(&s.execute_headless_in(Target::Active, id, params.clone()).unwrap()).unwrap();
        assert_eq!(ui, headless, "{id} {params}");
        assert_eq!(ui, targeted, "{id} {params}");
    }
    let ui = s.execute("vba.source", json!({"module": "nope"})).unwrap_err().to_string();
    let headless = s.execute_headless("vba.source", json!({"module": "nope"})).unwrap_err().to_string();
    assert_eq!(ui, headless);
}

#[test]
fn saving_as_xlsx_drops_the_project() {
    let dir = TempDir::new("xlsx");
    let path = dir.file("Macros.xlsm");
    write_xlsm(&fixture(), &path);
    let mut s = open(&path);
    let plain = dir.file("Plain.xlsx");
    s.execute("file.saveAs", json!({"path": plain})).unwrap();
    assert!(Path::new(&plain).exists());
    let mut reopened = open(&plain);
    assert!(reopened.doc().unwrap().wb.vba_project.is_none());
    let r = reopened.execute("vba.modules", json!({})).unwrap();
    assert_eq!(r, json!({"project": {"name": null, "codePage": null, "signed": false, "diagnostics": []}, "modules": []}));
    let e = reopened.execute("vba.source", json!({"module": "Module1"})).unwrap_err();
    assert!(matches!(&e, EngineError::Other(m) if m == "unknown module: Module1"), "{e:?}");
}

#[test]
fn vba_source_errors_after_open() {
    let dir = TempDir::new("errors");
    let path = dir.file("Macros.xlsm");
    write_xlsm(&fixture(), &path);
    let mut s = open(&path);
    let e = s.execute("vba.source", json!({"module": "nope"})).unwrap_err();
    assert!(matches!(&e, EngineError::Other(m) if m == "unknown module: nope"), "{e:?}");
    assert_eq!(e.to_string(), "unknown module: nope");
    for params in [json!({}), json!({"module": ""}), json!({"module": 7})] {
        let e = s.execute("vba.source", params.clone()).unwrap_err();
        assert!(matches!(&e, EngineError::BadParams { cmd, .. } if cmd == "vba.source"), "{params}: {e:?}");
    }
}
