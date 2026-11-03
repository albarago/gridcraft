//! `vba.modules` and `vba.source` give the same JSON through every automation surface (Req 9.8):
//! `Session::execute` (the UI's path), the `gridcraft-cli run` binary, the MCP `execute_command`
//! tool, and the engine-level control-channel methods (`engine.execute`, and the command id used
//! as a method) as the MCP `Headless` backend answers them. Results are compared as compact JSON
//! text after parsing, so only whitespace may differ.
//!
//! Not covered here: the desktop app's TCP control channel (`gridcraft --control`) and the MCP
//! `Remote` backend, which need a running window.
use gridcraft_engine::model::Workbook;
use gridcraft_engine::{Session, io};
use gridcraft_mcp::{Backend, Headless, call_tool};
use gridcraft_vba::ModuleKind;
use gridcraft_vba::fixture::{Fixture, FixtureModule};
use serde_json::{Value, json};
use std::path::PathBuf;
use std::process::Command;
use std::sync::Arc;

const MODULE1: &str = "Attribute VB_Name = \"Module1\"\r\nSub Hello()\r\n  MsgBox \"Café — 5 €\"\r\nEnd Sub\r\n";

struct TempDir(PathBuf);
impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

/// A temp dir holding `Macros.xlsm` with a three-module project, and that file's path.
fn xlsm(name: &str) -> (TempDir, String) {
    let dir = std::env::temp_dir().join(format!("gridcraft-cli-vba-{}-{name}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("Macros.xlsm").to_str().unwrap().to_string();
    let fixture = Fixture::new(vec![
        FixtureModule::new("Module1", ModuleKind::Standard, MODULE1),
        FixtureModule::new("Klasse", ModuleKind::Class, "Attribute VB_Name = \"Klasse\"\nPublic Größe As Long"),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, ""),
    ]);
    let mut wb = Workbook::new();
    wb.vba_project = Some(Arc::new(fixture.to_vba_project_bin().unwrap()));
    io::write_file(&path, &io::save_bytes(&wb, &path).unwrap()).unwrap();
    (TempDir(dir), path)
}

fn open(path: &str) -> Session {
    let mut s = Session::new();
    s.execute("file.open", json!({"path": path})).unwrap();
    s
}

fn compact(v: &Value) -> String {
    serde_json::to_string(v).unwrap()
}

const CALLS: [(&str, &str); 3] = [("vba.modules", "{}"), ("vba.source", r#"{"module":"Module1"}"#), ("vba.source", r#"{"module":"klasse"}"#)];

#[test]
fn vba_commands_match_across_cli_mcp_and_session() {
    let (_dir, path) = xlsm("same");
    let mut session = open(&path);
    let expected: Vec<String> = CALLS.iter().map(|(id, p)| compact(&session.execute(id, serde_json::from_str(p).unwrap()).unwrap())).collect();
    assert!(expected[0].contains(r#""kind":"standard","lines":4,"name":"Module1""#), "{}", expected[0]);
    assert!(expected[1].contains("Café — 5 €"), "{}", expected[1]);

    // CLI: one `{"command", "result"}` line per `--cmd`.
    let mut cli = Command::new(env!("CARGO_BIN_EXE_gridcraft-cli"));
    cli.args(["run", "--in", &path]);
    for (id, p) in CALLS {
        cli.args(["--cmd", &format!("{id}={p}")]);
    }
    let o = cli.output().unwrap();
    assert!(o.status.success(), "{}", String::from_utf8_lossy(&o.stderr));
    let lines: Vec<Value> = String::from_utf8(o.stdout).unwrap().lines().map(|l| serde_json::from_str(l).unwrap()).collect();
    assert_eq!(lines.len(), CALLS.len());
    for ((line, (id, _)), want) in lines.iter().zip(CALLS).zip(&expected) {
        assert_eq!(line["command"], id);
        assert_eq!(&compact(&line["result"]), want, "cli {id}");
    }

    // MCP tool, control-channel `engine.execute`, and the command id as a method.
    let mut mcp = Headless::with_session(open(&path));
    for ((id, p), want) in CALLS.iter().zip(&expected) {
        let params: Value = serde_json::from_str(p).unwrap();
        let r = call_tool(&mut mcp, "execute_command", &json!({"command": id, "params": params}));
        assert!(!r.is_error, "{r:?}");
        let text = r.content[0]["text"].as_str().unwrap();
        assert_eq!(&compact(&serde_json::from_str(text).unwrap()), want, "mcp {id}");
        let control = mcp.call("engine.execute", json!({"command": id, "params": params})).unwrap();
        assert_eq!(&compact(&control), want, "engine.execute {id}");
        assert_eq!(&compact(&mcp.call(id, params).unwrap()), want, "method {id}");
    }
}

#[test]
fn unknown_module_errors_on_cli_and_mcp() {
    let (_dir, path) = xlsm("errors");
    let o =
        Command::new(env!("CARGO_BIN_EXE_gridcraft-cli")).args(["run", "--in", &path, "--cmd", r#"vba.source={"module":"nope"}"#]).output().unwrap();
    assert!(!o.status.success());
    let stderr = String::from_utf8_lossy(&o.stderr);
    assert!(stderr.contains("vba.source): unknown module: nope"), "{stderr}");

    let mut mcp = Headless::with_session(open(&path));
    let r = call_tool(&mut mcp, "execute_command", &json!({"command": "vba.source", "params": {"module": "nope"}}));
    assert!(r.is_error);
    assert_eq!(r.content[0]["text"], "execute_command: unknown module: nope");
    assert_eq!(mcp.call("engine.execute", json!({"command": "vba.source", "params": {"module": "nope"}})), Err("unknown module: nope".to_string()));
}
