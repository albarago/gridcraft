//! The VBA cache parses on request only, once per live bytes (Req 6.1–6.5).
use std::sync::Arc;

use gridcraft_model::Workbook;
use gridcraft_vba::fixture::{Fixture, FixtureModule};
use gridcraft_vba::{DiagCode, ModuleKind};
use serde_json::json;

use super::{MAX_SLOTS, VbaCache};
use crate::{DocState, Session, Trust};

fn bin() -> Vec<u8> {
    Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, "Attribute VB_Name = \"Module1\"\r\nSub A()\r\nEnd Sub\r\n")])
        .to_vba_project_bin()
        .unwrap()
}

fn with_project(bytes: Vec<u8>) -> Workbook {
    let mut wb = Workbook::new();
    wb.vba_project = Some(Arc::new(bytes));
    wb
}

/// A session with an .xlsm carrying a fixture project, opened through `file.open`.
fn opened() -> Session {
    let xlsm = gridcraft_xlsx::write_xlsx_as(&with_project(bin()), gridcraft_xlsx::XlsxKind::MacroWorkbook).unwrap();
    let mut s = Session::new();
    s.execute("file.open", json!({"name": "m.xlsm", "base64": crate::io::base64_encode(&xlsm)})).unwrap();
    s
}

fn project(s: &mut Session) -> Arc<gridcraft_vba::VbaProject> {
    let wb = s.doc().unwrap().wb.clone();
    s.vba_cache.project(&wb)
}

#[test]
fn opening_a_workbook_parses_nothing() {
    let mut s = opened();
    assert!(s.doc().unwrap().wb.vba_project.is_some(), "the bin was kept");
    s.execute("cell.set", json!({"cell": "A1", "input": "1"})).unwrap();
    assert_eq!((s.vba_cache.parses, s.vba_cache.len()), (0, 0));
    let p = project(&mut s);
    assert_eq!(p.modules.len(), 1);
    assert_eq!(s.vba_cache.parses, 1);
}

#[test]
fn the_same_bytes_parse_once_across_edits_undo_and_redo() {
    let mut s = opened();
    let first = project(&mut s);
    assert!(Arc::ptr_eq(&first, &project(&mut s)));
    s.execute("cell.set", json!({"cell": "A1", "input": "1"})).unwrap();
    project(&mut s);
    s.execute("edit.undo", json!({})).unwrap();
    project(&mut s);
    s.execute("edit.redo", json!({})).unwrap();
    assert!(Arc::ptr_eq(&first, &project(&mut s)));
    assert_eq!(s.vba_cache.parses, 1);
    // New bytes, even equal ones, are a new Arc: read again.
    Arc::make_mut(&mut s.active_mut().unwrap().wb).vba_project = Some(Arc::new(bin()));
    let second = project(&mut s);
    assert_eq!(s.vba_cache.parses, 2);
    assert_eq!(*second, *first);
    // The signature part flag changes `signed`, so it is part of the key.
    Arc::make_mut(&mut s.active_mut().unwrap().wb).vba_signature_part = true;
    assert!(project(&mut s).signed);
    assert_eq!(s.vba_cache.parses, 3);
}

#[test]
fn a_workbook_without_a_project_gives_the_empty_project_without_parsing() {
    let mut cache = VbaCache::default();
    let p = cache.project(&Workbook::new());
    assert_eq!(*p, gridcraft_vba::VbaProject::default());
    assert!(p.diagnostics.is_empty());
    assert_eq!((cache.parses, cache.len()), (0, 0));
}

#[test]
fn unreadable_bytes_give_diagnostics_and_the_workbook_stays_editable() {
    let mut s = Session::new();
    s.add_document(DocState::new(with_project(b"not a compound file".to_vec()), None, "Book1".into(), Trust::FromFile));
    let p = project(&mut s);
    assert!(p.modules.is_empty());
    assert_eq!(p.diagnostics.iter().map(|d| d.code).collect::<Vec<_>>(), [DiagCode::DirCorrupt]);
    s.execute("cell.set", json!({"cell": "A1", "input": "2"})).unwrap();
    assert!(Arc::ptr_eq(&p, &project(&mut s)), "still cached after the edit");
    assert_eq!(s.vba_cache.parses, 1);
}

#[test]
fn slots_are_bounded_and_the_oldest_goes_first() {
    let mut cache = VbaCache::default();
    let bytes = bin();
    let wbs: Vec<Workbook> = (0..MAX_SLOTS + 4).map(|_| with_project(bytes.clone())).collect();
    for wb in &wbs {
        cache.project(wb);
    }
    assert_eq!((cache.parses, cache.len()), (MAX_SLOTS + 4, MAX_SLOTS));
    // The newest is still kept; the first was evicted and reads again.
    cache.project(&wbs[MAX_SLOTS + 3]);
    assert_eq!(cache.parses, MAX_SLOTS + 4);
    cache.project(&wbs[0]);
    assert_eq!((cache.parses, cache.len()), (MAX_SLOTS + 5, MAX_SLOTS));
}

#[test]
fn a_dropped_workbook_releases_its_slot() {
    let mut cache = VbaCache::default();
    let kept = with_project(bin());
    cache.project(&kept);
    {
        let dropped = with_project(bin());
        cache.project(&dropped);
        assert_eq!(cache.len(), 2);
    }
    cache.project(&kept);
    assert_eq!((cache.parses, cache.len()), (2, 1));
    // Through the session: closing the document frees its bytes and the slot.
    let mut s = opened();
    project(&mut s);
    s.close_document(0);
    assert_eq!(s.vba_cache.project(&Workbook::new()).modules.len(), 0);
    assert_eq!(s.vba_cache.len(), 0);
}
