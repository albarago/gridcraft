//! Properties of the VBA query commands (design P8, P9), on seeded projects opened through
//! `file.open`: clean fixtures of 0 to 8 modules (sometimes signed, sometimes with a stomped
//! module), corrupted fixtures (`testutil::p4_case`), fuzzed compound files
//! (`testutil::fuzz_bin`), bytes that aren't a compound file at all, and no project.
//!
//! Knobs: `GRIDCRAFT_FUZZ_SEEDS` (default 64) gives P8 and P9 one case per seed;
//! `GRIDCRAFT_FUZZ_SEED=17` replays one case.
use crate::{EngineError, Session};
use gridcraft_model::Workbook;
use gridcraft_vba::ModuleKind;
use gridcraft_vba::fixture::{Fixture, FixtureModule, SignatureKind};
use gridcraft_vba::testutil::{Rng, fuzz_bin, guarded, p4_case, seed_list, source};
use gridcraft_xlsx::XlsxKind;
use serde_json::{Value as Json, json};
use std::sync::Arc;
/// Distinct from the `testutil` streams, so a seed's case doesn't reuse their sequences.
const PROPS_STREAM: u64 = 0x5038_5039_5038_5039;
fn replay(seed: u64, test: &str) -> String {
    format!("replay: GRIDCRAFT_FUZZ_SEED={seed} cargo test -p gridcraft-engine --lib cmd::vba::props::{test}")
}
/// What a case's `vbaProject.bin` is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Class {
    None,
    Clean,
    Corrupt,
    Fuzzed,
    Garbage,
}
/// A seeded case: the project bytes the workbook carries, the module names a command may ask
/// for, and what the bytes are.
struct Case {
    bin: Option<Vec<u8>>,
    names: Vec<String>,
    class: Class,
}
/// A VBA identifier: a letter, then letters, digits and `_`, up to 31 characters.
fn identifier(rng: &mut Rng) -> String {
    const CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_";
    let len = rng.range(1, 31);
    let mut name = String::with_capacity(len);
    name.push(char::from(CHARS.get(rng.below(52)).copied().unwrap_or(b'M')));
    while name.len() < len {
        name.push(char::from(CHARS.get(rng.below(CHARS.len())).copied().unwrap_or(b'_')));
    }
    name
}
fn kind(rng: &mut Rng) -> ModuleKind {
    match rng.below(4) {
        0 => ModuleKind::Standard,
        1 => ModuleKind::Class,
        2 => ModuleKind::Form,
        _ => ModuleKind::Document,
    }
}
/// 0 to 8 modules with case-insensitively unique names (never `dir`); now and then one holds
/// compiled code and only an `Attribute` line (possible stomping), and some projects are signed.
fn clean_fixture(rng: &mut Rng) -> Fixture {
    let count = rng.range(0, 8);
    let mut modules: Vec<FixtureModule> = Vec::with_capacity(count);
    for _ in 0..count.saturating_mul(20) {
        if modules.len() >= count {
            break;
        }
        let name = identifier(rng);
        if name.eq_ignore_ascii_case("dir") || modules.iter().any(|m| m.name.eq_ignore_ascii_case(&name)) {
            continue;
        }
        let k = kind(rng);
        let module = if rng.chance(15) {
            FixtureModule::new(&name, k, &format!("Attribute VB_Name = \"{name}\"\r\n")).with_pcode(rng.range(1, 256))
        } else {
            let len = if rng.chance(85) { rng.range(0, 1500) } else { rng.range(1501, 6000) };
            FixtureModule::new(&name, k, &source(rng, len))
        };
        modules.push(module);
    }
    let mut f = Fixture::new(modules);
    if rng.chance(30) {
        let n = rng.range(1, SignatureKind::ALL.len());
        f.signatures = SignatureKind::ALL.iter().copied().take(n).collect();
    }
    f
}
/// Random bytes that aren't a compound file: 1 to 4,096 of them, never starting with the CFB
/// signature's first byte.
fn garbage(rng: &mut Rng) -> Vec<u8> {
    let mut out = Vec::new();
    let n = rng.range(1, 4096);
    rng.fill(&mut out, n);
    if let Some(first) = out.first_mut()
        && *first == 0xD0
    {
        *first = 0x00;
    }
    out
}
fn case(seed: u64) -> Case {
    let mut rng = Rng::new(seed, PROPS_STREAM);
    let names = |f: &Fixture| f.modules.iter().map(|m| m.name.clone()).collect::<Vec<_>>();
    match rng.below(16) {
        0 => Case { bin: None, names: Vec::new(), class: Class::None },
        1..=8 => {
            let f = clean_fixture(&mut rng);
            match f.to_vba_project_bin() {
                Ok(bin) => Case { bin: Some(bin), names: names(&f), class: Class::Clean },
                Err(_) => Case { bin: Some(garbage(&mut rng)), names: names(&f), class: Class::Garbage },
            }
        }
        9..=11 => {
            let c = p4_case(seed);
            match c.fixture.to_vba_project_bin() {
                Ok(bin) => Case { bin: Some(bin), names: names(&c.fixture), class: Class::Corrupt },
                // A cyclic chain on a one-sector stream can't be built: use bytes that aren't a
                // compound file instead.
                Err(_) => Case { bin: Some(garbage(&mut rng)), names: names(&c.fixture), class: Class::Garbage },
            }
        }
        12..=13 => Case { bin: Some(fuzz_bin(seed)), names: vec!["Module1".to_owned()], class: Class::Fuzzed },
        _ => Case { bin: Some(garbage(&mut rng)), names: vec!["Module1".to_owned()], class: Class::Garbage },
    }
}
/// An .xlsm (or .xltm) package of a new workbook carrying `bin` as its `vbaProject.bin`.
fn package(bin: Option<&[u8]>, kind: XlsxKind) -> Result<Vec<u8>, String> {
    let mut wb = Workbook::new();
    wb.vba_project = bin.map(|b| Arc::new(b.to_vec()));
    gridcraft_xlsx::write_xlsx_as(&wb, kind).map_err(|e| format!("the package didn't write: {e}"))
}
/// A session with `name` (whose extension picks the kind) opened from `bytes` through
/// `file.open`, as the UI and the control channel open a download.
fn open(name: &str, bytes: &[u8]) -> Result<Session, String> {
    let mut s = Session::new();
    s.execute("file.open", json!({"name": name, "base64": crate::io::base64_encode(bytes)})).map_err(|e| format!("file.open failed: {e}"))?;
    Ok(s)
}
/// `s` with each ASCII letter's case flipped at random, sometimes.
fn recase(rng: &mut Rng, s: &str) -> String {
    if rng.chance(40) {
        return s.to_owned();
    }
    s.chars().map(|c| if rng.chance(50) { if c.is_ascii_lowercase() { c.to_ascii_uppercase() } else { c.to_ascii_lowercase() } } else { c }).collect()
}
/// One `vba.modules` or `vba.source` call: a module's name (sometimes in another case), an
/// unknown name, an empty, missing, null or non-text `module`, or a non-object params value.
fn query(rng: &mut Rng, names: &[String]) -> (&'static str, Json) {
    if rng.chance(30) {
        return ("vba.modules", if rng.chance(80) { json!({}) } else { json!({"module": "ignored"}) });
    }
    let params = match rng.below(10) {
        0..=4 => match rng.pick(names) {
            Some(n) => json!({"module": recase(rng, n)}),
            None => json!({"module": "Module1"}),
        },
        5 => json!({"module": format!("{}Nope", rng.below(1000))}),
        6 => json!({"module": ""}),
        7 => json!({}),
        8 => json!({"module": if rng.chance(50) { Json::Null } else { json!(rng.below(100)) }}),
        _ => json!([1, "x"]),
    };
    ("vba.source", params)
}
/// Everything a query must leave alone on the active document and the session.
#[derive(Debug, PartialEq)]
struct State {
    wb: Workbook,
    undo: usize,
    redo: usize,
    dirty: bool,
    revision: u64,
    journal: usize,
    selection: crate::Selection,
    docs: usize,
}
fn state(s: &Session) -> Result<(State, Arc<Workbook>), String> {
    let d = s.doc().map_err(|e| format!("no document: {e}"))?;
    let st = State {
        wb: (*d.wb).clone(),
        undo: d.undo.len(),
        redo: d.redo.len(),
        dirty: d.is_dirty(),
        revision: d.revision,
        journal: s.journal.len(),
        selection: d.selection.clone(),
        docs: s.docs.len(),
    };
    Ok((st, d.wb.clone()))
}
/// Gives the document some history first, so the lengths and the dirty flag are not all zero:
/// a few cell edits, then sometimes an undo (a redo step) or a save point.
fn prelude(rng: &mut Rng, s: &mut Session) -> Result<(), String> {
    for i in 0..rng.range(0, 3) {
        s.execute("cell.set", json!({"cell": format!("A{}", i + 1), "input": format!("{}", rng.below(1000))}))
            .map_err(|e| format!("cell.set: {e}"))?;
    }
    if rng.chance(40) && s.doc().is_ok_and(|d| !d.undo.is_empty()) {
        s.execute("edit.undo", json!({})).map_err(|e| format!("edit.undo: {e}"))?;
    }
    Ok(())
}
/// P8 for one seed.
fn p8_case(seed: u64) -> Result<(), String> {
    let c = case(seed);
    let mut rng = Rng::new(seed, PROPS_STREAM ^ 0x8);
    let mut s = open("m.xlsm", &package(c.bin.as_deref(), XlsxKind::MacroWorkbook)?)?;
    prelude(&mut rng, &mut s)?;
    let (before, wb) = state(&s)?;
    for step in 0..rng.range(1, 12) {
        let (id, params) = query(&mut rng, &c.names);
        let r = s.execute(id, params.clone());
        let what = || format!("step {step}: {id} {params} ({:?} project) gave {r:?}", c.class);
        if let Err(EngineError::Internal(..)) = &r {
            return Err(format!("{}: a command panicked", what()));
        }
        if id == "vba.modules" && r.is_err() {
            return Err(format!("{}: vba.modules never fails on an open workbook", what()));
        }
        let known = params.get("module").and_then(Json::as_str).is_some_and(|m| c.names.iter().any(|n| n.eq_ignore_ascii_case(m)));
        if c.class == Class::Clean && id == "vba.source" && known && r.is_err() {
            return Err(format!("{}: a clean project's module wasn't found", what()));
        }
        let (after, now) = state(&s)?;
        if after != before {
            return Err(format!(
                "{}: the state changed\n  before: undo {} redo {} dirty {} revision {} journal {} docs {}\n  after:  undo {} redo {} dirty {} revision {} journal {} docs {}{}",
                what(),
                before.undo,
                before.redo,
                before.dirty,
                before.revision,
                before.journal,
                before.docs,
                after.undo,
                after.redo,
                after.dirty,
                after.revision,
                after.journal,
                after.docs,
                if after.wb == before.wb { "" } else { "\n  and the workbook differs" }
            ));
        }
        if !Arc::ptr_eq(&now, &wb) {
            return Err(format!("{}: the workbook was replaced by an equal copy", what()));
        }
    }
    Ok(())
}
/// P8: `vba.modules` and `vba.source`, valid or not, leave the workbook (by value and by
/// identity), the undo and redo lengths, the dirty flag, the revision, the selection and the
/// journal exactly as they were.
///
/// **Validates: Requirements 9.10**
#[test]
fn p8_vba_queries_change_nothing() {
    let (seeds, _) = seed_list(1);
    for &seed in &seeds {
        if let Ok(Err(why)) | Err(why) = guarded(|| p8_case(seed)) {
            panic!("P8 failed for seed {seed}: {why}\n  {}", replay(seed, "p8_vba_queries_change_nothing"));
        }
    }
}
/// The `xl/vbaProject.bin` part of the package `bytes`, read straight from the zip.
fn vba_part(bytes: &[u8]) -> Result<Option<Vec<u8>>, String> {
    use std::io::Read;
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).map_err(|e| format!("the saved package isn't a zip: {e}"))?;
    let Ok(mut f) = z.by_name("xl/vbaProject.bin") else { return Ok(None) };
    let mut out = Vec::new();
    f.read_to_end(&mut out).map_err(|e| format!("xl/vbaProject.bin didn't read: {e}"))?;
    Ok(Some(out))
}
/// Saves the active workbook in memory as each macro-enabled kind and checks that the saved
/// `xl/vbaProject.bin` is `want`, byte for byte (no part when there's no project).
fn check_saved(s: &Session, want: Option<&[u8]>, stage: &str) -> Result<(), String> {
    let wb = &s.doc().map_err(|e| format!("no document: {e}"))?.wb;
    for path in ["Saved.xlsm", "Saved.xltm"] {
        let bytes = crate::io::save_bytes(wb, path).map_err(|e| format!("{stage}: saving {path} failed: {e}"))?;
        let got = vba_part(&bytes)?;
        if got.as_deref() != want {
            let len = |b: Option<&[u8]>| b.map_or_else(|| "no part".to_owned(), |b| format!("{} bytes", b.len()));
            let first = match (got.as_deref(), want) {
                (Some(g), Some(w)) => {
                    g.iter().zip(w).position(|(a, b)| a != b).map_or_else(String::new, |i| format!(", first differing byte at {i}"))
                }
                _ => String::new(),
            };
            return Err(format!("{stage}: {path} saved {} of vbaProject.bin, the input had {}{first}", len(got.as_deref()), len(want)));
        }
    }
    Ok(())
}
/// P9 for one seed.
fn p9_case(seed: u64) -> Result<(), String> {
    let c = case(seed);
    let mut rng = Rng::new(seed, PROPS_STREAM ^ 0x9);
    let (name, kind) = if rng.chance(50) { ("m.xlsm", XlsxKind::MacroWorkbook) } else { ("m.xltm", XlsxKind::MacroTemplate) };
    let want = c.bin.as_deref();
    let mut s = open(name, &package(want, kind)?)?;
    let opened = s.doc().map_err(|e| format!("no document: {e}"))?.wb.vba_project.as_deref().map(Vec::as_slice);
    if opened != want {
        return Err(format!("opening {name} ({:?} project) didn't keep the bytes in the model", c.class));
    }
    check_saved(&s, want, &format!("open → save ({:?} project, from {name})", c.class))?;
    for _ in 0..rng.range(1, 6) {
        let (id, params) = query(&mut rng, &c.names);
        let _ = s.execute(id, params);
    }
    check_saved(&s, want, &format!("after vba queries ({:?} project, from {name})", c.class))?;
    let cell = format!("{}{}", ["A", "B", "C"].get(rng.below(3)).copied().unwrap_or("A"), rng.range(1, 50));
    s.execute("cell.set", json!({"cell": cell, "input": format!("{}", rng.below(1000))})).map_err(|e| format!("cell.set: {e}"))?;
    s.execute("vba.modules", json!({})).map_err(|e| format!("vba.modules: {e}"))?;
    s.execute("edit.undo", json!({})).map_err(|e| format!("edit.undo: {e}"))?;
    let _ = s.execute("vba.source", json!({"module": c.names.first().map_or("Module1", String::as_str)}));
    check_saved(&s, want, &format!("after a cell edit and undo ({:?} project, from {name})", c.class))
}
/// P9: opening a workbook with a `vbaProject.bin` and saving it as .xlsm or .xltm writes the
/// same bytes back, also after the VBA queries read (and cached) the project and after an
/// unrelated edit and its undo. Bytes that aren't a readable project are kept just the same.
///
/// **Validates: Requirements 5.12**
#[test]
fn p9_vba_project_bytes_survive_open_and_save() {
    let (seeds, _) = seed_list(1);
    for &seed in &seeds {
        if let Ok(Err(why)) | Err(why) = guarded(|| p9_case(seed)) {
            panic!("P9 failed for seed {seed}: {why}\n  {}", replay(seed, "p9_vba_project_bytes_survive_open_and_save"));
        }
    }
}
/// The generator reaches every class within the default seeds, so a property can't pass by
/// only ever seeing one kind of project.
#[test]
fn props_cases_cover_every_class() {
    let classes: Vec<Class> = (0..64).map(|seed| case(seed).class).collect();
    for want in [Class::None, Class::Clean, Class::Corrupt, Class::Fuzzed, Class::Garbage] {
        assert!(classes.contains(&want), "no {want:?} case in seeds 0..64");
    }
}
