//! Property P10 (Req 17.5, 17.6): a formula calling a VBA function keeps the result saved in the
//! file while nothing it reads has changed, through any number of full recalculations, unrelated
//! edits, volatile recalculation and structural edits away from it; once a direct or transitive
//! precedent changes it is `#BLOCKED!`, and stays so through full recalculation; undo back to the
//! load state brings the saved result back exactly, and redo blocks it again. No VBA ever runs, so
//! nothing but the saved result or `#BLOCKED!` appears (Req 17.10). Saving writes the saved result
//! as `<v>` for a clean call and no `<v>` for a blocked one (Req 17.11, 17.12).
//!
//! Each seed builds a workbook in code, written as an `.xlsx` (`_xludf.` calls), an `.xlsm`
//! (calls to the project's public functions) or an `.xls` (calls to a Lbl macro name, through
//! `gridcraft_xlsx::testing::xls_calling_vba_function`), and opens it through `file.open`. The
//! layout keeps the generator inside the limits of the foreign-cell design:
//!
//! - inputs in `A1:A8`, plain intermediate formulas in column B (reading column A and earlier B
//!   rows), calls in column C (reading A, B, earlier C rows, ranges `A2:A5` and `SUM(A1:A3)`);
//! - unrelated values in `H1:J8` and volatile formulas in column K, which no call reads;
//! - no call is *dynamic* (no names, tables, INDIRECT/OFFSET/NOW or unknown functions in its
//!   formula or its precedents): a dynamic call blocks on any edit, which is not this property;
//! - structural edits only at row 30 and below or column Y and right, so they never touch a
//!   precedent (a structural edit recalculates everything, which keeps saved results).
//!
//! The model is the undo stack of "cells each edit wrote" (re-entering the same value counts as a
//! change): a call is blocked exactly when one of its precedents (transitively, on the graph as
//! loaded) is among the cells written by the edits on the undo stack. Edits that change a formula
//! only ever change cells that are themselves precedents or unrelated, so the graph as loaded is
//! enough (an edit can't give a call a new precedent without writing one of its old ones first).
//!
//! Knobs: `GRIDCRAFT_FUZZ_SEEDS` (default 64) cases; `GRIDCRAFT_FUZZ_SEED=17` replays one.
use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use gridcraft_core::{CellError, CellRef, Value};
use gridcraft_model::{Cell, Formula, Workbook};
use gridcraft_vba::ModuleKind;
use gridcraft_vba::fixture::{Fixture, FixtureModule};
use gridcraft_vba::testutil::{Rng, guarded, seed_list};
use gridcraft_xlsx::testing::{XlsCell, xls_calling_vba_function};
use serde_json::{Value as Json, json};

use crate::Session;

/// Distinct from the `testutil` streams and `cmd::vba::props`.
const P10_STREAM: u64 = 0x5031_3050_3130_5031;
const BLOCKED: Value = Value::Error(CellError::Blocked);
/// The public functions of the `.xlsm` project (and the `_xludf.` names of the `.xlsx`); none is
/// a built-in.
const VBA_NAMES: [&str; 4] = ["Twice", "Halve", "Mixer", "AddUp"];
/// The `.xls` macro name, as in #2139.
const XLS_NAME: &str = "DIAGONALWERT";
/// Rows of the input, intermediate and unrelated columns.
const ROWS: u32 = 8;
const COL_A: u32 = 0;
const COL_B: u32 = 1;
const COL_C: u32 = 2;
const COL_H: u32 = 7;
const COL_J: u32 = 9;
const COL_K: u32 = 10;
/// Saved errors: everything a file can hold as a cached error.
const ERRORS: [CellError; 7] = [CellError::Null, CellError::Div0, CellError::Value, CellError::Ref, CellError::Name, CellError::Num, CellError::NA];

fn replay(seed: u64) -> String {
    format!(
        "replay: GRIDCRAFT_FUZZ_SEED={seed} cargo test -p gridcraft-engine --lib tests::udf_props::p10_saved_udf_results_survive_until_a_precedent_changes"
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Flavor {
    Xlsx,
    Xlsm,
    Xls,
}

impl Flavor {
    fn file_name(self) -> &'static str {
        match self {
            Flavor::Xlsx => "p10.xlsx",
            Flavor::Xlsm => "p10.xlsm",
            Flavor::Xls => "p10.xls",
        }
    }
    /// The format a save spot-check writes: the same package, or `.xlsx` for an `.xls`.
    fn save_format(self) -> &'static str {
        match self {
            Flavor::Xlsm => "xlsm",
            Flavor::Xlsx | Flavor::Xls => "xlsx",
        }
    }
}

/// A cell calling a VBA function, with the result the file saved for it.
struct Udf {
    at: CellRef,
    formula: String,
    saved: Value,
    /// Every cell it reads, directly or through other formulas, on the graph as loaded.
    closure: BTreeSet<CellRef>,
}

/// A seeded workbook, as file bytes, with its calls.
struct Case {
    flavor: Flavor,
    bytes: Vec<u8>,
    udfs: Vec<Udf>,
}

fn at(row: u32, col: u32) -> CellRef {
    CellRef::new(row, col)
}

fn row(rng: &mut Rng, below: u32) -> u32 {
    u32::try_from(rng.below(usize::try_from(below).unwrap_or(0))).unwrap_or(0)
}

/// `lo..=hi` as a u32.
fn urange(rng: &mut Rng, lo: u32, hi: u32) -> u32 {
    let (lo, hi) = (usize::try_from(lo).unwrap_or(0), usize::try_from(hi).unwrap_or(0));
    u32::try_from(rng.range(lo, hi)).unwrap_or(0)
}

/// A number that every writer prints exactly: a quarter between -500 and 500.
fn number(rng: &mut Rng) -> f64 {
    let quarters = f64::from(u32::try_from(rng.below(4001)).unwrap_or(0));
    (quarters - 2000.0) / 4.0
}

/// A saved result of any type a file holds; the `.xlsx` and `.xlsm` sometimes have none (a
/// formula saved without `<v>`, `#BLOCKED!` from the start).
fn saved(rng: &mut Rng, flavor: Flavor) -> Value {
    match rng.below(if flavor == Flavor::Xls { 4 } else { 5 }) {
        0 => Value::Number(number(rng)),
        1 => Value::text(format!("s{}", rng.below(1000))),
        2 => Value::Bool(rng.chance(50)),
        3 => Value::Error(rng.pick(&ERRORS).copied().unwrap_or(CellError::NA)),
        _ if rng.chance(40) => Value::Empty,
        _ => Value::Number(number(rng)),
    }
}

/// `A{i}:A{j}` with `i <= j`, and its cells.
fn a_range(rng: &mut Rng) -> (String, Vec<CellRef>) {
    let i = row(rng, ROWS);
    let j = urange(rng, i, ROWS - 1);
    (format!("{}:{}", at(i, COL_A).a1(), at(j, COL_A).a1()), (i..=j).map(|r| at(r, COL_A)).collect())
}

/// A formula cell in the generated layout: its text (with `=`) and the cells it reads directly.
struct Plain {
    at: CellRef,
    formula: String,
    reads: Vec<CellRef>,
}

/// One argument of a call: a cell of column A, an intermediate, an earlier call, a range of
/// column A, or `SUM` of one.
fn argument(rng: &mut Rng, intermediates: u32, call_row: u32, flavor: Flavor) -> (String, Vec<CellRef>) {
    let pick = if flavor == Flavor::Xls { rng.below(3) } else { rng.below(5) };
    let cell = match pick {
        1 if intermediates > 0 => at(row(rng, intermediates), COL_B),
        2 if call_row > 0 => at(row(rng, call_row), COL_C),
        3 => return a_range(rng),
        4 => {
            let (text, cells) = a_range(rng);
            return (format!("SUM({text})"), cells);
        }
        _ => at(row(rng, ROWS), COL_A),
    };
    (cell.a1(), vec![cell])
}

fn case(seed: u64) -> Result<Case, String> {
    let mut rng = Rng::new(seed, P10_STREAM);
    let flavor = match rng.below(3) {
        0 => Flavor::Xlsx,
        1 => Flavor::Xlsm,
        _ => Flavor::Xls,
    };
    // Inputs and unrelated values: small integers, so re-entered values collide.
    let mut values: Vec<(CellRef, f64)> = (0..ROWS).map(|r| (at(r, COL_A), f64::from(row(&mut rng, 20)))).collect();
    for col in COL_H..=COL_J {
        for r in 0..ROWS {
            if rng.chance(60) {
                values.push((at(r, col), f64::from(row(&mut rng, 20))));
            }
        }
    }
    // Intermediates in column B, reading column A and earlier B rows.
    let intermediates = urange(&mut rng, 0, 5);
    let mut plain: Vec<Plain> = Vec::new();
    let mut xls_times: Vec<(CellRef, CellRef, u16)> = Vec::new();
    for r in 0..intermediates {
        let here = at(r, COL_B);
        let a = at(row(&mut rng, ROWS), COL_A);
        if flavor == Flavor::Xls {
            let k = u16::try_from(rng.range(1, 9)).unwrap_or(2);
            xls_times.push((here, a, k));
            plain.push(Plain { at: here, formula: format!("={}*{k}", a.a1()), reads: vec![a] });
            continue;
        }
        let (formula, reads) = match rng.below(4) {
            0 => (format!("={}+{}", a.a1(), rng.below(10)), vec![a]),
            1 => {
                let (text, cells) = a_range(&mut rng);
                (format!("=SUM({text})"), cells)
            }
            2 if r > 0 => {
                let b = at(row(&mut rng, r), COL_B);
                (format!("={}*2+{}", b.a1(), a.a1()), vec![b, a])
            }
            _ => {
                let a2 = at(row(&mut rng, ROWS), COL_A);
                (format!("={}*{}", a.a1(), a2.a1()), vec![a, a2])
            }
        };
        plain.push(Plain { at: here, formula, reads });
    }
    // Calls in column C.
    let calls = urange(&mut rng, 1, 6);
    let mut udf_cells: Vec<(Plain, Value)> = Vec::new();
    let mut xls_calls: Vec<(CellRef, CellRef)> = Vec::new();
    for r in 0..calls {
        let here = at(r, COL_C);
        let value = saved(&mut rng, flavor);
        if flavor == Flavor::Xls {
            let (_, reads) = argument(&mut rng, intermediates, r, flavor);
            let arg = reads.first().copied().unwrap_or(at(0, COL_A));
            xls_calls.push((here, arg));
            udf_cells.push((Plain { at: here, formula: format!("={XLS_NAME}({})", arg.a1()), reads }, value));
            continue;
        }
        let name = rng.pick(&VBA_NAMES).copied().unwrap_or("Twice");
        let prefix = if flavor == Flavor::Xlsx { "_xludf." } else { "" };
        let mut reads = Vec::new();
        let mut args = Vec::new();
        for _ in 0..rng.range(1, 3) {
            let (text, cells) = argument(&mut rng, intermediates, r, flavor);
            args.push(text);
            reads.extend(cells);
        }
        let call = format!("{prefix}{name}({})", args.join(","));
        let formula = match rng.below(6) {
            4 => {
                let a = at(row(&mut rng, ROWS), COL_A);
                reads.push(a);
                format!("={call}+{}", a.a1())
            }
            5 => format!("=1+{call}"),
            _ => format!("={call}"),
        };
        udf_cells.push((Plain { at: here, formula, reads }, value));
    }
    // What each formula reads, for the closures.
    let mut reads: BTreeMap<CellRef, Vec<CellRef>> = BTreeMap::new();
    for p in plain.iter().chain(udf_cells.iter().map(|(p, _)| p)) {
        reads.insert(p.at, p.reads.clone());
    }
    let closure = |start: CellRef| {
        let mut seen = BTreeSet::new();
        let mut todo: Vec<CellRef> = reads.get(&start).cloned().unwrap_or_default();
        while let Some(c) = todo.pop() {
            if seen.insert(c) {
                todo.extend(reads.get(&c).cloned().unwrap_or_default());
            }
        }
        seen
    };
    let udfs: Vec<Udf> =
        udf_cells.iter().map(|(p, v)| Udf { at: p.at, formula: p.formula.clone(), saved: v.clone(), closure: closure(p.at) }).collect();

    let bytes = if flavor == Flavor::Xls {
        let mut cells: Vec<(CellRef, XlsCell)> = values.iter().map(|&(c, n)| (c, XlsCell::Number(n))).collect();
        cells.extend(xls_times.iter().map(|&(c, of, k)| (c, XlsCell::Times { of, k, saved: Value::Number(0.0) })));
        cells.extend(xls_calls.iter().zip(&udfs).map(|(&(c, arg), u)| (c, XlsCell::Call { arg, saved: u.saved.clone() })));
        guarded(|| xls_calling_vba_function(XLS_NAME, &cells))?
    } else {
        let mut wb = Workbook::new();
        let sh = wb.sheet_mut(0).ok_or("a new workbook has no sheet")?;
        for &(c, n) in &values {
            sh.cells.set(c, Cell::value(Value::Number(n)));
        }
        for p in &plain {
            sh.cells.set(p.at, Cell::formula(Formula::new(&p.formula)));
        }
        for u in &udfs {
            sh.cells.set(u.at, Cell { value: u.saved.clone(), ..Cell::formula(Formula::new(&u.formula)) });
        }
        // Volatile formulas, and one reading a volatile one, that no call reads.
        sh.cells.set(at(0, COL_K), Cell::formula(Formula::new("=NOW()")));
        sh.cells.set(at(1, COL_K), Cell::formula(Formula::new("=RAND()+K1")));
        sh.cells.set(at(2, COL_K), Cell::formula(Formula::new("=H2*2")));
        let kind = if flavor == Flavor::Xlsm {
            let functions: String =
                VBA_NAMES.iter().map(|n| format!("Public Function {n}(ParamArray xs())\r\n  {n} = 0\r\nEnd Function\r\n")).collect();
            let src = format!("Attribute VB_Name = \"Module1\"\r\n{functions}");
            let bin = Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, &src)])
                .to_vba_project_bin()
                .map_err(|e| format!("the project didn't build: {e:?}"))?;
            wb.vba_project = Some(Arc::new(bin));
            // As the application that saved the file sees them: a range passed to one of these
            // is the range itself, so the fixture's calls are written as plain formulas.
            wb.udf_names = VBA_NAMES.iter().map(|n| n.to_ascii_uppercase()).collect();
            gridcraft_xlsx::XlsxKind::MacroWorkbook
        } else {
            gridcraft_xlsx::XlsxKind::Workbook
        };
        gridcraft_xlsx::write_xlsx_as(&wb, kind).map_err(|e| format!("the workbook didn't write: {e}"))?
    };
    Ok(Case { flavor, bytes, udfs })
}

fn open(name: &str, bytes: &[u8]) -> Result<Session, String> {
    let mut s = Session::new();
    s.execute("file.open", json!({"name": name, "base64": crate::io::base64_encode(bytes)})).map_err(|e| format!("file.open {name} failed: {e}"))?;
    Ok(s)
}

fn value(s: &Session, c: CellRef) -> Result<Value, String> {
    let d = s.doc().map_err(|e| format!("no document: {e}"))?;
    Ok(d.wb.sheet(0).map(|sh| sh.value(c)).unwrap_or(Value::Empty))
}

fn exists(s: &Session, c: CellRef) -> bool {
    s.doc().is_ok_and(|d| d.wb.sheet(0).and_then(|sh| sh.cell(c)).is_some())
}

fn history(s: &Session) -> Result<(usize, usize), String> {
    let d = s.doc().map_err(|e| format!("no document: {e}"))?;
    Ok((d.undo.len(), d.redo.len()))
}

/// The `<c>` element of `at` in the first sheet's XML of the package `bytes`.
fn cell_xml(bytes: &[u8], at: &str) -> Result<String, String> {
    use std::io::Read;
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).map_err(|e| format!("the saved package isn't a zip: {e}"))?;
    let mut xml = String::new();
    z.by_name("xl/worksheets/sheet1.xml")
        .map_err(|e| format!("no sheet1.xml: {e}"))?
        .read_to_string(&mut xml)
        .map_err(|e| format!("sheet1.xml didn't read: {e}"))?;
    let start = xml.find(&format!("<c r=\"{at}\"")).ok_or_else(|| format!("{at} was not written"))?;
    let rest = xml.get(start..).unwrap_or_default();
    let open_end = rest.find('>').ok_or("an unterminated <c>")?;
    let end = if rest.get(..open_end).is_some_and(|t| t.ends_with('/')) { open_end + 1 } else { rest.find("</c>").ok_or("no </c>")? + 4 };
    Ok(rest.get(..end).unwrap_or_default().to_owned())
}

/// The `t` attribute and `<v>` text a saved result is written with.
fn written(v: &Value) -> (&'static str, String) {
    match v {
        Value::Number(n) => ("", format!("<v>{n}</v>")),
        Value::Text(t) => (" t=\"str\"", format!("<v>{t}</v>")),
        Value::Bool(b) => (" t=\"b\"", format!("<v>{}</v>", u8::from(*b))),
        Value::Error(e) => (" t=\"e\"", format!("<v>{e}</v>")),
        _ => ("", String::new()),
    }
}

/// The model: the cells each edit on the undo (and redo) stack wrote.
#[derive(Default)]
struct Model {
    undo: Vec<BTreeSet<CellRef>>,
    redo: Vec<BTreeSet<CellRef>>,
}

impl Model {
    fn expected(&self, u: &Udf) -> Value {
        let touched = self.undo.iter().any(|w| w.iter().any(|c| u.closure.contains(c)));
        if u.saved.is_empty() || touched { BLOCKED } else { u.saved.clone() }
    }
}

/// What the default seeds must reach, so the property can't pass by never blocking or never
/// restoring.
#[derive(Default)]
struct Coverage {
    clean_after_edits: usize,
    blocked: usize,
    restored: usize,
    flavors: BTreeSet<&'static str>,
}

struct Run<'a> {
    case: &'a Case,
    s: Session,
    model: Model,
    log: Vec<String>,
}

impl Run<'_> {
    fn fail(&self, why: String) -> String {
        let start = self.log.len().saturating_sub(12);
        format!(
            "{why}\n  {:?} case, calls {:?}\n  last steps: {}",
            self.case.flavor,
            self.calls(),
            self.log.get(start..).unwrap_or_default().join(" | ")
        )
    }

    fn calls(&self) -> Vec<String> {
        self.case.udfs.iter().map(|u| format!("{} {} saved {:?}", u.at.a1(), u.formula, u.saved)).collect()
    }

    fn exec(&mut self, id: &str, params: Json) -> Result<Json, String> {
        self.log.push(format!("{id} {params}"));
        self.s.execute(id, params.clone()).map_err(|e| self.fail(format!("{id} {params} failed: {e}")))
    }

    /// Runs an edit that writes `wrote`, and pushes it on the model.
    fn edit(&mut self, id: &str, params: Json, wrote: BTreeSet<CellRef>) -> Result<(), String> {
        let (undo, _) = history(&self.s)?;
        self.exec(id, params)?;
        let (now, redo) = history(&self.s)?;
        if now != undo + 1 || redo != 0 {
            return Err(self.fail(format!("the edit left undo {undo} → {now}, redo {redo}: expected one new step and no redo")));
        }
        self.model.undo.push(wrote);
        self.model.redo.clear();
        Ok(())
    }

    /// (a), (b), (d): every call shows what the model says, which is its saved result or
    /// `#BLOCKED!`; and the engine's history has the model's shape.
    fn check(&self, stage: &str, cov: &mut Coverage) -> Result<(), String> {
        let (undo, redo) = history(&self.s)?;
        if (undo, redo) != (self.model.undo.len(), self.model.redo.len()) {
            return Err(self.fail(format!(
                "{stage}: the history is undo {undo} redo {redo}, the model's {} {}",
                self.model.undo.len(),
                self.model.redo.len()
            )));
        }
        for u in &self.case.udfs {
            let got = value(&self.s, u.at)?;
            if got != u.saved && got != BLOCKED {
                return Err(self.fail(format!(
                    "{stage}: {} shows {got:?}, neither its saved result {:?} nor #BLOCKED! (Req 17.10)",
                    u.at.a1(),
                    u.saved
                )));
            }
            let want = self.model.expected(u);
            if got != want {
                let touched: Vec<String> = self.model.undo.iter().flatten().filter(|c| u.closure.contains(c)).map(CellRef::a1).collect();
                return Err(self.fail(format!(
                    "{stage}: {} shows {got:?}, expected {want:?} (precedents {:?}, written on the undo stack: {touched:?})",
                    u.at.a1(),
                    u.closure.iter().map(CellRef::a1).collect::<Vec<_>>()
                )));
            }
            if want == BLOCKED && !u.saved.is_empty() {
                cov.blocked += 1;
            } else if !u.saved.is_empty() && !self.model.undo.is_empty() {
                cov.clean_after_edits += 1;
            }
        }
        Ok(())
    }

    /// (e): the saved package holds each clean call's saved result as `<v>`, in its type, and
    /// no `<v>` for a blocked one; reopened (`.xlsx`, `.xlsm`), each call shows the same.
    fn check_save(&mut self, reopen: bool, stage: &str) -> Result<(), String> {
        let format = self.case.flavor.save_format();
        let r = self.exec("file.saveBytes", json!({"format": format}))?;
        let bytes =
            crate::io::base64_decode(r["base64"].as_str().unwrap_or_default()).ok_or_else(|| self.fail("saveBytes gave bad base64".to_owned()))?;
        for u in &self.case.udfs {
            let xml = cell_xml(&bytes, &u.at.a1()).map_err(|e| self.fail(format!("{stage}: {e}")))?;
            let want = self.model.expected(u);
            // A call with a bare range argument (`AddUp(A7:A8)`) passes the range itself: it is
            // a plain formula, never a dynamic-array one (`cm`, `<f t="array">`), Req 17.11.
            let tag = xml.split('>').next().unwrap_or_default();
            let plain = !tag.contains(" cm=\"") && !xml.contains("<f t=\"array\"") && xml.contains("<f>");
            let ok = plain
                && if want == BLOCKED {
                    !xml.contains("<v>")
                } else {
                    // The type is the `<c>` element's own `t`.
                    let (t, v) = written(&want);
                    let typed = if t.is_empty() { !tag.contains(" t=\"") } else { tag.contains(t) };
                    xml.contains(&v) && typed
                };
            if !ok {
                return Err(self.fail(format!("{stage}: {} is expected {want:?} but was saved as {xml}", u.at.a1())));
            }
        }
        if reopen && self.case.flavor != Flavor::Xls {
            let again = open(self.case.flavor.file_name(), &bytes).map_err(|e| self.fail(e))?;
            for u in &self.case.udfs {
                let (got, want) = (value(&again, u.at)?, self.model.expected(u));
                if got != want {
                    return Err(self.fail(format!("{stage}: reopened, {} shows {got:?}, expected {want:?}", u.at.a1())));
                }
            }
        }
        Ok(())
    }
}

/// A cell an edit writes: an input, an intermediate (or an empty B row), or an unrelated cell.
fn target(rng: &mut Rng) -> CellRef {
    let col = match rng.below(10) {
        0..=4 => COL_A,
        5..=6 => COL_B,
        _ => urange(rng, COL_H, COL_J),
    };
    at(row(rng, ROWS), col)
}

fn input(rng: &mut Rng) -> String {
    if rng.chance(80) { format!("{}", rng.below(20)) } else { format!("t{}", rng.below(5)) }
}

fn column(rng: &mut Rng) -> u32 {
    match rng.below(3) {
        0 => COL_A,
        1 => COL_B,
        _ => COL_H,
    }
}

/// One random step.
fn step(rng: &mut Rng, run: &mut Run) -> Result<(), String> {
    match rng.below(20) {
        // A value edit (sometimes re-entering the value already there).
        0..=3 => {
            let c = target(rng);
            let text = if c.col == COL_B && rng.chance(30) { format!("=A{}+1", rng.range(1, 8)) } else { input(rng) };
            run.edit("cell.set", json!({"cell": c.a1(), "input": text}), BTreeSet::from([c]))?;
        }
        // Clear Contents on 1 to 3 cells of a column: it writes the cells that exist.
        4 => {
            let (col, r0) = (column(rng), row(rng, ROWS));
            let r1 = urange(rng, r0, (r0 + 2).min(ROWS - 1));
            let wrote: BTreeSet<CellRef> = (r0..=r1).map(|r| at(r, col)).filter(|&c| exists(&run.s, c)).collect();
            let range = format!("{}:{}", at(r0, col).a1(), at(r1, col).a1());
            run.edit("edit.clearContents", json!({"range": range}), wrote)?;
        }
        // A text paste: 1 to 3 values down a column, or two across H:J.
        5..=6 => {
            if rng.chance(25) {
                let c = at(row(rng, ROWS), urange(rng, COL_H, COL_J - 1));
                let text = format!("{}\t{}", input(rng), input(rng));
                run.edit("edit.paste", json!({"at": c.a1(), "text": text}), BTreeSet::from([c, at(c.row, c.col + 1)]))?;
            } else {
                let (col, r0) = (column(rng), row(rng, ROWS));
                let n = urange(rng, 1, 3).min(ROWS - r0);
                let text: Vec<String> = (0..n).map(|_| input(rng)).collect();
                let wrote = (r0..r0 + n).map(|r| at(r, col)).collect();
                run.edit("edit.paste", json!({"at": at(r0, col).a1(), "text": text.join("\n")}), wrote)?;
            }
        }
        // Fill Down from a cell that exists over the 1 to 3 cells below it.
        7..=8 => {
            let (col, r0) = (column(rng), row(rng, ROWS - 1));
            let r1 = urange(rng, r0 + 1, (r0 + 3).min(ROWS - 1));
            if exists(&run.s, at(r0, col)) {
                let range = format!("{}:{}", at(r0, col).a1(), at(r1, col).a1());
                run.edit("edit.fillDown", json!({"range": range}), (r0 + 1..=r1).map(|r| at(r, col)).collect())?;
            } else {
                run.exec("formulas.calculateNow", json!({}))?;
            }
        }
        // A volatile or unrelated formula where no call reads.
        9 => {
            let c = at(urange(rng, 3, ROWS - 1), COL_K);
            let f = ["=NOW()", "=RAND()*H1", "=H2+J3", "=TODAY()+K1"].get(rng.below(4)).copied().unwrap_or("=NOW()");
            run.edit("cell.set", json!({"cell": c.a1(), "input": f}), BTreeSet::from([c]))?;
        }
        10..=11 => drop(run.exec("formulas.calculateNow", json!({}))?),
        // A structural edit away from every precedent: it recalculates everything.
        12 => {
            let (id, params) = match rng.below(4) {
                0 => ("home.insertRows", json!({"rows": "30:31"})),
                1 => ("home.deleteRows", json!({"rows": "40:41"})),
                2 => ("home.insertColumns", json!({"cols": "Z"})),
                _ => ("home.deleteColumns", json!({"cols": "Y:Y"})),
            };
            run.edit(id, params, BTreeSet::new())?;
        }
        // Undo and redo (unavailable with nothing to undo or redo: then Calculate Now).
        13..=15 => match run.model.undo.pop() {
            Some(w) => {
                run.exec("edit.undo", json!({}))?;
                run.model.redo.push(w);
            }
            None => drop(run.exec("formulas.calculateNow", json!({}))?),
        },
        16..=17 => match run.model.redo.pop() {
            Some(w) => {
                run.exec("edit.redo", json!({}))?;
                run.model.undo.push(w);
            }
            None => drop(run.exec("formulas.calculateNow", json!({}))?),
        },
        _ => {
            let reopen = rng.chance(30);
            run.check_save(reopen, "a save spot-check")?;
        }
    }
    Ok(())
}

/// P10 for one seed.
fn p10_case(seed: u64, cov: &mut Coverage) -> Result<(), String> {
    let case = case(seed)?;
    let mut rng = Rng::new(seed, P10_STREAM ^ 0x10);
    let s = open(case.flavor.file_name(), &case.bytes)?;
    let mut run = Run { case: &case, s, model: Model::default(), log: Vec::new() };
    cov.flavors.insert(case.flavor.file_name());
    run.check("as opened", cov)?;
    for i in 0..rng.range(8, 30) {
        step(&mut rng, &mut run)?;
        run.check(&format!("step {i}"), cov)?;
    }
    run.check_save(true, "after the steps")?;
    // (c): undo back to the load state brings every saved result back exactly, through a full
    // recalculation too; redo blocks again what the first edit blocked.
    let (undo, _) = history(&run.s)?;
    if undo > 0 {
        run.exec("edit.undo", json!({"steps": undo}))?;
        while let Some(w) = run.model.undo.pop() {
            run.model.redo.push(w);
        }
        run.check("undone to the load state", cov)?;
        for u in &case.udfs {
            let got = value(&run.s, u.at)?;
            let want = if u.saved.is_empty() { BLOCKED } else { u.saved.clone() };
            if got != want {
                return Err(run.fail(format!("undone to the load state, {} shows {got:?}, saved {:?}", u.at.a1(), u.saved)));
            }
        }
        cov.restored += 1;
        run.exec("formulas.calculateNow", json!({}))?;
        run.check("undone, then Calculate Now", cov)?;
        run.check_save(false, "undone to the load state")?;
        run.exec("edit.redo", json!({}))?;
        if let Some(w) = run.model.redo.pop() {
            run.model.undo.push(w);
        }
        run.check("redone one step", cov)?;
        run.exec("formulas.calculateNow", json!({}))?;
        run.check("redone, then Calculate Now", cov)?;
    }
    Ok(())
}

/// P10: a call to a VBA function keeps its saved result while no precedent has changed and is
/// `#BLOCKED!` once one has, through full recalculation, undo and redo, and is saved that way.
///
/// **Validates: Requirements 17.5, 17.6**
#[test]
fn p10_saved_udf_results_survive_until_a_precedent_changes() {
    let (seeds, replaying) = seed_list(1);
    let mut cov = Coverage::default();
    for &seed in &seeds {
        if let Ok(Err(why)) | Err(why) = guarded(|| p10_case(seed, &mut cov)) {
            panic!("P10 failed for seed {seed}: {why}\n  {}", replay(seed));
        }
    }
    if !replaying {
        assert!(cov.flavors.len() == 3, "the seeds reached only {:?}", cov.flavors);
        assert!(
            cov.clean_after_edits > 0 && cov.blocked > 0 && cov.restored > 0,
            "the seeds never reached a clean call after edits ({}), a blocked call ({}) or an undo to the load state ({})",
            cov.clean_after_edits,
            cov.blocked,
            cov.restored
        );
    }
}
