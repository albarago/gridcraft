//! Formulas that call VBA functions keep the results saved in the file while nothing they read
//! changes, and show `#BLOCKED!` once something does (Req 17.3–17.10). The workbooks are built
//! here and written with the xlsx writer, so the saved results are read back from real files.

use std::sync::Arc;

use gridcraft_core::{CellError, CellRef, Value};
use gridcraft_model::{Cell, Formula, Workbook};
use gridcraft_vba::ModuleKind;
use gridcraft_vba::fixture::{Fixture, FixtureModule};
use serde_json::{Value as Json, json};

use crate::Session;

/// A project with a standard module declaring `Public Function Twice(x)`.
fn twice_project() -> Vec<u8> {
    let src = "Attribute VB_Name = \"Module1\"\r\nPublic Function Twice(x)\r\n  Twice = x * 2\r\nEnd Function\r\n";
    Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, src)]).to_vba_project_bin().unwrap()
}

/// Sheet1 with `cells` (A1 notation, value, formula or none), every formula keeping the value
/// given as its saved result.
fn workbook(cells: &[(&str, Value, Option<&str>)]) -> Workbook {
    let mut wb = Workbook::new();
    let sh = wb.sheet_mut(0).unwrap();
    for (at, v, f) in cells {
        let cell = match f {
            Some(f) => Cell { value: v.clone(), ..Cell::formula(Formula::new(f)) },
            None => Cell::value(v.clone()),
        };
        sh.cells.set(CellRef::parse(at).unwrap(), cell);
    }
    wb
}

/// `wb` written as `name` (.xlsx, or .xlsm with its project) and opened through `file.open`.
fn open(wb: &Workbook, name: &str) -> Session {
    let kind = if wb.vba_project.is_some() { gridcraft_xlsx::XlsxKind::MacroWorkbook } else { gridcraft_xlsx::XlsxKind::Workbook };
    let bytes = gridcraft_xlsx::write_xlsx_as(wb, kind).unwrap();
    let mut s = Session::new();
    s.execute("file.open", json!({"name": name, "base64": crate::io::base64_encode(&bytes)})).unwrap();
    s
}

/// The `.xlsm` of the task: A1 = 2, B1 `=Twice(A1)` saved as 4, C1 `=B1+1` saved as 5.
fn twice_xlsm() -> Session {
    let mut wb =
        workbook(&[("A1", Value::Number(2.0), None), ("B1", Value::Number(4.0), Some("=Twice(A1)")), ("C1", Value::Number(5.0), Some("=B1+1"))]);
    wb.vba_project = Some(Arc::new(twice_project()));
    open(&wb, "twice.xlsm")
}

fn get(s: &mut Session, at: &str) -> Json {
    s.execute("cell.get", json!({"cell": at})).unwrap()
}

fn value(s: &mut Session, at: &str) -> Value {
    s.doc().unwrap().wb.sheet(0).unwrap().value(CellRef::parse(at).unwrap())
}

fn set(s: &mut Session, at: &str, input: &str) {
    s.execute("cell.set", json!({"cell": at, "input": input})).unwrap();
}

const BLOCKED: Value = Value::Error(CellError::Blocked);

/// Req 17.3, 17.4: an `.xlsx` with `_xludf.` calls shows each saved result with its type, in
/// `cell.get` and `sheet.read` alike; a formula saved without a result is `#BLOCKED!`.
#[test]
fn an_xlsx_shows_the_saved_results_of_xludf_calls() {
    let wb = workbook(&[
        ("A1", Value::Number(21.0), None),
        ("B1", Value::Number(42.0), Some("=_xludf.Twice(A1)")),
        ("B2", Value::text("forty-two"), Some("=_xludf.Name(A1)")),
        ("B3", Value::Bool(true), Some("=_xludf.Flag(A1)")),
        ("B4", Value::Error(CellError::Div0), Some("=_xludf.Fails(A1)")),
        ("B5", Value::Empty, Some("=_xludf.Twice(A1)")),
    ]);
    let mut s = open(&wb, "udf.xlsx");
    let want = [
        ("B1", json!(42.0), "number"),
        ("B2", json!("forty-two"), "text"),
        ("B3", json!(true), "boolean"),
        ("B4", json!({"error": "#DIV/0!"}), "error"),
        ("B5", json!({"error": "#BLOCKED!"}), "error"),
    ];
    for (at, v, ty) in &want {
        let g = get(&mut s, at);
        assert_eq!((&g["value"], g["type"].as_str()), (v, Some(*ty)), "{at}: {g}");
    }
    assert_eq!(get(&mut s, "B1")["formula"], "=_xludf.TWICE(A1)");
    assert_eq!(get(&mut s, "B1")["udfCalls"], json!(["_xludf.TWICE"]));
    assert_eq!(get(&mut s, "A1")["udfCalls"], Json::Null);
    let rows = s.execute("sheet.read", json!({"range": "B1:B5"})).unwrap();
    assert_eq!(rows.to_string().matches("forty-two").count(), 1, "{rows}");
}

/// Req 17.3, 17.5: an `.xlsm` whose project declares `Twice` keeps `=Twice(A1)`'s saved 4
/// through Calculate Now and volatile recalculation elsewhere; nothing about it is dynamic.
#[test]
fn an_xlsm_keeps_saved_results_through_full_and_volatile_recalculation() {
    let mut s = twice_xlsm();
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (Value::Number(4.0), Value::Number(5.0)));
    assert_eq!(get(&mut s, "B1")["udfCalls"], json!(["TWICE"]));
    s.execute("formulas.calculateNow", json!({})).unwrap();
    assert_eq!(value(&mut s, "B1"), Value::Number(4.0));
    set(&mut s, "D1", "=NOW()");
    set(&mut s, "D2", "=RAND()");
    set(&mut s, "E1", "7");
    s.execute("formulas.calculateNow", json!({})).unwrap();
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (Value::Number(4.0), Value::Number(5.0)));
    // A structural edit elsewhere recalculates everything; the saved result stays.
    s.execute("home.insertRows", json!({"rows": "10:10"})).unwrap();
    assert_eq!(value(&mut s, "B1"), Value::Number(4.0));
}

/// Req 17.6, 17.9: a precedent edit blocks the cell and its reader; undo brings both back.
#[test]
fn a_precedent_edit_blocks_and_undo_restores_the_saved_result() {
    let mut s = twice_xlsm();
    set(&mut s, "A1", "3");
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (BLOCKED, BLOCKED));
    s.execute("formulas.calculateNow", json!({})).unwrap();
    assert_eq!(value(&mut s, "B1"), BLOCKED, "a full recalculation doesn't bring it back");
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (Value::Number(4.0), Value::Number(5.0)));
    s.execute("formulas.calculateNow", json!({})).unwrap();
    assert_eq!(value(&mut s, "B1"), Value::Number(4.0));
    s.execute("edit.redo", json!({})).unwrap();
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (BLOCKED, BLOCKED));
}

/// Req 17.7, 17.9: an edited formula that still calls `Twice` is blocked; one that doesn't is
/// calculated; undoing both steps brings back the saved result.
#[test]
fn an_edited_formula_is_blocked_while_it_calls_a_udf() {
    let mut s = twice_xlsm();
    set(&mut s, "B1", "=Twice(A1)+0");
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (BLOCKED, BLOCKED));
    set(&mut s, "B1", "=A1*2");
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (Value::Number(4.0), Value::Number(5.0)));
    assert_eq!(get(&mut s, "B1")["udfCalls"], Json::Null);
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(value(&mut s, "B1"), BLOCKED);
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (Value::Number(4.0), Value::Number(5.0)));
    assert_eq!(get(&mut s, "B1")["formula"], "=TWICE(A1)");
}

/// Req 17.2: a defined name `Twice` (a LAMBDA) wins over the VBA function of the same name.
#[test]
fn a_defined_name_wins_over_a_vba_function() {
    let mut wb = workbook(&[("A1", Value::Number(2.0), None), ("B1", Value::Number(99.0), Some("=Twice(A1)"))]);
    wb.vba_project = Some(Arc::new(twice_project()));
    wb.names.push(gridcraft_model::DefinedName {
        name: "Twice".into(),
        scope: None,
        formula: "LAMBDA(x,x*2)".into(),
        comment: String::new(),
        hidden: false,
    });
    let mut s = open(&wb, "named.xlsm");
    assert_eq!(value(&mut s, "B1"), Value::Number(4.0), "evaluated, not the saved 99");
    assert_eq!(get(&mut s, "B1")["udfCalls"], Json::Null);
    assert_eq!(s.vba_cache.parses, 0, "no formula needed the project");
}

/// Lazy names: a macro workbook whose formulas call only built-ins is opened without reading its
/// project; typing a call to one of its functions reads it then, so the call is `#BLOCKED!`.
#[test]
fn typing_a_vba_function_call_reads_the_project_names_then() {
    let mut wb = workbook(&[("A1", Value::Number(2.0), None), ("B1", Value::Number(2.0), Some("=SUM(A1)"))]);
    wb.vba_project = Some(Arc::new(twice_project()));
    let mut s = open(&wb, "lazy.xlsm");
    assert_eq!(s.vba_cache.parses, 0);
    set(&mut s, "C1", "=Twice(1)");
    assert_eq!(value(&mut s, "C1"), BLOCKED);
    assert_eq!(s.vba_cache.parses, 1);
    assert_eq!(get(&mut s, "C1")["udfCalls"], json!(["TWICE"]));
    // A typo is #NAME? as before, and doesn't read the project again.
    set(&mut s, "C2", "=TWICEE(1)");
    assert_eq!(value(&mut s, "C2"), Value::Error(CellError::Name));
    assert_eq!(s.vba_cache.parses, 1);
    // Without a project, an unknown call stays #NAME?.
    let mut s = open(&workbook(&[("A1", Value::Number(2.0), None)]), "plain.xlsx");
    set(&mut s, "C1", "=Twice(1)");
    assert_eq!(value(&mut s, "C1"), Value::Error(CellError::Name));
}

/// Req 17.10: across a run of edits, undos and recalculations, `B1` only ever shows its saved
/// result or `#BLOCKED!`: no VBA runs, so nothing else can appear.
#[test]
fn a_foreign_cell_only_shows_its_saved_result_or_blocked() {
    let mut s = twice_xlsm();
    let mut state = 0x9E37_79B9_7F4A_7C15u64;
    for step in 0..200 {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        match state % 6 {
            0 => set(&mut s, "A1", &format!("{}", state % 10)),
            1 => set(&mut s, "E1", "=NOW()"),
            2 => drop(s.execute("edit.undo", json!({}))),
            3 => drop(s.execute("edit.redo", json!({}))),
            4 => drop(s.execute("formulas.calculateNow", json!({}))),
            _ => set(&mut s, "A2", "x"),
        }
        let v = value(&mut s, "B1");
        assert!(v == Value::Number(4.0) || v == BLOCKED, "step {step}: {v:?}");
    }
}

/// The active workbook encoded through `file.saveBytes` (the extension picks the package kind).
fn save(s: &mut Session, format: &str) -> Vec<u8> {
    let r = s.execute("file.saveBytes", json!({"format": format})).unwrap();
    crate::io::base64_decode(r["base64"].as_str().unwrap()).unwrap()
}

fn reopen(bytes: &[u8], name: &str) -> Session {
    let mut s = Session::new();
    s.execute("file.open", json!({"name": name, "base64": crate::io::base64_encode(bytes)})).unwrap();
    s
}

/// The `<c>` element of `at` in the first sheet's XML, as written.
fn cell_xml(bytes: &[u8], at: &str) -> String {
    use std::io::Read;
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut xml = String::new();
    z.by_name("xl/worksheets/sheet1.xml").unwrap().read_to_string(&mut xml).unwrap();
    let start = xml.find(&format!("<c r=\"{at}\"")).unwrap_or_else(|| panic!("{at} not written: {xml}"));
    let rest = &xml[start..];
    let open_end = rest.find('>').unwrap();
    let end = if rest[..open_end].ends_with('/') { open_end + 1 } else { rest.find("</c>").unwrap() + 4 };
    rest[..end].to_string()
}

/// `(formula, value, type)` of `at` through `cell.get`.
fn shown(s: &mut Session, at: &str) -> (Json, Json, Json) {
    let g = get(s, at);
    (g["formula"].clone(), g["value"].clone(), g["type"].clone())
}

/// Req 17.11, 17.13 (.xlsx): `_xludf.` calls saved with a number, text, boolean and error come
/// back from open → save → open with the same formula text and the same result in its type.
#[test]
fn an_xlsx_saves_xludf_calls_with_their_saved_results() {
    let wb = workbook(&[
        ("A1", Value::Number(21.0), None),
        ("B1", Value::Number(42.0), Some("=_xludf.Twice(A1)")),
        ("B2", Value::text("forty-two"), Some("=_xludf.Name(A1)")),
        ("B3", Value::Bool(false), Some("=_xludf.Flag(A1)")),
        ("B4", Value::Error(CellError::NA), Some("=_xludf.Fails(A1)")),
    ]);
    let mut first = open(&wb, "udf.xlsx");
    let bytes = save(&mut first, "xlsx");
    assert_eq!(cell_xml(&bytes, "B1"), "<c r=\"B1\"><f>_xludf.TWICE(A1)</f><v>42</v></c>");
    assert_eq!(cell_xml(&bytes, "B2"), "<c r=\"B2\" t=\"str\"><f>_xludf.NAME(A1)</f><v>forty-two</v></c>");
    assert_eq!(cell_xml(&bytes, "B3"), "<c r=\"B3\" t=\"b\"><f>_xludf.FLAG(A1)</f><v>0</v></c>");
    assert_eq!(cell_xml(&bytes, "B4"), "<c r=\"B4\" t=\"e\"><f>_xludf.FAILS(A1)</f><v>#N/A</v></c>");
    let mut again = reopen(&bytes, "udf.xlsx");
    for at in ["B1", "B2", "B3", "B4"] {
        assert_eq!(shown(&mut again, at), shown(&mut first, at), "{at}");
    }
    assert_eq!(shown(&mut again, "B2"), (json!("=_xludf.NAME(A1)"), json!("forty-two"), json!("text")));
}

/// Req 17.11, 17.13 (.xlsm): `=Twice(A1)`, a function of the project, is saved bare with its 4,
/// and the project's bytes travel unchanged.
#[test]
fn an_xlsm_saves_vba_calls_bare_with_their_saved_results_and_the_project_unchanged() {
    let mut s = twice_xlsm();
    let bytes = save(&mut s, "xlsm");
    assert_eq!(cell_xml(&bytes, "B1"), "<c r=\"B1\"><f>TWICE(A1)</f><v>4</v></c>");
    assert_eq!(cell_xml(&bytes, "C1"), "<c r=\"C1\"><f>B1+1</f><v>5</v></c>");
    let mut again = reopen(&bytes, "twice.xlsm");
    assert_eq!(shown(&mut again, "B1"), (json!("=TWICE(A1)"), json!(4.0), json!("number")));
    assert_eq!(get(&mut again, "B1")["udfCalls"], json!(["TWICE"]));
    assert_eq!(again.doc().unwrap().wb.vba_project.as_deref().map(Vec::as_slice), Some(twice_project().as_slice()));
}

/// Req 17.12, 17.4: after a precedent edit the VBA call and the cell reading it are saved as
/// formulas with no cached value; reopened, the call has no saved result, so it is `#BLOCKED!`
/// and so is its reader. Undo before saving brings the saved results back into the file.
#[test]
fn a_blocked_vba_call_and_its_reader_are_saved_with_no_cached_value() {
    let mut s = twice_xlsm();
    set(&mut s, "A1", "3");
    let bytes = save(&mut s, "xlsm");
    assert_eq!(cell_xml(&bytes, "B1"), "<c r=\"B1\"><f>TWICE(A1)</f></c>");
    assert_eq!(cell_xml(&bytes, "C1"), "<c r=\"C1\"><f>B1+1</f></c>");
    let mut again = reopen(&bytes, "twice.xlsm");
    assert_eq!((value(&mut again, "B1"), value(&mut again, "C1")), (BLOCKED, BLOCKED));
    assert_eq!(get(&mut again, "B1")["formula"], "=TWICE(A1)");
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!(cell_xml(&save(&mut s, "xlsm"), "B1"), "<c r=\"B1\"><f>TWICE(A1)</f><v>4</v></c>");
}

/// #2139's case (Req 17.1, 17.3, 17.6, 17.11–17.13, .xls): an Excel 97-2003 workbook whose cells
/// call the VBA function `DIAGONALWERT` (a Lbl flagged fFunc|fProc) shows the saved number and
/// text instead of `#NAME?`; editing an argument blocks that call and its reader only; saving as
/// `.xlsx` writes `DIAGONALWERT(...)` with the saved results, and no value for the blocked ones.
#[test]
fn xls_vba_function_cells_keep_saved_results_issue_2139() {
    use gridcraft_xlsx::testing::{XlsCell, xls_calling_vba_function};
    let c = |a1: &str| CellRef::parse(a1).unwrap();
    let bytes = xls_calling_vba_function(
        "DIAGONALWERT",
        &[
            (c("A1"), XlsCell::Number(3.0)),
            (c("A2"), XlsCell::Number(4.0)),
            (c("B1"), XlsCell::Call { arg: c("A1"), saved: Value::Number(6.0) }),
            (c("B2"), XlsCell::Call { arg: c("A2"), saved: Value::text("d4") }),
            (c("C1"), XlsCell::Times { of: c("B1"), k: 10, saved: Value::Number(60.0) }),
        ],
    );
    let mut s = reopen(&bytes, "issue2139.xls");
    assert_eq!(shown(&mut s, "B1"), (json!("=DIAGONALWERT(A1)"), json!(6.0), json!("number")));
    assert_eq!(shown(&mut s, "B2"), (json!("=DIAGONALWERT(A2)"), json!("d4"), json!("text")));
    assert_eq!(value(&mut s, "C1"), Value::Number(60.0));
    assert_eq!(get(&mut s, "B1")["udfCalls"], json!(["DIAGONALWERT"]));
    s.execute("formulas.calculateNow", json!({})).unwrap();
    assert_eq!((value(&mut s, "B1"), value(&mut s, "B2")), (Value::Number(6.0), Value::text("d4")));

    let clean = save(&mut s, "xlsx");
    assert_eq!(cell_xml(&clean, "B1"), "<c r=\"B1\"><f>DIAGONALWERT(A1)</f><v>6</v></c>");
    assert_eq!(cell_xml(&clean, "B2"), "<c r=\"B2\" t=\"str\"><f>DIAGONALWERT(A2)</f><v>d4</v></c>");

    set(&mut s, "A1", "5");
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (BLOCKED, BLOCKED));
    assert_eq!(value(&mut s, "B2"), Value::text("d4"), "a call whose argument didn't change keeps its result");
    let edited = save(&mut s, "xlsx");
    assert_eq!(cell_xml(&edited, "B1"), "<c r=\"B1\"><f>DIAGONALWERT(A1)</f></c>");
    assert_eq!(cell_xml(&edited, "C1"), "<c r=\"C1\"><f>B1*10</f></c>");
    assert_eq!(cell_xml(&edited, "B2"), "<c r=\"B2\" t=\"str\"><f>DIAGONALWERT(A2)</f><v>d4</v></c>");
    s.execute("edit.undo", json!({})).unwrap();
    assert_eq!((value(&mut s, "B1"), value(&mut s, "C1")), (Value::Number(6.0), Value::Number(60.0)));
}
/// Req 17.11: a VBA function given a bare range (`AddUp(A7:A8,B1)`, `_xludf.AddUp(A1:A3)`)
/// receives the range itself, so the call is saved as the plain formula it was read as, never
/// as a dynamic-array formula (`cm`, `<f t="array">`), with its saved result; reopened, it shows
/// the same formula and result.
#[test]
fn a_vba_call_with_a_range_argument_is_saved_as_a_plain_formula() {
    let src = "Attribute VB_Name = \"Module1\"\r\nPublic Function AddUp(r, x)\r\n  AddUp = 0\r\nEnd Function\r\n";
    let bin = Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, src)]).to_vba_project_bin().unwrap();
    let mut wb = workbook(&[
        ("A7", Value::Number(1.0), None),
        ("A8", Value::Number(2.0), None),
        ("B1", Value::Number(3.0), None),
        ("C3", Value::Number(7.0), Some("=1+AddUp(A7:A8,B1)")),
    ]);
    wb.vba_project = Some(Arc::new(bin));
    // The file as Excel saves it: a plain formula.
    wb.udf_names = vec!["ADDUP".into()];
    let mut s = open(&wb, "addup.xlsm");
    assert_eq!(value(&mut s, "C3"), Value::Number(7.0));
    let bytes = save(&mut s, "xlsm");
    assert_eq!(cell_xml(&bytes, "C3"), "<c r=\"C3\"><f>1+ADDUP(A7:A8,B1)</f><v>7</v></c>");
    let mut again = reopen(&bytes, "addup.xlsm");
    assert_eq!(shown(&mut again, "C3"), (json!("=1+ADDUP(A7:A8,B1)"), json!(7.0), json!("number")));

    let wb = workbook(&[("A1", Value::Number(1.0), None), ("B1", Value::Number(6.0), Some("=_xludf.AddUp(A1:A3)"))]);
    let mut s = open(&wb, "addup.xlsx");
    let bytes = save(&mut s, "xlsx");
    assert_eq!(cell_xml(&bytes, "B1"), "<c r=\"B1\"><f>_xludf.ADDUP(A1:A3)</f><v>6</v></c>");
    let mut again = reopen(&bytes, "addup.xlsx");
    assert_eq!(shown(&mut again, "B1"), (json!("=_xludf.ADDUP(A1:A3)"), json!(6.0), json!("number")));
}
