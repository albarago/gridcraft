//! Capability_Manifest tests (Requirement 14).
use super::triggers::TRIGGERS;
use super::{CapSpan, Capability, CompatibilityReport, EmptyCatalog, IncompleteReason, analyze};
use crate::fixture::{Fixture, FixtureModule};
use crate::project::{Limits, VbaModule, VbaProject, read_project};
use crate::project_stream::ModuleKind;

fn module(name: &str, kind: ModuleKind, source: &str) -> VbaModule {
    VbaModule {
        name: name.to_owned(),
        code_name: name.to_owned(),
        stream_name: name.to_owned(),
        kind,
        text_offset: 0,
        source: source.to_owned(),
        diagnostics: Vec::new(),
    }
}

fn report(modules: Vec<VbaModule>) -> CompatibilityReport {
    analyze(&VbaProject { modules, ..VbaProject::default() }, &EmptyCatalog)
}

/// The present categories of `body` run inside a `Sub` in a standard module, which must parse
/// and use no dynamic dispatch.
fn present(body: &str) -> Vec<Capability> {
    let r = report(vec![module("M", ModuleKind::Standard, &format!("Sub T()\n{body}\nEnd Sub\n"))]);
    assert!(r.unparsed.is_empty() && !r.manifest.incomplete, "{body}: {:?}", r.manifest.incomplete_reasons);
    r.manifest.categories.iter().filter(|c| c.present).map(|c| c.category).collect()
}

fn cap_span(module: &str, start: (u32, u32), end: (u32, u32), construct: &'static str) -> CapSpan {
    CapSpan { module: module.to_owned(), start_line: start.0, start_col: start.1, end_line: end.0, end_col: end.1, construct }
}

#[test]
fn nine_categories_in_fixed_order_with_stable_names() {
    let r = report(vec![module("M", ModuleKind::Standard, "Sub A()\nx = 1\nEnd Sub\n")]);
    let categories: Vec<Capability> = r.manifest.categories.iter().map(|c| c.category).collect();
    assert_eq!(categories, Capability::ALL);
    assert!(r.manifest.categories.iter().all(|c| !c.present && c.spans.is_empty()));
    assert!(!r.manifest.incomplete && r.manifest.incomplete_reasons.is_empty() && r.manifest.dynamic.is_empty());
    let names: Vec<&str> = Capability::ALL.iter().map(|c| c.as_str()).collect();
    assert_eq!(names, ["fileRead", "fileWrite", "network", "process", "registry", "clipboard", "uiDialogs", "eventHandlers", "autoExec"]);
    let shell = report(vec![module("M", ModuleKind::Standard, "Sub A()\nShell \"x\"\nEnd Sub\n")]);
    let categories: Vec<Capability> = shell.manifest.categories.iter().map(|c| c.category).collect();
    assert_eq!(categories, Capability::ALL, "present categories don't change the order");
}

#[test]
fn every_category_has_trigger_rows() {
    for category in Capability::ALL {
        assert!(TRIGGERS.iter().any(|t| t.category == category), "{category:?}");
    }
}

#[test]
fn each_category_has_positive_and_negative_triggers() {
    use Capability::{Clipboard, FileRead, FileWrite, Network, Process, Registry, UiDialogs};
    let cases: &[(&str, &[Capability])] = &[
        // File read and write.
        ("Open \"a\" For Input As #1", &[FileRead]),
        ("Open \"a\" For Output As #1", &[FileWrite]),
        ("Open \"a\" For Append As #1", &[FileWrite]),
        ("Open \"a\" For Binary Access Read As #1", &[FileRead]),
        ("Open \"a\" For Binary Access Read Write Lock Write As #1", &[FileRead, FileWrite]),
        ("Open \"a\" As #1", &[FileRead]),
        ("Line Input #1, s", &[FileRead]),
        ("Get #1, , x", &[FileRead]),
        ("Print #1, \"x\"", &[FileWrite]),
        ("Put #1, , x", &[FileWrite]),
        ("Close #1", &[]),
        ("x = Dir(\"*.txt\")", &[FileRead]),
        ("x = Dir", &[FileRead]),
        ("n = VBA.FileSystem.FileLen(\"a\")", &[FileRead]),
        ("Workbooks.Open \"a.xlsx\"", &[FileRead]),
        ("Set wb = Application.Workbooks.Open(\"a.xlsx\")", &[FileRead]),
        ("rs.Open sql", &[]),
        ("Kill \"a\"", &[FileWrite]),
        ("MkDir \"d\"", &[FileWrite]),
        ("FileCopy \"a\", \"b\"", &[FileWrite]),
        ("ActiveWorkbook.SaveAs \"b.xlsx\"", &[FileWrite]),
        ("Set fso = CreateObject(\"Scripting.FileSystemObject\")", &[FileRead, FileWrite]),
        // Network.
        ("Set h = CreateObject(\"MSXML2.XMLHTTP.6.0\")", &[Network]),
        ("Set h = CreateObject(\"winhttp.WinHttpRequest.5.1\")", &[Network]),
        ("Set h = CreateObject(\"Microsoft.XMLHTTP\")", &[Network]),
        ("Set d = CreateObject(\"Scripting.Dictionary\")", &[]),
        ("Set d = CreateObject(\"MSXML2X.Thing\")", &[]),
        // Process.
        ("Shell \"calc.exe\"", &[Process]),
        ("Set sh = CreateObject(\"WScript.Shell\")", &[Process]),
        ("Set sh = CreateObject(\"WScript.ShellX\")", &[]),
        ("MyLib.Shell \"x\"", &[]),
        // Registry.
        ("s = GetSetting(\"A\", \"B\", \"C\")", &[Registry]),
        ("SaveSetting \"A\", \"B\", \"C\", \"D\"", &[Registry]),
        ("DeleteSetting \"A\"", &[Registry]),
        ("v = sh.RegRead(\"HKCU\\x\")", &[Registry]),
        ("s = Settings(\"A\")", &[]),
        // Clipboard.
        ("Set d = New DataObject", &[Clipboard]),
        ("Dim d As MSForms.DataObject", &[Clipboard]),
        ("d.PutInClipboard", &[Clipboard]),
        ("Dim d As DataObjectX", &[]),
        // UI dialogs.
        ("MsgBox \"hi\"", &[UiDialogs]),
        ("x = InputBox(\"q\")", &[UiDialogs]),
        ("x = Application.InputBox(\"q\")", &[UiDialogs]),
        ("UserForm1.Show", &[UiDialogs]),
        ("f = Application.GetOpenFilename()", &[UiDialogs]),
        ("With UserForm1\n.Show\nEnd With", &[UiDialogs]),
        ("Debug.Print \"x\"", &[]),
        ("x = Show", &[]),
    ];
    for (body, expected) in cases {
        assert_eq!(present(body), *expected, "{body}");
    }
}

#[test]
fn declare_is_process() {
    let src = "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long\nSub A()\nEnd Sub\n";
    let r = report(vec![module("M", ModuleKind::Standard, src)]);
    let process = &r.manifest.categories[Capability::Process.index()];
    assert_eq!(process.category, Capability::Process);
    assert_eq!(process.spans, [cap_span("M", (1, 1), (1, 71), "Declare")]);
}

#[test]
fn event_handlers_and_auto_exec_come_from_the_marks() {
    let wb = "Private Sub Workbook_Open()\nEnd Sub\nPrivate Sub Workbook_BeforeClose(Cancel As Boolean)\nEnd Sub\n";
    let mut this_workbook = module("ThisWorkbook", ModuleKind::Document, wb);
    this_workbook.code_name = "ThisWorkbook".to_owned();
    let r = report(vec![
        this_workbook,
        module("Sheet1", ModuleKind::Document, "Sub Worksheet_Change(ByVal t As Range)\nEnd Sub\n"),
        module("M", ModuleKind::Standard, "Sub Auto_Close()\nEnd Sub\nSub Workbook_Open()\nEnd Sub\n"),
    ]);
    let spans = |c: Capability| r.manifest.categories[c.index()].spans.clone();
    assert_eq!(
        spans(Capability::EventHandlers),
        [
            cap_span("ThisWorkbook", (1, 1), (2, 7), "event handler"),
            cap_span("ThisWorkbook", (3, 1), (4, 7), "event handler"),
            cap_span("Sheet1", (1, 1), (2, 7), "event handler"),
        ]
    );
    assert_eq!(spans(Capability::AutoExec), [cap_span("ThisWorkbook", (1, 1), (2, 7), "auto-exec"), cap_span("M", (1, 1), (2, 7), "auto-exec")]);
    let none = report(vec![module("M", ModuleKind::Standard, "Sub Workbook_Open()\nEnd Sub\n")]);
    assert!(!none.manifest.categories[Capability::EventHandlers.index()].present);
    assert!(!none.manifest.categories[Capability::AutoExec.index()].present);
}

#[test]
fn spans_give_lines_and_character_columns_of_multi_line_calls() {
    let src = "Sub T()\n    s = \"\u{e9}\": MsgBox \"a\", _\n        vbOKOnly\n    Open \"a\" For Output Access Write As #1\nEnd Sub\n";
    let r = report(vec![module("M", ModuleKind::Standard, src)]);
    assert_eq!(r.manifest.categories[Capability::UiDialogs.index()].spans, [cap_span("M", (2, 14), (3, 16), "MsgBox")]);
    // Two rows match the `Open`; the occurrence counts once, under the first row.
    assert_eq!(r.manifest.categories[Capability::FileWrite.index()].spans, [cap_span("M", (4, 5), (4, 42), "Open For Output")]);
}

#[test]
fn dynamic_dispatch_makes_the_manifest_incomplete() {
    let src = "\
Sub T()
    CallByName o, \"M\", VbMethod
    Application.Run \"Macro1\"
    Application.Run macroName
    Run \"Book.xlsm!M\" & n
    x = Evaluate(\"A1\")
    y = Application.Evaluate(expr)
    z = [A1]
    Set a = CreateObject(progId)
    Set b = GetObject(, cls)
    Set c = GetObject(\"C:\\x.xls\")
    Set d = CreateObject(\"Scripting.Dictionary\")
End Sub
";
    let r = report(vec![module("M", ModuleKind::Standard, src)]);
    let dynamic: Vec<(&str, u32)> = r.manifest.dynamic.iter().map(|d| (d.construct, d.span.start_line)).collect();
    assert_eq!(dynamic, [("CallByName", 2), ("Application.Run", 4), ("Application.Run", 5), ("Evaluate", 7), ("CreateObject", 9), ("GetObject", 10)]);
    let first = &r.manifest.dynamic[0];
    assert_eq!(first.module, "M");
    assert_eq!(first.span, cap_span("M", (2, 5), (2, 31), "CallByName"));
    assert!(r.manifest.incomplete);
    assert_eq!(r.manifest.incomplete_reasons, [IncompleteReason::Dynamic]);
    assert_eq!(IncompleteReason::Dynamic.description(), "dynamic: unknown capabilities");
}

#[test]
fn literal_dispatch_is_not_dynamic() {
    let r = report(vec![module("M", ModuleKind::Standard, "Sub T()\nApplication.Run \"Macro1\"\nx = Evaluate((\"1+1\"))\nEnd Sub\n")]);
    assert!(r.manifest.dynamic.is_empty() && !r.manifest.incomplete);
}

#[test]
fn unparsed_and_unreadable_modules_make_the_manifest_incomplete() {
    let mut stomped = module("Stomped", ModuleKind::Standard, "");
    stomped.diagnostics.push(crate::diag::Diag::module(crate::diag::DiagCode::PossibleStomping, "Stomped", "x"));
    let r = report(vec![
        stomped,
        module("B", ModuleKind::Standard, "Sub U()\nShell \"x\"\nIf\nEnd Sub\n"),
        module("C", ModuleKind::Standard, "Sub V()\nCallByName o, \"M\", VbGet\nEnd Sub\n"),
    ]);
    assert!(r.manifest.incomplete);
    assert_eq!(
        r.manifest.incomplete_reasons,
        [IncompleteReason::Dynamic, IncompleteReason::UnparsedModule("B".to_owned()), IncompleteReason::UnreadableModule("Stomped".to_owned())]
    );
    assert_eq!(r.manifest.incomplete_reasons[1].module(), Some("B"));
    assert_eq!(r.manifest.incomplete_reasons[1].description(), "unparsed module: capabilities unknown");
    assert!(r.manifest.categories[Capability::Process.index()].present, "statements before the error still count");
}

#[test]
fn projects_from_the_fixture_builder_have_a_manifest() {
    let fixture = Fixture::new(vec![
        FixtureModule::new("Module1", ModuleKind::Standard, "Sub Main()\r\n    Shell \"calc\"\r\nEnd Sub\r\n"),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, "Private Sub Workbook_Open()\r\nEnd Sub\r\n"),
    ]);
    let r = analyze(&read_project(&fixture.to_streams(), &Limits::default()), &EmptyCatalog);
    let present: Vec<Capability> = r.manifest.categories.iter().filter(|c| c.present).map(|c| c.category).collect();
    assert_eq!(present, [Capability::Process, Capability::EventHandlers, Capability::AutoExec]);
    assert_eq!(r.manifest.categories[Capability::Process.index()].spans, [cap_span("Module1", (2, 5), (2, 16), "Shell")]);
    assert!(!r.manifest.incomplete);
}
