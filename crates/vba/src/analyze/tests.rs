use super::{
    CompatibilityReport, EmptyCatalog, Flag, FlagKind, LineIndex, Mark, MemberCatalog, MemberStatus, MemberUse, ParamInfo, ProgId, Reason, Verdict,
    VerdictRule, analyze,
};
use crate::ast::ProcKind;
use crate::diag::{Diag, DiagCode};
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

fn project(modules: Vec<VbaModule>) -> VbaProject {
    VbaProject { modules, ..VbaProject::default() }
}

fn report(modules: Vec<VbaModule>) -> CompatibilityReport {
    analyze(&project(modules), &EmptyCatalog)
}

/// `(name, marks)` of every procedure.
fn marks_of(report: &CompatibilityReport) -> Vec<(&str, Vec<Mark>)> {
    report.procedures.iter().map(|p| (p.name.as_str(), p.marks.clone())).collect()
}

/// A catalog with fixed answers, keyed by lowercased `object.member`.
struct FixedCatalog(Vec<(&'static str, MemberStatus)>);

impl MemberCatalog for FixedCatalog {
    fn status(&self, object: Option<&str>, member: &str) -> Option<MemberStatus> {
        let key = format!("{}.{member}", object.unwrap_or("")).to_lowercase();
        self.0.iter().find(|(k, _)| *k == key).map(|(_, s)| *s)
    }
}

fn reason(rule: VerdictRule, module: Option<&str>, line: Option<u32>, member: Option<&str>) -> Reason {
    Reason { rule, module: module.map(str::to_owned), line, member: member.map(str::to_owned) }
}

#[test]
fn procedures_report_kind_params_and_lines() {
    let src = "Option Explicit\r\nSub A()\r\nEnd Sub\r\n\r\nFunction F(ByVal n As Long, s$, Optional v, arr() As Integer) As Long\r\n  F = 1\r\nEnd Function\r\n";
    let class = "Property Get P() As Long\nEnd Property\rProperty Let P(x As Long)\nEnd Property\nProperty Set Q(o As Excel.Range)\nEnd Property\n";
    let r = report(vec![module("M", ModuleKind::Standard, src), module("C", ModuleKind::Class, class)]);
    let rows: Vec<_> = r.procedures.iter().map(|p| (p.module.as_str(), p.name.as_str(), p.kind, p.start_line, p.end_line)).collect();
    assert_eq!(
        rows,
        [
            ("M", "A", ProcKind::Sub, 2, 3),
            ("M", "F", ProcKind::Function, 5, 7),
            ("C", "P", ProcKind::PropertyGet, 1, 2),
            ("C", "P", ProcKind::PropertyLet, 3, 4),
            ("C", "Q", ProcKind::PropertySet, 5, 6),
        ]
    );
    let param = |name: &str, ty: &str| ParamInfo { name: name.to_owned(), ty: ty.to_owned() };
    assert_eq!(r.procedures[1].params, [param("n", "Long"), param("s", "String"), param("v", "Variant"), param("arr", "Integer()")]);
    assert_eq!(r.procedures[4].params, [param("o", "Excel.Range")]);
}

#[test]
fn entry_points_need_a_public_parameterless_sub_in_a_standard_module() {
    let src =
        "Sub Main()\nEnd Sub\nPublic Sub Pub()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub\nSub WithArgs(x)\nEnd Sub\nProperty Get P()\nEnd Property\n";
    let r = report(vec![module("M", ModuleKind::Standard, src), module("C", ModuleKind::Class, "Sub InClass()\nEnd Sub\n")]);
    assert_eq!(
        marks_of(&r),
        [
            ("Main", vec![Mark::EntryPoint]),
            ("Pub", vec![Mark::EntryPoint]),
            ("Hidden", vec![]),
            ("WithArgs", vec![]),
            ("P", vec![]),
            ("InClass", vec![]),
        ]
    );
}

#[test]
fn udf_candidates_need_a_public_function_in_a_standard_module() {
    let src = "Function Twice(x)\nEnd Function\nPrivate Function Hidden()\nEnd Function\n";
    let r = report(vec![module("M", ModuleKind::Standard, src), module("C", ModuleKind::Class, "Function InClass()\nEnd Function\n")]);
    assert_eq!(marks_of(&r), [("Twice", vec![Mark::UdfCandidate]), ("Hidden", vec![]), ("InClass", vec![])]);
}
#[test]
fn udf_names_are_the_udf_candidates_upper_cased() {
    let src = "Function Twice(x)\nEnd Function\nPublic Function diagonalWert(n As String)\nEnd Function\nPrivate Function Hidden()\nEnd Function\nSub NotAFunction()\nEnd Sub\n";
    let modules = vec![module("M", ModuleKind::Standard, src), module("C", ModuleKind::Class, "Function InClass()\nEnd Function\n")];
    let names = super::udf_names(&project(modules.clone()));
    assert_eq!(names.into_iter().collect::<Vec<_>>(), ["DIAGONALWERT", "TWICE"]);
    // The same procedures `analyze` marks.
    let marked: Vec<String> =
        report(modules).procedures.iter().filter(|p| p.marks.contains(&Mark::UdfCandidate)).map(|p| p.name.to_ascii_uppercase()).collect();
    assert_eq!(marked, ["TWICE", "DIAGONALWERT"]);
    assert!(super::udf_names(&VbaProject::default()).is_empty());
}

#[test]
fn event_handlers_follow_the_document_module() {
    let wb = "Private Sub Workbook_Open()\nEnd Sub\nPrivate Sub workbook_BeforeClose(Cancel As Boolean)\nEnd Sub\nSub Worksheet_Change(ByVal t As Range)\nEnd Sub\nSub Workbook_()\nEnd Sub\n";
    let sheet = "Private Sub Worksheet_Change(ByVal t As Range)\nEnd Sub\nSub Workbook_Open()\nEnd Sub\n";
    let mut this_workbook = module("Diese Arbeitsmappe", ModuleKind::Document, wb);
    this_workbook.code_name = "thisworkbook".to_owned();
    let r = report(vec![
        this_workbook,
        module("Sheet1", ModuleKind::Document, sheet),
        module("M", ModuleKind::Standard, "Sub Workbook_Open()\nEnd Sub\n"),
    ]);
    assert_eq!(
        marks_of(&r),
        [
            ("Workbook_Open", vec![Mark::EventHandler, Mark::AutoExec]),
            ("workbook_BeforeClose", vec![Mark::EventHandler]),
            ("Worksheet_Change", vec![]),
            ("Workbook_", vec![]),
            ("Worksheet_Change", vec![Mark::EventHandler]),
            ("Workbook_Open", vec![]),
            ("Workbook_Open", vec![Mark::EntryPoint]),
        ]
    );
    assert_eq!(r.procedures[0].module, "Diese Arbeitsmappe");
}

#[test]
fn auto_exec_adds_to_other_marks() {
    let src = "Sub Auto_Open()\nEnd Sub\nPrivate Sub auto_close()\nEnd Sub\n";
    let r = report(vec![module("M", ModuleKind::Standard, src), module("C", ModuleKind::Class, "Sub Auto_Open()\nEnd Sub\n")]);
    assert_eq!(marks_of(&r), [("Auto_Open", vec![Mark::EntryPoint, Mark::AutoExec]), ("auto_close", vec![Mark::AutoExec]), ("Auto_Open", vec![])]);
}

const FLAGGED: &str = "\
Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long
Private WithEvents conn As ADODB.Connection
Private WithEvents btn As MSForms.CommandButton
Implements Foo.IBar
Implements IPlain
Sub T()
    Set a = CreateObject(\"Scripting.Dictionary\")
    Set b = VBA.CreateObject(progId)
    Set c = Interaction.GetObject(, (\"Excel.Application\"))
    Set d = GetObject(\"C:\\x.xls\")
    Set e = VBA.Interaction.CreateObject(Class:=\"WScript.Shell\")
    Shell \"calc.exe\"
    x = VBA.Shell(\"notepad\", _
        1)
    SendKeys \"{ENTER}\"
    Application.SendKeys \"%{F4}\", True
    v = Application.ExecuteExcel4Macro(\"GET.WORKSPACE(1)\")
    ExecuteExcel4Macro \"x\"
    With Application
        .SendKeys \"y\"
    End With
    y = MyLib.Shell(\"x\")
    z = Application.Shell(\"x\")
    Set f = CreateObject(\"A.\" & \"B\")
End Sub
";

#[test]
fn flags_cover_each_construct_with_prog_ids() {
    let r = report(vec![module("M", ModuleKind::Standard, FLAGGED)]);
    let flag = |kind, start_line, end_line, prog_id: Option<ProgId>| Flag { kind, module: "M".to_owned(), start_line, end_line, prog_id };
    let lit = |s: &str| Some(ProgId::Literal(s.to_owned()));
    assert_eq!(
        r.flags,
        [
            flag(FlagKind::Declare, 1, 1, None),
            flag(FlagKind::ActiveX, 2, 2, None),
            flag(FlagKind::ActiveX, 4, 4, None),
            flag(FlagKind::CreateObject, 7, 7, lit("Scripting.Dictionary")),
            flag(FlagKind::CreateObject, 8, 8, Some(ProgId::NonLiteral)),
            flag(FlagKind::GetObject, 9, 9, lit("Excel.Application")),
            flag(FlagKind::GetObject, 10, 10, None),
            flag(FlagKind::CreateObject, 11, 11, lit("WScript.Shell")),
            flag(FlagKind::Shell, 12, 12, None),
            flag(FlagKind::Shell, 13, 14, None),
            flag(FlagKind::SendKeys, 15, 15, None),
            flag(FlagKind::SendKeys, 16, 16, None),
            flag(FlagKind::ExecuteExcel4Macro, 17, 17, None),
            flag(FlagKind::ExecuteExcel4Macro, 18, 18, None),
            flag(FlagKind::SendKeys, 20, 20, None),
            flag(FlagKind::CreateObject, 24, 24, Some(ProgId::NonLiteral)),
        ]
    );
    assert_eq!(ProgId::NonLiteral.as_str(), "non-literal");
    assert_eq!(r.verdict, Verdict::UsesUnsupportedFeatures);
}

#[test]
fn members_count_chains_and_with_targets() {
    let src = "\
Sub T()
    Worksheets(\"A\").Range(\"B1\").Value = 1
    x = ws.range(\"C1\").value
    With ws.Range(\"A1\")
        .Value = 2
        .Font.Bold = True
    End With
    With New Collection
        .Add 1
    End With
    y = rs!Field
    Me.Caption = \"x\"
End Sub
";
    let r = report(vec![module("M", ModuleKind::Standard, src)]);
    let m = |object: Option<&str>, member: &str, count| MemberUse {
        object: object.map(str::to_owned),
        member: member.to_owned(),
        count,
        status: MemberStatus::Unknown,
    };
    assert_eq!(
        r.members,
        [
            m(Some("Range"), "Value", 3),
            m(Some("ws"), "range", 2),
            m(None, "Add", 1),
            m(Some("Font"), "Bold", 1),
            m(Some("Me"), "Caption", 1),
            m(Some("Range"), "Font", 1),
            m(Some("Worksheets"), "Range", 1),
        ]
    );
    assert_eq!(r.verdict, Verdict::AnalysisOnly);
}

#[test]
fn catalog_status_is_used_and_missing_members_are_unknown() {
    let catalog = FixedCatalog(vec![("range.value", MemberStatus::Supported), ("ws.range", MemberStatus::Planned)]);
    let src = "Sub T()\nx = ws.Range(\"A1\").Value\ny = ws.Other\nEnd Sub\n";
    let r = analyze(&project(vec![module("M", ModuleKind::Standard, src)]), &catalog);
    let statuses: Vec<_> = r.members.iter().map(|m| (m.member.as_str(), m.status)).collect();
    assert_eq!(statuses, [("Value", MemberStatus::Supported), ("Other", MemberStatus::Unknown), ("Range", MemberStatus::Planned)]);
}

#[test]
fn verdict_no_macros_for_an_empty_project() {
    let r = analyze(&VbaProject::default(), &EmptyCatalog);
    assert_eq!(r.verdict, Verdict::NoMacros);
    assert!(r.procedures.is_empty() && r.members.is_empty() && r.flags.is_empty() && r.reasons.is_empty() && r.unparsed.is_empty());
    assert!(r.unreadable.is_empty() && !r.manifest.incomplete && r.manifest.categories.iter().all(|c| !c.present));
}

#[test]
fn verdict_would_run_when_everything_is_known() {
    let catalog = FixedCatalog(vec![("ws.range", MemberStatus::Supported)]);
    let src = "Sub T()\nx = 1\nSet r = ws.Range(\"A1\")\nEnd Sub\n";
    let r = analyze(&project(vec![module("M", ModuleKind::Standard, src)]), &catalog);
    assert_eq!(r.verdict, Verdict::WouldRunInFuturePhase);
    assert_eq!(r.reasons, [reason(VerdictRule::AllKnown, None, None, None)]);
}

#[test]
fn verdict_analysis_only_for_unknown_members() {
    let r = report(vec![module("M", ModuleKind::Standard, "Sub T()\nx = ws.Range(\"A1\")\nEnd Sub\n")]);
    assert_eq!(r.verdict, Verdict::AnalysisOnly);
    assert_eq!(r.reasons, [reason(VerdictRule::MemberUnknown, None, None, Some("ws.Range"))]);
}

#[test]
fn verdict_unsupported_for_not_planned_members() {
    let catalog = FixedCatalog(vec![("application.vbe", MemberStatus::NotPlanned)]);
    let r = analyze(&project(vec![module("M", ModuleKind::Standard, "Sub T()\nSet v = Application.VBE\nEnd Sub\n")]), &catalog);
    assert_eq!(r.verdict, Verdict::UsesUnsupportedFeatures);
    assert_eq!(r.reasons, [reason(VerdictRule::MemberNotPlanned, None, None, Some("Application.VBE"))]);
}

#[test]
fn verdict_unsupported_for_flags() {
    let r = report(vec![module("M", ModuleKind::Standard, "Sub T()\n\nShell \"x\"\nEnd Sub\n")]);
    assert_eq!(r.verdict, Verdict::UsesUnsupportedFeatures);
    assert_eq!(r.reasons, [reason(VerdictRule::Flagged(FlagKind::Shell), Some("M"), Some(3), None)]);
}

#[test]
fn unparsed_module_still_reports_good_procedures() {
    let src = "Sub Good()\nEnd Sub\nSub Bad()\nShell \"x\"\nx = (\nEnd Sub\nSub Good2()\nEnd Sub\n";
    let r = report(vec![module("Broken", ModuleKind::Standard, src)]);
    let names: Vec<_> = r.procedures.iter().map(|p| p.name.as_str()).collect();
    assert_eq!(names, ["Good", "Good2"]);
    assert_eq!(r.unparsed, ["Broken"]);
    assert_eq!(r.flags.len(), 1, "statements before the error are still scanned");
    assert_eq!(r.verdict, Verdict::UsesUnsupportedFeatures);
    assert_eq!(
        r.reasons,
        [
            reason(VerdictRule::Flagged(FlagKind::Shell), Some("Broken"), Some(4), None),
            reason(VerdictRule::UnparsedModule, Some("Broken"), Some(5), None)
        ]
    );
}

#[test]
fn reasons_cover_every_condition_in_rule_order() {
    let catalog = FixedCatalog(vec![("application.vbe", MemberStatus::NotPlanned)]);
    let good = "Sub T()\nSet v = Application.VBE\nx = ws.Cells\nSendKeys \"a\"\nEnd Sub\n";
    let r = analyze(&project(vec![module("A", ModuleKind::Standard, good), module("B", ModuleKind::Standard, "Sub U()\nIf\nEnd Sub\n")]), &catalog);
    assert_eq!(r.verdict, Verdict::UsesUnsupportedFeatures);
    assert_eq!(
        r.reasons,
        [
            reason(VerdictRule::Flagged(FlagKind::SendKeys), Some("A"), Some(4), None),
            reason(VerdictRule::MemberNotPlanned, None, None, Some("Application.VBE")),
            reason(VerdictRule::UnparsedModule, Some("B"), Some(2), None),
            reason(VerdictRule::MemberUnknown, None, None, Some("ws.Cells")),
        ]
    );
    let r = analyze(
        &project(vec![module("B", ModuleKind::Standard, "Sub U()\nIf\nEnd Sub\n"), module("C", ModuleKind::Standard, "Sub V()\nx = a.b\nEnd Sub\n")]),
        &catalog,
    );
    assert_eq!(r.verdict, Verdict::AnalysisOnly);
    assert_eq!(
        r.reasons,
        [reason(VerdictRule::UnparsedModule, Some("B"), Some(2), None), reason(VerdictRule::MemberUnknown, None, None, Some("a.b"))]
    );
}

#[test]
fn stomped_or_unreadable_module_is_analysis_only_with_a_reason() {
    let mut stomped = module("Stomped", ModuleKind::Standard, "");
    stomped.diagnostics.push(Diag::module(DiagCode::PossibleStomping, "Stomped", "possible stomping"));
    let mut broken = module("Broken", ModuleKind::Class, "");
    broken.diagnostics.push(Diag::module(DiagCode::ModuleCorrupt, "Broken", "corrupt"));
    let r = report(vec![
        stomped,
        module("M", ModuleKind::Standard, "Sub A()\nEnd Sub\n"),
        broken,
        module("U", ModuleKind::Standard, "Sub U()\nIf\nEnd Sub\n"),
    ]);
    assert!(r.flags.is_empty() && r.members.is_empty());
    assert_eq!(marks_of(&r), [("A", vec![Mark::EntryPoint])]);
    assert_eq!(r.unreadable, ["Stomped", "Broken"]);
    assert_eq!(r.verdict, Verdict::AnalysisOnly);
    assert_eq!(
        r.reasons,
        [
            reason(VerdictRule::UnparsedModule, Some("U"), Some(2), None),
            reason(VerdictRule::UnreadableModule, Some("Stomped"), None, None),
            reason(VerdictRule::UnreadableModule, Some("Broken"), None, None),
        ]
    );
}

#[test]
fn empty_module_with_no_diagnostics_contributes_nothing() {
    let mut info_only = module("Info", ModuleKind::Standard, "");
    info_only.diagnostics.push(Diag::module(DiagCode::CompiledCodeNotChecked, "Info", "not checked"));
    let r = report(vec![module("Empty", ModuleKind::Standard, ""), info_only, module("M", ModuleKind::Standard, "Sub A()\nEnd Sub\n")]);
    assert!(r.unparsed.is_empty() && r.unreadable.is_empty() && r.flags.is_empty() && r.members.is_empty());
    assert_eq!(marks_of(&r), [("A", vec![Mark::EntryPoint])]);
    assert_eq!(r.verdict, Verdict::WouldRunInFuturePhase);
    assert!(!r.manifest.incomplete);
}

#[test]
fn stomped_module_with_only_attribute_lines_is_unreadable() {
    let mut stomped = module("Stomped", ModuleKind::Standard, "Attribute VB_Name = \"Stomped\"\r\n");
    stomped.diagnostics.push(Diag::module(DiagCode::PossibleStomping, "Stomped", "possible stomping"));
    let r = report(vec![stomped, module("M", ModuleKind::Standard, "Sub A()\nEnd Sub\n")]);
    assert_eq!(r.unreadable, ["Stomped"]);
    assert_eq!(r.verdict, Verdict::AnalysisOnly);
    assert!(r.manifest.incomplete);
}

#[test]
fn empty_module_whose_only_warning_is_kind_from_module_type_is_not_unreadable() {
    let mut kind_only = module("KindOnly", ModuleKind::Standard, "");
    kind_only.diagnostics.push(Diag::module(DiagCode::KindFromModuleType, "KindOnly", "kind from MODULETYPE"));
    let r = report(vec![kind_only, module("M", ModuleKind::Standard, "Sub A()\nEnd Sub\n")]);
    assert!(r.unreadable.is_empty());
    assert_eq!(r.verdict, Verdict::WouldRunInFuturePhase);
    assert!(!r.manifest.incomplete);
}

#[test]
fn projects_from_the_fixture_builder_analyse() {
    let fixture = Fixture::new(vec![
        FixtureModule::new("Module1", ModuleKind::Standard, "Sub Main()\r\n    Shell \"calc\"\r\nEnd Sub\r\n"),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, "Private Sub Workbook_Open()\r\nEnd Sub\r\n"),
    ]);
    let r = analyze(&read_project(&fixture.to_streams(), &Limits::default()), &EmptyCatalog);
    assert_eq!(marks_of(&r), [("Main", vec![Mark::EntryPoint]), ("Workbook_Open", vec![Mark::EventHandler, Mark::AutoExec])]);
    assert_eq!(r.flags.iter().map(|f| (f.kind, f.module.as_str())).collect::<Vec<_>>(), [(FlagKind::Shell, "Module1")]);
}

#[test]
fn line_index_follows_the_lexer_terminators() {
    let index = LineIndex::new("a\r\nb\nc\rd\u{2028}e");
    let lines: Vec<u32> = [0, 3, 5, 7, 9, 11].iter().map(|&pos| index.line_of(pos)).collect();
    assert_eq!(lines, [1, 2, 3, 4, 4, 5]);
}

#[test]
fn line_index_end_columns_count_characters() {
    let src = "ab\n\u{e9}x \u{e9}\u{2028}y";
    let index = LineIndex::new(src);
    let span = |line, col, start, end| crate::lexer::Span { line, col, start, end };
    // "éx é" ends on a two-byte character: line 2, column 4.
    assert_eq!(index.end(span(2, 1, 3, 9)), (2, 4));
    // After U+2028 the column restarts: "y" is line 3, column 1.
    assert_eq!(index.end(span(1, 1, 0, src.len())), (3, 1));
    // An empty span ends where it starts.
    assert_eq!(index.end(span(2, 2, 5, 5)), (2, 2));
}

#[test]
fn a_long_member_chain_on_a_small_stack_does_not_overflow() {
    let src = format!("Sub T()\nx = a{}\ny = f{}.z\nEnd Sub\n", ".b".repeat(200_000), "(1)".repeat(200_000));
    let worker = std::thread::Builder::new().stack_size(1 << 20).spawn(move || {
        let r = report(vec![module("M", ModuleKind::Standard, &src)]);
        assert_eq!(r.manifest.categories.len(), 9);
        assert!(!r.manifest.incomplete && r.manifest.categories.iter().all(|c| !c.present));
        r.members.iter().map(|m| (m.object.clone(), m.member.clone(), m.count)).collect::<Vec<_>>()
    });
    let members = worker.expect("spawn").join().expect("no stack overflow");
    assert_eq!(
        members,
        [(Some("b".to_owned()), "b".to_owned(), 199_999), (Some("a".to_owned()), "b".to_owned(), 1), (Some("f".to_owned()), "z".to_owned(), 1)]
    );
}
