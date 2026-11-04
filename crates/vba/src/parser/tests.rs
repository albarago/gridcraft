use super::{MAX_DEPTH, ParseDiag, ParseDiagKind, parse};
use crate::ast::{
    AssignKind, BinaryOp, CaseTest, CompareMode, Decl, DeclareKind, Expr, FileOp, Literal, Module, OptionKind, Passing, PrintSep, ProcKind,
    ResumeTarget, Stmt, UnaryOp, VarKeyword, Visibility,
};
use crate::lexer::MAX_SOURCE;

/// Parses `src` and requires no diagnostics.
fn parse_ok(src: &str) -> Module {
    let (module, diags) = parse(src);
    assert!(diags.is_empty(), "unexpected diagnostics for {src:?}: {diags:#?}");
    module.expect("a tree")
}

/// The body of the only procedure of `src` with no diagnostics.
fn body(src: &str) -> Vec<Stmt> {
    let mut module = parse_ok(src);
    assert_eq!(module.procs.len(), 1, "{src:?}");
    module.procs.remove(0).body
}

/// The body of `Sub T() … End Sub` around `lines`.
fn sub_body(lines: &str) -> Vec<Stmt> {
    body(&format!("Sub T()\n{lines}\nEnd Sub\n"))
}

/// The one statement of `line` inside a Sub.
fn stmt(line: &str) -> Stmt {
    let mut stmts = sub_body(line);
    assert_eq!(stmts.len(), 1, "{line:?}: {stmts:#?}");
    stmts.remove(0)
}

/// The value of `x = <src>`.
fn expr(src: &str) -> Expr {
    match stmt(&format!("x = {src}")) {
        Stmt::Assign { value, .. } => value,
        other => panic!("not an assignment: {other:#?}"),
    }
}

/// A compact rendering of an expression's shape for precedence tests.
fn shape(e: &Expr) -> String {
    match e {
        Expr::Literal { value, .. } => match value {
            Literal::Integer(t) | Literal::Float(t) | Literal::Hex(t) | Literal::Octal(t) | Literal::Date(t) => t.clone(),
            Literal::String(s) => format!("{s:?}"),
            Literal::Boolean(b) => b.to_string(),
            Literal::Nothing => "Nothing".into(),
            Literal::Empty => "Empty".into(),
            Literal::Null => "Null".into(),
        },
        Expr::Ident { name, .. } => name.clone(),
        Expr::Me { .. } => "Me".into(),
        Expr::Member { object, bang, name, .. } => {
            format!("{}{}{name}", object.as_deref().map(shape).unwrap_or_default(), if *bang { "!" } else { "." })
        }
        Expr::CallOrIndex { target, args, .. } => {
            let args: Vec<String> = args
                .iter()
                .map(|a| {
                    let value = a.value.as_ref().map(shape).unwrap_or_else(|| "_".into());
                    a.name.as_ref().map_or(value.clone(), |n| format!("{n}:={value}"))
                })
                .collect();
            format!("{}[{}]", shape(target), args.join(", "))
        }
        Expr::Paren { inner, .. } => format!("({})", shape(inner)),
        Expr::Unary { op, operand, .. } => format!("{}{}", if *op == UnaryOp::Neg { "-" } else { "Not " }, shape(operand)),
        Expr::Binary { op, left, right, .. } => format!("({} {op:?} {})", shape(left), shape(right)),
        Expr::New { class, .. } => format!("New {class}"),
        Expr::TypeOfIs { object, class, .. } => format!("TypeOf {} Is {class}", shape(object)),
        Expr::AddressOf { target, .. } => format!("AddressOf {}", shape(target)),
    }
}

fn diags(src: &str) -> (Module, Vec<ParseDiag>) {
    let (module, diags) = parse(src);
    (module.expect("a tree"), diags)
}

// ----- module level ----------------------------------------------------------------------

#[test]
fn empty_and_declarations_only_modules_have_no_procs_or_diags() {
    for src in ["", "\n\n", "' just a comment\r\n", "Option Explicit\nDim a As Long\nPrivate Const B = 2\n"] {
        let module = parse_ok(src);
        assert!(module.procs.is_empty(), "{src:?}");
    }
}

#[test]
fn attributes_become_module_attrs_without_diags() {
    let module = parse_ok("Attribute VB_Name = \"Module1\"\r\nAttribute VB_Base = \"0{000}\"\r\nAttribute VB_Ext_KEY = \"a\" ,\"b\"\r\n");
    let names: Vec<&str> = module.attrs.iter().map(|a| a.name.as_str()).collect();
    assert_eq!(names, ["VB_Name", "VB_Base", "VB_Ext_KEY"]);
    assert!(matches!(&module.attrs[0].values[..], [Expr::Literal { value: Literal::String(s), .. }] if s == "Module1"));
    assert_eq!(module.attrs[2].values.len(), 2);
    assert_eq!((module.attrs[1].span.line, module.attrs[1].span.col), (2, 1));
    let in_proc = sub_body("Attribute T.VB_UserMemId = 0");
    assert!(matches!(&in_proc[..], [Stmt::Attribute(a)] if a.name == "T.VB_UserMemId"));
}

#[test]
fn options() {
    let module = parse_ok("Option Explicit\nOption Base 1\nOption Compare Text\noption compare database\nOption Private Module\n");
    let kinds: Vec<OptionKind> = module.options.into_iter().map(|o| o.kind).collect();
    assert_eq!(
        kinds,
        [
            OptionKind::Explicit,
            OptionKind::Base("1".into()),
            OptionKind::Compare(CompareMode::Text),
            OptionKind::Compare(CompareMode::Database),
            OptionKind::PrivateModule
        ]
    );
}

#[test]
fn variable_declarations() {
    let module = parse_ok(
        "Dim a, b(1 To 10, 5) As Long\nPrivate WithEvents app As Excel.Application\nPublic s$, c As New Collection\nGlobal f As String * 10\nDim d() As Variant\n",
    );
    let Decl::Var(dim) = &module.decls[0] else { panic!() };
    assert_eq!(dim.keyword, VarKeyword::Dim);
    assert_eq!(dim.vars[0].name, "a");
    let bounds = dim.vars[1].bounds.as_ref().unwrap();
    assert_eq!(bounds.len(), 2);
    assert!(bounds[0].lower.is_some() && bounds[1].lower.is_none());
    assert_eq!(dim.vars[1].ty.as_ref().unwrap().name, "Long");
    let Decl::Var(private) = &module.decls[1] else { panic!() };
    assert_eq!(private.keyword, VarKeyword::Private);
    assert!(private.vars[0].with_events);
    assert_eq!(private.vars[0].ty.as_ref().unwrap().name, "Excel.Application");
    let Decl::Var(public) = &module.decls[2] else { panic!() };
    assert_eq!((public.keyword, public.vars[0].suffix), (VarKeyword::Public, Some('$')));
    assert!(public.vars[1].ty.as_ref().unwrap().is_new);
    let Decl::Var(global) = &module.decls[3] else { panic!() };
    assert!(global.vars[0].ty.as_ref().unwrap().fixed_len.is_some());
    let Decl::Var(empty) = &module.decls[4] else { panic!() };
    assert_eq!(empty.vars[0].bounds.as_deref(), Some(&[][..]));
}

#[test]
fn consts_types_and_enums() {
    let module = parse_ok(
        "Private Const A As Long = 1, B = \"x\"\nPublic Type Pt\n  X As Long\n  Name As String * 8\n  Items(1 To 3) As Integer\nEnd Type\nEnum Color\n  Red = 1\n  Green\nEnd Enum\n",
    );
    let Decl::Const(c) = &module.decls[0] else { panic!() };
    assert_eq!(c.visibility, Some(Visibility::Private));
    assert_eq!(c.consts.len(), 2);
    let Decl::Type(t) = &module.decls[1] else { panic!() };
    assert_eq!((t.name.as_str(), t.members.len(), t.visibility), ("Pt", 3, Some(Visibility::Public)));
    assert!(t.members[2].bounds.is_some());
    let Decl::Enum(e) = &module.decls[2] else { panic!() };
    assert_eq!(e.members.iter().map(|m| (m.name.as_str(), m.value.is_some())).collect::<Vec<_>>(), [("Red", true), ("Green", false)]);
}

#[test]
fn declare_event_implements_def() {
    let module = parse_ok(concat!(
        "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" Alias \"GetTickCount64\" () As LongLong\n",
        "Declare Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\n",
        "Public Event Changed(ByVal value As Long)\n",
        "Implements IShape\n",
        "DefInt A-Z, I\n",
        "#If VBA7 Then\n",
    ));
    let Decl::Declare(d) = &module.decls[0] else { panic!() };
    assert!(d.ptr_safe);
    assert_eq!(
        (d.kind, d.name.as_str(), d.lib.as_str(), d.alias.as_deref()),
        (DeclareKind::Function, "GetTickCount", "kernel32", Some("GetTickCount64"))
    );
    assert_eq!(d.return_type.as_ref().unwrap().name, "LongLong");
    let Decl::Declare(s) = &module.decls[1] else { panic!() };
    assert_eq!((s.kind, s.params.len(), s.params[0].passing), (DeclareKind::Sub, 1, Some(Passing::ByVal)));
    assert!(matches!(&module.decls[2], Decl::Event(e) if e.name == "Changed" && e.params.len() == 1));
    assert!(matches!(&module.decls[3], Decl::Implements { name, .. } if name == "IShape"));
    let Decl::Def(def) = &module.decls[4] else { panic!() };
    assert_eq!(def.ranges.len(), 2);
    assert_eq!((def.ranges[0].from.as_str(), def.ranges[0].to.as_deref()), ("A", Some("Z")));
    assert!(matches!(module.decls[5], Decl::Directive { .. }));
}

// ----- procedures ------------------------------------------------------------------------

#[test]
fn procedures_and_params() {
    let module = parse_ok(concat!(
        "Public Sub A()\nEnd Sub\n",
        "Private Static Function B$(Optional ByVal x As Long = 5, ByRef y(), ParamArray rest() As Variant) As String()\nEnd Function\n",
        "Friend Property Get C() As Long\nEnd Property\n",
        "Property Let C(ByVal v As Long)\nEnd Property\n",
        "Property Set D(ByVal o As Object)\nEnd Property\n",
    ));
    let kinds: Vec<ProcKind> = module.procs.iter().map(|p| p.kind).collect();
    assert_eq!(kinds, [ProcKind::Sub, ProcKind::Function, ProcKind::PropertyGet, ProcKind::PropertyLet, ProcKind::PropertySet]);
    let b = &module.procs[1];
    assert_eq!((b.visibility, b.is_static, b.suffix), (Some(Visibility::Private), true, Some('$')));
    assert!(b.return_type.as_ref().unwrap().is_array);
    assert!(b.params[0].optional && b.params[0].default.is_some());
    assert!(b.params[1].is_array && b.params[1].passing == Some(Passing::ByRef));
    assert!(b.params[2].param_array);
    assert_eq!(module.procs[2].visibility, Some(Visibility::Friend));
    assert_eq!((module.procs[1].span.line, module.procs[1].span.col), (3, 1));
    assert!(module.procs.iter().all(|p| !p.is_recovered()));
}

// ----- statements ------------------------------------------------------------------------

#[test]
fn assignments_and_calls() {
    let stmts = sub_body(concat!(
        "x = 1\nLet y = 2\nSet o = New Collection\nLSet s = t\n",
        "Call f(x)\nf x, y\nobj.Method a\nf(x)\nf (x), y\nRange(\"A1\").Value = 5\nDoEvents\n",
    ));
    assert!(matches!(&stmts[0], Stmt::Assign { kind: AssignKind::Implicit, .. }));
    assert!(matches!(&stmts[1], Stmt::Assign { kind: AssignKind::Let, .. }));
    assert!(matches!(&stmts[2], Stmt::Assign { kind: AssignKind::Set, value: Expr::New { .. }, .. }));
    assert!(matches!(&stmts[3], Stmt::Assign { kind: AssignKind::LSet, .. }));
    let calls: Vec<String> = stmts[4..]
        .iter()
        .map(|s| match s {
            Stmt::Call { explicit, target, args, .. } => {
                let args: Vec<String> = args.iter().map(|a| shape(a.value.as_ref().unwrap())).collect();
                format!("{}{} {}", if *explicit { "Call " } else { "" }, shape(target), args.join(","))
            }
            Stmt::Assign { target, value, .. } => format!("{} = {}", shape(target), shape(value)),
            other => panic!("{other:?}"),
        })
        .collect();
    assert_eq!(calls, ["Call f x", "f x,y", "obj.Method a", "f[x] ", "f (x),y", "Range[\"A1\"].Value = 5", "DoEvents "]);
}

#[test]
fn block_and_single_line_if() {
    let stmts = sub_body(
        "If a Then\n  x = 1\nElseIf b Then\n  x = 2\nElse\n  x = 3\nEnd If\nIf a Then x = 1 Else x = 2: y = 3\nIf a Then b: c\nIf a Then 10\n",
    );
    let Stmt::If(block) = &stmts[0] else { panic!() };
    assert!(!block.single_line);
    assert_eq!(block.branches.len(), 2);
    assert_eq!(block.else_body.as_ref().map(Vec::len), Some(1));
    let Stmt::If(inline) = &stmts[1] else { panic!() };
    assert!(inline.single_line);
    assert_eq!(inline.branches[0].body.len(), 1);
    assert_eq!(inline.else_body.as_ref().map(Vec::len), Some(2));
    let Stmt::If(two) = &stmts[2] else { panic!() };
    assert_eq!((two.branches[0].body.len(), two.else_body.is_none()), (2, true));
    let Stmt::If(goto) = &stmts[3] else { panic!() };
    assert!(matches!(&goto.branches[0].body[..], [Stmt::GoTo { label, .. }] if label == "10"));
    assert_eq!(stmts.len(), 4);
}

#[test]
fn select_case() {
    let Stmt::Select(s) = stmt("Select Case n\nCase 1, 3 To 5\n  a\nCase Is > 10, < 0\n  b\nCase Else\n  c\nEnd Select") else { panic!() };
    assert_eq!(s.cases.len(), 2);
    assert!(matches!(&s.cases[0].tests[..], [CaseTest::Value(_), CaseTest::Range { .. }]));
    assert!(matches!(&s.cases[1].tests[..], [CaseTest::Is { op: BinaryOp::Gt, .. }, CaseTest::Is { op: BinaryOp::Lt, .. }]));
    assert_eq!(s.else_body.as_ref().map(Vec::len), Some(1));
}

#[test]
fn loops() {
    let stmts = sub_body(concat!(
        "For i = 1 To 10 Step 2\n  x = i\nNext i\n",
        "For Each c In Range(\"A1:B2\").Cells\nNext\n",
        "For i = 1 To 2\nFor j = 1 To 2\nNext j, i\n",
        "Do While x < 5\n  x = x + 1\nLoop\n",
        "Do\nLoop Until x\n",
        "While a\nWend\n",
        "With ws\n  .Range(\"A1\").Value = 1\n  .Name = \"x\"\nEnd With\n",
    ));
    let Stmt::For(f) = &stmts[0] else { panic!() };
    assert!(f.step.is_some() && f.next_var.is_some());
    let Stmt::ForEach(each) = &stmts[1] else { panic!() };
    assert_eq!(shape(&each.collection), "Range[\"A1:B2\"].Cells");
    let Stmt::For(outer) = &stmts[2] else { panic!() };
    let [Stmt::For(inner)] = &outer.body[..] else { panic!("{:#?}", outer.body) };
    assert_eq!((inner.next_var.as_ref().map(shape), outer.next_var.as_ref().map(shape)), (Some("j".into()), Some("i".into())));
    let Stmt::Do(pre) = &stmts[3] else { panic!() };
    assert!(pre.pre.as_ref().is_some_and(|c| !c.until) && pre.post.is_none());
    let Stmt::Do(post) = &stmts[4] else { panic!() };
    assert!(post.post.as_ref().is_some_and(|c| c.until));
    assert!(matches!(&stmts[5], Stmt::While { .. }));
    let Stmt::With { body, .. } = &stmts[6] else { panic!() };
    let Stmt::Assign { target, .. } = &body[0] else { panic!() };
    assert_eq!(shape(target), ".Range[\"A1\"].Value");
}

#[test]
fn on_error_and_resume() {
    let stmts = sub_body("On Error GoTo Handler\nOn Error GoTo 0\nOn Error GoTo -1\nOn Error Resume Next\nResume\nResume Next\nResume done\n");
    use crate::ast::OnErrorAction::{GoTo, GoToMinusOne, GoToZero, ResumeNext};
    let actions: Vec<_> = stmts[..4].iter().map(|s| if let Stmt::OnError { action, .. } = s { action.clone() } else { panic!() }).collect();
    assert_eq!(actions, [GoTo("Handler".into()), GoToZero, GoToMinusOne, ResumeNext]);
    let resumes: Vec<_> = stmts[4..].iter().map(|s| if let Stmt::Resume { target, .. } = s { target.clone() } else { panic!() }).collect();
    assert_eq!(resumes, [ResumeTarget::Same, ResumeTarget::Next, ResumeTarget::Label("done".into())]);
}

#[test]
fn jumps_and_labels() {
    let stmts = sub_body("GoTo done\nGoSub 100\nReturn\nExit Sub\nExit For\nOn n GoTo a, b\nHandler:\n100 x = 1\ndone: Stop\nEnd\n");
    let described: Vec<String> = stmts
        .iter()
        .map(|s| match s {
            Stmt::GoTo { label, .. } => format!("GoTo {label}"),
            Stmt::GoSub { label, .. } => format!("GoSub {label}"),
            Stmt::Return { .. } => "Return".into(),
            Stmt::Exit { kind, .. } => format!("Exit {kind:?}"),
            Stmt::OnJump { expr, gosub, labels, .. } => format!("On {} {} {}", shape(expr), if *gosub { "GoSub" } else { "GoTo" }, labels.join(",")),
            Stmt::Label { name, is_line_number, .. } => format!("{}label {name}", if *is_line_number { "line " } else { "" }),
            Stmt::Assign { .. } => "assign".into(),
            Stmt::Stop { .. } => "Stop".into(),
            Stmt::End { .. } => "End".into(),
            other => format!("{other:?}"),
        })
        .collect();
    assert_eq!(
        described,
        [
            "GoTo done",
            "GoSub 100",
            "Return",
            "Exit Sub",
            "Exit For",
            "On n GoTo a,b",
            "label Handler",
            "line label 100",
            "assign",
            "label done",
            "Stop",
            "End"
        ]
    );
}

#[test]
fn local_declarations_arrays_and_events() {
    let stmts = sub_body(
        "Dim a As Long, b\nStatic n%\nConst K = 3\nReDim Preserve arr(1 To n, 2) As String\nReDim Me.v(5)\nErase arr, v\nRaiseEvent Changed(1)\n",
    );
    assert!(matches!(&stmts[0], Stmt::Var(v) if v.keyword == VarKeyword::Dim && v.vars.len() == 2));
    assert!(matches!(&stmts[1], Stmt::Var(v) if v.keyword == VarKeyword::Static && v.vars[0].suffix == Some('%')));
    assert!(matches!(&stmts[2], Stmt::Const(_)));
    let Stmt::ReDim { preserve, vars, .. } = &stmts[3] else { panic!() };
    assert!(*preserve && vars[0].bounds.len() == 2 && vars[0].ty.is_some());
    let Stmt::ReDim { vars, .. } = &stmts[4] else { panic!() };
    assert_eq!(shape(&vars[0].target), "Me.v");
    assert!(matches!(&stmts[5], Stmt::Erase { targets, .. } if targets.len() == 2));
    assert!(matches!(&stmts[6], Stmt::RaiseEvent { name, args, .. } if name == "Changed" && args.len() == 1));
}

#[test]
fn debug_print_lists() {
    let Stmt::DebugPrint { items, .. } = stmt("Debug.Print a; b, c") else { panic!() };
    let seps: Vec<_> = items.iter().map(|i| (i.expr.as_ref().map(shape), i.sep)).collect();
    assert_eq!(seps, [(Some("a".into()), Some(PrintSep::Semicolon)), (Some("b".into()), Some(PrintSep::Comma)), (Some("c".into()), None)]);
    let Stmt::DebugPrint { items, .. } = stmt("Debug.Print") else { panic!() };
    assert!(items.is_empty());
    let Stmt::DebugPrint { items, .. } = stmt("Debug.Print Tab(5); x;") else { panic!() };
    assert_eq!(items.len(), 2);
    assert!(matches!(stmt("Debug.Assert x > 0"), Stmt::Call { .. }));
}

#[test]
fn file_statements() {
    let stmts = sub_body(concat!(
        "Open \"C:\\a.txt\" For Binary Access Read Lock Write As #1 Len = 128\n",
        "Print #1, a; b\nWrite #1, a, b\nInput #1, a, b\nLine Input #1, s\nGet #1, , rec\nPut #f, 3, rec\n",
        "Seek #1, 10\nLock #1, 1 To 5\nClose #1, #2\nClose\n",
    ));
    let files: Vec<_> = stmts.iter().map(|s| if let Stmt::File(f) = s { f } else { panic!("{s:?}") }).collect();
    let ops: Vec<FileOp> = files.iter().map(|f| f.op).collect();
    use FileOp::{Close, Get, Input, LineInput, Lock, Open, Print, Put, Seek, Write};
    assert_eq!(ops, [Open, Print, Write, Input, LineInput, Get, Put, Seek, Lock, Close, Close]);
    assert_eq!(files[0].mode, ["Binary", "Access", "Read", "Lock", "Write"]);
    assert_eq!(files[0].exprs().count(), 3, "path, number and Len");
    assert_eq!(files[1].items.len(), 2);
    assert!(files[5].args[0].is_none() && files[5].args[1].is_some());
    assert_eq!(files[8].args.len(), 2);
    assert_eq!((files[9].args.len(), files[10].args.len()), (2, 0));
    assert!(files.iter().all(|f| f.op == Close || f.number.is_some()));
}

#[test]
fn separators_continuations_and_directives() {
    let stmts = sub_body("a = 1: b = 2\nc = 1 + _\n    2\nRange(\"A1\") _\n    .Value = 3\n#If VBA7 Then\nd = 4\n#End If");
    assert_eq!(stmts.len(), 7, "{stmts:#?}");
    let Stmt::Assign { target, .. } = &stmts[3] else { panic!() };
    assert_eq!(shape(target), "Range[\"A1\"].Value");
    assert!(matches!(stmts[4], Stmt::Directive { .. }) && matches!(stmts[6], Stmt::Directive { .. }));
}

// ----- expressions -----------------------------------------------------------------------

#[test]
fn operator_precedence() {
    let cases = [
        ("a + b * c ^ -d", "(a Add (b Mul (c Pow -d)))"),
        ("Not a = b And c", "(Not (a Eq b) And c)"),
        ("a & b + c", "(a Concat (b Add c))"),
        ("x Like \"*a\" Or y Is Nothing", "((x Like \"*a\") Or (y Is Nothing))"),
        ("-a ^ 2", "-(a Pow 2)"),
        ("a - b - c", "((a Sub b) Sub c)"),
        ("a \\ b Mod c * d", "((a IntDiv b) Mod (c Mul d))"),
        ("a Imp b Eqv c Xor d Or e", "(a Imp (b Eqv (c Xor (d Or e))))"),
        ("a <= b <> c >= d", "(((a Le b) Ne c) Ge d)"),
        ("(a + b) * c", "(((a Add b)) Mul c)"),
    ];
    for (src, want) in cases {
        assert_eq!(shape(&expr(src)), want, "{src}");
    }
    let (_, diags) = diags("Sub T()\nx = a < = b\nEnd Sub\n");
    assert_eq!(diags.len(), 1, "a spaced '< =' is two operators: {diags:?}");
}

#[test]
fn special_expressions() {
    assert_eq!(shape(&expr("TypeOf o Is Excel.Range And Not b")), "(TypeOf o Is Excel.Range And Not b)");
    assert_eq!(shape(&expr("AddressOf Module1.Callback")), "AddressOf Module1.Callback");
    assert_eq!(shape(&expr("Me.Name & rs!Field & rs![Other Name]")), "((Me.Name Concat rs!Field) Concat rs![Other Name])");
    assert_eq!(shape(&expr("f(a, , c:=1)")), "f[a, _, c:=1]");
    assert_eq!(shape(&expr("f()(1)")), "f[][1]");
    assert_eq!(shape(&expr("Array(True, Empty, Null, #1/2/2003#, &HFF, 1.5)")), "Array[true, Empty, Null, #1/2/2003#, &HFF, 1.5]");
    assert_eq!(shape(&expr("\"say \"\"hi\"\"\"")), "\"say \\\"hi\\\"\"");
    assert_eq!(shape(&expr("Left$(s, 2) + UBound(a)")), "(Left[s, 2] Add UBound[a])");
    assert!(matches!(expr("s$"), Expr::Ident { suffix: Some('$'), .. }));
    let Stmt::Call { args, .. } = stmt("MsgBox Prompt:=\"x\", Buttons:=vbOKOnly") else { panic!() };
    assert_eq!(args.iter().map(|a| a.name.as_deref()).collect::<Vec<_>>(), [Some("Prompt"), Some("Buttons")]);
    let Stmt::Call { args, .. } = stmt("f , x") else { panic!() };
    assert!(args[0].value.is_none() && args[1].value.is_some());
}

// ----- recovery, depth and size ----------------------------------------------------------

#[test]
fn bad_procedure_between_good_ones_is_recovered() {
    let src = "Sub Good1()\n  x = 1\nEnd Sub\n\nSub Bad()\n  y = 1\n  If y Then\n    z = )\n  End If\n  w = 2\nEnd Sub\n\nFunction Good2() As Long\n  Good2 = 2\nEnd Function\n";
    let (module, diags) = diags(src);
    assert_eq!(diags.len(), 1, "{diags:#?}");
    assert_eq!(diags[0].kind, ParseDiagKind::Syntax);
    assert_eq!((diags[0].span.line, diags[0].span.col), (8, 9));
    let names: Vec<&str> = module.procs.iter().map(|p| p.name.as_str()).collect();
    assert_eq!(names, ["Good1", "Bad", "Good2"]);
    assert!(!module.procs[0].is_recovered() && !module.procs[2].is_recovered());
    let bad = &module.procs[1];
    assert!(bad.is_recovered());
    assert!(matches!(&bad.body[..], [Stmt::Assign { .. }, Stmt::Unknown { .. }]));
    assert_eq!(module.procs[2].body.len(), 1);
}

#[test]
fn missing_end_sub_stops_at_next_header() {
    let (module, diags) = diags("Sub A()\n  x = 1\nPrivate Sub B()\nEnd Sub\n");
    assert_eq!(diags.len(), 1);
    assert_eq!((diags[0].span.line, diags[0].span.col), (3, 1));
    assert_eq!(module.procs.iter().map(|p| (p.name.as_str(), p.is_recovered())).collect::<Vec<_>>(), [("A", true), ("B", false)]);
}

#[test]
fn module_level_errors_skip_the_line() {
    let (module, diags) = diags("Dim As\nx = 1\nDim ok As Long\nSub A()\nEnd Sub\n");
    assert_eq!(diags.len(), 2, "{diags:#?}");
    assert_eq!(module.decls.len(), 1);
    assert_eq!(module.procs.len(), 1);
}

#[test]
fn unterminated_blocks_report_at_end() {
    let (module, diags) = diags("Sub A()\n  If x Then\n");
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0].span.line, 3);
    assert!(module.procs[0].is_recovered());
}

/// `Sub T()` with `n` nested block `If`s.
fn nested_ifs(n: usize) -> String {
    format!("Sub T()\n{}{}End Sub\n", "If x Then\n".repeat(n), "End If\n".repeat(n))
}

/// `Sub T()` with `x = ` and `n` nested parentheses.
fn nested_parens(n: usize) -> String {
    format!("Sub T()\nx = {}1{}\nEnd Sub\n", "(".repeat(n), ")".repeat(n))
}

#[test]
fn nesting_limit_for_blocks_and_parens() {
    let depth = usize::from(MAX_DEPTH);
    parse_ok(&nested_ifs(depth));
    parse_ok(&nested_parens(depth));
    let (module, diags) = diags(&nested_ifs(depth + 1));
    assert_eq!(diags.len(), 1, "{diags:?}");
    assert_eq!(diags[0].kind, ParseDiagKind::NestingTooDeep);
    assert_eq!((diags[0].span.line, diags[0].span.col), (u32::from(MAX_DEPTH) + 2, 1), "the 257th If");
    assert!(module.procs[0].is_recovered());
    let (_, diags) = parse(&nested_parens(depth + 1));
    assert_eq!(diags.len(), 1, "{diags:?}");
    assert_eq!(diags[0].kind, ParseDiagKind::NestingTooDeep);
    assert_eq!((diags[0].span.line, diags[0].span.col), (2, u32::from(MAX_DEPTH) + 5), "the 257th '('");
    let (_, diags) = parse(&format!("Sub T()\nx = {}1\nEnd Sub\nSub U()\nEnd Sub\n", "-".repeat(depth + 1)));
    assert_eq!(diags.iter().map(|d| d.kind).collect::<Vec<_>>(), [ParseDiagKind::NestingTooDeep]);
}

#[test]
fn deep_nesting_on_a_small_stack_does_not_overflow() {
    let worker = std::thread::Builder::new().stack_size(1 << 20).spawn(|| {
        let mut kinds = Vec::new();
        for src in [nested_parens(10_000), nested_ifs(10_000), nested_parens(usize::from(MAX_DEPTH)), nested_ifs(usize::from(MAX_DEPTH))] {
            let (module, diags) = parse(&src);
            assert!(module.is_some());
            kinds.push(diags.iter().map(|d| d.kind).collect::<Vec<_>>());
        }
        kinds
    });
    let kinds = worker.expect("spawn").join().expect("no stack overflow");
    assert_eq!(kinds, [vec![ParseDiagKind::NestingTooDeep], vec![ParseDiagKind::NestingTooDeep], vec![], vec![]]);
}

/// Regression: flat chains have no nesting level per link, so their trees are as deep as the
/// source is long; dropping one used to recurse per link and overflow a 1 MiB stack.
#[test]
fn long_flat_chains_drop_on_a_small_stack() {
    let n = 200_000;
    let sources = [
        format!("Sub T()\nx = {}\nEnd Sub\n", vec!["a"; n].join(" + ")),
        format!("Sub T()\nx = a{}\nEnd Sub\n", ".b".repeat(n)),
        format!("Sub T()\nx = f{}\nEnd Sub\n", "(1)".repeat(n)),
    ];
    let worker = std::thread::Builder::new().stack_size(1 << 20).spawn(move || {
        sources
            .iter()
            .map(|src| {
                let (module, diags) = parse(src);
                let parsed = module.is_some() && diags.is_empty();
                drop(module);
                parsed
            })
            .collect::<Vec<_>>()
    });
    assert_eq!(worker.expect("spawn").join().expect("no stack overflow"), [true, true, true]);
}

#[test]
fn source_over_the_cap_gives_no_tree() {
    let big = " ".repeat(MAX_SOURCE + 1);
    let (module, diags) = parse(&big);
    assert!(module.is_none());
    assert_eq!(diags.iter().map(|d| d.kind).collect::<Vec<_>>(), [ParseDiagKind::SourceTooLarge]);
    let (module, diags) = parse(&" ".repeat(MAX_SOURCE));
    assert!(module.is_some() && diags.is_empty());
}
