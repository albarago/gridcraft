use super::{print_expr, print_module, print_stmt};
use crate::ast::testeq::{equivalent, normalized};
use crate::ast::{Arg, BinaryOp, Expr, IfBranch, IfStmt, Literal, Module, Stmt, UnaryOp};
use crate::lexer::Span;
use crate::parser::parse;

const Z: Span = Span { line: 0, col: 0, start: 0, end: 0 };

/// Parses `src` with no diagnostics, prints it, and checks that the printed source parses with
/// no diagnostics (Req 12.7) to an equivalent tree (Req 12.8) and prints the same again.
fn round_trip(src: &str) -> String {
    let (tree, diags) = parse(src);
    assert!(diags.is_empty(), "source has diagnostics: {src:?}\n{diags:#?}");
    let tree = tree.expect("a tree");
    let printed = print_module(&tree);
    let (again, diags) = parse(&printed);
    assert!(diags.is_empty(), "printed source has diagnostics:\n{printed}\n{diags:#?}");
    let again = again.expect("a tree");
    assert_eq!(normalized(&again), normalized(&tree), "not equivalent after printing:\n{printed}");
    assert!(equivalent(&tree, &again));
    assert_eq!(print_module(&again), printed, "printing isn't stable");
    printed
}

/// Round-trips `lines` inside `Sub T()` and returns the printed body lines.
fn round_trip_body(lines: &str) -> String {
    let printed = round_trip(&format!("Sub T()\n{lines}\nEnd Sub\n"));
    printed.strip_prefix("Sub T()\n").and_then(|rest| rest.strip_suffix("End Sub\n")).expect("one Sub").to_owned()
}

/// Round-trips `x = <src>` and returns the printed expression.
fn round_trip_expr(src: &str) -> String {
    let body = round_trip_body(&format!("x = {src}"));
    body.strip_prefix("    x = ").and_then(|rest| rest.strip_suffix('\n')).expect("one assignment").to_owned()
}

/// Prints a hand-built expression and checks it parses back with no diagnostics (Req 12.7).
fn print_built(e: &Expr) -> String {
    let text = print_expr(e);
    let (_, diags) = parse(&format!("Sub T()\nx = {text}\nEnd Sub\n"));
    assert!(diags.is_empty(), "{text}: {diags:#?}");
    text
}

fn id(name: &str) -> Expr {
    Expr::Ident { name: name.into(), suffix: None, span: Z }
}

fn bin(op: BinaryOp, left: Expr, right: Expr) -> Expr {
    Expr::Binary { op, left: Box::new(left), right: Box::new(right), span: Z }
}

fn un(op: UnaryOp, operand: Expr) -> Expr {
    Expr::Unary { op, operand: Box::new(operand), span: Z }
}

// ----- one source per construct family ----------------------------------------------------

#[test]
fn module_level_constructs_round_trip() {
    round_trip(concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Attribute VB_Ext_KEY = \"a\" ,\"b\"\r\n",
        "Option Explicit\nOption Base 1\nOption Compare Text\noption compare database\nOption Private Module\n",
        "Dim a, b(1 To 10, 5) As Long\nPrivate WithEvents app As Excel.Application\nPublic s$, c As New Collection\n",
        "Global f As String * 10\nDim d() As Variant\nStatic st As Integer\n",
        "Private Const A As Long = 1, B = \"x\"\n",
        "Public Type Pt\n  X As Long\n  Name As String * 8\n  Items(1 To 3) As Integer\nEnd Type\n",
        "Enum Color\n  Red = 1\n  Green\n  Blue = Red Or 4\nEnd Enum\n",
        "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" Alias \"GetTickCount64\" () As LongLong\n",
        "Declare Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\n",
        "Declare Function Odd Lib \"a \"\"b\"\".dll\" Alias \"#1\" (ByRef p As Any) As Long\n",
        "Public Event Changed(ByVal value As Long)\nEvent Plain\n",
        "Implements IShape\nImplements Lib1.IFace\nDefInt A-Z, I\nDefStr S\n",
        "#If VBA7 Then\nDim p As LongPtr\n#Else\nDim p As Long\n#End If\n",
    ));
}

#[test]
fn procedures_and_params_round_trip() {
    round_trip(concat!(
        "Public Sub A()\nEnd Sub\n",
        "Private Static Function B$(Optional ByVal x As Long = 5, ByRef y(), ParamArray rest() As Variant) As String()\nEnd Function\n",
        "Friend Property Get C() As Long\n  C = 1\nEnd Property\n",
        "Property Let C(ByVal v As Long)\nEnd Property\n",
        "Property Set D(ByVal o As Object)\nEnd Property\n",
        "Static Public Sub E(a, Optional b As String = \"q\")\nEnd Sub\n",
        "Global Function F(ByVal s As String * 4) As Variant\nEnd Function\n",
    ));
}

#[test]
fn statements_round_trip() {
    round_trip_body(concat!(
        "x = 1\nLet y = 2\nSet o = New Collection\nLSet s = t\nRSet s = t\n",
        "Call f(x)\nCall g\nCall f()(1)\nCall f(1)()\nCall obj.M(a, , c:=1)\n",
        "f x, y\nobj.Method a\nf(x)\nf (x), y\nf (x) + 1, y\nf (x) = 1\nf , x\nf a,\n",
        "Range(\"A1\").Value = 5\nDoEvents\nMsgBox Prompt:=\"x\", Buttons:=vbOKOnly\nf -1\nf .5\n",
        "Dim a As Long, b\nStatic n%\nConst K = 3, L As String = \"z\"\n",
        "ReDim Preserve arr(1 To n, 2) As String\nReDim Me.v(5)\nReDim e()\nErase arr, v\n",
        "RaiseEvent Changed(1)\nRaiseEvent Done\n",
        "Attribute T.VB_UserMemId = 0\n",
        "On Error GoTo Handler\nOn Error GoTo 0\nOn Error GoTo -1\nOn Error Resume Next\nResume\nResume Next\nResume done\n",
        "GoTo done\nGoSub 100\nReturn\nExit Sub\nExit For\nExit Do\nExit Function\nExit Property\n",
        "On n GoTo a, b\nOn n + 1 GoSub 10, 20\n",
        "Handler:\n100 x = 1\n200: y = 2\ndone: Stop\nEnd\n",
        "#If VBA7 Then\nd = 4\n#End If\n",
    ));
}

#[test]
fn blocks_round_trip() {
    round_trip_body(concat!(
        "If a Then\n  x = 1\nElseIf b Then\n  x = 2\nElse\n  x = 3\nEnd If\n",
        "If a Then\nEnd If\nIf a Then\nElse\nEndIf\n",
        "Select Case n\nCase 1, 3 To 5\n  a\nCase Is > 10, < 0, Is = 7\n  b\nCase Else\n  c\nEnd Select\n",
        "Select Case s\nEnd Select\n",
        "For i = 1 To 10 Step 2\n  x = i\nNext i\nFor i = 1 To 2\nNext\n",
        "For Each c In Range(\"A1:B2\").Cells\nNext\nFor Each c In coll\nNext c\n",
        "Do While x < 5\n  x = x + 1\nLoop\nDo\nLoop Until x\nDo Until a\nLoop While b\nDo\nLoop\n",
        "While a\n  a = a - 1\nWend\n",
        "With ws\n  .Range(\"A1\").Value = 1\n  .Name = \"x\"\n  .Add .Item(1), !Key\n  f .x\n  With .Child\n    !y = .z\n  End With\nEnd With\n",
    ));
}

#[test]
fn single_line_if_prints_as_a_block() {
    let body = round_trip_body("If a Then x = 1 Else x = 2: y = 3\nIf a Then b: c\nIf a Then 10\nIf a Then If b Then c Else d\n");
    assert_eq!(
        body,
        concat!(
            "    If a Then\n        x = 1\n    Else\n        x = 2\n        y = 3\n    End If\n",
            "    If a Then\n        b\n        c\n    End If\n",
            "    If a Then\n        GoTo 10\n    End If\n",
            "    If a Then\n        If b Then\n            c\n        Else\n            d\n        End If\n    End If\n",
        )
    );
}

#[test]
fn next_with_two_variables_prints_as_two_nexts() {
    let body =
        round_trip_body("For i = 1 To 2\nFor j = 1 To 2\nx = i * j\nNext j, i\nFor a = 1 To 2: For b = 1 To 2: For c = 1 To 2\nNext c, b, a\n");
    assert!(body.contains("        Next j\n    Next i\n"), "{body}");
    assert!(body.contains("            Next c\n        Next b\n    Next a\n"), "{body}");
}

#[test]
fn file_and_print_statements_round_trip() {
    round_trip_body(concat!(
        "Open \"C:\\a.txt\" For Binary Access Read Lock Write As #1 Len = 128\nOpen p As 2\nOpen p For Input As #f\n",
        "Print #1, a; b\nPrint #1,\nPrint #1\nWrite #1, a, b\nWrite #1, , a\nInput #1, a, b\nLine Input #1, s\n",
        "Get #1, , rec\nPut #f, 3, rec\nPut #1, ,\nSeek #1, 10\nLock #1, 1 To 5\nUnlock #1\nWidth #1, 80\nClose #1, #2\nClose 3\nClose\n",
        "Print #1, #1/2/2003#; \"x\"\n",
        "Debug.Print a; b, c\nDebug.Print\nDebug.Print Tab(5); x;\nDebug.Print , x\nDebug.Assert x > 0\n",
    ));
}

#[test]
fn debug_print_member_call_keeps_its_line_continuation() {
    let body = round_trip_body("Debug _\n.Print x\nDebug _\n  .print.Foo = 1\n");
    assert_eq!(body, "    Debug _\n        .Print x\n    Debug _\n        .print.Foo = 1\n");
}

#[test]
fn literals_round_trip_with_their_text() {
    assert_eq!(round_trip_expr("\"say \"\"hi\"\"\""), "\"say \"\"hi\"\"\"");
    assert_eq!(round_trip_expr("\"\"\"\""), "\"\"\"\"");
    assert_eq!(round_trip_expr("\"\""), "\"\"");
    assert_eq!(
        round_trip_expr(
            "Array(True, false, Empty, Null, Nothing, #1/2/2003#, # 10:30 PM #, &HFF, &hff&, &O17, &17, 1.5, 1.5E+10, 2#, 3!, 4@, 5%, 6&, .5)"
        ),
        "Array(True, False, Empty, Null, Nothing, #1/2/2003#, # 10:30 PM #, &HFF, &hff&, &O17, &17, 1.5, 1.5E+10, 2#, 3!, 4@, 5%, 6&, .5)"
    );
}

#[test]
fn names_keep_their_spelling() {
    assert_eq!(round_trip_expr("Me.Name & rs!Field & rs![Other Name] & [A B].c"), "Me.Name & rs!Field & rs![Other Name] & [A B].c");
    assert_eq!(
        round_trip_expr("Left$(s, 2) + UBound(a) + Len(s) + CStr(x) + s$ + n% + obj.Print + x.End"),
        "Left$(s, 2) + UBound(a) + Len(s) + CStr(x) + s$ + n% + obj.Print + x.End"
    );
    round_trip_body("Text = 1\nBinary = Compare\nDim Database As Long\n[If] = [Dim]\n");
}

// ----- precedence -------------------------------------------------------------------------

#[test]
fn precedence_edge_cases_print_without_extra_parentheses() {
    let cases = [
        "-a ^ b",
        "(-a) ^ b",
        "a - (b - c)",
        "(a - b) - c",
        "a - b - c",
        "a ^ b ^ c",
        "a ^ (b ^ c)",
        "Not a = b",
        "(Not a) = b",
        "Not a And b",
        "Not (a And b)",
        "a & b + c",
        "(a & b) + c",
        "a + b * c ^ -d",
        "a * -b + c",
        "a * Not b",
        "a + Not b = c",
        "-Not a",
        "--a",
        "a - -1",
        "x Like \"*a\" Or y Is Nothing",
        "a \\ b Mod c * d",
        "a Imp b Eqv c Xor d Or e",
        "a <= b <> c >= d",
        "a < b = c > d",
        "TypeOf o Is Excel.Range And Not b",
        "TypeOf (Not o) Is C",
        "Not TypeOf o Is C",
        "AddressOf Module1.Callback",
        "f(a, , c:=1) + f() + f()(1) + g(-1, (2))",
        "New Collection",
        "a&b",
    ];
    for src in cases {
        let want = if src == "a&b" { "a & b" } else { src };
        assert_eq!(round_trip_expr(src), want, "{src}");
    }
}

#[test]
fn hand_built_trees_get_the_parentheses_they_need() {
    use BinaryOp::{Add, And, Eq, Mul, Pow, Sub};
    use UnaryOp::{Neg, Not};
    let (a, b, c) = (|| id("a"), || id("b"), || id("c"));
    let cases = [
        (bin(Pow, un(Neg, a()), b()), "(-a) ^ b"),
        (bin(Sub, a(), bin(Sub, b(), c())), "a - (b - c)"),
        (bin(Sub, bin(Sub, a(), b()), c()), "a - b - c"),
        (bin(Pow, a(), bin(Pow, b(), c())), "a ^ (b ^ c)"),
        (bin(Eq, un(Not, a()), b()), "(Not a) = b"),
        (un(Not, bin(And, a(), b())), "Not (a And b)"),
        (un(Neg, bin(Add, a(), b())), "-(a + b)"),
        (bin(Mul, bin(Add, a(), b()), c()), "(a + b) * c"),
        (bin(Add, bin(Mul, a(), un(Not, b())), c()), "(a * Not b) + c"),
        (bin(Pow, bin(Pow, a(), un(Neg, b())), c()), "(a ^ -b) ^ c"),
        (Expr::TypeOfIs { object: Box::new(un(Not, a())), class: "C".into(), span: Z }, "TypeOf (Not a) Is C"),
        (Expr::TypeOfIs { object: Box::new(bin(And, a(), b())), class: "C".into(), span: Z }, "TypeOf (a And b) Is C"),
        (bin(Add, id("If"), id("Rem")), "[If] + [Rem]"),
        (bin(Add, id("two words"), id("Len")), "[two words] + Len"),
        (Expr::Literal { value: Literal::String("a\"b".into()), span: Z }, "\"a\"\"b\""),
    ];
    for (e, want) in cases {
        assert_eq!(print_built(&e), want);
    }
}

#[test]
fn hand_built_statements_parse_back() {
    let call = |explicit, target: Expr, args: Vec<Arg>| Stmt::Call { explicit, target, args, span: Z };
    let arg = |value: Option<Expr>| Arg { name: None, value, span: Z };
    let paren = |e: Expr| Expr::Paren { inner: Box::new(e), span: Z };
    let index = |target: Expr, args: Vec<Arg>| Expr::CallOrIndex { target: Box::new(target), args, span: Z };
    let cases = [
        (call(false, id("f"), vec![arg(Some(paren(id("x")))), arg(Some(id("y")))]), "f (x), y\n"),
        (call(false, index(id("f"), vec![arg(Some(id("x")))]), vec![]), "f(x)\n"),
        (call(true, id("f"), vec![arg(None), arg(Some(id("y")))]), "Call f(, y)\n"),
        (call(true, index(id("f"), vec![arg(Some(id("x")))]), vec![]), "Call f(x)()\n"),
        (call(false, id("f"), vec![arg(Some(Expr::Member { object: None, bang: false, name: "x".into(), suffix: None, span: Z }))]), "f .x\n"),
        (Stmt::Unknown { span: Z }, "' (unparsed)\n"),
        (
            Stmt::If(IfStmt {
                branches: vec![IfBranch { cond: id("a"), body: vec![call(false, id("b"), vec![])], span: Z }],
                else_body: None,
                single_line: true,
                span: Z,
            }),
            "If a Then\n    b\nEnd If\n",
        ),
    ];
    for (s, want) in cases {
        let text = print_stmt(&s);
        assert_eq!(text, want);
        let (tree, diags) = parse(&format!("Sub T()\n{text}End Sub\n"));
        assert!(diags.is_empty(), "{text}: {diags:#?}");
        let body = tree.map(|m| m.procs.into_iter().flat_map(|p| p.body).count());
        assert_eq!(body, Some(if matches!(s, Stmt::Unknown { .. }) { 0 } else { 1 }), "{text}");
    }
}

// ----- formatting -------------------------------------------------------------------------

#[test]
fn canonical_layout() {
    let src = concat!(
        "attribute VB_Name = \"M\"\r\n",
        "option explicit\r\n",
        "dim  x   as long ' a comment\r\n",
        "public sub Go( byval n as long )\r\n",
        "  if n>0 then debug.print n:else _\r\n     n = -n\r\n",
        "  select case n\r\n  case 1: x=1\r\n  end select\r\n",
        "end sub\r\n",
        "private x2 as string\r\n",
    );
    let printed = round_trip(src);
    assert_eq!(
        printed,
        concat!(
            "Attribute VB_Name = \"M\"\n",
            "Option Explicit\n",
            "Dim x As long\n",
            "\n",
            "Public Sub Go(ByVal n As long)\n",
            "    If n > 0 Then\n",
            "        Debug.Print n\n",
            "    Else\n",
            "        n = -n\n",
            "    End If\n",
            "    Select Case n\n",
            "        Case 1\n",
            "            x = 1\n",
            "    End Select\n",
            "End Sub\n",
            "\n",
            "Private x2 As string\n",
        )
    );
}

#[test]
fn empty_module_prints_nothing() {
    assert_eq!(round_trip(""), "");
    assert_eq!(round_trip("' only a comment\n\n"), "");
}

#[test]
fn recovered_procedures_print_a_comment_that_parses() {
    let (tree, diags) = parse("Sub Bad()\n  y = 1\n  z = )\nEnd Sub\nSub Good()\nEnd Sub\n");
    assert_eq!(diags.len(), 1);
    let printed = print_module(&tree.expect("a tree"));
    assert_eq!(printed, "Sub Bad()\n    y = 1\n    ' (unparsed)\nEnd Sub\n\nSub Good()\nEnd Sub\n");
    let (again, diags) = parse(&printed);
    assert!(diags.is_empty(), "{diags:#?}");
    assert_eq!(again.map(|m| m.procs.iter().map(|p| p.body.len()).collect::<Vec<_>>()), Some(vec![1, 0]));
}

#[test]
fn module_items_keep_source_order() {
    let printed = round_trip("Sub A()\nEnd Sub\n#If X Then\nDim v\nSub B()\nEnd Sub\n#End If\n");
    assert_eq!(printed, "Sub A()\nEnd Sub\n\n#If X Then\nDim v\n\nSub B()\nEnd Sub\n\n#End If\n");
}

// ----- size -------------------------------------------------------------------------------

#[test]
fn deep_and_long_trees_print_on_a_small_stack() {
    let depth = usize::from(crate::parser::MAX_DEPTH);
    let n = 200_000;
    let sources = [
        format!("Sub T()\nx = {}1{}\nEnd Sub\n", "(".repeat(depth), ")".repeat(depth)),
        format!("Sub T()\n{}{}End Sub\n", "If x Then\n".repeat(depth), "End If\n".repeat(depth)),
        format!("Sub T()\nx = {}\nEnd Sub\n", vec!["-a ^ b"; n].join(" + ")),
        format!("Sub T()\nx = a{}\nEnd Sub\n", ".b(1)".repeat(n)),
    ];
    // The derived `Clone`/`PartialEq` behind `normalized` recurse, so on the small stack this
    // checks diagnostics and that printing is stable; equivalence is covered above.
    let worker = std::thread::Builder::new().stack_size(1 << 20).spawn(move || {
        sources
            .iter()
            .map(|src| {
                let (tree, diags) = parse(src);
                let printed = tree.as_ref().map(print_module).unwrap_or_default();
                let (again, again_diags) = parse(&printed);
                let stable = again.as_ref().map(print_module) == Some(printed);
                (diags.is_empty(), again_diags.is_empty(), stable)
            })
            .collect::<Vec<_>>()
    });
    let results = worker.expect("spawn").join().expect("no stack overflow");
    assert_eq!(results, vec![(true, true, true); 4]);
}

#[test]
fn modules_from_parse_are_reused_by_the_equivalence() {
    // `Module` is part of the helper's contract for 10.3: case and spans don't matter.
    let (a, _) = parse("SUB t()\nX = A + 1\nEND SUB\n");
    let (b, _) = parse("Sub T()\n    x   =   a + 1\nEnd Sub\n");
    let (a, b): (Module, Module) = (a.expect("a"), b.expect("b"));
    assert!(equivalent(&a, &b));
    let (c, _) = parse("Sub T()\nx = a + 2\nEnd Sub\n");
    assert!(!equivalent(&a, &c.expect("c")));
}
