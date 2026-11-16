use super::keywords::{KEYWORDS, cmp_ignore_ascii_case};
use super::{Keyword, LexError, MAX_SOURCE, Span, Token, TokenKind, lex, text};
use std::cmp::Ordering;

use TokenKind::{
    Comment, Date, Float, ForeignIdentifier, Hex, Identifier, Integer, LineContinuation, Newline, Octal, Operator, Punct, TypeSuffix, Whitespace,
};

/// Lexes `src` and checks the tokens cover it exactly: in order, contiguous, non-empty, and
/// joining to the input.
fn lex_ok(src: &str) -> Vec<Token> {
    let tokens = match lex(src) {
        Ok(tokens) => tokens,
        Err(err) => panic!("lex({src:?}) failed: {err}"),
    };
    let mut pos = 0;
    let mut joined = String::new();
    for token in &tokens {
        assert_eq!(token.span.start, pos, "gap or overlap before {token:?} in {src:?}");
        assert!(token.span.end > token.span.start, "empty token {token:?} in {src:?}");
        joined.push_str(text(src, token));
        pos = token.span.end;
    }
    assert_eq!(pos, src.len(), "tokens stop early in {src:?}");
    assert_eq!(joined, src);
    tokens
}

/// The non-whitespace tokens of `src` (newlines and comments kept) with their text.
fn sig(src: &str) -> Vec<(TokenKind, &str)> {
    lex_ok(src).iter().filter(|t| t.kind != Whitespace).map(|t| (t.kind, text(src, t))).collect()
}

fn kw(keyword: Keyword) -> TokenKind {
    TokenKind::Keyword(keyword)
}

fn err(error: LexError) -> TokenKind {
    TokenKind::Error(error)
}

#[test]
fn keyword_table_is_sorted_and_round_trips() {
    for pair in KEYWORDS.windows(2) {
        if let [(a, _), (b, _)] = pair {
            assert_eq!(cmp_ignore_ascii_case(a, b), Ordering::Less, "{a} must sort before {b}");
        }
    }
    assert_eq!(Keyword::all().count(), KEYWORDS.len());
    assert_eq!(Keyword::all().filter(|k| k.is_reserved()).count(), 172, "MS-VBAL §3.3.5.2 lists 172 reserved identifiers");
    for keyword in Keyword::all() {
        let spelling = keyword.as_str();
        assert_eq!(Keyword::lookup(spelling), Some(keyword));
        assert_eq!(Keyword::lookup(&spelling.to_ascii_uppercase()), Some(keyword));
        assert_eq!(Keyword::lookup(&spelling.to_ascii_lowercase()), Some(keyword));
    }
    assert_eq!(Keyword::lookup("Dimension"), None);
    assert_eq!(Keyword::lookup(""), None);
    assert!(!Keyword::Property.is_reserved());
    assert!(Keyword::Attribute.is_reserved());
}

#[test]
fn keywords_in_any_case() {
    assert_eq!(
        sig("DIM dim Dim dIm"),
        vec![(kw(Keyword::Dim), "DIM"), (kw(Keyword::Dim), "dim"), (kw(Keyword::Dim), "Dim"), (kw(Keyword::Dim), "dIm")]
    );
    assert_eq!(sig("end IF"), vec![(kw(Keyword::End), "end"), (kw(Keyword::If), "IF")]);
    assert_eq!(sig("Dimension"), vec![(Identifier, "Dimension")]);
}

#[test]
fn identifiers_and_type_suffixes() {
    let src = "a% b& c^ d! e# f@ g$";
    let expected: Vec<(TokenKind, &str)> = src
        .split(' ')
        .flat_map(|word| {
            let (name, suffix) = word.split_at(1);
            [(Identifier, name), (TypeSuffix, suffix)]
        })
        .collect();
    assert_eq!(sig(src), expected);
    assert_eq!(
        sig("Mid$(s, 1)"),
        vec![(kw(Keyword::Mid), "Mid"), (TypeSuffix, "$"), (Punct, "("), (Identifier, "s"), (Punct, ","), (Integer, "1"), (Punct, ")")]
    );
    assert_eq!(sig("my_var1"), vec![(Identifier, "my_var1")]);
    assert_eq!(sig("niño"), vec![(Identifier, "niño")]);
}

#[test]
fn suffix_characters_that_continue_a_term_are_operators() {
    assert_eq!(sig("rs!Name"), vec![(Identifier, "rs"), (Punct, "!"), (Identifier, "Name")]);
    assert_eq!(sig("rs![Field Name]"), vec![(Identifier, "rs"), (Punct, "!"), (ForeignIdentifier, "[Field Name]")]);
    assert_eq!(sig("a&b"), vec![(Identifier, "a"), (Operator, "&"), (Identifier, "b")]);
    assert_eq!(sig("a^2"), vec![(Identifier, "a"), (Operator, "^"), (Integer, "2")]);
}

#[test]
fn bracketed_identifiers() {
    assert_eq!(
        sig("x = [A1] + [my name]"),
        vec![(Identifier, "x"), (Operator, "="), (ForeignIdentifier, "[A1]"), (Operator, "+"), (ForeignIdentifier, "[my name]")]
    );
    assert_eq!(sig("[]"), vec![(err(LexError::EmptyForeignName), "[]")]);
    assert_eq!(sig("x [abc\ny"), vec![(Identifier, "x"), (err(LexError::UnterminatedForeignName), "[abc"), (Newline, "\n"), (Identifier, "y")]);
}

#[test]
fn integer_hex_and_octal_literals() {
    let cases = [
        ("123", Integer),
        ("123%", Integer),
        ("123&", Integer),
        ("123^", Integer),
        ("&HFF", Hex),
        ("&hff&", Hex),
        ("&H7FFF%", Hex),
        ("&O17", Octal),
        ("&o17^", Octal),
        ("&17", Octal),
    ];
    for (src, kind) in cases {
        assert_eq!(sig(src), vec![(kind, src)], "{src}");
    }
    assert_eq!(sig("&H"), vec![(Operator, "&"), (Identifier, "H")]);
    assert_eq!(sig("&O9"), vec![(Operator, "&"), (Identifier, "O9")]);
    assert_eq!(sig("&8"), vec![(Operator, "&"), (Integer, "8")]);
    assert_eq!(sig("1&H10"), vec![(Integer, "1"), (Hex, "&H10")]);
}

#[test]
fn float_literals() {
    for src in ["1.5", "1.", ".5", "1e10", "1E+10", "1.5E-3", "2D+4", ".5e2", "3!", "3#", "3@", "1.5!", "2.5e3#"] {
        assert_eq!(sig(src), vec![(Float, src)], "{src}");
    }
    assert_eq!(sig("1e"), vec![(Integer, "1"), (Identifier, "e")]);
    assert_eq!(sig("x.5"), vec![(Identifier, "x"), (Float, ".5")]);
    assert_eq!(sig("a.b"), vec![(Identifier, "a"), (Punct, "."), (Identifier, "b")]);
}

#[test]
fn strings_with_escaped_quotes() {
    assert_eq!(sig(r#"s = "say ""hi""""#), vec![(Identifier, "s"), (Operator, "="), (TokenKind::String, r#""say ""hi""""#)]);
    assert_eq!(sig(r#""""#), vec![(TokenKind::String, r#""""#)]);
    assert_eq!(sig(r#""a _ b""#), vec![(TokenKind::String, r#""a _ b""#)]);
    assert_eq!(sig("\"ñ€\""), vec![(TokenKind::String, "\"ñ€\"")]);
}

#[test]
fn unterminated_string_is_an_error_to_the_end_of_the_line() {
    assert_eq!(
        sig("x = \"abc ' not a comment\r\ny = 1"),
        vec![
            (Identifier, "x"),
            (Operator, "="),
            (err(LexError::UnterminatedString), "\"abc ' not a comment"),
            (Newline, "\r\n"),
            (Identifier, "y"),
            (Operator, "="),
            (Integer, "1"),
        ]
    );
    assert_eq!(sig("\"abc\"\""), vec![(err(LexError::UnterminatedString), "\"abc\"\"")]);
    // A trailing continuation stays its own token, so the logical line still joins up.
    assert_eq!(sig("\"abc _\nx"), vec![(err(LexError::UnterminatedString), "\"abc"), (LineContinuation, " _\n"), (Identifier, "x")]);
}

#[test]
fn date_literals() {
    for src in [
        "#1/2/2003#",
        "#12:30 PM#",
        "#12:30:15#",
        "#10 am#",
        "#January 1, 2003#",
        "#1-Jan-2003#",
        "# 2003-01-02 10:00:00 #",
        "#1/2/2003 12:00#",
        "#1/2 12:00#",
    ] {
        assert_eq!(sig(src), vec![(Date, src)], "{src}");
    }
    assert_eq!(
        sig("If d > #1/1/2000# Then"),
        vec![(kw(Keyword::If), "If"), (Identifier, "d"), (Operator, ">"), (Date, "#1/1/2000#"), (kw(Keyword::Then), "Then")]
    );
    assert_eq!(sig("Debug.Print #1/1/2000#"), vec![(kw(Keyword::Debug), "Debug"), (Punct, "."), (kw(Keyword::Print), "Print"), (Date, "#1/1/2000#")]);
}

#[test]
fn unterminated_date_is_an_error_to_the_end_of_the_line() {
    assert_eq!(
        sig("d = #1/2/2003 ' oops\nx"),
        vec![(Identifier, "d"), (Operator, "="), (err(LexError::UnterminatedDate), "#1/2/2003 ' oops"), (Newline, "\n"), (Identifier, "x")]
    );
    assert_eq!(sig("t = #12:30"), vec![(Identifier, "t"), (Operator, "="), (err(LexError::UnterminatedDate), "#12:30")]);
}

#[test]
fn hash_as_file_number_suffix_and_directive() {
    assert_eq!(sig("Print #1, x"), vec![(kw(Keyword::Print), "Print"), (Punct, "#"), (Integer, "1"), (Punct, ","), (Identifier, "x")]);
    assert_eq!(
        sig("Write #1, 1/2"),
        vec![(kw(Keyword::Write), "Write"), (Punct, "#"), (Integer, "1"), (Punct, ","), (Integer, "1"), (Operator, "/"), (Integer, "2")]
    );
    assert_eq!(sig("Print #1, 2#"), vec![(kw(Keyword::Print), "Print"), (Punct, "#"), (Integer, "1"), (Punct, ","), (Float, "2#")]);
    assert_eq!(
        sig("Open f For Input As #1"),
        vec![
            (kw(Keyword::Open), "Open"),
            (Identifier, "f"),
            (kw(Keyword::For), "For"),
            (kw(Keyword::Input), "Input"),
            (kw(Keyword::As), "As"),
            (Punct, "#"),
            (Integer, "1")
        ]
    );
    assert_eq!(sig("Close #1, #2"), vec![(kw(Keyword::Close), "Close"), (Punct, "#"), (Integer, "1"), (Punct, ","), (Punct, "#"), (Integer, "2")]);
    assert_eq!(sig("x# = 1"), vec![(Identifier, "x"), (TypeSuffix, "#"), (Operator, "="), (Integer, "1")]);
    assert_eq!(sig("#If Win64 Then"), vec![(Punct, "#"), (kw(Keyword::If), "If"), (Identifier, "Win64"), (kw(Keyword::Then), "Then")]);
    assert_eq!(sig("#1#"), vec![(Punct, "#"), (Float, "1#")]);
}

#[test]
fn line_continuations() {
    for terminator in ["\r\n", "\n", "\r", "\u{2028}"] {
        let src = format!("a = 1 _{terminator}  + 2");
        assert_eq!(
            sig(&src),
            vec![
                (Identifier, "a"),
                (Operator, "="),
                (Integer, "1"),
                (LineContinuation, format!(" _{terminator}").as_str()),
                (Operator, "+"),
                (Integer, "2")
            ]
        );
    }
    assert_eq!(sig("a \t_  \t\nb"), vec![(Identifier, "a"), (LineContinuation, " \t_  \t\n"), (Identifier, "b")]);
    assert_eq!(sig("a _"), vec![(Identifier, "a"), (LineContinuation, " _")]);
    assert_eq!(sig("a_\nb"), vec![(Identifier, "a_"), (Newline, "\n"), (Identifier, "b")]);
    assert_eq!(sig("a _b"), vec![(Identifier, "a"), (err(LexError::InvalidCharacter), "_"), (Identifier, "b")]);
    assert_eq!(sig("1_"), vec![(Integer, "1"), (err(LexError::InvalidCharacter), "_")]);
}

#[test]
fn comments() {
    assert_eq!(
        sig("x = 1 ' note\ny"),
        vec![(Identifier, "x"), (Operator, "="), (Integer, "1"), (Comment, "' note"), (Newline, "\n"), (Identifier, "y")]
    );
    assert_eq!(
        sig("Rem a \"note\"\r\nREM\nrem"),
        vec![(Comment, "Rem a \"note\""), (Newline, "\r\n"), (Comment, "REM"), (Newline, "\n"), (Comment, "rem")]
    );
    assert_eq!(sig("x = 1: Rem done"), vec![(Identifier, "x"), (Operator, "="), (Integer, "1"), (Punct, ":"), (Comment, "Rem done")]);
    assert_eq!(sig("10 Rem label"), vec![(Integer, "10"), (Comment, "Rem label")]);
    assert_eq!(sig("Remark = 1"), vec![(Identifier, "Remark"), (Operator, "="), (Integer, "1")]);
    assert_eq!(sig("obj.Rem = 1"), vec![(Identifier, "obj"), (Punct, "."), (kw(Keyword::Rem), "Rem"), (Operator, "="), (Integer, "1")]);
}

#[test]
fn comments_carry_on_past_a_line_continuation() {
    let src = "' one _\r\ntwo\nx";
    let tokens = lex_ok(src);
    assert_eq!(sig(src), vec![(Comment, "' one _\r\ntwo"), (Newline, "\n"), (Identifier, "x")]);
    assert_eq!(tokens.last().map(|t| t.span), Some(Span { line: 3, col: 1, start: 13, end: 14 }));
    assert_eq!(sig("Rem a _\nb"), vec![(Comment, "Rem a _\nb")]);
}

#[test]
fn attribute_lines() {
    assert_eq!(
        sig("Attribute VB_Name = \"Module1\"\r\n"),
        vec![
            (kw(Keyword::Attribute), "Attribute"),
            (kw(Keyword::VbName), "VB_Name"),
            (Operator, "="),
            (TokenKind::String, "\"Module1\""),
            (Newline, "\r\n")
        ]
    );
}

#[test]
fn operators_and_punctuation() {
    let src = "& * + - / \\ ^ < = > , . ! ( ) : ; ?";
    let kinds: Vec<TokenKind> = sig(src).into_iter().map(|(kind, _)| kind).collect();
    let mut expected = vec![Operator; 10];
    expected.extend([Punct; 8]);
    assert_eq!(kinds, expected);
    assert_eq!(sig("a<>b"), vec![(Identifier, "a"), (Operator, "<"), (Operator, ">"), (Identifier, "b")]);
    assert_eq!(sig("f x:=1"), vec![(Identifier, "f"), (Identifier, "x"), (Punct, ":"), (Operator, "="), (Integer, "1")]);
}

#[test]
fn invalid_characters_group_into_one_error() {
    assert_eq!(sig("a ~~` b"), vec![(Identifier, "a"), (err(LexError::InvalidCharacter), "~~`"), (Identifier, "b")]);
    assert_eq!(
        sig("% $ @"),
        vec![(err(LexError::InvalidCharacter), "%"), (err(LexError::InvalidCharacter), "$"), (err(LexError::InvalidCharacter), "@")]
    );
    // U+00A0 is a CP2 character, not WSC; non-letter non-ASCII characters are invalid.
    assert_eq!(sig("a\u{a0}b"), vec![(Identifier, "a"), (err(LexError::InvalidCharacter), "\u{a0}"), (Identifier, "b")]);
    assert_eq!(sig("\u{0}\u{1}"), vec![(err(LexError::InvalidCharacter), "\u{0}\u{1}")]);
}

#[test]
fn whitespace_kinds() {
    let tokens = lex_ok("a\u{3000}\t\u{19} b");
    let kinds: Vec<TokenKind> = tokens.iter().map(|t| t.kind).collect();
    assert_eq!(kinds, vec![Identifier, Whitespace, Identifier]);
}

#[test]
fn spans_count_lines_and_character_columns() {
    let src = "Dim café As String\r\nx = \"ñ\" ' é\rEnd\u{2029}y";
    let tokens = lex_ok(src);
    let find = |needle: &str| match tokens.iter().find(|t| text(src, t) == needle) {
        Some(token) => token.span,
        None => panic!("no token {needle:?}"),
    };
    assert_eq!(find("Dim"), Span { line: 1, col: 1, start: 0, end: 3 });
    assert_eq!(find("café"), Span { line: 1, col: 5, start: 4, end: 9 });
    // `é` is two bytes, so the byte offset runs one ahead of the column from here.
    assert_eq!(find("As"), Span { line: 1, col: 10, start: 10, end: 12 });
    assert_eq!(find("\r\n"), Span { line: 1, col: 19, start: 19, end: 21 });
    assert_eq!(find("x"), Span { line: 2, col: 1, start: 21, end: 22 });
    assert_eq!(find("\"ñ\""), Span { line: 2, col: 5, start: 25, end: 29 });
    assert_eq!(find("' é"), Span { line: 2, col: 9, start: 30, end: 34 });
    assert_eq!(find("End"), Span { line: 3, col: 1, start: 35, end: 38 });
    assert_eq!(find("y"), Span { line: 4, col: 1, start: 41, end: 42 });
}

#[test]
fn lone_cr_then_lf_tokens_count_as_one_break_each() {
    let tokens = lex_ok("a\n\nb\r\rc");
    let lines: Vec<u32> = tokens.iter().filter(|t| t.kind == Identifier).map(|t| t.span.line).collect();
    assert_eq!(lines, vec![1, 3, 5]);
}

#[test]
fn source_size_limit() {
    let at_limit = " ".repeat(MAX_SOURCE);
    assert_eq!(lex_ok(&at_limit).len(), 1);
    drop(at_limit);
    let over = " ".repeat(MAX_SOURCE + 1);
    let result = lex(&over);
    drop(over);
    match result {
        Err(error) => {
            assert_eq!(error.len, MAX_SOURCE + 1);
            assert!(error.to_string().contains("source too large"));
        }
        Ok(tokens) => panic!("expected SourceTooLarge, got {} tokens", tokens.len()),
    }
}

#[test]
fn empty_source_has_no_tokens() {
    assert!(lex_ok("").is_empty());
}

#[test]
fn small_module_round_trips() {
    let src = "Attribute VB_Name = \"Module1\"\r\nOption Explicit\r\n\r\nPublic Sub Test(ByVal n As Long)\r\n    Dim s$, d As Date\r\n    d = #1/2/2003#\r\n    s = \"a\" & _\r\n        \"b\" ' join\r\n    If n > &HFF Then MsgBox s: Rem done\r\nEnd Sub\r\n";
    let tokens = lex_ok(src);
    assert!(tokens.iter().all(|t| !matches!(t.kind, TokenKind::Error(_))), "{tokens:?}");
    assert_eq!(tokens.iter().filter(|t| t.kind == Newline).count(), 9);
    assert_eq!(tokens.iter().filter(|t| t.kind == LineContinuation).count(), 1);
}
