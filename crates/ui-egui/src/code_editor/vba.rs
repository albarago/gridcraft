//! VBA colouring for View Code (Phase 1): a static keyword list, double-quoted strings with `""`
//! escapes, comments (`'` and `Rem`) and numbers (with `&H` and `&O` literals and type
//! suffixes), all case-insensitive. It reads no grammar: the VBA lexer's spans replace it later
//! (VBA-2.1), so it stays in this one module.

use std::ops::Range;

use super::Token;

/// VBA's reserved words and the statement and type keywords, lower case ([MS-VBAL] §3.3.5).
const KEYWORDS: &[&str] = &[
    "addressof",
    "alias",
    "and",
    "as",
    "attribute",
    "boolean",
    "byref",
    "byte",
    "byval",
    "call",
    "case",
    "const",
    "currency",
    "date",
    "declare",
    "dim",
    "do",
    "double",
    "each",
    "else",
    "elseif",
    "empty",
    "end",
    "enum",
    "eqv",
    "erase",
    "error",
    "event",
    "exit",
    "explicit",
    "false",
    "for",
    "friend",
    "function",
    "get",
    "global",
    "gosub",
    "goto",
    "if",
    "imp",
    "implements",
    "in",
    "integer",
    "is",
    "let",
    "lib",
    "like",
    "long",
    "longlong",
    "longptr",
    "loop",
    "me",
    "mod",
    "new",
    "next",
    "not",
    "nothing",
    "null",
    "object",
    "on",
    "option",
    "optional",
    "or",
    "paramarray",
    "preserve",
    "private",
    "property",
    "ptrsafe",
    "public",
    "raiseevent",
    "redim",
    "resume",
    "return",
    "select",
    "set",
    "single",
    "static",
    "step",
    "stop",
    "string",
    "sub",
    "then",
    "to",
    "true",
    "type",
    "typeof",
    "until",
    "variant",
    "wend",
    "while",
    "with",
    "withevents",
    "xor",
];

/// The type characters that may end a number literal (`10&`, `1.5#`).
const SUFFIXES: &[char] = &['%', '&', '^', '!', '#', '@'];

/// Whether `word` is a VBA keyword, ignoring case.
pub fn is_keyword(word: &str) -> bool {
    KEYWORDS.iter().any(|k| k.eq_ignore_ascii_case(word))
}

/// The coloured spans of VBA `text`: byte ranges on char boundaries, in order, not overlapping.
/// Never panics, whatever the input.
pub fn highlight(text: &str) -> Vec<(Range<usize>, Token)> {
    let chars: Vec<(usize, char)> = text.char_indices().collect();
    let end_of = |i: usize| chars.get(i).map(|(b, _)| *b).unwrap_or(text.len());
    let char_at = |i: usize| chars.get(i).map(|(_, c)| *c);
    // End (exclusive, in chars) of the line holding char `i`, before its line break.
    let line_end = |mut j: usize| {
        while char_at(j).is_some_and(|c| c != '\n' && c != '\r') {
            j += 1;
        }
        j
    };
    let mut out = Vec::new();
    let mut i = 0;
    while let Some(&(start, c)) = chars.get(i) {
        if c == '\'' {
            let j = line_end(i);
            out.push((start..end_of(j), Token::Comment));
            i = j;
        } else if c == '"' {
            // `""` inside a string is one quote; a string ends at its line's end when unclosed.
            let mut j = i + 1;
            loop {
                match char_at(j) {
                    None | Some('\n' | '\r') => break,
                    Some('"') if char_at(j + 1) == Some('"') => j += 2,
                    Some('"') => {
                        j += 1;
                        break;
                    }
                    Some(_) => j += 1,
                }
            }
            out.push((start..end_of(j), Token::String));
            i = j;
        } else if c == '&'
            && let Some(radix) = char_at(i + 1).and_then(|r| match r.to_ascii_uppercase() {
                'H' => Some(16),
                'O' => Some(8),
                _ => None,
            })
            && char_at(i + 2).is_some_and(|d| d.is_digit(radix))
        {
            let mut j = i + 2;
            while char_at(j).is_some_and(|d| d.is_digit(radix)) {
                j += 1;
            }
            if char_at(j).is_some_and(|s| SUFFIXES.contains(&s)) {
                j += 1;
            }
            out.push((start..end_of(j), Token::Number));
            i = j;
        } else if c.is_ascii_digit() || (c == '.' && char_at(i + 1).is_some_and(|d| d.is_ascii_digit())) {
            let mut j = i + 1;
            loop {
                match char_at(j) {
                    Some(d) if d.is_ascii_digit() || d == '.' => j += 1,
                    // An exponent (`1E10`, `2.5D-3`), with its sign.
                    Some('e' | 'E' | 'd' | 'D') => {
                        j += 1;
                        if char_at(j).is_some_and(|s| s == '+' || s == '-') && char_at(j + 1).is_some_and(|d| d.is_ascii_digit()) {
                            j += 1;
                        }
                    }
                    _ => break,
                }
            }
            if char_at(j).is_some_and(|s| SUFFIXES.contains(&s)) {
                j += 1;
            }
            out.push((start..end_of(j), Token::Number));
            i = j;
        } else if c.is_alphabetic() || c == '_' {
            let mut j = i + 1;
            while char_at(j).is_some_and(|d| d.is_alphanumeric() || d == '_') {
                j += 1;
            }
            let word = text.get(start..end_of(j)).unwrap_or("");
            if word.eq_ignore_ascii_case("rem") && char_at(j).is_none_or(|d| d == ' ' || d == '\t' || d == '\n' || d == '\r' || d == ':') {
                // `Rem` comments out the rest of its line, as `'` does.
                let k = line_end(j);
                out.push((start..end_of(k), Token::Comment));
                i = k;
                continue;
            }
            if is_keyword(word) {
                out.push((start..end_of(j), Token::Keyword));
            }
            i = j;
        } else {
            i += 1;
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn spans(text: &str) -> Vec<(&str, Token)> {
        highlight(text).into_iter().map(|(r, k)| (text.get(r).unwrap_or("?"), k)).collect()
    }

    #[test]
    fn keywords_ignore_case() {
        assert_eq!(
            spans("Sub Main()\n  dim x AS Long\nEND SUB"),
            vec![
                ("Sub", Token::Keyword),
                ("dim", Token::Keyword),
                ("AS", Token::Keyword),
                ("Long", Token::Keyword),
                ("END", Token::Keyword),
                ("SUB", Token::Keyword),
            ]
        );
        // Names that only contain a keyword stay plain.
        assert_eq!(spans("Subtotal = Endless + MyDim"), vec![]);
        assert_eq!(spans("Option Explicit"), vec![("Option", Token::Keyword), ("Explicit", Token::Keyword)]);
    }

    #[test]
    fn strings_keep_doubled_quotes() {
        assert_eq!(spans(r#"s = "say ""hi"" now""#), vec![(r#""say ""hi"" now""#, Token::String)]);
        assert_eq!(spans(r#"s = """""#), vec![(r#""""""#, Token::String)]);
        // An apostrophe inside a string is not a comment, and an unclosed string stops at the line's end.
        assert_eq!(spans("MsgBox \"it's\" ' note"), vec![("\"it's\"", Token::String), ("' note", Token::Comment)]);
        assert_eq!(spans("x = \"open\nIf"), vec![("\"open", Token::String), ("If", Token::Keyword)]);
    }

    #[test]
    fn comments_with_apostrophe_and_rem() {
        assert_eq!(spans("' Dim not code\r\nDim"), vec![("' Dim not code", Token::Comment), ("Dim", Token::Keyword)]);
        assert_eq!(spans("REM If then\nrem\nRemark = 1"), vec![("REM If then", Token::Comment), ("rem", Token::Comment), ("1", Token::Number)]);
        assert_eq!(spans("x = 1: Rem done"), vec![("1", Token::Number), ("Rem done", Token::Comment)]);
    }

    #[test]
    fn numbers() {
        assert_eq!(
            spans("a = 10& + 1.5# + 2E-3 + &HFF + &o17 + .5"),
            vec![
                ("10&", Token::Number),
                ("1.5#", Token::Number),
                ("2E-3", Token::Number),
                ("&HFF", Token::Number),
                ("&o17", Token::Number),
                (".5", Token::Number)
            ]
        );
        // `&` as the join operator is plain, also before a name starting with H.
        assert_eq!(spans("s = a & Hello"), vec![]);
    }

    #[test]
    fn spans_are_ordered_on_char_boundaries() {
        for text in ["Ünïcode \"ß\"\"é\" ' 日本", "\"", "'", "&H", "1E", "Rem", "é1.2.3e+", "\r\n\"\r"] {
            let mut pos = 0;
            for (r, _) in highlight(text) {
                assert!(r.start >= pos && r.start < r.end && r.end <= text.len(), "{text:?} {r:?}");
                assert!(text.is_char_boundary(r.start) && text.is_char_boundary(r.end), "{text:?} {r:?}");
                pos = r.end;
            }
        }
    }
}
