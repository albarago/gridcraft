//! VBA colouring for View Code, from the VBA lexer's tokens (`gridcraft_engine::vba::lex`, the
//! same lexer the parser and the compatibility report read), so the colours follow the language
//! as the analyzer sees it:
//!
//! - keywords: every reserved identifier and contextual word in the lexer's table (`Sub`, `Dim`,
//!   `Long`, `Explicit`, `Property`, …), except a word right after `.` or `!`, which names a
//!   member (`ws.Range("A1").Text`) and stays plain;
//! - strings: closed `"…"` literals, with `""` escapes;
//! - numbers: integer, `&H`/`&O` and floating-point literals with their type suffixes, and
//!   `#…#` date literals;
//! - comments: `'` and `Rem`, up to the end of the logical line (a ` _` continuation carries a
//!   comment on to the next line);
//! - errors: what the lexer couldn't read (an unclosed string, date or `[name]`, a character VBA
//!   doesn't allow).
//!
//! Identifiers, operators, punctuation and a type suffix after a name (`s$`) are plain. Source
//! longer than the lexer's limit (16 MiB) is shown plain.
//!
//! The editor lays the text out on every frame, so [`highlight_cached`] keeps the spans of the
//! last few texts it coloured, keyed by a hash and the length of the text: the shown source is
//! lexed once, not once a frame.

use std::cell::RefCell;
use std::ops::Range;
use std::rc::Rc;

use gridcraft_engine::vba::{TokenKind, lex};

use super::Token;

/// Coloured spans, shared between frames.
pub type Spans = Rc<[(Range<usize>, Token)]>;

/// How many texts [`highlight_cached`] keeps: one View Code source, with room for a few others.
const CACHE_SLOTS: usize = 4;

thread_local! {
    /// (hash, length, spans), least recently used first.
    static CACHE: RefCell<Vec<(u64, usize, Spans)>> = const { RefCell::new(Vec::new()) };
}

/// The colour role of a token of `kind`; `None` for plain text. `after_member_dot` says the
/// token comes right after `.` or `!`.
fn role(kind: TokenKind, after_member_dot: bool) -> Option<Token> {
    match kind {
        TokenKind::Keyword(_) if !after_member_dot => Some(Token::Keyword),
        TokenKind::Integer | TokenKind::Hex | TokenKind::Octal | TokenKind::Float | TokenKind::Date => Some(Token::Number),
        TokenKind::String => Some(Token::String),
        TokenKind::Comment => Some(Token::Comment),
        TokenKind::Error(_) => Some(Token::Error),
        TokenKind::Keyword(_)
        | TokenKind::Identifier
        | TokenKind::ForeignIdentifier
        | TokenKind::TypeSuffix
        | TokenKind::Operator
        | TokenKind::Punct
        | TokenKind::LineContinuation
        | TokenKind::Newline
        | TokenKind::Whitespace => None,
    }
}

/// The coloured spans of VBA `text`: byte ranges on char boundaries, in order, not overlapping.
/// Never panics, whatever the input. See the module docs for the roles.
pub fn highlight(text: &str) -> Vec<(Range<usize>, Token)> {
    let Ok(tokens) = lex(text) else { return Vec::new() };
    let mut out = Vec::new();
    let mut after_member_dot = false;
    for t in tokens {
        let range = t.span.start..t.span.end;
        let Some(slice) = text.get(range.clone()) else { continue };
        if let Some(role) = role(t.kind, after_member_dot)
            && !slice.is_empty()
        {
            out.push((range, role));
        }
        after_member_dot = t.kind == TokenKind::Punct && (slice == "." || slice == "!");
    }
    out
}

/// [`highlight`], kept for the last [`CACHE_SLOTS`] texts (per thread), so a text shown frame
/// after frame is lexed once.
pub fn highlight_cached(text: &str) -> Spans {
    let key = (egui::util::hash(text), text.len());
    let hit = CACHE.with(|c| {
        let mut c = c.try_borrow_mut().ok()?;
        let at = c.iter().position(|(h, len, _)| (*h, *len) == key)?;
        let slot = c.remove(at);
        let spans = slot.2.clone();
        c.push(slot);
        Some(spans)
    });
    if let Some(spans) = hit {
        return spans;
    }
    let spans: Spans = highlight(text).into();
    CACHE.with(|c| {
        if let Ok(mut c) = c.try_borrow_mut() {
            if c.len() >= CACHE_SLOTS {
                c.remove(0);
            }
            c.push((key.0, key.1, spans.clone()));
        }
    });
    spans
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
    fn a_keyword_after_a_member_dot_is_a_member_name() {
        assert_eq!(spans("x = ws.Range(\"A1\").Text"), vec![("\"A1\"", Token::String)]);
        assert_eq!(spans("Set r = rs!Line"), vec![("Set", Token::Keyword)]);
        assert_eq!(
            spans("With ws\n.Text = 1\nEnd With"),
            vec![("With", Token::Keyword), ("1", Token::Number), ("End", Token::Keyword), ("With", Token::Keyword)]
        );
    }

    #[test]
    fn strings_keep_doubled_quotes() {
        assert_eq!(spans(r#"s = "say ""hi"" now""#), vec![(r#""say ""hi"" now""#, Token::String)]);
        assert_eq!(spans(r#"s = """""#), vec![(r#""""""#, Token::String)]);
        // An apostrophe inside a string is not a comment.
        assert_eq!(spans("MsgBox \"it's\" ' note"), vec![("\"it's\"", Token::String), ("' note", Token::Comment)]);
    }

    #[test]
    fn what_the_lexer_cannot_read_is_an_error() {
        // An unclosed string stops at the line's end, and the next line is read as usual.
        assert_eq!(spans("x = \"open\nIf"), vec![("\"open", Token::Error), ("If", Token::Keyword)]);
        assert!(spans("x = [open").iter().any(|(_, k)| *k == Token::Error));
    }

    #[test]
    fn comments_with_apostrophe_and_rem() {
        assert_eq!(spans("' Dim not code\r\nDim"), vec![("' Dim not code", Token::Comment), ("Dim", Token::Keyword)]);
        assert_eq!(spans("REM If then\nrem\nRemark = 1"), vec![("REM If then", Token::Comment), ("rem", Token::Comment), ("1", Token::Number)]);
        assert_eq!(spans("x = 1: Rem done"), vec![("1", Token::Number), ("Rem done", Token::Comment)]);
        // A line continuation carries a comment on to the next line.
        assert_eq!(spans("' one _\ntwo\nDim"), vec![("' one _\ntwo", Token::Comment), ("Dim", Token::Keyword)]);
    }

    #[test]
    fn numbers_and_dates() {
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
        assert_eq!(spans("d = #1/2/2024#"), vec![("#1/2/2024#", Token::Number)]);
        // A type suffix after a name is part of the name, not a number.
        assert_eq!(spans("s$ = n%"), vec![]);
    }

    #[test]
    fn spans_are_ordered_on_char_boundaries() {
        for text in ["Ünïcode \"ß\"\"é\" ' 日本", "\"", "'", "&H", "1E", "Rem", "é1.2.3e+", "\r\n\"\r", "#", "[", "x\u{2028}y", "\u{0}"] {
            let mut pos = 0;
            for (r, _) in highlight(text) {
                assert!(r.start >= pos && r.start < r.end && r.end <= text.len(), "{text:?} {r:?}");
                assert!(text.is_char_boundary(r.start) && text.is_char_boundary(r.end), "{text:?} {r:?}");
                pos = r.end;
            }
        }
    }

    #[test]
    fn the_same_text_is_lexed_once() {
        let a = "Sub A()\nEnd Sub\n".repeat(50);
        let first = highlight_cached(&a);
        assert!(Rc::ptr_eq(&first, &highlight_cached(&a)), "a second frame reuses the spans");
        assert_eq!(&*first, highlight(&a).as_slice());
        let b = format!("{a}' changed");
        let other = highlight_cached(&b);
        assert!(!Rc::ptr_eq(&first, &other));
        assert_eq!(other.last().map(|(_, k)| *k), Some(Token::Comment));
        // More texts than the cache keeps push the oldest out; the spans are still right.
        for i in 0..CACHE_SLOTS {
            let _ = highlight_cached(&format!("x = {i}"));
        }
        assert_eq!(&*highlight_cached(&a), highlight(&a).as_slice());
    }
}
