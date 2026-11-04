//! A lossless VBA lexer following [MS-VBAL] §3 "VBA Program Text".
//!
//! [`lex`] turns module source into [`Token`]s that cover the input exactly: every byte belongs
//! to one token, tokens are in order, and each starts where the previous one ends. Trivia
//! (whitespace, line terminators, line continuations, comments) are tokens too, and text the
//! lexer can't read becomes [`TokenKind::Error`] tokens, so joining the token texts always gives
//! back the input byte for byte (design property P5).
//!
//! # How the grammar maps to tokens
//!
//! - Physical lines (§3.2.1) end in CRLF, CR, LF, U+2028 or U+2029. Each terminator is one
//!   [`TokenKind::Newline`]; CRLF counts as one line break in [`Span::line`].
//! - `line-continuation = 1*WSC underscore line-terminator` (§3.2.2) is one
//!   [`TokenKind::LineContinuation`] that includes the leading whitespace and the terminator.
//!   Whitespace between the `_` and the terminator is accepted too, and a continuation may end at
//!   the end of the input (§3.2.1 lets a last unterminated line act as if terminated). `a_` is an
//!   identifier, not a continuation, because the `_` must follow whitespace.
//! - WSC (§3.2.2) is tab, U+0019, space, U+3000 and the other Unicode `Zs` characters except
//!   U+00A0, which is a CP2 character and so excluded by `most-Unicode-class-Zs`.
//! - Comments: `'` up to the end of the logical line, and `Rem` (a whole word) up to the end of
//!   the logical line, are each one [`TokenKind::Comment`]. `comment-body` (§3.3.1) includes
//!   line continuations, so a comment ending in ` _` carries on onto the next physical line and
//!   the one token spans both. `Rem` is reserved, so it starts a comment wherever it appears
//!   except right after `.` or `!` (a member named `Rem`). That includes after a line-number
//!   label (`10 Rem x`), which is wider than "at statement start or after `:`".
//! - Words: a Latin identifier (§3.3.5) is `A-Z a-z` followed by letters, digits and `_`. Non-ASCII
//!   letters may start a word and non-ASCII letters and digits may continue it, an approximation
//!   of the code-page identifier sets (§3.3.5.1) that needs no code-page tables. A word that
//!   matches the keyword table case-insensitively is a [`TokenKind::Keyword`] (§3.3.5.2), others
//!   are [`TokenKind::Identifier`]. `Attribute VB_Name = "M"` lexes as keyword, keyword, operator,
//!   string; the parser handles `Attribute` lines.
//! - `type-suffix` (§3.3.5.3: `% & ^ ! # @ $`) right after a word is its own
//!   [`TokenKind::TypeSuffix`] token, but only when the character after the suffix can't continue
//!   an expression term: `rs!Name` and `rs![Name]` are bang member access, `a&b` is a
//!   concatenation, `a^2` is a power. Number literals keep their suffix inside the literal token,
//!   as the INTEGER and FLOAT rules do (§3.3.2), under the same rule.
//! - Numbers (§3.3.2): decimal, `&H` hex, `&O` and bare `&` octal ([`TokenKind::Integer`],
//!   [`TokenKind::Hex`], [`TokenKind::Octal`], with an optional `% & ^` suffix) and floats
//!   ([`TokenKind::Float`]: `1.5`, `1.`, `.5`, exponents with `E` or `D`, or a decimal with a
//!   `! # @` suffix). Range checks (the §3.3.2 validity table) are left to later phases.
//! - Strings (§3.3.4): `"…"` with `""` for a quote. A string with no closing quote before the end
//!   of the physical line is an `Error(UnterminatedString)` up to the end of that line (Req 11.5),
//!   even though §3.3.4 lets a STRING end at LINE-END or a line continuation. If the line ends in
//!   a continuation, that continuation stays its own token so the logical line still joins up.
//! - Dates (§3.3.3): `"#" *WSC date-or-time *WSC "#"`. `#` is also a file-number prefix
//!   (`Print #1, x`), a type suffix (`x#`) and the conditional-compilation marker (`#If`), so a
//!   `#` is a date only when the text after it reads as a `date-or-time` followed by a closing
//!   `#`, and never right after a statement keyword that takes a file number (`Print`, `Write`,
//!   `Input`, `Get`, `Put`, `Seek`, `Close`, `Lock`, `Unlock`, `As`, `Width`). It is an
//!   `Error(UnterminatedDate)` up to the end of the line when the text after it reads as a
//!   date-or-time that uses `/`, `-`, `:` or a month or am/pm word and no `#` follows on that
//!   line. Otherwise it is a [`TokenKind::Punct`].
//! - `[name]` (FOREIGN-NAME, §3.3.5.3) is one [`TokenKind::ForeignIdentifier`]; a `[` with no `]`
//!   on its line is an `Error(UnterminatedForeignName)` to the end of the line, and `[]` is an
//!   `Error(EmptyForeignName)` because `foreign-identifier` needs at least one character.
//! - `special-token` (§3.3.1) characters are one token each. The arithmetic, comparison and
//!   concatenation characters (`& * + - / \ ^ < = >`) are [`TokenKind::Operator`], the rest
//!   (`, . ! # ( ) : ; ?`) are [`TokenKind::Punct`]. The grammar spells `<=`, `<>` and `:=` as two
//!   tokens (§5.6.9, §5.6.13.1), so the parser joins them.
//! - Any other character (`_` not in a continuation, a lone `%`, `$`, `@`, `]`, control
//!   characters, non-letter non-ASCII characters) is an `Error(InvalidCharacter)`; consecutive ones
//!   share one token.
//!
//! # Limits
//!
//! Sources longer than [`MAX_SOURCE`] bytes give [`SourceTooLarge`] and no tokens (Req 11.6).
//! There is at most one token per input byte.
//!
//! [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/

mod date;
mod keywords;

pub use keywords::Keyword;

/// The longest source [`lex`] accepts, in bytes (16 MiB, Req 11.6).
pub const MAX_SOURCE: usize = 16 * 1024 * 1024;

/// Where a token sits in the source: the 1-based line and column (in characters) of its first
/// character, and its byte range `start..end`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Span {
    /// 1-based physical line. CRLF counts as one line break.
    pub line: u32,
    /// 1-based column, counted in characters (not bytes).
    pub col: u32,
    /// Byte offset of the first byte.
    pub start: usize,
    /// Byte offset just past the last byte.
    pub end: usize,
}

/// Why a span of source couldn't be read as a token.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, thiserror::Error)]
pub enum LexError {
    /// A `"` with no closing `"` before the end of its line.
    #[error("the string has no closing quote on its line")]
    UnterminatedString,
    /// A `#` followed by a date with no closing `#` on its line.
    #[error("the date literal has no closing # on its line")]
    UnterminatedDate,
    /// A `[` with no closing `]` on its line.
    #[error("the bracketed name has no closing ] on its line")]
    UnterminatedForeignName,
    /// `[]`: a bracketed name needs at least one character.
    #[error("the bracketed name is empty")]
    EmptyForeignName,
    /// Characters that can't start any VBA token.
    #[error("the character can't appear here in VBA source")]
    InvalidCharacter,
}

/// What a token is. Trivia ([`TokenKind::is_trivia`]) are kept so the tokens reproduce the input.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TokenKind {
    /// A reserved identifier or contextual word, matched case-insensitively.
    Keyword(Keyword),
    /// A Latin (or non-ASCII letter) identifier that isn't in the keyword table.
    Identifier,
    /// `[any text]`, used as an identifier.
    ForeignIdentifier,
    /// A `% & ^ ! # @ $` directly after a word.
    TypeSuffix,
    /// A decimal integer, with an optional `% & ^` suffix.
    Integer,
    /// `&H` followed by hex digits, with an optional `% & ^` suffix.
    Hex,
    /// `&O` or `&` followed by octal digits, with an optional `% & ^` suffix.
    Octal,
    /// A floating-point literal, or a decimal integer with a `! # @` suffix.
    Float,
    /// A closed `"…"` string.
    String,
    /// A closed `#…#` date or time.
    Date,
    /// One of `& * + - / \ ^ < = >`.
    Operator,
    /// One of `, . ! # ( ) : ; ?`.
    Punct,
    /// Whitespace, `_`, optional whitespace and a line terminator.
    LineContinuation,
    /// A `'` or `Rem` comment up to the end of its logical line (terminator excluded).
    Comment,
    /// One line terminator: CRLF, CR, LF, U+2028 or U+2029.
    Newline,
    /// A run of WSC characters.
    Whitespace,
    /// Source that couldn't be read as a token.
    Error(LexError),
}

impl TokenKind {
    /// `true` for whitespace, line terminators, line continuations and comments.
    #[must_use]
    pub const fn is_trivia(self) -> bool {
        matches!(self, TokenKind::Whitespace | TokenKind::Newline | TokenKind::LineContinuation | TokenKind::Comment)
    }
}

/// One token and where it is.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Token {
    /// What the token is.
    pub kind: TokenKind,
    /// Where it is in the source.
    pub span: Span,
}

/// The source is longer than [`MAX_SOURCE`]; no tokens were produced.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[error("source too large: {len} bytes, more than the {MAX_SOURCE}-byte limit")]
pub struct SourceTooLarge {
    /// The source length in bytes.
    pub len: usize,
}

/// Splits `source` into tokens that cover it exactly (see the module docs).
///
/// # Errors
///
/// Returns [`SourceTooLarge`] (and no tokens) when `source` is longer than [`MAX_SOURCE`] bytes.
pub fn lex(source: &str) -> Result<Vec<Token>, SourceTooLarge> {
    if source.len() > MAX_SOURCE {
        return Err(SourceTooLarge { len: source.len() });
    }
    let mut lexer = Lexer::new(source);
    while lexer.pos < source.len() {
        lexer.step();
    }
    Ok(lexer.tokens)
}

/// The source text of `token`, or `""` if its span doesn't fit `source`.
#[must_use]
pub fn text<'a>(source: &'a str, token: &Token) -> &'a str {
    source.get(token.span.start..token.span.end).unwrap_or("")
}

/// The last non-trivia token on the current logical line, as far as the lexer's context rules
/// need it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Prev {
    /// Nothing yet on this logical line.
    LineStart,
    /// A keyword that isn't a member name (not right after `.` or `!`).
    Keyword(Keyword),
    /// A punctuation character.
    Punct(u8),
    /// Anything else.
    Other,
}

struct Lexer<'a> {
    src: &'a str,
    tokens: Vec<Token>,
    /// Byte offset of the next token.
    pos: usize,
    /// Line and column of `pos`.
    line: u32,
    col: u32,
    /// The last character consumed was a CR, so a following LF doesn't start another line.
    after_cr: bool,
    prev: Prev,
}

impl<'a> Lexer<'a> {
    fn new(src: &'a str) -> Self {
        Self { src, tokens: Vec::new(), pos: 0, line: 1, col: 1, after_cr: false, prev: Prev::LineStart }
    }

    /// Reads one token (two for a word with a type suffix) at `pos`.
    fn step(&mut self) {
        let start = self.pos;
        let Some(c) = char_at(self.src, start) else {
            self.push(TokenKind::Error(LexError::InvalidCharacter), self.src.len());
            return;
        };
        if let Some(len) = terminator_len(self.src, start) {
            self.push(TokenKind::Newline, start.saturating_add(len));
        } else if is_wsc(c) {
            self.whitespace(start);
        } else {
            self.token(start, c);
        }
    }

    /// Reads a non-trivia token (or a comment) starting with `c`.
    fn token(&mut self, start: usize, c: char) {
        match c {
            '\'' => self.push(TokenKind::Comment, scan_to_line_end(self.src, start, true)),
            '"' => self.string(start),
            '[' => self.foreign_name(start),
            '#' => self.hash(start),
            '&' => self.ampersand(start),
            '0'..='9' => self.number(start),
            '.' if byte_at(self.src, start.saturating_add(1)).is_some_and(|byte| byte.is_ascii_digit()) => {
                self.number(start);
            }
            _ if is_ident_start(c) => self.word(start),
            _ => match special_kind(c) {
                Some(kind) => self.push(kind, start.saturating_add(1)),
                None => self.push(TokenKind::Error(LexError::InvalidCharacter), invalid_run_end(self.src, start)),
            },
        }
    }

    /// Whitespace, or a line continuation if the run is followed by `_` and the line's end.
    fn whitespace(&mut self, start: usize) {
        let run_end = skip_wsc(self.src, start);
        match continuation_after_run(self.src, run_end) {
            Some(end) => self.push(TokenKind::LineContinuation, end),
            None => self.push(TokenKind::Whitespace, run_end),
        }
    }

    /// A `"…"` string, or an unterminated-string error up to the end of the line.
    fn string(&mut self, start: usize) {
        let mut pos = start.saturating_add(1);
        while let Some(c) = char_at(self.src, pos) {
            if terminator_len(self.src, pos).is_some() {
                break;
            }
            let next = pos.saturating_add(c.len_utf8());
            if c == '"' {
                if byte_at(self.src, next) != Some(b'"') {
                    self.push(TokenKind::String, next);
                    return;
                }
                pos = next.saturating_add(1);
            } else {
                pos = next;
            }
        }
        self.push(TokenKind::Error(LexError::UnterminatedString), scan_to_line_end(self.src, start, false));
    }

    /// A `[name]`, or an error for `[]` or a missing `]`.
    fn foreign_name(&mut self, start: usize) {
        let body = start.saturating_add(1);
        let mut pos = body;
        while let Some(c) = char_at(self.src, pos) {
            if terminator_len(self.src, pos).is_some() {
                break;
            }
            let next = pos.saturating_add(c.len_utf8());
            if c == ']' {
                let kind = if pos == body { TokenKind::Error(LexError::EmptyForeignName) } else { TokenKind::ForeignIdentifier };
                self.push(kind, next);
                return;
            }
            pos = next;
        }
        self.push(TokenKind::Error(LexError::UnterminatedForeignName), scan_to_line_end(self.src, start, false));
    }

    /// A `#`: a date literal, an unterminated date, or a plain `#` (see the module docs).
    fn hash(&mut self, start: usize) {
        let after = start.saturating_add(1);
        if matches!(self.prev, Prev::Keyword(keyword) if keyword.takes_file_number()) {
            self.push(TokenKind::Punct, after);
            return;
        }
        let body = skip_wsc(self.src, after);
        let readings = date::date_or_time(self.src, body);
        let closed = readings.ends().iter().map(|&end| skip_wsc(self.src, end)).filter(|&close| byte_at(self.src, close) == Some(b'#')).max();
        if let Some(close) = closed {
            self.push(TokenKind::Date, close.saturating_add(1));
            return;
        }
        let line_end = scan_to_line_end(self.src, start, false);
        let unterminated = readings.ends().iter().any(|&end| {
            self.src.get(body..end).is_some_and(date::looks_like_date) && self.src.get(end..line_end).is_some_and(|rest| !rest.contains('#'))
        });
        if unterminated {
            self.push(TokenKind::Error(LexError::UnterminatedDate), line_end);
        } else {
            self.push(TokenKind::Punct, after);
        }
    }

    /// `&H…` hex, `&O…` or `&…` octal, or the `&` operator.
    fn ampersand(&mut self, start: usize) {
        let after = start.saturating_add(1);
        let (kind, digits_start, is_digit): (_, _, fn(u8) -> bool) = match byte_at(self.src, after) {
            Some(b'h' | b'H') => (TokenKind::Hex, after.saturating_add(1), |byte| byte.is_ascii_hexdigit()),
            Some(b'o' | b'O') => (TokenKind::Octal, after.saturating_add(1), is_octal_digit),
            _ => (TokenKind::Octal, after, is_octal_digit),
        };
        let digits = digits_end(self.src, digits_start, is_digit);
        if digits > digits_start {
            self.push(kind, suffix_end(self.src, digits, INTEGER_SUFFIXES).unwrap_or(digits));
        } else {
            self.push(TokenKind::Operator, after);
        }
    }

    /// A decimal integer or float, starting with a digit or with `.` and a digit.
    fn number(&mut self, start: usize) {
        let int_end = digits_end(self.src, start, |byte| byte.is_ascii_digit());
        let float_end = if byte_at(self.src, int_end) == Some(b'.') {
            let fraction_end = digits_end(self.src, int_end.saturating_add(1), |byte| byte.is_ascii_digit());
            Some(exponent_end(self.src, fraction_end).unwrap_or(fraction_end))
        } else {
            exponent_end(self.src, int_end)
        };
        let (kind, end) = match float_end {
            Some(end) => (TokenKind::Float, suffix_end(self.src, end, FLOAT_SUFFIXES).unwrap_or(end)),
            None => match suffix_end(self.src, int_end, INTEGER_SUFFIXES) {
                Some(end) => (TokenKind::Integer, end),
                None => match suffix_end(self.src, int_end, FLOAT_SUFFIXES) {
                    Some(end) => (TokenKind::Float, end),
                    None => (TokenKind::Integer, int_end),
                },
            },
        };
        self.push(kind, end);
    }

    /// A keyword or identifier with an optional type suffix, or a `Rem` comment.
    fn word(&mut self, start: usize) {
        let end = ident_end(self.src, start);
        let keyword = self.src.get(start..end).and_then(Keyword::lookup);
        if keyword == Some(Keyword::Rem) && !matches!(self.prev, Prev::Punct(b'.' | b'!')) {
            self.push(TokenKind::Comment, scan_to_line_end(self.src, start, true));
            return;
        }
        self.push(keyword.map_or(TokenKind::Identifier, TokenKind::Keyword), end);
        if let Some(suffix_end) = suffix_end(self.src, end, IDENT_SUFFIXES) {
            self.push(TokenKind::TypeSuffix, suffix_end);
        }
    }

    /// Appends a token from `pos` to `end` and moves past it. An `end` that wouldn't advance or
    /// isn't a character boundary is replaced by the end of the next character, so the loop in
    /// [`lex`] always terminates and never splits a character.
    fn push(&mut self, kind: TokenKind, end: usize) {
        let start = self.pos;
        let end = if end > start && self.src.is_char_boundary(end) { end } else { char_end(self.src, start) };
        self.tokens.push(Token { kind, span: Span { line: self.line, col: self.col, start, end } });
        self.advance_position(start, end);
        self.pos = end;
        self.prev = match kind {
            TokenKind::Newline => Prev::LineStart,
            kind if kind.is_trivia() => self.prev,
            TokenKind::Keyword(_) | TokenKind::Identifier if matches!(self.prev, Prev::Punct(b'.' | b'!')) => Prev::Other,
            TokenKind::Keyword(keyword) => Prev::Keyword(keyword),
            TokenKind::Punct => byte_at(self.src, start).map_or(Prev::Other, Prev::Punct),
            _ => Prev::Other,
        };
    }

    /// Moves the line and column over `src[start..end]`.
    fn advance_position(&mut self, start: usize, end: usize) {
        for c in self.src.get(start..end).unwrap_or("").chars() {
            if c == '\n' && self.after_cr {
                self.after_cr = false;
            } else if is_line_terminator(c) {
                self.line = self.line.saturating_add(1);
                self.col = 1;
                self.after_cr = c == '\r';
            } else {
                self.col = self.col.saturating_add(1);
                self.after_cr = false;
            }
        }
    }
}

/// Suffixes an INTEGER may carry (§3.3.2).
const INTEGER_SUFFIXES: &[u8] = b"%&^";
/// Suffixes a FLOAT may carry (§3.3.2).
const FLOAT_SUFFIXES: &[u8] = b"!#@";
/// `type-suffix` (§3.3.5.3).
const IDENT_SUFFIXES: &[u8] = b"%&^!#@$";

/// The character at byte `pos`, or `None` at the end or off a character boundary.
fn char_at(src: &str, pos: usize) -> Option<char> {
    src.get(pos..)?.chars().next()
}

/// The byte at `pos`.
fn byte_at(src: &str, pos: usize) -> Option<u8> {
    src.as_bytes().get(pos).copied()
}

/// The end of the character at `pos` (the end of `src` if there is none).
fn char_end(src: &str, pos: usize) -> usize {
    char_at(src, pos).map_or(src.len(), |c| pos.saturating_add(c.len_utf8()))
}

/// `line-terminator` characters (§3.2.1).
const fn is_line_terminator(c: char) -> bool {
    matches!(c, '\r' | '\n' | '\u{2028}' | '\u{2029}')
}

/// The byte length of the line terminator at `pos` (CRLF is one terminator of two bytes).
fn terminator_len(src: &str, pos: usize) -> Option<usize> {
    let c = char_at(src, pos)?;
    if c == '\r' && byte_at(src, pos.saturating_add(1)) == Some(b'\n') { Some(2) } else { is_line_terminator(c).then_some(c.len_utf8()) }
}

/// WSC (§3.2.2): tab, U+0019, space, U+3000 and Unicode `Zs` except U+00A0 (a CP2 character).
const fn is_wsc(c: char) -> bool {
    matches!(c, '\t' | '\u{19}' | ' ' | '\u{3000}' | '\u{1680}' | '\u{2000}'..='\u{200A}' | '\u{202F}' | '\u{205F}')
}

/// The end of the WSC run starting at `pos` (`pos` itself if there is none).
fn skip_wsc(src: &str, pos: usize) -> usize {
    let mut end = pos;
    while let Some(c) = char_at(src, end).filter(|&c| is_wsc(c)) {
        end = end.saturating_add(c.len_utf8());
    }
    end
}

/// If a WSC run ending at `run_end` is followed by `_`, optional WSC and a line terminator (or
/// the end of input), the end of that line continuation.
fn continuation_after_run(src: &str, run_end: usize) -> Option<usize> {
    if byte_at(src, run_end) != Some(b'_') {
        return None;
    }
    let after = skip_wsc(src, run_end.saturating_add(1));
    if after >= src.len() {
        return Some(after);
    }
    terminator_len(src, after).map(|len| after.saturating_add(len))
}

/// The end of the line starting before `pos`, excluding its terminator. With `follow`, line
/// continuations are crossed (a comment runs to the end of its logical line); without it, the
/// scan stops before a trailing continuation so that stays its own token.
fn scan_to_line_end(src: &str, pos: usize, follow: bool) -> usize {
    let mut at = pos;
    while let Some(c) = char_at(src, at) {
        if terminator_len(src, at).is_some() {
            break;
        }
        if !is_wsc(c) {
            at = at.saturating_add(c.len_utf8());
            continue;
        }
        let run_end = skip_wsc(src, at);
        match continuation_after_run(src, run_end) {
            Some(_) if !follow => break,
            Some(end) => at = end,
            None => at = run_end,
        }
    }
    at
}

/// The end of a run of bytes matching `is_digit` starting at `pos`.
fn digits_end(src: &str, pos: usize, is_digit: impl Fn(u8) -> bool) -> usize {
    let len = src.as_bytes().get(pos..).map_or(0, |rest| rest.iter().take_while(|&&byte| is_digit(byte)).count());
    pos.saturating_add(len)
}

const fn is_octal_digit(byte: u8) -> bool {
    matches!(byte, b'0'..=b'7')
}

/// `exponent = exponent-letter [sign] decimal-literal` at `pos`, if there is one.
fn exponent_end(src: &str, pos: usize) -> Option<usize> {
    if !matches!(byte_at(src, pos), Some(b'd' | b'D' | b'e' | b'E')) {
        return None;
    }
    let mut digits = pos.saturating_add(1);
    if matches!(byte_at(src, digits), Some(b'+' | b'-')) {
        digits = digits.saturating_add(1);
    }
    let end = digits_end(src, digits, |byte| byte.is_ascii_digit());
    (end > digits).then_some(end)
}

/// If one of `suffixes` sits at `pos` and the next character can't continue a term (so it isn't
/// `rs!Name`, `rs![Name]`, `a&b` or `a^2`), the end of the suffix.
fn suffix_end(src: &str, pos: usize, suffixes: &[u8]) -> Option<usize> {
    let byte = byte_at(src, pos)?;
    if !suffixes.contains(&byte) {
        return None;
    }
    let end = pos.saturating_add(1);
    let continues_term = char_at(src, end).is_some_and(|c| is_ident_continue(c) || c == '[');
    (!continues_term).then_some(end)
}

/// `first-Latin-identifier-character`, or a non-ASCII letter (see the module docs).
fn is_ident_start(c: char) -> bool {
    c.is_ascii_alphabetic() || (!c.is_ascii() && c.is_alphabetic())
}

/// `subsequent-Latin-identifier-character`, or a non-ASCII letter or digit.
fn is_ident_continue(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || (!c.is_ascii() && c.is_alphanumeric())
}

/// The end of the word starting at `pos`.
fn ident_end(src: &str, pos: usize) -> usize {
    let mut end = pos;
    while let Some(c) = char_at(src, end).filter(|&c| is_ident_continue(c)) {
        end = end.saturating_add(c.len_utf8());
    }
    end
}

/// The token kind of a single-character `special-token` (§3.3.1).
const fn special_kind(c: char) -> Option<TokenKind> {
    match c {
        '&' | '*' | '+' | '-' | '/' | '\\' | '^' | '<' | '=' | '>' => Some(TokenKind::Operator),
        ',' | '.' | '!' | '#' | '(' | ')' | ':' | ';' | '?' => Some(TokenKind::Punct),
        _ => None,
    }
}

/// `true` for a character that starts no token: [`Lexer::token`]'s fallback case.
fn is_invalid(c: char) -> bool {
    !(is_line_terminator(c) || is_wsc(c) || matches!(c, '\'' | '"' | '[' | '0'..='9') || is_ident_start(c) || special_kind(c).is_some())
}

/// The end of the run of invalid characters starting at `pos` (at least one character).
fn invalid_run_end(src: &str, pos: usize) -> usize {
    let mut end = char_end(src, pos);
    while let Some(c) = char_at(src, end).filter(|&c| is_invalid(c)) {
        end = end.saturating_add(c.len_utf8());
    }
    end
}

#[cfg(test)]
mod props;
#[cfg(test)]
mod tests;
