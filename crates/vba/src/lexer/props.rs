//! Property test P5 for [`lex`]: the tokens reproduce the input byte for byte, including input
//! that produces error tokens, driven by a seeded xorshift64 generator so every failure replays.
//!
//! For every input of at most [`MAX_SOURCE`] bytes, `lex` returns `Ok` without panicking, and:
//! - joining `text(src, t)` over the tokens gives the input byte for byte;
//! - the tokens are contiguous (the first starts at 0, each starts where the previous one ends,
//!   the last ends at `src.len()`), non-empty, and start and end on character boundaries;
//! - each span's line and column match a recount from the source alone: CR, LF, U+2028 and
//!   U+2029 each end a line, CRLF ends one line, and columns count characters from 1.
//!
//! Inputs come from four classes, picked per case:
//! - realistic VBA from a small grammar: keywords in random case, identifiers (non-ASCII ones
//!   too) with type suffixes, decimal, `&H`, `&O` and float literals with exponents and
//!   suffixes, strings with `""`, `#…#` dates, `#1` file numbers, `[bracketed]` names, bang
//!   access, operators, `'` and `Rem` comments, ` _` continuations before CRLF, LF, CR, U+2028
//!   and U+2029, `Attribute` and `#If` lines;
//! - realistic VBA with characters inserted, deleted and replaced, closing `"`, `#` and `]`
//!   dropped, and the text cut short;
//! - random strings of ASCII, Latin-1, CJK, emoji, U+2028/U+2029, U+00A0, control characters,
//!   arbitrary scalar values and the characters the lexer treats specially (`_ " # [ ] ' &`);
//! - large inputs (1 to 4 MiB) built by repeating realistic and mutated pieces, on one case in
//!   [`BIG_EVERY`]. [`p5_holds_at_the_size_limit`] checks an input of exactly [`MAX_SOURCE`]
//!   bytes and that one more byte is refused.
//!
//! Knobs (environment variables):
//! - `GRIDCRAFT_FUZZ_SEEDS` (default 64): runs 16 cases per seed (1,024 by default). Soak with
//!   `GRIDCRAFT_FUZZ_SEEDS=500 cargo test --release -p gridcraft-vba --lib lexer::props::`.
//! - `GRIDCRAFT_FUZZ_SEED=17`: replay that one case (failure messages print the command line).
use super::{Keyword, LexError, MAX_SOURCE, SourceTooLarge, Token, TokenKind, lex, text};
use std::panic::{AssertUnwindSafe, catch_unwind};

/// Default `GRIDCRAFT_FUZZ_SEEDS`: small, so `cargo xtask ci` stays fast.
const DEFAULT_SEEDS: u64 = 64;
/// Cases per seed.
const CASES_PER_SEED: u64 = 16;
/// One case in this many is a large input (two in the default 1,024 cases).
const BIG_EVERY: u64 = 512;
/// Large inputs are this many bytes, inclusive.
const BIG_LEN: (usize, usize) = (1 << 20, 4 << 20);
/// Keeps P5's random sequences apart from other property tests'.
const P5_STREAM: u64 = 0x5035_5035_5035_5035;

// ---------------------------------------------------------------- deterministic randomness

/// xorshift64: tiny and deterministic.
struct Rng(u64);

impl Rng {
    fn new(seed: u64, stream: u64) -> Self {
        // Spread consecutive seeds over the state space; xorshift needs a nonzero state.
        let mixed = seed.wrapping_add(1).wrapping_mul(0x9E37_79B9_7F4A_7C15) ^ stream;
        let mut rng = Self(if mixed == 0 { 0x2545_F491_4F6C_DD1D } else { mixed });
        for _ in 0..4 {
            rng.next();
        }
        rng
    }

    fn next(&mut self) -> u64 {
        let mut x = self.0;
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        self.0 = x;
        x
    }

    fn below(&mut self, n: usize) -> usize {
        if n == 0 { 0 } else { (self.next() % n as u64) as usize }
    }

    /// Inclusive range.
    fn range(&mut self, lo: usize, hi: usize) -> usize {
        lo + self.below(hi.saturating_sub(lo) + 1)
    }

    fn chance(&mut self, percent: u64) -> bool {
        self.next() % 100 < percent
    }

    fn pick<T: Copy>(&mut self, items: &[T]) -> T {
        items[self.below(items.len())]
    }

    /// A character in `lo..=hi`, or `fallback` if that lands on a surrogate.
    fn char_in(&mut self, lo: u32, hi: u32, fallback: char) -> char {
        char::from_u32(lo + self.below((hi - lo + 1) as usize) as u32).unwrap_or(fallback)
    }
}

fn env_num(name: &str) -> Option<u64> {
    std::env::var(name).ok().and_then(|v| v.trim().parse().ok())
}

/// The cases to run: `GRIDCRAFT_FUZZ_SEED` alone, or `0..count * per_seed`.
fn seed_list(per_seed: u64) -> (Vec<u64>, bool) {
    match env_num("GRIDCRAFT_FUZZ_SEED") {
        Some(one) => (vec![one], true),
        None => {
            let seeds = env_num("GRIDCRAFT_FUZZ_SEEDS").unwrap_or(DEFAULT_SEEDS);
            ((0..seeds.saturating_mul(per_seed)).collect(), false)
        }
    }
}

fn replay(seed: u64, test: &str) -> String {
    format!("replay: GRIDCRAFT_FUZZ_SEED={seed} cargo test -p gridcraft-vba --lib lexer::props::{test}")
}

// ---------------------------------------------------------------- realistic VBA

/// Line terminators, CRLF most often.
const NEWLINES: [&str; 7] = ["\r\n", "\r\n", "\r\n", "\n", "\r", "\u{2028}", "\u{2029}"];
/// Separators between atoms: WSC characters (tab, space, U+3000, U+0019) or nothing.
const SPACES: [&str; 7] = [" ", " ", " ", "", "\t", "\u{3000}", "  \u{19}"];
const NAMES: [&str; 14] = ["x", "i", "total", "Größe", "数値", "café", "Ñandú", "rs", "Sheet1", "my_var", "a_", "Ωmega", "名前", "Rem2"];
const TYPE_SUFFIXES: [char; 7] = ['%', '&', '^', '!', '#', '@', '$'];
const TYPES: [&str; 7] = ["Long", "String", "Variant", "Double", "Date", "Object", "Integer"];
const OPERATORS: [&str; 18] = ["+", "-", "*", "/", "\\", "^", "&", "=", "<>", "<=", ">=", "<", ">", "Mod", "And", "Or", "Like", "Is"];
const STRING_PARTS: [&str; 12] = ["hello", "\"\"", "日本語", "😀", "a\tb", "", " ", "x' not a comment", "Rem", "#1/2/2000#", "[a]", " _"];
const DATES: [&str; 10] = [
    "#1/2/2000#",
    "#12/31/1999 11:59:59 PM#",
    "# 1-2-2000 #",
    "#13:45#",
    "#1:30 am#",
    "#Jan 1, 2000#",
    "#february 29 2000#",
    "#2000/1/2#",
    "#10 PM#",
    "#1/2#",
];
const FILE_STATEMENTS: [&str; 10] = ["Print", "Write", "Input", "Get", "Put", "Close", "Seek", "Lock", "Unlock", "Width"];
const MEMBERS: [&str; 5] = [".Value", "!Name", "![Field Name]", ".Cells(1, 2)", "!Rem"];
const ATTRIBUTES: [&str; 4] =
    ["Attribute VB_Name = \"Module1\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Exposed = True", "Attribute Value.VB_UserMemId = 0"];
const DIRECTIVES: [&str; 5] = ["#If Win64 Then", "#Else", "#End If", "#Const DEBUG_MODE = 1", "#ElseIf VBA7 Then"];
const COMMENT_WORDS: [&str; 10] = ["note", "Größe", "日本語", "😀", "\"open", "#1/2", "[x", "rem", "'", "\u{a0}"];

/// Builds realistic VBA source. Holds the keyword spellings so they're collected once.
struct Gen {
    keywords: Vec<&'static str>,
}

impl Gen {
    fn new() -> Self {
        Self { keywords: Keyword::all().map(Keyword::as_str).collect() }
    }

    /// Several lines of realistic source, about `target` bytes long.
    fn source(&self, rng: &mut Rng, target: usize) -> String {
        let mut out = String::with_capacity(target + 256);
        while out.len() < target {
            self.line(rng, &mut out);
        }
        // Sometimes the last line has no terminator.
        if rng.chance(30) {
            while out.ends_with(['\r', '\n', '\u{2028}', '\u{2029}']) {
                out.pop();
            }
        }
        out
    }

    fn line(&self, rng: &mut Rng, out: &mut String) {
        for _ in 0..rng.below(3) {
            out.push_str(rng.pick(&["    ", "\t"]));
        }
        self.statement(rng, out);
        if rng.chance(15) {
            out.push_str(": ");
            self.statement(rng, out);
        }
        if rng.chance(20) {
            out.push_str(rng.pick(&[" ", "\t", ""]));
            comment(rng, out, "'");
        }
        out.push_str(rng.pick(&NEWLINES));
    }

    fn statement(&self, rng: &mut Rng, out: &mut String) {
        match rng.below(10) {
            0 => {
                random_case(rng, "Dim ", out);
                identifier(rng, out);
                random_case(rng, " As ", out);
                random_case_pick(rng, &TYPES, out);
            }
            1 => {
                identifier(rng, out);
                if rng.chance(40) {
                    out.push_str(rng.pick(&MEMBERS));
                }
                out.push_str(" = ");
                self.expression(rng, out);
            }
            2 => {
                identifier(rng, out);
                out.push('(');
                self.expression(rng, out);
                out.push_str(", ");
                self.expression(rng, out);
                out.push(')');
            }
            3 => self.file_statement(rng, out),
            4 => {
                random_case(rng, "If ", out);
                self.expression(rng, out);
                random_case(rng, " Then ", out);
                self.statement(rng, out);
            }
            5 => random_case_pick(rng, &ATTRIBUTES, out),
            6 => out.push_str(rng.pick(&DIRECTIVES)),
            7 => {
                if rng.chance(30) {
                    out.push_str(&rng.below(1000).to_string());
                    out.push(' ');
                }
                let opener = rng.pick(&["Rem ", "rem\t", "REM", "Rem:"]);
                comment(rng, out, opener);
            }
            8 => {
                for _ in 0..rng.range(1, 8) {
                    self.atom(rng, out);
                    out.push_str(rng.pick(&SPACES));
                }
            }
            _ => random_case_pick(rng, &["Sub Main()", "End Sub", "Function F(ByVal a As Long) As Long", "End Function", "Option Explicit"], out),
        }
    }

    fn file_statement(&self, rng: &mut Rng, out: &mut String) {
        if rng.chance(25) {
            random_case(rng, "Open ", out);
            string_literal(rng, out);
            random_case(rng, " For Input As #", out);
            out.push_str(&rng.range(1, 511).to_string());
            return;
        }
        random_case_pick(rng, &FILE_STATEMENTS, out);
        out.push_str(rng.pick(&[" #", " # ", "#"]));
        out.push_str(&rng.range(1, 511).to_string());
        if rng.chance(60) {
            out.push_str(", ");
            self.expression(rng, out);
        }
    }

    /// Atoms joined by operators, sometimes in parentheses or split by a line continuation.
    fn expression(&self, rng: &mut Rng, out: &mut String) {
        let parens = rng.chance(20);
        if parens {
            out.push('(');
        }
        self.atom(rng, out);
        for _ in 0..rng.below(4) {
            if rng.chance(10) {
                continuation(rng, out);
            }
            let spaced = rng.chance(70);
            let op = rng.pick(&OPERATORS);
            // Word operators need space around them.
            let space = if spaced || op.as_bytes().first().is_some_and(u8::is_ascii_alphabetic) { " " } else { "" };
            out.push_str(space);
            random_case(rng, op, out);
            out.push_str(space);
            self.atom(rng, out);
        }
        if parens {
            out.push(')');
        }
    }

    fn atom(&self, rng: &mut Rng, out: &mut String) {
        match rng.below(9) {
            0 | 1 => identifier(rng, out),
            2 => random_case_pick(rng, self.keywords.as_slice(), out),
            3 | 4 => number(rng, out),
            5 => string_literal(rng, out),
            6 => out.push_str(rng.pick(&DATES)),
            7 => {
                out.push('[');
                out.push_str(rng.pick(&["Field Name", "a", "日本", "x]y", "#1", "\"q", ""]));
                out.push(']');
            }
            _ => {
                identifier(rng, out);
                out.push_str(rng.pick(&MEMBERS));
            }
        }
    }
}

/// One of `words`, each letter's case picked at random.
fn random_case_pick(rng: &mut Rng, words: &[&str], out: &mut String) {
    let word = rng.pick(words);
    random_case(rng, word, out);
}

/// `word` with each letter's case picked at random.
fn random_case(rng: &mut Rng, word: &str, out: &mut String) {
    for c in word.chars() {
        if rng.chance(50) {
            out.extend(c.to_uppercase());
        } else {
            out.extend(c.to_lowercase());
        }
    }
}

fn identifier(rng: &mut Rng, out: &mut String) {
    if rng.chance(50) {
        out.push_str(rng.pick(&NAMES));
    } else {
        out.push(rng.char_in(u32::from('a'), u32::from('z'), 'a'));
        for _ in 0..rng.below(8) {
            out.push(rng.pick(&['a', 'Z', 'q', '0', '7', '_', 'é']));
        }
    }
    if rng.chance(20) {
        out.push(rng.pick(&TYPE_SUFFIXES));
    }
}

fn digits(rng: &mut Rng, out: &mut String, alphabet: &[char], max: usize) {
    for _ in 0..rng.range(1, max) {
        out.push(rng.pick(alphabet));
    }
}

const DECIMAL: [char; 10] = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'];
const OCTAL: [char; 8] = ['0', '1', '2', '3', '4', '5', '6', '7'];
const HEX: [char; 16] = ['0', '1', '2', '9', 'a', 'b', 'c', 'd', 'e', 'f', 'A', 'B', 'C', 'D', 'E', 'F'];

fn number(rng: &mut Rng, out: &mut String) {
    match rng.below(5) {
        0 => digits(rng, out, &DECIMAL, 6),
        1 => {
            out.push_str(rng.pick(&["&H", "&h"]));
            digits(rng, out, &HEX, 8);
        }
        2 => {
            out.push_str(rng.pick(&["&O", "&o", "&"]));
            digits(rng, out, &OCTAL, 8);
        }
        3 => {
            digits(rng, out, &DECIMAL, 4);
            out.push('.');
            digits(rng, out, &DECIMAL, 4);
        }
        _ => out.push_str(rng.pick(&[".5", "1.", "0.25", "123"])),
    }
    if rng.chance(30) {
        out.push(rng.pick(&['e', 'E', 'd', 'D']));
        out.push_str(rng.pick(&["", "+", "-"]));
        digits(rng, out, &DECIMAL, 3);
    }
    if rng.chance(25) {
        out.push(rng.pick(&['%', '&', '^', '!', '#', '@']));
    }
}

fn string_literal(rng: &mut Rng, out: &mut String) {
    out.push('"');
    for _ in 0..rng.below(4) {
        out.push_str(rng.pick(&STRING_PARTS));
    }
    out.push('"');
}

/// ` _` (sometimes with WSC after the `_`), a line terminator and an indent.
fn continuation(rng: &mut Rng, out: &mut String) {
    out.push_str(rng.pick(&[" _", "\t_", " _ ", "  _\t"]));
    out.push_str(rng.pick(&NEWLINES));
    out.push_str(rng.pick(&["", "    ", "\t"]));
}

/// A comment starting with `opener`, sometimes carried onto the next line by a continuation.
fn comment(rng: &mut Rng, out: &mut String, opener: &str) {
    out.push_str(opener);
    for _ in 0..rng.below(5) {
        out.push(' ');
        out.push_str(rng.pick(&COMMENT_WORDS));
    }
    if rng.chance(15) {
        continuation(rng, out);
        out.push_str(rng.pick(&COMMENT_WORDS));
    }
}

// ---------------------------------------------------------------- mutations and noise

/// One random character from a mix of ranges the lexer must cope with.
fn noise_char(rng: &mut Rng) -> char {
    const SPECIAL: [char; 22] = [
        '\u{2028}', '\u{2029}', '\u{a0}', '\u{3000}', '\u{19}', '\t', '\r', '\n', '_', '"', '#', '[', ']', '&', '\'', ' ', '\0', '\u{7f}',
        '\u{feff}', '\u{200b}', '!', '.',
    ];
    match rng.below(8) {
        0 => rng.char_in(0x20, 0x7E, ' '),
        1 => rng.char_in(0xA0, 0xFF, 'é'),
        2 => rng.char_in(0x4E00, 0x9FFF, '数'),
        3 => rng.char_in(0x1F600, 0x1F64F, '😀'),
        4 => rng.char_in(0, 0x1F, '\0'),
        5 => rng.char_in(0, 0x10_FFFF, '\u{fffd}'),
        _ => rng.pick(&SPECIAL),
    }
}

fn unicode_input(rng: &mut Rng) -> String {
    (0..rng.range(0, 200)).map(|_| noise_char(rng)).collect()
}

/// Realistic source with a few random edits.
fn mutated_input(rng: &mut Rng, generator: &Gen) -> String {
    let target = rng.range(0, 600);
    let base = generator.source(rng, target);
    let mut chars: Vec<char> = base.chars().collect();
    for _ in 0..rng.range(1, 8) {
        mutate(rng, &mut chars);
    }
    chars.into_iter().collect()
}

fn mutate(rng: &mut Rng, chars: &mut Vec<char>) {
    let at = rng.below(chars.len() + 1);
    match rng.below(6) {
        0 => chars.insert(at, noise_char(rng)),
        1 if at < chars.len() => {
            chars.remove(at);
        }
        2 if at < chars.len() => chars[at] = noise_char(rng),
        3 => {
            // Drop a closing delimiter, so strings, dates and bracketed names run unterminated.
            let closers: Vec<usize> = chars.iter().enumerate().filter(|&(_, c)| matches!(c, '"' | '#' | ']')).map(|(i, _)| i).collect();
            if !closers.is_empty() {
                chars.remove(rng.pick(&closers));
            }
        }
        4 => chars.truncate(at),
        _ => {
            // Duplicate a short slice somewhere else.
            let len = rng.below(chars.len().saturating_sub(at).min(16) + 1);
            let piece: Vec<char> = chars[at..at + len].to_vec();
            let to = rng.below(chars.len() + 1);
            chars.splice(to..to, piece);
        }
    }
}

/// 1 to 4 MiB of repeated realistic and mutated pieces.
fn big_input(rng: &mut Rng, generator: &Gen) -> String {
    let pieces: Vec<String> = (0..8)
        .map(|i| {
            if i % 4 == 3 {
                mutated_input(rng, generator)
            } else {
                let target = rng.range(256, 4096);
                generator.source(rng, target)
            }
        })
        .collect();
    let target = rng.range(BIG_LEN.0, BIG_LEN.1);
    let mut out = String::with_capacity(target + 8192);
    while out.len() < target {
        out.push_str(&pieces[rng.below(pieces.len())]);
        // Keep pieces on separate lines most of the time.
        if rng.chance(80) {
            out.push_str("\r\n");
        }
    }
    out
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Class {
    Realistic,
    Mutated,
    Unicode,
    Big,
}

const CLASSES: [Class; 4] = [Class::Realistic, Class::Mutated, Class::Unicode, Class::Big];

fn class_of(case: u64) -> Class {
    if case % BIG_EVERY == 3 {
        return Class::Big;
    }
    match case % 3 {
        0 => Class::Realistic,
        1 => Class::Mutated,
        _ => Class::Unicode,
    }
}

fn input(generator: &Gen, case: u64, class: Class) -> String {
    let mut rng = Rng::new(case, P5_STREAM);
    match class {
        Class::Realistic => {
            let target = rng.range(0, 2000);
            generator.source(&mut rng, target)
        }
        Class::Mutated => mutated_input(&mut rng, generator),
        Class::Unicode => unicode_input(&mut rng),
        Class::Big => big_input(&mut rng, generator),
    }
}

// ---------------------------------------------------------------- the property

/// A few characters around byte `at`, for failure messages.
fn context(src: &str, at: usize) -> String {
    let start = (0..=at.saturating_sub(24).min(src.len())).rev().find(|&i| src.is_char_boundary(i)).unwrap_or(0);
    let end = (at.saturating_add(24).min(src.len())..=src.len()).find(|&i| src.is_char_boundary(i)).unwrap_or(src.len());
    format!("{:?}", &src[start..end])
}

/// The tokens are contiguous, non-empty, on character boundaries and join to `src`.
fn check_cover(src: &str, tokens: &[Token]) -> Result<(), String> {
    let mut pos = 0;
    let mut joined = String::with_capacity(src.len());
    for (i, token) in tokens.iter().enumerate() {
        let span = token.span;
        if span.start != pos {
            return Err(format!("token {i} {token:?} starts at {}, the previous token ended at {pos} (near {})", span.start, context(src, pos)));
        }
        if span.end <= span.start || span.end > src.len() {
            return Err(format!("token {i} {token:?} is empty or runs past the {}-byte input", src.len()));
        }
        if !src.is_char_boundary(span.start) || !src.is_char_boundary(span.end) {
            return Err(format!("token {i} {token:?} splits a character (near {})", context(src, span.start)));
        }
        joined.push_str(text(src, token));
        pos = span.end;
    }
    if pos != src.len() {
        return Err(format!("the tokens end at byte {pos}, the input is {} bytes (near {})", src.len(), context(src, pos)));
    }
    if joined != src {
        let at = joined.bytes().zip(src.bytes()).position(|(a, b)| a != b).unwrap_or(joined.len().min(src.len()));
        return Err(format!("the joined tokens differ from the input at byte {at} (near {})", context(src, at)));
    }
    Ok(())
}

/// Each token's line and column match a recount over the source text before it.
fn check_positions(src: &str, tokens: &[Token]) -> Result<(), String> {
    let mut chars = src.char_indices().peekable();
    let (mut line, mut col) = (1u64, 1u64);
    let mut previous = None;
    for (i, token) in tokens.iter().enumerate() {
        while let Some(&(at, c)) = chars.peek() {
            if at >= token.span.start {
                break;
            }
            if c == '\n' && previous == Some('\r') {
                // The LF of a CRLF: the CR already ended the line.
            } else if matches!(c, '\r' | '\n' | '\u{2028}' | '\u{2029}') {
                line += 1;
                col = 1;
            } else {
                col += 1;
            }
            previous = Some(c);
            chars.next();
        }
        if (u64::from(token.span.line), u64::from(token.span.col)) != (line, col) {
            return Err(format!(
                "token {i} {token:?} is at line {} col {}, the recount says line {line} col {col} (near {})",
                token.span.line,
                token.span.col,
                context(src, token.span.start)
            ));
        }
    }
    Ok(())
}

/// How many tokens of each [`LexError`] a case produced.
type ErrorCounts = [usize; 5];

fn error_slot(error: LexError) -> usize {
    match error {
        LexError::UnterminatedString => 0,
        LexError::UnterminatedDate => 1,
        LexError::UnterminatedForeignName => 2,
        LexError::EmptyForeignName => 3,
        LexError::InvalidCharacter => 4,
    }
}

/// Runs `lex` on `src` (catching panics) and checks P5. Returns the error-token counts.
fn check(src: &str) -> Result<ErrorCounts, String> {
    let result = catch_unwind(AssertUnwindSafe(|| lex(src))).map_err(|_| "lex panicked".to_owned())?;
    let tokens = result.map_err(|e| format!("lex refused an input within the limit: {e}"))?;
    check_cover(src, &tokens)?;
    check_positions(src, &tokens)?;
    let mut counts = [0; 5];
    for token in &tokens {
        if let TokenKind::Error(error) = token.kind {
            counts[error_slot(error)] += 1;
        }
    }
    Ok(counts)
}

/// P5: the tokens reproduce the input byte for byte, with contiguous spans on character
/// boundaries whose line and column match a recount, and `lex` never panics.
///
/// **Validates: Requirements 11.7**
#[test]
fn p5_tokens_reproduce_the_input() {
    let (cases, replaying) = seed_list(CASES_PER_SEED);
    let generator = Gen::new();
    let mut per_class = [0usize; 4];
    let mut errors: ErrorCounts = [0; 5];
    let mut cases_with_errors = 0;
    for &case in &cases {
        let class = class_of(case);
        let src = input(&generator, case, class);
        match check(&src) {
            Ok(counts) => {
                per_class[CLASSES.iter().position(|&c| c == class).unwrap_or(0)] += 1;
                cases_with_errors += usize::from(counts.iter().any(|&n| n > 0));
                for (total, n) in errors.iter_mut().zip(counts) {
                    *total += n;
                }
            }
            Err(why) => {
                let shown = if src.len() <= 600 { format!("\n  input: {src:?}") } else { String::new() };
                panic!(
                    "P5 failed for seed {case} ({class:?}, {} bytes): {why}{shown}\n  {}",
                    src.len(),
                    replay(case, "p5_tokens_reproduce_the_input")
                );
            }
        }
    }
    // Every generator class must run, and the invalid inputs must reach every error token.
    if !replaying && cases.len() >= 1024 {
        assert!(per_class.iter().all(|&n| n > 0), "P5 cases per class [realistic, mutated, unicode, big] = {per_class:?}");
        assert!(cases_with_errors > 0, "P5 produced no error tokens in {} cases", cases.len());
        assert!(
            errors.iter().all(|&n| n > 0),
            "P5 error tokens [string, date, foreign name, empty foreign name, invalid character] = {errors:?} in {} cases",
            cases.len()
        );
    }
}

/// P5 at the size limit: realistic source of exactly [`MAX_SOURCE`] bytes is lexed and covered,
/// and one more byte is refused with no tokens.
///
/// **Validates: Requirements 11.7**
#[test]
fn p5_holds_at_the_size_limit() {
    // Mostly long comment lines, so 16 MiB stays a few tens of thousands of tokens.
    let code = "x = &H1F& + 1.5E3 * Größe# & \"a\"\"b\" ' note _\r\n  more\r\nPrint #1, #1/2/2000#, [a b]\n";
    let comment_line = format!("' {}\r\n", "Größe 数値 😀 ".repeat(40));
    let mut src = String::with_capacity(MAX_SOURCE + 1);
    while src.len() + code.len() + comment_line.len() < MAX_SOURCE {
        src.push_str(code);
        for _ in 0..16 {
            if src.len() + comment_line.len() >= MAX_SOURCE {
                break;
            }
            src.push_str(&comment_line);
        }
    }
    // Pad with an unterminated string so the exact boundary also ends in an error token.
    src.push('"');
    while src.len() < MAX_SOURCE {
        src.push('x');
    }
    assert_eq!(src.len(), MAX_SOURCE);
    let counts = match check(&src) {
        Ok(counts) => counts,
        Err(why) => panic!("P5 failed at exactly MAX_SOURCE ({MAX_SOURCE} bytes): {why}"),
    };
    assert!(counts[error_slot(LexError::UnterminatedString)] > 0, "the padding should end in an unterminated string, got {counts:?}");
    src.push('x');
    match catch_unwind(AssertUnwindSafe(|| lex(&src))) {
        Ok(Err(SourceTooLarge { len })) => assert_eq!(len, MAX_SOURCE + 1),
        Ok(Ok(tokens)) => panic!("MAX_SOURCE + 1 bytes should be refused, got {} tokens", tokens.len()),
        Err(_) => panic!("lex panicked on MAX_SOURCE + 1 bytes"),
    }
}
