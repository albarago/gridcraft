//! Property tests P6 and P7 for [`parse`], driven by a seeded xorshift64 generator so every
//! failure replays.
//!
//! - **P6** (Req 12.8): for every generated module that parses with zero diagnostics,
//!   `parse(print_module(tree))` has zero diagnostics and is equivalent to `tree`
//!   ([`crate::ast::testeq::equivalent`]), and printing is idempotent:
//!   `print_module(parse(print_module(tree))) == print_module(tree)`. The same holds for the
//!   parser-test-style sources in [`CORPUS`]. (`print::tests` checks the same round trip on its
//!   own hand-written sources.)
//! - **P7** (Req 12.9): for arbitrary input of at most [`MAX_SOURCE`] bytes, `parse` returns a
//!   tree and/or diagnostics without panicking, overflowing the stack or aborting, on a thread
//!   with a 1 MiB stack. A default run parses at least [`P7_MIN_CASES`] inputs. The small-stack
//!   thread only parses and drops: derived `Clone`, `PartialEq` and `Debug` on the tree recurse,
//!   so nothing there clones, compares or formats a tree, and coverage is counted from the
//!   diagnostics alone.
//!
//! The P6 generator writes source text from a grammar, not trees: modules with attributes,
//! options, declarations (variables with bounds, `New`, `WithEvents`, fixed-length strings,
//! `Const`, `Type`, `Enum`, `Declare`, `Event`, `Implements`, `Def…`), procedures of every kind
//! with every parameter form, and bodies that use every statement family (assignments, calls,
//! `Debug.Print`, file statements, jumps, labels and line numbers, error handling, local
//! declarations, `ReDim`/`Erase`, `RaiseEvent`, single-line `If`, `With`-relative members, named
//! and missing arguments, directives and attributes) with blocks nested up to eight deep, all in
//! random keyword and identifier case with random spacing, line continuations, comments and
//! line terminators. Expressions stay shallow (the derived tree traits recurse).
//!
//! P7 mixes: P6-style source; the same source mutated (block closers dropped, words deleted or
//! swapped, tokens, unbalanced parentheses and stray `:`/`_` inserted, lines swapped or
//! duplicated, text cut short); token soup from the keyword table and punctuation; Unicode
//! noise; and on one case in [`BIG_EVERY`], 1 to 4 MiB of repeated pieces. Pathological shapes
//! (10,000-deep parentheses, blocks, prefix operators and calls, 200,000-link flat chains, a
//! 4 MiB module, very long lines, an input of exactly [`MAX_SOURCE`] bytes and one byte more)
//! run in [`p7_pathological_shapes_never_panic_on_a_small_stack`], one round per 64 seeds.
//!
//! Knobs (environment variables):
//! - `GRIDCRAFT_FUZZ_SEEDS` (default 64): P6 runs 16 cases per seed (1,024 by default), P7 runs
//!   160 per seed (10,240 by default). Soak with
//!   `GRIDCRAFT_FUZZ_SEEDS=500 cargo test --release -p gridcraft-vba --lib parser::props::`.
//! - `GRIDCRAFT_FUZZ_SEED=17`: replay that one case (failure messages print the command line).
//! - `GRIDCRAFT_FUZZ_TRACE=1`: P7 prints each case number before parsing it. A stack overflow
//!   aborts the whole test binary, so run with `-- --nocapture` to see the last case.
use super::{ParseDiag, ParseDiagKind, parse};
use crate::ast::testeq::{equivalent, normalized};
use crate::lexer::{Keyword, MAX_SOURCE};
use crate::print::print_module;
use std::panic::{AssertUnwindSafe, catch_unwind};
/// Default `GRIDCRAFT_FUZZ_SEEDS`: small enough that `cargo xtask ci` stays fast.
const DEFAULT_SEEDS: u64 = 64;
/// P6 cases per seed.
const P6_CASES_PER_SEED: u64 = 16;
/// P7 cases per seed.
const P7_CASES_PER_SEED: u64 = 160;
/// A default P7 run parses at least this many inputs (Req 12.9).
const P7_MIN_CASES: usize = 10_000;
/// One P7 case in this many is a large input (five in the default 10,240 cases).
const BIG_EVERY: u64 = 2048;
/// Large P7 inputs are this many bytes, inclusive.
const BIG_LEN: (usize, usize) = (1 << 20, 4 << 20);
/// The stack the P7 parser threads run on.
const SMALL_STACK: usize = 1 << 20;
/// The share of generated modules that must parse with zero diagnostics, in percent.
const MIN_ACCEPT_PERCENT: usize = 90;
/// Keeps the property tests' random sequences apart from each other and from P5's.
const P6_STREAM: u64 = 0x5036_5036_5036_5036;
const P7_STREAM: u64 = 0x5037_5037_5037_5037;
const PATHOLOGICAL_STREAM: u64 = 0x5037_7061_7468_6f6c;
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
fn seed_count() -> u64 {
    env_num("GRIDCRAFT_FUZZ_SEEDS").unwrap_or(DEFAULT_SEEDS)
}
/// The cases to run: `GRIDCRAFT_FUZZ_SEED` alone, or `0..seeds * per_seed`.
fn seed_list(per_seed: u64) -> (Vec<u64>, bool) {
    match env_num("GRIDCRAFT_FUZZ_SEED") {
        Some(one) => (vec![one], true),
        None => ((0..seed_count().saturating_mul(per_seed)).collect(), false),
    }
}
fn replay(seed: u64, test: &str) -> String {
    format!("replay: GRIDCRAFT_FUZZ_SEED={seed} cargo test -p gridcraft-vba --lib parser::props::{test}")
}
/// `src` for a failure message, if it's short enough to read.
fn shown(src: &str) -> String {
    if src.len() <= 4000 { format!("\n----- input -----\n{src}\n-----------------") } else { String::new() }
}
// ---------------------------------------------------------------- the VBA source generator
/// Plain identifiers: none is a keyword, some are not ASCII.
const NAMES: [&str; 20] = [
    "x", "i", "j", "n", "total", "s", "rs", "ws", "obj", "coll", "arr", "result", "Größe", "数値", "café", "my_var", "a_", "Ωmega", "Sheet1",
    "counter",
];
/// Contextual words the grammar quotes but doesn't reserve: valid identifiers in expressions.
const CONTEXTUAL: [&str; 10] = ["Text", "Binary", "Compare", "Explicit", "Base", "Module", "Access", "Output", "Random", "Append"];
/// Member names, keywords included (`x.Print`, `x.End` are member accesses).
const MEMBERS: [&str; 12] = ["Value", "Name", "Count", "Item", "Cells", "Range", "Print", "End", "Text", "Add", "Close", "Größe"];
const FUNCTIONS: [&str; 10] = ["Len", "Left$", "Mid$", "UBound", "LBound", "CStr", "CLng", "Abs", "Array", "Format"];
const TYPE_SUFFIXES: [char; 7] = ['%', '&', '^', '!', '#', '@', '$'];
const TYPES: [&str; 14] = [
    "Long",
    "String",
    "Variant",
    "Double",
    "Integer",
    "Boolean",
    "Byte",
    "Currency",
    "Date",
    "Object",
    "Single",
    "LongLong",
    "Collection",
    "Excel.Range",
];
const LITERALS: [&str; 29] = [
    "0",
    "1",
    "42",
    "32767",
    "5%",
    "6&",
    "&HFF",
    "&hff&",
    "&O17",
    "&17",
    "1.5",
    "1.5E+10",
    "2#",
    "3!",
    "4@",
    "\"\"",
    "\"x\"",
    "\"say \"\"hi\"\"\"",
    "\"日本語 😀\"",
    "\"a'b _\"",
    "#1/2/2003#",
    "# 10:30 PM #",
    "#Jan 1, 2000#",
    "True",
    "False",
    "Nothing",
    "Empty",
    "Null",
    "\"C:\\a.txt\"",
];
const BINARY_OPS: [&str; 21] =
    ["+", "-", "*", "/", "\\", "^", "&", "=", "<>", "<", ">", "<=", ">=", "Mod", "And", "Or", "Xor", "Eqv", "Imp", "Like", "Is"];
const COMPARISONS: [&str; 6] = ["=", "<>", "<", ">", "<=", ">="];
const LABELS: [&str; 5] = ["done", "Handler", "100", "ErrH", "l_1"];
const COMMENT_WORDS: [&str; 11] = ["note", "Größe", "日本語", "😀", "\"open", "#1/2", "[x", "rem", "'", "Then", "End Sub"];
const DIRECTIVES: [&str; 6] = ["#If VBA7 Then", "#Else", "#End If", "#Const DEBUG_MODE = 1", "#ElseIf Win64 Then", "#if  Mac then"];
/// Line terminators, CRLF and LF most often.
const NEWLINES: [&str; 8] = ["\r\n", "\r\n", "\r\n", "\n", "\n", "\n", "\r", "\u{2028}"];
/// What a statement may do where it's written.
#[derive(Debug, Clone, Copy)]
struct Ctx {
    /// Blocks that may still open inside this one.
    depth: u8,
    indent: usize,
    /// Inside a `With` block: `.x` and `!x` are allowed.
    with: bool,
    in_for: bool,
    in_do: bool,
    /// The procedure kind for `Exit`: `Sub`, `Function` or `Property`.
    exit: &'static str,
    /// The first statement of each body opens a block while depth remains.
    spine: bool,
}
/// Writes realistic VBA source into `out`.
struct Gen {
    rng: Rng,
    out: String,
    /// Statements left for this module, so a module stays small whatever the dice say.
    budget: usize,
}
impl Gen {
    fn new(rng: Rng) -> Self {
        Self { rng, out: String::new(), budget: 0 }
    }
    // ----- tokens and trivia
    /// `word` in random case: mostly as written, sometimes lower, upper or mixed (ASCII only).
    fn word(&mut self, word: &str) {
        match self.rng.below(20) {
            0..=13 => self.out.push_str(word),
            14 | 15 => self.out.push_str(&word.to_ascii_lowercase()),
            16 | 17 => self.out.push_str(&word.to_ascii_uppercase()),
            _ => {
                for c in word.chars() {
                    let upper = self.rng.chance(50);
                    self.out.push(if upper { c.to_ascii_uppercase() } else { c.to_ascii_lowercase() });
                }
            }
        }
    }
    /// Space-separated keywords, each in random case, with random spacing between.
    fn words(&mut self, text: &str) {
        for (i, w) in text.split(' ').enumerate() {
            if i > 0 {
                self.sp();
            }
            self.word(w);
        }
    }
    /// Whitespace between two tokens.
    fn sp(&mut self) {
        let s = self.rng.pick(&[" ", " ", " ", " ", " ", "  ", "\t", " \t"]);
        self.out.push_str(s);
    }
    /// Whitespace between two tokens where a line continuation may go.
    fn op_sp(&mut self) {
        if self.rng.chance(4) {
            let nl = self.rng.pick(&["\r\n", "\n"]);
            let lead = self.rng.pick(&["    ", "\t", " "]);
            self.out.push_str(" _");
            self.out.push_str(nl);
            self.out.push_str(lead);
        } else {
            self.sp();
        }
    }
    fn indent(&mut self, n: usize) {
        for _ in 0..n {
            let s = self.rng.pick(&["    ", "    ", "  ", "\t"]);
            self.out.push_str(s);
        }
    }
    fn comment_text(&mut self) {
        for _ in 0..self.rng.range(0, 4) {
            self.out.push(' ');
            let w = self.rng.pick(&COMMENT_WORDS);
            self.out.push_str(w);
        }
    }
    /// Ends a line, sometimes after a `'` comment.
    fn nl(&mut self) {
        if self.rng.chance(12) {
            self.sp();
            self.out.push('\'');
            self.comment_text();
        }
        self.newline();
    }
    fn newline(&mut self) {
        let nl = self.rng.pick(&NEWLINES);
        self.out.push_str(nl);
    }
    // ----- names and expressions
    fn ident(&mut self) {
        let name = self.rng.pick(&NAMES);
        self.word(name);
    }
    fn label(&mut self) {
        let label = self.rng.pick(&LABELS);
        self.word(label);
    }
    fn type_name(&mut self) {
        let ty = self.rng.pick(&TYPES);
        self.word(ty);
    }
    fn member_name(&mut self) {
        let m = self.rng.pick(&MEMBERS);
        self.word(m);
    }
    fn expr(&mut self, depth: u8, with: bool) {
        if depth == 0 || self.rng.chance(40) {
            self.term(depth, with);
            return;
        }
        match self.rng.below(7) {
            0..=2 => {
                for k in 0..self.rng.range(2, 4) {
                    if k > 0 {
                        self.op_sp();
                        let op = self.rng.pick(&BINARY_OPS);
                        self.word(op);
                        self.op_sp();
                    }
                    self.expr(depth - 1, with);
                }
            }
            3 => {
                self.out.push('-');
                self.expr(depth - 1, with);
            }
            4 => {
                self.word("Not");
                self.sp();
                self.expr(depth - 1, with);
            }
            _ => self.paren(depth - 1, with),
        }
    }
    fn paren(&mut self, depth: u8, with: bool) {
        self.out.push('(');
        if self.rng.chance(20) {
            self.sp();
        }
        self.expr(depth, with);
        self.out.push(')');
    }
    fn term(&mut self, depth: u8, with: bool) {
        match self.rng.below(12) {
            0 | 1 => {
                self.ident();
                if self.rng.chance(20) {
                    let suffix = self.rng.pick(&TYPE_SUFFIXES);
                    self.out.push(suffix);
                }
            }
            2..=4 => {
                let lit = self.rng.pick(&LITERALS);
                self.word(lit);
            }
            5 | 6 => self.member_chain(depth, with),
            7 => self.function_call(depth, with),
            8 => {
                let word = self.rng.pick(&CONTEXTUAL);
                self.word(word);
            }
            9 if with => self.with_member(depth),
            10 if depth > 0 => {
                self.out.push('(');
                self.words("TypeOf");
                self.sp();
                self.ident();
                self.sp();
                self.word("Is");
                self.sp();
                self.type_name();
                self.out.push(')');
            }
            _ => self.ident(),
        }
    }
    /// `.x`, `!x` or `![x y]`, maybe followed by more links: a member of the `With` object.
    fn with_member(&mut self, depth: u8) {
        match self.rng.below(4) {
            0 => self.out.push_str("![Field Name]"),
            1 => {
                self.out.push('!');
                self.ident();
            }
            _ => {
                self.out.push('.');
                self.member_name();
            }
        }
        self.links(depth, true, 2);
    }
    /// An object (an identifier, `Me`, a foreign name, a `With` member) and one to three links.
    fn member_chain(&mut self, depth: u8, with: bool) {
        match self.rng.below(6) {
            0 => self.word("Me"),
            1 => self.out.push_str(self.rng.pick(&["[A B]", "[If]", "[x]"])),
            2 if with => {
                self.out.push('.');
                self.member_name();
            }
            _ => self.ident(),
        }
        let max = self.rng.range(1, 3);
        self.links(depth, with, max);
    }
    /// Up to `max` of `.name`, `!name`, `![a b]` and `(args)`.
    fn links(&mut self, depth: u8, with: bool, max: usize) {
        for _ in 0..self.rng.below(max + 1) {
            match self.rng.below(6) {
                0 => {
                    self.out.push('!');
                    self.ident();
                }
                1 => self.out.push_str("![Other Name]"),
                2 if depth > 0 => {
                    self.out.push('(');
                    self.args(depth - 1, with);
                    self.out.push(')');
                }
                _ => {
                    self.out.push('.');
                    self.member_name();
                }
            }
        }
    }
    fn function_call(&mut self, depth: u8, with: bool) {
        if self.rng.chance(50) {
            let f = self.rng.pick(&FUNCTIONS);
            self.word(f);
        } else {
            self.ident();
        }
        self.out.push('(');
        if depth > 0 {
            self.args(depth - 1, with);
        }
        self.out.push(')');
        if self.rng.chance(10) {
            self.out.push_str("(1)");
        }
    }
    /// A comma-separated argument list: values, missing arguments (not last) and named
    /// arguments (last).
    fn args(&mut self, depth: u8, with: bool) {
        let n = self.rng.range(0, 3);
        for k in 0..n {
            if k > 0 {
                self.out.push(',');
                self.op_sp();
            }
            let last = k + 1 == n;
            if !last && self.rng.chance(10) {
                continue;
            }
            self.arg(depth, with, last);
        }
    }
    fn arg(&mut self, depth: u8, with: bool, last: bool) {
        if last && self.rng.chance(15) {
            self.ident();
            self.out.push_str(":=");
        } else if self.rng.chance(4) {
            self.word("AddressOf");
            self.sp();
            self.ident();
            if self.rng.chance(50) {
                self.out.push('.');
                self.ident();
            }
            return;
        }
        self.expr(depth, with);
    }
    /// A condition for `If`, `While`, `Do` and `Loop`.
    fn cond(&mut self, with: bool) {
        match self.rng.below(8) {
            0 => {
                self.word("TypeOf");
                self.sp();
                self.ident();
                self.sp();
                self.word("Is");
                self.sp();
                self.type_name();
            }
            1 => {
                self.word("Not");
                self.sp();
                self.ident();
                self.sp();
                self.words("Is Nothing");
            }
            _ => self.expr(3, with),
        }
    }
    // ----- statements
    /// An assignment target or call target.
    fn target(&mut self, ctx: Ctx) {
        match self.rng.below(6) {
            0 if ctx.with => self.with_member(1),
            1 => {
                self.ident();
                self.out.push('(');
                self.args(1, ctx.with);
                self.out.push(')');
            }
            2 | 3 => self.member_chain(1, ctx.with),
            _ => self.ident(),
        }
    }
    fn assign(&mut self, ctx: Ctx) {
        match self.rng.below(8) {
            0 => {
                self.word("Let");
                self.sp();
            }
            1 => {
                self.word("Set");
                self.sp();
                self.target(ctx);
                self.out.push_str(" = ");
                match self.rng.below(3) {
                    0 => self.words("New Collection"),
                    1 => self.word("Nothing"),
                    _ => self.member_chain(1, ctx.with),
                }
                return;
            }
            2 => {
                let w = self.rng.pick(&["LSet", "RSet"]);
                self.word(w);
                self.sp();
                self.ident();
                self.out.push_str(" = ");
                self.ident();
                return;
            }
            _ => {}
        }
        self.target(ctx);
        self.sp();
        self.out.push('=');
        self.op_sp();
        self.expr(3, ctx.with);
    }
    fn call(&mut self, ctx: Ctx) {
        match self.rng.below(7) {
            0 => {
                self.word("Call");
                self.sp();
                if self.rng.chance(50) {
                    self.member_chain(1, ctx.with);
                } else {
                    self.ident();
                }
                if self.rng.chance(75) {
                    self.out.push('(');
                    self.args(2, ctx.with);
                    self.out.push(')');
                }
            }
            1 => {
                self.ident();
                self.out.push('(');
                self.args(2, ctx.with);
                self.out.push(')');
            }
            2 => {
                let w = self.rng.pick(&["DoEvents", "Beep", "Initialize"]);
                self.word(w);
            }
            3 => {
                self.word("MsgBox");
                self.out.push(' ');
                self.words("Prompt:=");
                self.expr(1, ctx.with);
                self.out.push_str(", Buttons:=vbOKOnly");
            }
            _ => {
                if self.rng.chance(40) {
                    self.member_chain(0, ctx.with);
                } else {
                    self.ident();
                }
                if self.rng.chance(80) {
                    self.out.push(' ');
                    if self.rng.chance(10) {
                        self.out.push(',');
                        self.sp();
                    }
                    self.args(2, ctx.with);
                }
            }
        }
    }
    fn debug_print(&mut self, ctx: Ctx) {
        self.word("Debug");
        self.out.push('.');
        self.word("Print");
        let n = self.rng.range(0, 4);
        for k in 0..n {
            if k == 0 {
                self.sp();
            }
            if self.rng.chance(10) {
                self.out.push_str("Tab(5)");
            } else if k > 0 || self.rng.chance(90) {
                self.expr(2, ctx.with);
            }
            if k + 1 < n || self.rng.chance(20) {
                let sep = self.rng.pick(&[";", ",", "; "]);
                self.out.push_str(sep);
                self.sp();
            }
        }
    }
    fn file_number(&mut self) {
        self.out.push('#');
        match self.rng.below(3) {
            0 => self.out.push('f'),
            _ => {
                let n = self.rng.range(1, 511).to_string();
                self.out.push_str(&n);
            }
        }
    }
    fn open(&mut self, ctx: Ctx) {
        self.word("Open");
        self.sp();
        self.expr(1, ctx.with);
        if self.rng.chance(80) {
            self.sp();
            self.word("For");
            self.sp();
            let mode = self.rng.pick(&["Input", "Output", "Append", "Binary", "Random"]);
            self.word(mode);
        }
        if self.rng.chance(30) {
            self.sp();
            let access = self.rng.pick(&["Access Read", "Access Write", "Access Read Write"]);
            self.words(access);
        }
        if self.rng.chance(30) {
            self.sp();
            let lock = self.rng.pick(&["Shared", "Lock Read", "Lock Write", "Lock Read Write"]);
            self.words(lock);
        }
        self.sp();
        self.word("As");
        self.sp();
        if self.rng.chance(80) {
            self.file_number();
        } else {
            self.out.push('2');
        }
        if self.rng.chance(20) {
            self.out.push_str(" Len = 128");
        }
    }
    fn file_statement(&mut self, ctx: Ctx) {
        match self.rng.below(11) {
            0 => self.open(ctx),
            1 | 2 => {
                let op = self.rng.pick(&["Print", "Write", "Input"]);
                self.word(op);
                self.sp();
                self.file_number();
                if self.rng.chance(85) {
                    self.out.push(',');
                    for k in 0..self.rng.range(0, 3) {
                        if k > 0 {
                            // `Input #` takes a comma-delimited list; `Print #` and `Write #` take `;` too.
                            let sep = if op == "Input" { "," } else { self.rng.pick(&[",", ";"]) };
                            self.out.push_str(sep);
                        }
                        self.sp();
                        self.expr(1, ctx.with);
                    }
                }
            }
            3 => {
                self.words("Line Input");
                self.sp();
                self.file_number();
                self.out.push_str(", ");
                self.ident();
            }
            4 | 5 => {
                let op = self.rng.pick(&["Get", "Put"]);
                self.word(op);
                self.sp();
                self.file_number();
                self.out.push_str(self.rng.pick(&[", , ", ", 3, "]));
                self.ident();
            }
            6 => {
                let op = self.rng.pick(&["Seek", "Width"]);
                self.word(op);
                self.sp();
                self.file_number();
                self.out.push_str(", 10");
            }
            7 | 8 => {
                let op = self.rng.pick(&["Lock", "Unlock"]);
                self.word(op);
                self.sp();
                self.file_number();
                if self.rng.chance(50) {
                    self.out.push_str(self.rng.pick(&[", 1 To 5", ", 3"]));
                }
            }
            _ => {
                self.word("Close");
                if self.rng.chance(70) {
                    self.sp();
                    self.file_number();
                    if self.rng.chance(30) {
                        self.out.push_str(", ");
                        self.file_number();
                    }
                }
            }
        }
    }
    fn jump(&mut self, ctx: Ctx) {
        match self.rng.below(9) {
            0 | 1 => {
                self.word("GoTo");
                self.sp();
                self.label();
            }
            2 => {
                self.word("GoSub");
                self.sp();
                self.label();
            }
            3 => self.word("Return"),
            4 => {
                self.word("On");
                self.sp();
                self.expr(1, ctx.with);
                self.sp();
                let w = self.rng.pick(&["GoTo", "GoSub"]);
                self.word(w);
                self.sp();
                self.label();
                self.out.push_str(", ");
                self.label();
            }
            5 | 6 => {
                self.words("On Error");
                self.sp();
                let action = self.rng.pick(&["GoTo Handler", "GoTo 0", "GoTo -1", "Resume Next"]);
                self.words(action);
            }
            _ => {
                self.word("Resume");
                match self.rng.below(4) {
                    0 => {}
                    1 => {
                        self.sp();
                        self.word("Next");
                    }
                    _ => {
                        self.sp();
                        self.label();
                    }
                }
            }
        }
    }
    fn exit(&mut self, ctx: Ctx) {
        let mut kinds = vec![ctx.exit];
        if ctx.in_for {
            kinds.push("For");
        }
        if ctx.in_do {
            kinds.push("Do");
        }
        let kind = self.rng.pick(&kinds);
        self.word("Exit");
        self.sp();
        self.word(kind);
    }
    fn bounds(&mut self) {
        let n = self.rng.range(0, 2);
        for k in 0..n {
            if k > 0 {
                self.out.push_str(", ");
            }
            if self.rng.chance(50) {
                self.out.push_str(self.rng.pick(&["0", "1", "n"]));
                self.sp();
                self.word("To");
                self.sp();
            }
            self.out.push_str(self.rng.pick(&["10", "5", "UBound(a)", "n + 1"]));
        }
    }
    /// One variable in a `Dim`-style list.
    fn var_decl(&mut self, with_events: bool) {
        if with_events {
            self.word("WithEvents");
            self.sp();
        }
        self.ident();
        if !with_events && self.rng.chance(15) {
            let suffix = self.rng.pick(&TYPE_SUFFIXES);
            self.out.push(suffix);
            return;
        }
        if !with_events && self.rng.chance(25) {
            self.out.push('(');
            self.bounds();
            self.out.push(')');
        }
        if with_events || self.rng.chance(70) {
            self.sp();
            self.word("As");
            self.sp();
            match self.rng.below(8) {
                0 if !with_events => self.words("New Collection"),
                1 => self.out.push_str("Excel.Application"),
                2 => {
                    self.word("String");
                    self.out.push_str(" * ");
                    self.out.push_str(self.rng.pick(&["10", "8", "n"]));
                }
                _ => self.type_name(),
            }
        }
    }
    fn var_list(&mut self, with_events: bool) {
        for k in 0..self.rng.range(1, 3) {
            if k > 0 {
                self.out.push(',');
                self.op_sp();
            }
            self.var_decl(with_events && k == 0);
        }
    }
    fn const_list(&mut self) {
        self.word("Const");
        self.sp();
        for k in 0..self.rng.range(1, 2) {
            if k > 0 {
                self.out.push_str(", ");
            }
            self.ident();
            if self.rng.chance(40) {
                self.sp();
                self.word("As");
                self.sp();
                self.type_name();
            }
            self.out.push_str(" = ");
            self.expr(2, false);
        }
    }
    fn local_decl(&mut self) {
        match self.rng.below(5) {
            0 | 1 => {
                self.word("Dim");
                self.sp();
                self.var_list(false);
            }
            2 => {
                self.word("Static");
                self.sp();
                self.var_list(false);
            }
            3 => self.const_list(),
            _ => {
                self.word("ReDim");
                self.sp();
                if self.rng.chance(40) {
                    self.word("Preserve");
                    self.sp();
                }
                if self.rng.chance(20) {
                    self.word("Me");
                    self.out.push('.');
                }
                self.ident();
                self.out.push('(');
                self.bounds();
                self.out.push(')');
                if self.rng.chance(30) {
                    self.sp();
                    self.word("As");
                    self.sp();
                    self.type_name();
                }
            }
        }
    }
    /// A statement that fits on one line, next to others after `:` and inside a single-line
    /// `If` (`inline`, where declarations, `Stop` and `End` aren't generated).
    fn simple(&mut self, ctx: Ctx, inline: bool) {
        match self.rng.below(20) {
            0..=4 => self.assign(ctx),
            5..=7 => self.call(ctx),
            8 => self.debug_print(ctx),
            9 | 10 => self.file_statement(ctx),
            11 | 12 => self.jump(ctx),
            13 => self.exit(ctx),
            14 if !inline => self.local_decl(),
            15 => {
                self.word("Erase");
                self.sp();
                self.ident();
                if self.rng.chance(50) {
                    self.out.push_str(", ");
                    self.ident();
                }
            }
            16 => {
                self.word("RaiseEvent");
                self.sp();
                self.ident();
                if self.rng.chance(60) {
                    self.out.push('(');
                    self.args(1, ctx.with);
                    self.out.push(')');
                }
            }
            17 if !inline => {
                let w = self.rng.pick(&["Stop", "End"]);
                self.word(w);
            }
            _ => self.assign(ctx),
        }
    }
    /// A body: zero to four statements, the first a block when `ctx.spine`.
    fn body(&mut self, ctx: Ctx) {
        let n = if ctx.spine && ctx.depth > 0 { self.rng.range(1, 3) } else { self.rng.range(0, 4) };
        for k in 0..n {
            if k == 0 && ctx.spine && ctx.depth > 0 {
                self.block(ctx);
            } else {
                self.stmt(ctx);
            }
        }
    }
    fn stmt(&mut self, ctx: Ctx) {
        self.budget = self.budget.saturating_sub(1);
        let roll = self.rng.below(100);
        if ctx.depth > 0 && self.budget > 0 && roll < 22 {
            self.block(ctx);
        } else if roll < 30 {
            self.special_line(ctx);
        } else if roll < 36 {
            self.indent(ctx.indent);
            self.single_line_if(ctx, 0);
            self.nl();
        } else {
            self.indent(ctx.indent);
            self.simple(ctx, false);
            if self.rng.chance(12) {
                self.out.push(':');
                self.sp();
                self.simple(ctx, false);
            }
            self.nl();
        }
    }
    /// Labels, line numbers, comments, blank lines, directives and attributes.
    fn special_line(&mut self, ctx: Ctx) {
        match self.rng.below(8) {
            0 => {
                self.label();
                self.out.push(':');
                if self.rng.chance(40) {
                    self.sp();
                    self.simple(ctx, false);
                }
                self.nl();
            }
            1 => {
                let n = self.rng.range(1, 9999).to_string();
                self.out.push_str(&n);
                if self.rng.chance(30) {
                    self.out.push(':');
                }
                self.sp();
                self.simple(ctx, false);
                self.nl();
            }
            2 => {
                self.indent(ctx.indent);
                self.out.push('\'');
                self.comment_text();
                self.newline();
            }
            3 => {
                self.indent(ctx.indent);
                self.word("Rem");
                self.comment_text();
                self.newline();
            }
            4 => self.newline(),
            5 => {
                let d = self.rng.pick(&DIRECTIVES);
                self.out.push_str(d);
                self.newline();
            }
            6 => {
                self.indent(ctx.indent);
                self.out.push_str("Attribute ");
                self.ident();
                self.out.push_str(".VB_UserMemId = 0");
                self.nl();
            }
            _ => {
                self.indent(ctx.indent);
                self.ident();
                self.out.push_str(" = ");
                let w = self.rng.pick(&CONTEXTUAL);
                self.word(w);
                self.nl();
            }
        }
    }
    /// `If c Then s [: s] [Else s [: s]]`, or `If c Then 10`, on one line.
    fn single_line_if(&mut self, ctx: Ctx, nested: u8) {
        self.word("If");
        self.sp();
        self.cond(ctx.with);
        self.sp();
        self.word("Then");
        self.sp();
        if self.rng.chance(10) {
            self.out.push_str(self.rng.pick(&["10", "200"]));
            return;
        }
        if nested == 0 && self.rng.chance(8) {
            self.single_line_if(ctx, 1);
            return;
        }
        self.simple(ctx, true);
        if self.rng.chance(20) {
            self.out.push(':');
            self.sp();
            self.simple(ctx, true);
        }
        if self.rng.chance(40) {
            self.sp();
            self.word("Else");
            self.sp();
            self.simple(ctx, true);
            if self.rng.chance(20) {
                self.out.push(':');
                self.sp();
                self.simple(ctx, true);
            }
        }
    }
    fn block(&mut self, ctx: Ctx) {
        let inner = Ctx { depth: ctx.depth.saturating_sub(1), indent: ctx.indent + 1, ..ctx };
        self.indent(ctx.indent);
        match self.rng.below(8) {
            0 | 1 => self.if_block(ctx, inner),
            2 => self.select_block(ctx, inner),
            3 => self.for_block(ctx, inner),
            4 => self.do_block(ctx, inner),
            5 => {
                self.word("While");
                self.sp();
                self.cond(ctx.with);
                self.nl();
                self.body(Ctx { in_do: false, ..inner });
                self.indent(ctx.indent);
                self.word("Wend");
            }
            _ => {
                self.word("With");
                self.sp();
                if ctx.with && self.rng.chance(50) {
                    self.with_member(0);
                } else {
                    self.member_chain(0, false);
                }
                self.nl();
                self.body(Ctx { with: true, ..inner });
                self.indent(ctx.indent);
                self.words("End With");
            }
        }
        self.nl();
    }
    fn if_block(&mut self, ctx: Ctx, inner: Ctx) {
        self.word("If");
        self.sp();
        self.cond(ctx.with);
        self.sp();
        self.word("Then");
        self.nl();
        self.body(inner);
        for _ in 0..self.rng.range(0, 2) {
            self.indent(ctx.indent);
            self.word("ElseIf");
            self.sp();
            self.cond(ctx.with);
            self.sp();
            self.word("Then");
            self.nl();
            self.body(inner);
        }
        if self.rng.chance(40) {
            self.indent(ctx.indent);
            self.word("Else");
            self.nl();
            self.body(inner);
        }
        self.indent(ctx.indent);
        let end = self.rng.pick(&["End If", "End If", "EndIf"]);
        self.words(end);
    }
    fn case_test(&mut self, with: bool) {
        match self.rng.below(5) {
            0 => {
                self.expr(1, with);
                self.sp();
                self.word("To");
                self.sp();
                self.expr(1, with);
            }
            1 => {
                self.word("Is");
                self.sp();
                let op = self.rng.pick(&COMPARISONS);
                self.out.push_str(op);
                self.sp();
                self.expr(1, with);
            }
            2 => {
                let op = self.rng.pick(&COMPARISONS);
                self.out.push_str(op);
                self.sp();
                self.expr(1, with);
            }
            _ => self.expr(2, with),
        }
    }
    fn select_block(&mut self, ctx: Ctx, inner: Ctx) {
        self.words("Select Case");
        self.sp();
        self.expr(2, ctx.with);
        self.nl();
        let body = Ctx { indent: inner.indent + 1, ..inner };
        for _ in 0..self.rng.range(0, 3) {
            self.indent(inner.indent);
            self.word("Case");
            self.sp();
            for k in 0..self.rng.range(1, 3) {
                if k > 0 {
                    self.out.push(',');
                    self.sp();
                }
                self.case_test(ctx.with);
            }
            self.nl();
            self.body(body);
        }
        if self.rng.chance(40) {
            self.indent(inner.indent);
            self.words("Case Else");
            self.nl();
            self.body(body);
        }
        self.indent(ctx.indent);
        self.words("End Select");
    }
    fn for_block(&mut self, ctx: Ctx, inner: Ctx) {
        self.word("For");
        self.sp();
        let each = self.rng.chance(35);
        if each {
            self.word("Each");
            self.sp();
        }
        let start = self.out.len();
        self.ident();
        let var = self.out[start..].to_owned();
        self.sp();
        if each {
            self.word("In");
            self.sp();
            self.expr(2, ctx.with);
        } else {
            self.out.push('=');
            self.sp();
            self.expr(2, ctx.with);
            self.sp();
            self.word("To");
            self.sp();
            self.expr(2, ctx.with);
            if self.rng.chance(30) {
                self.sp();
                self.word("Step");
                self.sp();
                self.expr(1, ctx.with);
            }
        }
        self.nl();
        self.body(Ctx { in_for: true, ..inner });
        self.indent(ctx.indent);
        self.word("Next");
        if self.rng.chance(50) {
            self.sp();
            self.out.push_str(&var);
        }
    }
    fn loop_cond(&mut self, with: bool) {
        self.sp();
        let w = self.rng.pick(&["While", "Until"]);
        self.word(w);
        self.sp();
        self.cond(with);
    }
    fn do_block(&mut self, ctx: Ctx, inner: Ctx) {
        self.word("Do");
        let at = self.rng.below(3);
        if at == 0 {
            self.loop_cond(ctx.with);
        }
        self.nl();
        self.body(Ctx { in_do: true, ..inner });
        self.indent(ctx.indent);
        self.word("Loop");
        if at == 1 {
            self.loop_cond(ctx.with);
        }
    }
    // ----- procedures and modules
    fn params(&mut self) {
        let n = self.rng.range(0, 4);
        let mut optional = false;
        for k in 0..n {
            if k > 0 {
                self.out.push(',');
                self.op_sp();
            }
            if k + 1 == n && !optional && self.rng.chance(15) {
                self.words("ParamArray");
                self.sp();
                self.ident();
                self.out.push_str("() As Variant");
                break;
            }
            optional |= self.rng.chance(20);
            if optional {
                self.word("Optional");
                self.sp();
            }
            if self.rng.chance(50) {
                let p = self.rng.pick(&["ByVal", "ByRef"]);
                self.word(p);
                self.sp();
            }
            self.ident();
            if self.rng.chance(15) {
                self.out.push_str("()");
            }
            if self.rng.chance(60) {
                self.sp();
                self.word("As");
                self.sp();
                self.type_name();
            }
            if optional && self.rng.chance(70) {
                self.out.push_str(" = ");
                let lit = self.rng.pick(&LITERALS);
                self.word(lit);
            }
        }
    }
    fn procedure(&mut self) {
        let static_first = self.rng.chance(5);
        if static_first {
            self.word("Static");
            self.sp();
        }
        if self.rng.chance(60) {
            let vis = self.rng.pick(&["Public", "Private", "Friend", "Public", "Private"]);
            self.word(vis);
            self.sp();
        }
        if !static_first && self.rng.chance(8) {
            self.word("Static");
            self.sp();
        }
        let (kind, end, exit, returns) = match self.rng.below(7) {
            0..=2 => ("Sub", "End Sub", "Sub", false),
            3 | 4 => ("Function", "End Function", "Function", true),
            5 => ("Property Get", "End Property", "Property", true),
            _ => (self.rng.pick(&["Property Let", "Property Set"]), "End Property", "Property", false),
        };
        self.words(kind);
        self.sp();
        self.ident();
        if kind == "Function" && self.rng.chance(15) {
            self.out.push('$');
        }
        self.out.push('(');
        self.params();
        self.out.push(')');
        if returns && self.rng.chance(70) {
            self.sp();
            self.word("As");
            self.sp();
            self.type_name();
            if self.rng.chance(15) {
                self.out.push_str("()");
            }
        }
        self.nl();
        let spine = self.rng.chance(10);
        let depth = if spine { 8 } else { self.rng.range(1, 4) as u8 };
        let ctx = Ctx { depth, indent: 1, with: false, in_for: false, in_do: false, exit, spine };
        self.body(ctx);
        self.words(end);
        self.nl();
    }
    fn type_block(&mut self) {
        self.word("Type");
        self.sp();
        self.ident();
        self.nl();
        for _ in 0..self.rng.range(1, 4) {
            self.indent(1);
            self.ident();
            if self.rng.chance(25) {
                self.out.push('(');
                self.bounds();
                self.out.push(')');
            }
            self.sp();
            self.word("As");
            self.sp();
            if self.rng.chance(20) {
                self.word("String");
                self.out.push_str(" * 8");
            } else {
                self.type_name();
            }
            self.nl();
        }
        self.words("End Type");
    }
    fn enum_block(&mut self) {
        self.word("Enum");
        self.sp();
        self.ident();
        self.nl();
        for _ in 0..self.rng.range(1, 4) {
            self.indent(1);
            self.ident();
            if self.rng.chance(50) {
                self.out.push_str(" = ");
                self.expr(2, false);
            }
            self.nl();
        }
        self.words("End Enum");
    }
    fn declare(&mut self) {
        self.word("Declare");
        self.sp();
        if self.rng.chance(50) {
            self.word("PtrSafe");
            self.sp();
        }
        let function = self.rng.chance(50);
        self.word(if function { "Function" } else { "Sub" });
        self.sp();
        self.ident();
        self.sp();
        self.word("Lib");
        self.out.push_str(self.rng.pick(&[" \"kernel32\"", " \"user32.dll\"", " \"a \"\"b\"\".dll\""]));
        if self.rng.chance(40) {
            self.sp();
            self.word("Alias");
            self.out.push_str(self.rng.pick(&[" \"GetTickCount64\"", " \"#1\""]));
        }
        self.sp();
        self.out.push('(');
        self.params();
        self.out.push(')');
        if function && self.rng.chance(70) {
            self.sp();
            self.word("As");
            self.sp();
            self.type_name();
        }
    }
    /// A module-level declaration line (or block).
    fn decl(&mut self) {
        let visibility = self.rng.chance(50);
        match self.rng.below(11) {
            0..=2 => {
                let kw = self.rng.pick(&["Dim", "Private", "Public", "Global"]);
                self.word(kw);
                self.sp();
                let with_events = kw != "Dim" && self.rng.chance(15);
                self.var_list(with_events);
            }
            3 => {
                if visibility {
                    let w = self.rng.pick(&["Private", "Public"]);
                    self.word(w);
                    self.sp();
                }
                self.const_list();
            }
            4 => {
                if visibility {
                    let w = self.rng.pick(&["Private", "Public"]);
                    self.word(w);
                    self.sp();
                }
                self.type_block();
            }
            5 => {
                if visibility {
                    let w = self.rng.pick(&["Private", "Public"]);
                    self.word(w);
                    self.sp();
                }
                self.enum_block();
            }
            6 => {
                if visibility {
                    let w = self.rng.pick(&["Private", "Public"]);
                    self.word(w);
                    self.sp();
                }
                self.declare();
            }
            7 => {
                if visibility {
                    self.word("Public");
                    self.sp();
                }
                self.word("Event");
                self.sp();
                self.ident();
                if self.rng.chance(70) {
                    self.out.push('(');
                    self.params();
                    self.out.push(')');
                }
            }
            8 => {
                self.word("Implements");
                self.sp();
                self.ident();
                if self.rng.chance(30) {
                    self.out.push_str(".IFace");
                }
            }
            9 => {
                let def = self.rng.pick(&["DefInt", "DefStr", "DefLng", "DefObj"]);
                self.word(def);
                self.sp();
                self.out.push_str(self.rng.pick(&["A-Z", "A-Z, I", "S", "a - c"]));
            }
            _ => {
                let d = self.rng.pick(&DIRECTIVES);
                self.out.push_str(d);
                self.newline();
                return;
            }
        }
        self.nl();
    }
    fn header(&mut self) {
        const ATTRS: [&str; 4] = [
            "Attribute VB_Name = \"Module1\"",
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Exposed = True",
            "Attribute VB_Ext_KEY = \"a\" ,\"b\"",
        ];
        const OPTIONS: [&str; 7] = [
            "Option Explicit",
            "Option Base 1",
            "Option Base 0",
            "Option Compare Text",
            "Option Compare Binary",
            "Option Compare Database",
            "Option Private Module",
        ];
        for _ in 0..self.rng.range(0, 3) {
            let a = self.rng.pick(&ATTRS);
            self.words(a);
            self.newline();
        }
        for _ in 0..self.rng.range(0, 2) {
            let o = self.rng.pick(&OPTIONS);
            self.words(o);
            self.nl();
        }
    }
    /// A whole module: header, declarations, procedures, with blank lines and comments between.
    fn module(&mut self) -> String {
        self.out.clear();
        self.budget = self.rng.range(4, 60);
        self.header();
        for _ in 0..self.rng.range(0, 4) {
            self.decl();
        }
        for _ in 0..self.rng.range(1, 4) {
            match self.rng.below(6) {
                0 => self.newline(),
                1 => {
                    self.out.push('\'');
                    self.comment_text();
                    self.newline();
                }
                _ => {}
            }
            self.procedure();
        }
        // Sometimes the last line has no terminator.
        if self.rng.chance(20) {
            while self.out.ends_with(['\r', '\n', '\u{2028}']) {
                self.out.pop();
            }
        }
        std::mem::take(&mut self.out)
    }
}
// ---------------------------------------------------------------- mutations and noise
/// The largest char boundary at or before a random byte of `s`.
fn boundary(rng: &mut Rng, s: &str) -> usize {
    let mut at = rng.below(s.len() + 1);
    while !s.is_char_boundary(at) {
        at -= 1;
    }
    at
}
/// Byte offsets of every ASCII-case-insensitive occurrence of `needle` in `s`.
fn find_all(s: &str, needle: &str) -> Vec<usize> {
    let lower = s.to_ascii_lowercase();
    lower.match_indices(needle).map(|(i, _)| i).collect()
}
const CLOSERS: [&str; 11] =
    ["end if", "next", "end sub", "end function", "end property", "loop", "wend", "end with", "end select", "end type", "end enum"];
const PUNCT: [&str; 26] =
    ["(", ")", ",", ":", ":=", ".", "!", "#", "&", "_", " _\n", "\n", "=", "<", ">", "+", "-", "*", "\\", "^", "\"", "'", "[", "]", ";", "#1"];
fn random_token(rng: &mut Rng, keywords: &[&'static str]) -> &'static str {
    match rng.below(4) {
        0 | 1 => rng.pick(keywords),
        2 => rng.pick(&PUNCT),
        _ => rng.pick(&["x", "f", "1", "&H", "1.5E", "#1/2/2000#", "\"s\"", "[a b]", "Größe", "100", "Sub T()", "End Sub"]),
    }
}
/// Applies one mutation to `src`: closers dropped, words deleted, tokens inserted, lines
/// swapped or duplicated, the text cut short.
fn mutate(rng: &mut Rng, src: &mut String, keywords: &[&'static str]) {
    match rng.below(9) {
        0 | 1 => {
            let closer = rng.pick(&CLOSERS);
            let hits = find_all(src, closer);
            if !hits.is_empty() {
                let at = rng.pick(&hits);
                src.replace_range(at..at + closer.len(), "");
            }
        }
        2 => {
            let at = boundary(rng, src);
            let end = src[at..].find(char::is_whitespace).map_or(src.len(), |n| at + n);
            src.replace_range(at..end, "");
        }
        3 => {
            let at = boundary(rng, src);
            let token = random_token(rng, keywords);
            src.insert_str(at, token);
        }
        4 => {
            let at = boundary(rng, src);
            src.insert_str(at, rng.pick(&["(", ")", "((", "))", "(1", ")("]));
        }
        5 => {
            let at = boundary(rng, src);
            src.insert_str(at, rng.pick(&[":", " _", "_", " _\n", "\n", ": :", " _\r\n_"]));
        }
        6 => {
            let at = boundary(rng, src);
            src.truncate(at);
        }
        7 => {
            let mut lines: Vec<&str> = src.split_inclusive('\n').collect();
            if lines.len() >= 2 {
                let (a, b) = (rng.below(lines.len()), rng.below(lines.len()));
                if rng.chance(50) {
                    lines.swap(a, b);
                } else {
                    let line = lines[a];
                    lines.insert(b, line);
                }
                *src = lines.concat();
            }
        }
        _ => {
            let words: Vec<(usize, &str)> = src.split_whitespace().map(|w| (w.as_ptr() as usize - src.as_ptr() as usize, w)).collect();
            if words.len() >= 2 {
                // Swap two words, keeping the text between them.
                let (mut a, mut b) = (rng.below(words.len()), rng.below(words.len()));
                if a > b {
                    std::mem::swap(&mut a, &mut b);
                }
                if a != b {
                    let ((sa, wa), (sb, wb)) = (words[a], words[b]);
                    let swapped = format!("{}{wb}{}{wa}{}", &src[..sa], &src[sa + wa.len()..sb], &src[sb + wb.len()..]);
                    *src = swapped;
                }
            }
        }
    }
}
fn token_soup(rng: &mut Rng, keywords: &[&'static str]) -> String {
    let mut out = String::new();
    for _ in 0..rng.range(1, 400) {
        out.push_str(random_token(rng, keywords));
        out.push_str(rng.pick(&[" ", " ", "", "\n", "\t", " _\n", "\r\n"]));
    }
    out
}
fn noise_char(rng: &mut Rng) -> char {
    match rng.below(12) {
        0 | 1 => rng.char_in(0x20, 0x7E, '?'),
        2 => rng.pick(&['_', '"', '#', '[', ']', '\'', '&', ':', '(', ')', '.', '!', ' ', '\n', '\r']),
        3 => rng.char_in(0, 0x1F, '\0'),
        4 => rng.char_in(0xA0, 0xFF, '\u{a0}'),
        5 => rng.char_in(0x4E00, 0x9FFF, '数'),
        6 => rng.char_in(0x1F300, 0x1FAFF, '😀'),
        7 => rng.pick(&['\u{2028}', '\u{2029}', '\u{a0}', '\u{3000}', '\u{feff}', '\u{200b}']),
        _ => rng.char_in(0, 0x10_FFFF, '\u{fffd}'),
    }
}
fn unicode_noise(rng: &mut Rng) -> String {
    (0..rng.range(0, 500)).map(|_| noise_char(rng)).collect()
}
// ---------------------------------------------------------------- P6
/// Parser-test-style sources: the hand-written modules of `parser::tests`, which P6 also
/// round-trips (the `print::tests` sources go through the same check there).
const CORPUS: [&str; 16] = [
    "",
    "' just a comment\r\n",
    "Option Explicit\nDim a As Long\nPrivate Const B = 2\n",
    "Attribute VB_Name = \"Module1\"\r\nAttribute VB_Base = \"0{000}\"\r\nAttribute VB_Ext_KEY = \"a\" ,\"b\"\r\n",
    "Option Explicit\nOption Base 1\nOption Compare Text\noption compare database\nOption Private Module\n",
    "Dim a, b(1 To 10, 5) As Long\nPrivate WithEvents app As Excel.Application\nPublic s$, c As New Collection\nGlobal f As String * 10\nDim d() As Variant\n",
    "Private Const A As Long = 1, B = \"x\"\nPublic Type Pt\n  X As Long\n  Name As String * 8\n  Items(1 To 3) As Integer\nEnd Type\nEnum Color\n  Red = 1\n  Green\nEnd Enum\n",
    "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" Alias \"GetTickCount64\" () As LongLong\nDeclare Sub Sleep Lib \"kernel32\" (ByVal ms As Long)\nPublic Event Changed(ByVal value As Long)\nImplements IShape\nDefInt A-Z, I\n#If VBA7 Then\n",
    "Public Sub A()\nEnd Sub\nPrivate Static Function B$(Optional ByVal x As Long = 5, ByRef y(), ParamArray rest() As Variant) As String()\nEnd Function\nFriend Property Get C() As Long\nEnd Property\nProperty Let C(ByVal v As Long)\nEnd Property\nProperty Set D(ByVal o As Object)\nEnd Property\n",
    "Sub T()\nx = 1\nLet y = 2\nSet o = New Collection\nLSet s = t\nCall f(x)\nf x, y\nobj.Method a\nf(x)\nf (x), y\nRange(\"A1\").Value = 5\nDoEvents\nEnd Sub\n",
    "Sub T()\nIf a Then\n  x = 1\nElseIf b Then\n  x = 2\nElse\n  x = 3\nEnd If\nIf a Then x = 1 Else x = 2: y = 3\nIf a Then b: c\nIf a Then 10\nSelect Case n\nCase 1, 3 To 5\n  a\nCase Is > 10, < 0\n  b\nCase Else\n  c\nEnd Select\nEnd Sub\n",
    "Sub T()\nFor i = 1 To 10 Step 2\n  x = i\nNext i\nFor Each c In Range(\"A1:B2\").Cells\nNext\nFor i = 1 To 2\nFor j = 1 To 2\nNext j, i\nDo While x < 5\n  x = x + 1\nLoop\nDo\nLoop Until x\nWhile a\nWend\nWith ws\n  .Range(\"A1\").Value = 1\n  .Name = \"x\"\nEnd With\nEnd Sub\n",
    "Sub T()\nOn Error GoTo Handler\nOn Error GoTo 0\nOn Error GoTo -1\nOn Error Resume Next\nResume\nResume Next\nResume done\nGoTo done\nGoSub 100\nReturn\nExit Sub\nExit For\nOn n GoTo a, b\nHandler:\n100 x = 1\ndone: Stop\nEnd\nEnd Sub\n",
    "Sub T()\nDim a As Long, b\nStatic n%\nConst K = 3\nReDim Preserve arr(1 To n, 2) As String\nReDim Me.v(5)\nErase arr, v\nRaiseEvent Changed(1)\nDebug.Print a; b, c\nDebug.Print\nDebug.Print Tab(5); x;\nDebug.Assert x > 0\nEnd Sub\n",
    "Sub T()\nOpen \"C:\\a.txt\" For Binary Access Read Lock Write As #1 Len = 128\nPrint #1, a; b\nWrite #1, a, b\nInput #1, a, b\nLine Input #1, s\nGet #1, , rec\nPut #f, 3, rec\nSeek #1, 10\nLock #1, 1 To 5\nClose #1, #2\nClose\na = 1: b = 2\nc = 1 + _\n    2\nRange(\"A1\") _\n    .Value = 3\n#If VBA7 Then\nd = 4\n#End If\nEnd Sub\n",
    "Sub T()\nx = a + b * c ^ -d\nx = Not a = b And c\nx = a & b + c\nx = x Like \"*a\" Or y Is Nothing\nx = -a ^ 2\nx = a - b - c\nx = a \\ b Mod c * d\nx = a Imp b Eqv c Xor d Or e\nx = a <= b <> c >= d\nx = (a + b) * c\nx = TypeOf o Is Excel.Range And Not b\nx = AddressOf Module1.Callback\nx = Me.Name & rs!Field & rs![Other Name]\nx = f(a, , c:=1)\nx = f()(1)\nx = Array(True, Empty, Null, #1/2/2003#, &HFF, 1.5)\nx = Left$(s, 2) + UBound(a)\nMsgBox Prompt:=\"x\", Buttons:=vbOKOnly\nf , x\nEnd Sub\n",
];
/// How a P6 case went.
enum P6Outcome {
    /// The source parsed with no diagnostics and the round trip held.
    RoundTripped,
    /// The source has diagnostics, so P6 doesn't apply (the first one, for tuning).
    Rejected(ParseDiag),
}
/// The first line where two multi-line renderings differ, for failure messages.
fn first_difference(a: &str, b: &str) -> String {
    for (n, (x, y)) in a.lines().zip(b.lines()).enumerate() {
        if x != y {
            return format!("line {}:\n  before: {x}\n  after:  {y}", n + 1);
        }
    }
    format!("one is a prefix of the other ({} vs {} lines)", a.lines().count(), b.lines().count())
}
/// Checks P6 on `src`.
fn check_p6(src: &str) -> Result<P6Outcome, String> {
    let (tree, diags) = catch_unwind(AssertUnwindSafe(|| parse(src))).map_err(|_| "parse panicked".to_owned())?;
    if let Some(first) = diags.into_iter().next() {
        return Ok(P6Outcome::Rejected(first));
    }
    let tree = tree.ok_or("no tree and no diagnostics")?;
    let printed = catch_unwind(AssertUnwindSafe(|| print_module(&tree))).map_err(|_| "print_module panicked".to_owned())?;
    let (again, diags) = catch_unwind(AssertUnwindSafe(|| parse(&printed))).map_err(|_| "parse of the printed source panicked".to_owned())?;
    if !diags.is_empty() {
        return Err(format!("the printed source has diagnostics: {diags:#?}\n----- printed -----\n{printed}"));
    }
    let again = again.ok_or("no tree for the printed source")?;
    if !equivalent(&tree, &again) {
        let (a, b) = (format!("{:#?}", normalized(&tree)), format!("{:#?}", normalized(&again)));
        return Err(format!("the reparsed tree is not equivalent; first difference at {}\n----- printed -----\n{printed}", first_difference(&a, &b)));
    }
    let reprinted = print_module(&again);
    if reprinted != printed {
        return Err(format!(
            "printing is not idempotent; first difference at {}\n----- printed -----\n{printed}",
            first_difference(&printed, &reprinted)
        ));
    }
    Ok(P6Outcome::RoundTripped)
}
/// P6 cases: generated modules, and one in four mutated (checked only when the mutant still
/// parses cleanly, and not counted in the acceptance rate).
fn p6_input(case: u64, keywords: &[&'static str]) -> (String, bool) {
    let mut g = Gen::new(Rng::new(case, P6_STREAM));
    let mut src = g.module();
    let mutated = case % 4 == 3;
    if mutated {
        for _ in 0..g.rng.range(1, 2) {
            mutate(&mut g.rng, &mut src, keywords);
        }
    }
    (src, mutated)
}
/// P6: a generated module that parses with zero diagnostics prints to source that parses with
/// zero diagnostics to an equivalent tree, and printing that tree gives the same text again.
///
/// **Validates: Requirements 12.8**
#[test]
fn p6_print_then_parse_gives_an_equivalent_tree() {
    let (cases, replaying) = seed_list(P6_CASES_PER_SEED);
    let keywords: Vec<&'static str> = Keyword::all().map(Keyword::as_str).collect();
    let (mut generated, mut accepted, mut mutants_checked) = (0usize, 0usize, 0usize);
    let mut first_rejected: Option<(u64, ParseDiag)> = None;
    for &case in &cases {
        let (src, mutated) = p6_input(case, &keywords);
        generated += usize::from(!mutated);
        match check_p6(&src) {
            Ok(P6Outcome::RoundTripped) if mutated => mutants_checked += 1,
            Ok(P6Outcome::RoundTripped) => accepted += 1,
            Ok(P6Outcome::Rejected(diag)) => {
                if replaying {
                    eprintln!("seed {case}: P6 doesn't apply, the source has diagnostics: {diag:?}{}", shown(&src));
                }
                if !mutated && first_rejected.is_none() {
                    first_rejected = Some((case, diag));
                }
            }
            Err(why) => panic!(
                "P6 failed for seed {case}{}: {why}{}\n  {}",
                if mutated { " (mutated)" } else { "" },
                shown(&src),
                replay(case, "p6_print_then_parse_gives_an_equivalent_tree")
            ),
        }
    }
    if !replaying && generated >= 100 {
        // The generator has to produce mostly valid VBA, or P6 checks little.
        assert!(
            accepted * 100 >= generated * MIN_ACCEPT_PERCENT,
            "only {accepted} of {generated} generated modules parsed cleanly (need {MIN_ACCEPT_PERCENT}%); first rejected: {first_rejected:?}"
        );
        assert!(mutants_checked > 0, "no mutated module parsed cleanly");
    }
}
/// P6 on the parser-test-style corpus.
///
/// **Validates: Requirements 12.8**
#[test]
fn p6_corpus_round_trips() {
    for (n, src) in CORPUS.iter().enumerate() {
        match check_p6(src) {
            Ok(P6Outcome::RoundTripped) => {}
            Ok(P6Outcome::Rejected(diag)) => panic!("corpus source {n} has diagnostics: {diag:?}{}", shown(src)),
            Err(why) => panic!("P6 failed for corpus source {n}: {why}{}", shown(src)),
        }
    }
}
// ---------------------------------------------------------------- P7
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Class {
    Valid,
    Mutated,
    Soup,
    Unicode,
    Big,
}
const CLASSES: [Class; 5] = [Class::Valid, Class::Mutated, Class::Soup, Class::Unicode, Class::Big];
fn class_of(case: u64) -> Class {
    if case % BIG_EVERY == 5 {
        return Class::Big;
    }
    match case % 4 {
        0 => Class::Valid,
        1 => Class::Mutated,
        2 => Class::Soup,
        _ => Class::Unicode,
    }
}
fn mutated_module(g: &mut Gen, keywords: &[&'static str]) -> String {
    let mut src = g.module();
    for _ in 0..g.rng.range(1, 4) {
        mutate(&mut g.rng, &mut src, keywords);
    }
    src
}
/// 1 to 4 MiB of repeated valid and mutated modules.
fn big_input(g: &mut Gen, keywords: &[&'static str]) -> String {
    let pieces: Vec<String> = (0..8).map(|i| if i % 4 == 3 { mutated_module(g, keywords) } else { g.module() }).collect();
    let target = g.rng.range(BIG_LEN.0, BIG_LEN.1);
    let mut out = String::with_capacity(target + 16_384);
    while out.len() < target {
        let piece = &pieces[g.rng.below(pieces.len())];
        out.push_str(piece);
        out.push_str("\r\n");
    }
    out
}
fn p7_input(case: u64, class: Class, keywords: &[&'static str]) -> String {
    let mut g = Gen::new(Rng::new(case, P7_STREAM));
    match class {
        Class::Valid => g.module(),
        Class::Mutated => mutated_module(&mut g, keywords),
        Class::Soup => token_soup(&mut g.rng, keywords),
        Class::Unicode => unicode_noise(&mut g.rng),
        Class::Big => big_input(&mut g, keywords),
    }
}
/// What one parse returned, without touching the tree beyond `is_some`.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
struct Seen {
    tree: bool,
    syntax: usize,
    nesting: usize,
    too_large: usize,
}
/// Parses `src` (catching panics), checks the diagnostics are sane, and drops the tree.
fn parse_and_drop(src: &str) -> Result<Seen, String> {
    let (tree, diags) = catch_unwind(AssertUnwindSafe(|| parse(src))).map_err(|_| "parse panicked".to_owned())?;
    let mut seen = Seen { tree: tree.is_some(), ..Seen::default() };
    drop(tree);
    for d in &diags {
        if d.span.start > d.span.end || d.span.end > src.len() || d.span.line == 0 || d.span.col == 0 {
            return Err(format!("diagnostic span out of range for a {}-byte input: {d:?}", src.len()));
        }
        match d.kind {
            ParseDiagKind::Syntax => seen.syntax += 1,
            ParseDiagKind::NestingTooDeep => seen.nesting += 1,
            ParseDiagKind::SourceTooLarge => seen.too_large += 1,
        }
    }
    let too_large = src.len() > MAX_SOURCE;
    if seen.tree == too_large || (seen.too_large > 0) != too_large {
        return Err(format!("a {}-byte input gave tree: {}, {} SourceTooLarge diagnostics", src.len(), seen.tree, seen.too_large));
    }
    Ok(seen)
}
/// A P7 failure, reported from the small-stack thread.
struct Failure {
    case: u64,
    what: String,
    why: String,
    input: String,
}
/// Runs `work` on a thread with a [`SMALL_STACK`] stack and returns its result. A panic that
/// escapes `work` fails the test; a stack overflow aborts the test binary.
fn on_small_stack<T: Send + 'static>(work: impl FnOnce() -> T + Send + 'static) -> T {
    let worker = std::thread::Builder::new().name("p7-small-stack".into()).stack_size(SMALL_STACK).spawn(work);
    match worker.map(std::thread::JoinHandle::join) {
        Ok(Ok(result)) => result,
        Ok(Err(_)) => panic!("the 1 MiB parser thread panicked"),
        Err(e) => panic!("couldn't spawn the parser thread: {e}"),
    }
}
#[derive(Debug, Default)]
struct P7Stats {
    per_class: [usize; 5],
    clean: usize,
    valid_clean: usize,
    with_syntax: usize,
}
fn run_p7(cases: Vec<u64>) -> Result<P7Stats, Failure> {
    let trace = env_num("GRIDCRAFT_FUZZ_TRACE").is_some_and(|n| n > 0);
    let keywords: Vec<&'static str> = Keyword::all().map(Keyword::as_str).collect();
    let mut stats = P7Stats::default();
    for case in cases {
        let class = class_of(case);
        if trace {
            eprintln!("P7 case {case} ({class:?})");
        }
        let src = p7_input(case, class, &keywords);
        let seen = parse_and_drop(&src).map_err(|why| Failure { case, what: format!("{class:?}"), why, input: shown(&src) })?;
        stats.per_class[CLASSES.iter().position(|&c| c == class).unwrap_or(0)] += 1;
        let clean = seen.syntax + seen.nesting + seen.too_large == 0;
        stats.clean += usize::from(clean);
        stats.valid_clean += usize::from(clean && class == Class::Valid);
        stats.with_syntax += usize::from(seen.syntax > 0);
    }
    Ok(stats)
}
/// P7: `parse` returns a tree and/or diagnostics for arbitrary input, on a 1 MiB stack, without
/// panicking, overflowing or aborting; a default run covers at least 10,000 inputs.
///
/// **Validates: Requirements 12.9**
#[test]
fn p7_arbitrary_input_never_panics_on_a_small_stack() {
    let (cases, replaying) = seed_list(P7_CASES_PER_SEED);
    let ran = cases.len();
    let stats = match on_small_stack(move || run_p7(cases)) {
        Ok(stats) => stats,
        Err(f) => panic!(
            "P7 failed for seed {} ({}): {}{}\n  {}",
            f.case,
            f.what,
            f.why,
            f.input,
            replay(f.case, "p7_arbitrary_input_never_panics_on_a_small_stack")
        ),
    };
    if replaying {
        return;
    }
    if seed_count() >= DEFAULT_SEEDS {
        assert!(ran >= P7_MIN_CASES, "P7 ran {ran} inputs, need at least {P7_MIN_CASES}");
    }
    if ran as u64 >= BIG_EVERY {
        let valid = stats.per_class[0];
        assert!(stats.per_class.iter().all(|&n| n > 0), "P7 cases per class [valid, mutated, soup, unicode, big] = {:?}", stats.per_class);
        assert!(stats.valid_clean * 100 >= valid * MIN_ACCEPT_PERCENT, "only {} of {valid} valid inputs parsed cleanly", stats.valid_clean);
        assert!(stats.with_syntax > 0 && stats.clean < ran, "P7 never produced a syntax diagnostic: {stats:?}");
    }
}
// ---------------------------------------------------------------- P7: pathological shapes
/// What a pathological shape must give, beyond not panicking.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Expect {
    /// A tree and no diagnostics.
    Clean,
    /// At least one [`ParseDiagKind::NestingTooDeep`].
    TooDeep,
    /// No tree and one [`ParseDiagKind::SourceTooLarge`].
    TooLarge,
    /// Anything but a panic.
    Any,
}
const SHAPES: usize = 21;
/// `n` with up to 10% added, so rounds differ.
fn jitter(rng: &mut Rng, n: usize) -> usize {
    n + rng.below(n / 10 + 1)
}
fn in_sub(body: &str) -> String {
    format!("Sub T()\n{body}\nEnd Sub\n")
}
/// About `target` bytes of whole generated modules.
fn repeated_modules(rng: &mut Rng, target: usize) -> String {
    let mut g = Gen::new(Rng::new(rng.next(), P6_STREAM));
    let pieces: Vec<String> = (0..8).map(|_| g.module()).collect();
    let mut out = String::with_capacity(target + 16_384);
    while out.len() < target {
        out.push_str(&pieces[rng.below(pieces.len())]);
        out.push_str("\r\n");
    }
    out
}
/// Exactly `len` bytes: code lines among long comment lines (cheap to lex), padded with `x`s in
/// a comment.
fn sized_input(len: usize) -> String {
    let code = "Sub T()\r\n  x = &H1F& + 1.5E3 * Größe# & \"a\"\"b\" ' note\r\n  Print #1, #1/2/2000#, [a b]\r\nEnd Sub\r\n";
    let comment = format!("' {}\r\n", "Größe 数値 😀 ".repeat(40));
    let mut src = String::with_capacity(len + 1);
    while src.len() + code.len() + 17 * comment.len() + 1 < len {
        src.push_str(code);
        for _ in 0..16 {
            src.push_str(&comment);
        }
    }
    src.push('\'');
    while src.len() < len {
        src.push('x');
    }
    src
}
/// The pathological shape `shape` for one round.
fn pathological(shape: usize, rng: &mut Rng) -> (&'static str, String, Expect) {
    let deep = jitter(rng, 10_000);
    let long = jitter(rng, 200_000);
    match shape {
        0 => ("10k nested parentheses", in_sub(&format!("x = {}1{}", "(".repeat(deep), ")".repeat(deep))), Expect::TooDeep),
        1 => ("10k nested block Ifs", in_sub(&format!("{}{}", "If x Then\n".repeat(deep), "End If\n".repeat(deep))), Expect::TooDeep),
        2 => ("200k flat a + a", in_sub(&format!("x = {}", vec!["a"; long].join(" + "))), Expect::Clean),
        3 => ("200k member chain", in_sub(&format!("x = a{}", ".b".repeat(long))), Expect::Clean),
        4 => ("200k call chain", in_sub(&format!("x = f{}", "(1)".repeat(long))), Expect::Clean),
        5 => ("4 MiB of modules", repeated_modules(rng, 4 << 20), Expect::Any),
        6 => ("2 MiB string on one line", in_sub(&format!("x = \"{}\"", "a\"\"b".repeat(jitter(rng, 500_000)))), Expect::Clean),
        7 => ("100k arguments on one line", in_sub(&format!("f {}", vec!["a"; long / 2].join(", "))), Expect::Clean),
        8 => ("10k prefix minus", in_sub(&format!("x = {}1", "- ".repeat(deep))), Expect::TooDeep),
        9 => ("10k prefix Not", in_sub(&format!("x = {}a", "Not ".repeat(deep))), Expect::TooDeep),
        10 => ("10k nested calls", in_sub(&format!("x = {}1{}", "f(".repeat(deep), ")".repeat(deep))), Expect::TooDeep),
        11 => ("10k nested single-line Ifs", in_sub(&format!("{}x = 1", "If a Then ".repeat(deep))), Expect::TooDeep),
        12 => (
            "10k unterminated blocks",
            in_sub(&["With o\n", "For i = 1 To 2\n", "Do\n", "Select Case a\nCase 1\n", "While a\n"].concat().repeat(deep / 5)),
            Expect::TooDeep,
        ),
        13 => ("10k unclosed parentheses", in_sub(&format!("x = {}", "(".repeat(deep))), Expect::TooDeep),
        14 => ("10k right-nested a + (a + …)", in_sub(&format!("x = {}a{}", "a + (".repeat(deep), ")".repeat(deep))), Expect::TooDeep),
        15 => ("200k a - -a links", in_sub(&format!("x = a{}", " - -a".repeat(long))), Expect::Clean),
        16 => (
            "100k ElseIf and Case branches",
            in_sub(&format!("If a Then\n{}End If\nSelect Case a\n{}End Select", "ElseIf b Then\n".repeat(long / 4), "Case 1\n".repeat(long / 4))),
            Expect::Clean,
        ),
        17 => ("100k statements on one line", in_sub(&vec!["x = 1"; long / 2].join(": ")), Expect::Clean),
        18 => ("100k line continuations", in_sub(&format!("x = a{}", " + _\n a".repeat(long / 2))), Expect::Clean),
        19 => ("exactly MAX_SOURCE bytes", sized_input(MAX_SOURCE), Expect::Any),
        _ => ("MAX_SOURCE + 1 bytes", sized_input(MAX_SOURCE + 1), Expect::TooLarge),
    }
}
fn meets(seen: Seen, expect: Expect) -> bool {
    match expect {
        Expect::Clean => seen.tree && seen == Seen { tree: true, ..Seen::default() },
        Expect::TooDeep => seen.tree && seen.nesting > 0,
        Expect::TooLarge => !seen.tree && seen.too_large == 1,
        Expect::Any => true,
    }
}
fn run_pathological(cases: Vec<u64>) -> Result<usize, Failure> {
    for &case in &cases {
        let shape = (case % SHAPES as u64) as usize;
        let mut rng = Rng::new(case / SHAPES as u64, PATHOLOGICAL_STREAM);
        let (name, src, expect) = pathological(shape, &mut rng);
        let fail = |why: String| Failure { case, what: name.to_owned(), why, input: String::new() };
        let seen = parse_and_drop(&src).map_err(fail)?;
        if !meets(seen, expect) {
            return Err(Failure { case, what: name.to_owned(), why: format!("expected {expect:?}, got {seen:?}"), input: String::new() });
        }
    }
    Ok(cases.len())
}
/// P7 on pathological shapes, each once per round; a default run is one round, and
/// `GRIDCRAFT_FUZZ_SEEDS` adds a round per 64 seeds.
///
/// **Validates: Requirements 12.9**
#[test]
fn p7_pathological_shapes_never_panic_on_a_small_stack() {
    let cases = match env_num("GRIDCRAFT_FUZZ_SEED") {
        Some(one) => vec![one],
        None => (0..(seed_count() / DEFAULT_SEEDS).max(1) * SHAPES as u64).collect(),
    };
    let ran = match on_small_stack(move || run_pathological(cases)) {
        Ok(ran) => ran,
        Err(f) => panic!(
            "P7 failed for seed {} ({}): {}\n  {}",
            f.case,
            f.what,
            f.why,
            replay(f.case, "p7_pathological_shapes_never_panic_on_a_small_stack")
        ),
    };
    assert!(ran >= 1);
}
