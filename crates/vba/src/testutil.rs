//! Seeded test helpers shared by the property tests of this crate and of the engine (which reads
//! the same fixtures through the real compound file reader). Compiled only for tests and the
//! `test-support` feature, like [`crate::fixture`].
//!
//! - [`Rng`]: xorshift64, so every case replays from its seed.
//! - [`seed_list`]: the `GRIDCRAFT_FUZZ_SEEDS` (default [`DEFAULT_SEEDS`]) and
//!   `GRIDCRAFT_FUZZ_SEED` (replay one seed) knobs.
//! - [`guarded`]: runs a case under `catch_unwind`, so a panic reports its seed.
//! - Generators and checkers for the design properties P3 (fixture round trip, [`p3_fixture`],
//!   [`check_round_trip`]) and P4 (corruption hooks give diagnostics, [`p4_case`],
//!   [`check_corrupt`]), and fuzz inputs ([`fuzz_source`], [`fuzz_bin`], [`check_limits`]).
//!
//! Checkers return `Err(why)` instead of panicking, so the caller can add the seed and a replay
//! command to the message.
use std::panic::{AssertUnwindSafe, catch_unwind};

use crate::diag::DiagCode;
use crate::encoding::{self, CP_WINDOWS_1252};
use crate::fixture::{Corruption, Fixture, FixtureModule};
use crate::ovba;
use crate::project::{Limits, MemSource, StreamSource, VbaProject};
use crate::project_stream::ModuleKind;

/// Default `GRIDCRAFT_FUZZ_SEEDS`: small, so `cargo xtask ci` stays fast.
pub const DEFAULT_SEEDS: u64 = 64;
/// A decompressed chunk ([MS-OVBA] §2.4.1.3): truncating a container between chunks leaves a
/// valid container of whole 4,096-byte chunks.
///
/// [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/
const CHUNK: usize = 4096;
/// The longest P3 source.
pub const MAX_P3_SOURCE: usize = 65_536;
/// The most P3 modules.
pub const MAX_P3_MODULES: usize = 64;
/// Distinct streams so the generators don't share random sequences for one seed.
const P3_STREAM: u64 = 0x5033_5033_5033_5033;
const P4_STREAM: u64 = 0x5034_5034_5034_5034;
const FUZZ_STREAM: u64 = 0x4655_5A5A_4655_5A5A;
/// The compound file signature ([MS-CFB] §2.2).
///
/// [MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/
const CFB_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

// ---------------------------------------------------------------- deterministic randomness

/// xorshift64: tiny and deterministic.
#[derive(Debug, Clone)]
pub struct Rng(u64);

impl Rng {
    /// A generator for `seed` in the sequence family `stream`.
    #[must_use]
    pub fn new(seed: u64, stream: u64) -> Self {
        // Spread consecutive seeds over the state space; xorshift needs a nonzero state.
        let mixed = seed.wrapping_add(1).wrapping_mul(0x9E37_79B9_7F4A_7C15) ^ stream;
        let mut rng = Self(if mixed == 0 { 0x2545_F491_4F6C_DD1D } else { mixed });
        for _ in 0..4 {
            rng.next_u64();
        }
        rng
    }

    pub fn next_u64(&mut self) -> u64 {
        let mut x = self.0;
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        self.0 = x;
        x
    }

    pub fn byte(&mut self) -> u8 {
        self.next_u64().to_be_bytes()[0]
    }

    /// `0..n`, or 0 when `n` is 0.
    pub fn below(&mut self, n: usize) -> usize {
        let n64 = u64::try_from(n).unwrap_or(u64::MAX);
        if n64 == 0 { 0 } else { usize::try_from(self.next_u64() % n64).unwrap_or(0) }
    }

    /// `lo..=hi` (just `lo` when `hi < lo`).
    pub fn range(&mut self, lo: usize, hi: usize) -> usize {
        lo.saturating_add(self.below(hi.saturating_sub(lo).saturating_add(1)))
    }

    pub fn chance(&mut self, percent: u64) -> bool {
        self.next_u64() % 100 < percent
    }

    /// Appends `n` random bytes.
    pub fn fill(&mut self, out: &mut Vec<u8>, n: usize) {
        out.extend((0..n).map(|_| self.byte()));
    }

    /// One element of `items`, or `None` when it is empty.
    pub fn pick<'a, T>(&mut self, items: &'a [T]) -> Option<&'a T> {
        items.get(self.below(items.len()))
    }
}

/// An environment variable as a number.
#[must_use]
pub fn env_num(name: &str) -> Option<u64> {
    std::env::var(name).ok().and_then(|v| v.trim().parse().ok())
}

/// The seeds to run: `GRIDCRAFT_FUZZ_SEED` alone (and `true`, replaying), or
/// `0..GRIDCRAFT_FUZZ_SEEDS * per_seed`.
#[must_use]
pub fn seed_list(per_seed: u64) -> (Vec<u64>, bool) {
    seed_list_scaled(per_seed, 1)
}

/// Like [`seed_list`], with `GRIDCRAFT_FUZZ_SEEDS / divisor` (at least 1) seeds, for slower
/// properties that run a subset.
#[must_use]
pub fn seed_list_scaled(per_seed: u64, divisor: u64) -> (Vec<u64>, bool) {
    match env_num("GRIDCRAFT_FUZZ_SEED") {
        Some(one) => (vec![one], true),
        None => {
            let seeds = env_num("GRIDCRAFT_FUZZ_SEEDS").unwrap_or(DEFAULT_SEEDS).checked_div(divisor).unwrap_or(0).max(1);
            ((0..seeds.saturating_mul(per_seed)).collect(), false)
        }
    }
}

/// Runs `f`, turning a panic into `Err` with its message.
///
/// # Errors
///
/// The panic message when `f` panics.
pub fn guarded<T>(f: impl FnOnce() -> T) -> Result<T, String> {
    catch_unwind(AssertUnwindSafe(f)).map_err(|payload| {
        let why = payload.downcast_ref::<&str>().map(|s| (*s).to_owned()).or_else(|| payload.downcast_ref::<String>().cloned());
        format!("panicked: {}", why.unwrap_or_else(|| "(no message)".to_owned()))
    })
}

// ---------------------------------------------------------------- source text and modules

const WORDS: [&str; 20] = [
    "Sub ",
    "End Sub",
    "Function F(ByVal x As Long) As Long",
    "End Function",
    "Dim ",
    " As Long",
    " As String",
    "Range(\"A1\")",
    ".Value",
    " = ",
    "If ",
    " Then",
    "Else",
    "MsgBox ",
    "    ",
    "' note",
    "Option Explicit",
    "Attribute VB_Name = \"Module1\"",
    "x",
    "123",
];

/// cp1252's non-ASCII characters: 0x80–0xFF minus the five bytes it leaves undefined (which
/// decode to C1 controls).
fn cp1252_extra() -> Vec<char> {
    (0x80..=0xFFu8).filter_map(|b| encoding::decode(CP_WINDOWS_1252, &[b]).0.chars().next()).filter(|c| !c.is_control()).collect()
}

/// A printable ASCII character or, rarely, a tab.
fn printable(rng: &mut Rng) -> char {
    if rng.chance(3) { '\t' } else { char::from(0x20u8.saturating_add(u8::try_from(rng.below(95)).unwrap_or(0))) }
}

/// Source text of exactly `len` characters, each one byte in Windows-1252 (so `len` is also the
/// encoded length), mixing VBA-like words, CRLF lines, cp1252 non-ASCII characters, long runs
/// and pseudo-random printable text. No NUL: VBA source holds none.
#[must_use]
pub fn source(rng: &mut Rng, len: usize) -> String {
    let extra = cp1252_extra();
    let mut out: Vec<char> = Vec::with_capacity(len);
    while out.len() < len {
        let room = len.saturating_sub(out.len());
        let n = rng.range(1, room.min(3000));
        let mut seg: Vec<char> = Vec::with_capacity(n.saturating_add(64));
        match rng.below(6) {
            // VBA-like words.
            0 => {
                while seg.len() < n {
                    seg.extend(rng.pick(&WORDS).copied().unwrap_or("x").chars());
                }
            }
            // CRLF lines of words.
            1 => {
                while seg.len() < n {
                    for _ in 0..rng.range(1, 6) {
                        seg.extend(rng.pick(&WORDS).copied().unwrap_or("x").chars());
                    }
                    seg.extend(['\r', '\n']);
                }
            }
            // cp1252 non-ASCII characters among letters.
            2 => {
                while seg.len() < n {
                    let c = if rng.chance(40) {
                        rng.pick(&extra).copied().unwrap_or('é')
                    } else {
                        char::from(b'a'.saturating_add(u8::try_from(rng.below(26)).unwrap_or(0)))
                    };
                    seg.push(c);
                }
            }
            // A long run of one character.
            3 => {
                let c = if rng.chance(30) { rng.pick(&extra).copied().unwrap_or('é') } else { printable(rng) };
                seg.resize(n, c);
            }
            // Pseudo-random printable text.
            4 => seg.extend((0..n).map(|_| printable(rng))),
            // Lines of mixed text.
            _ => {
                while seg.len() < n {
                    seg.extend((0..rng.range(0, 80)).map(|_| if rng.chance(10) { rng.pick(&extra).copied().unwrap_or('é') } else { printable(rng) }));
                    seg.extend(['\r', '\n']);
                }
            }
        }
        seg.truncate(n);
        out.extend(seg);
    }
    out.into_iter().collect()
}

/// A source length: mostly short, sometimes long, now and then on a chunk boundary.
fn source_len(rng: &mut Rng, max: usize) -> usize {
    const BOUNDARIES: [usize; 8] = [0, 1, 4095, 4096, 4097, 8192, 65_535, MAX_P3_SOURCE];
    let len = match rng.below(100) {
        0..=4 => BOUNDARIES.get(rng.below(BOUNDARIES.len())).copied().unwrap_or(0),
        5..=49 => rng.range(0, 64),
        50..=77 => rng.range(65, 1024),
        78..=93 => rng.range(1025, 8192),
        94..=98 => rng.range(8193, 32_768),
        _ => rng.range(32_769, MAX_P3_SOURCE),
    };
    len.min(max)
}

/// A VBA identifier of 1 to 31 ASCII characters: a letter, then letters, digits and `_`.
fn identifier(rng: &mut Rng) -> String {
    const REST: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_";
    let len = match rng.below(10) {
        0 => 1,
        1 => 31,
        _ => rng.range(1, 31),
    };
    let mut name = String::with_capacity(len);
    name.push(char::from(REST.get(rng.below(52)).copied().unwrap_or(b'M')));
    while name.len() < len {
        name.push(char::from(REST.get(rng.below(REST.len())).copied().unwrap_or(b'_')));
    }
    name
}

fn kind(rng: &mut Rng) -> ModuleKind {
    match rng.below(4) {
        0 => ModuleKind::Standard,
        1 => ModuleKind::Class,
        2 => ModuleKind::Form,
        _ => ModuleKind::Document,
    }
}

/// Up to `count` modules with names unique case-insensitively (as VBA requires; duplicates are
/// dropped and redrawn) and never `dir`, which the fixture would store over the `VBA/dir` stream.
fn modules(rng: &mut Rng, count: usize, mut len: impl FnMut(&mut Rng) -> usize) -> Vec<FixtureModule> {
    let mut out: Vec<FixtureModule> = Vec::with_capacity(count);
    for _ in 0..count.saturating_mul(20) {
        if out.len() >= count {
            break;
        }
        let name = identifier(rng);
        if name.eq_ignore_ascii_case("dir") || out.iter().any(|m| m.name.eq_ignore_ascii_case(&name)) {
            continue;
        }
        let k = kind(rng);
        let n = len(rng);
        out.push(FixtureModule::new(&name, k, &source(rng, n)));
    }
    out
}

// ---------------------------------------------------------------- P3: round trip

/// A P3 fixture: 0 to 64 modules of any kind with 0 to 65,536-byte sources. Most cases are
/// small; seed classes force the edges (seed mod 16): 0 has no modules, 1 has 64 modules, 2 has
/// a 65,536-byte module, 3 ends a module in 3,800+ incompressible bytes (a raw, padded last
/// chunk).
#[must_use]
pub fn p3_fixture(seed: u64) -> Fixture {
    let mut rng = Rng::new(seed, P3_STREAM);
    let mods = match seed % 16 {
        0 => Vec::new(),
        1 => modules(&mut rng, MAX_P3_MODULES, |r| source_len(r, 256)),
        2 => {
            let mut m = modules(&mut rng, 4, |r| source_len(r, 128));
            if let Some(first) = m.first_mut() {
                first.source = source(&mut rng, MAX_P3_SOURCE);
            }
            m
        }
        3 => {
            let mut m = modules(&mut rng, 2, |r| source_len(r, 128));
            if let Some(first) = m.first_mut() {
                let whole = rng.below(3).saturating_mul(CHUNK);
                let mut text = source(&mut rng, whole);
                text.extend((0..rng.range(3800, CHUNK - 1)).map(|_| printable(&mut rng)));
                first.source = text;
            }
            m
        }
        _ => {
            let count = match rng.below(10) {
                0..=2 => rng.range(0, 2),
                3..=7 => rng.range(3, 12),
                _ => rng.range(13, MAX_P3_MODULES),
            };
            // Many modules stay short, so one case's cost stays bounded.
            let max = if count > 12 { 1024 } else { MAX_P3_SOURCE };
            modules(&mut rng, count, |r| source_len(r, max))
        }
    };
    Fixture::new(mods)
}

/// Each module's (name, kind, source).
#[must_use]
pub fn summary(p: &VbaProject) -> Vec<(String, ModuleKind, String)> {
    p.modules.iter().map(|m| (m.name.clone(), m.kind, m.source.clone())).collect()
}

/// Every diagnostic code, project-level first, then each module's.
#[must_use]
pub fn codes(p: &VbaProject) -> Vec<DiagCode> {
    p.diagnostics.iter().chain(p.modules.iter().flat_map(|m| &m.diagnostics)).map(|d| d.code).collect()
}

/// Where two strings first differ, for messages.
fn first_difference(a: &str, b: &str) -> String {
    let at = a.chars().zip(b.chars()).position(|(x, y)| x != y).unwrap_or_else(|| a.chars().count().min(b.chars().count()));
    let tail = |s: &str| s.chars().skip(at).take(12).collect::<String>();
    format!("{} vs {} chars, first difference at char {at}: {:?} vs {:?}", a.chars().count(), b.chars().count(), tail(a), tail(b))
}

/// P3: `p` lists the fixture's modules with the same names, kinds and sources, in order, and has
/// no diagnostics.
///
/// # Errors
///
/// What differs.
pub fn check_round_trip(f: &Fixture, p: &VbaProject) -> Result<(), String> {
    check_limits(p, &Limits::default())?;
    if p.modules.len() != f.modules.len() {
        return Err(format!("{} modules read, {} built; diagnostics {:?}", p.modules.len(), f.modules.len(), codes(p)));
    }
    for (i, (got, want)) in p.modules.iter().zip(&f.modules).enumerate() {
        if got.name != want.name || got.kind != want.kind {
            return Err(format!("module {i}: read {:?} {:?}, built {:?} {:?}", got.name, got.kind, want.name, want.kind));
        }
        if got.source != want.source {
            return Err(format!("module {i} ({}): source differs, {}", want.name, first_difference(&got.source, &want.source)));
        }
    }
    let diags = codes(p);
    if !diags.is_empty() {
        return Err(format!("unexpected diagnostics {diags:?}"));
    }
    Ok(())
}

// ---------------------------------------------------------------- P4: corruption hooks

/// A P4 case: a fixture with corruptions, and its streams before them (for stream lengths).
#[derive(Debug, Clone)]
pub struct P4Case {
    pub fixture: Fixture,
    pub clean: MemSource,
}

fn strings(p: &[&str]) -> Vec<String> {
    p.iter().map(|s| (*s).to_owned()).collect()
}

/// `s` with each ASCII letter's case flipped at random, sometimes: names match
/// case-insensitively.
fn recase(rng: &mut Rng, s: &str) -> String {
    if !rng.chance(30) {
        return s.to_owned();
    }
    s.chars().map(|c| if rng.chance(50) { if c.is_ascii_lowercase() { c.to_ascii_uppercase() } else { c.to_ascii_lowercase() } } else { c }).collect()
}

fn stream_of(src: &MemSource, path: &[String]) -> Option<Vec<u8>> {
    let p: Vec<&str> = path.iter().map(String::as_str).collect();
    src.stream(&p).ok().flatten()
}

/// The offsets at which an OVBA container's chunks end (after the signature byte, then after
/// each chunk that fits in `container`).
fn chunk_ends(container: &[u8]) -> Vec<usize> {
    let mut ends = Vec::new();
    let mut at = 1usize;
    while at <= container.len() {
        ends.push(at);
        let Some(&[lo, hi]) = container.get(at..at.saturating_add(2)) else { break };
        at = at.saturating_add(usize::from(u16::from_le_bytes([lo, hi]) & 0x0FFF)).saturating_add(3);
    }
    ends
}

/// A truncation length for a stream of `len` bytes: a few bytes, anywhere, a chunk boundary, or
/// at least `len` (no change).
fn keep(rng: &mut Rng, stream: &[u8]) -> usize {
    let len = stream.len();
    match rng.below(10) {
        0..=2 => rng.range(0, len.min(8)),
        3..=6 => rng.below(len.max(1)),
        7..=8 => rng.pick(&chunk_ends(stream)).copied().unwrap_or(0),
        _ => len.saturating_add(rng.below(4096)),
    }
}

/// One random corruption of `f`.
fn corruption(rng: &mut Rng, f: &Fixture, clean: &MemSource) -> Corruption {
    let module = |rng: &mut Rng| match rng.pick(&f.modules) {
        Some(m) if rng.chance(92) => recase(rng, &m.name),
        _ => "Nope".to_owned(),
    };
    let path = match rng.below(20) {
        0..=11 => {
            let name = module(rng);
            vec!["VBA".to_owned(), name]
        }
        12..=14 => strings(&[if rng.chance(50) { "VBA" } else { "vba" }, if rng.chance(50) { "dir" } else { "DIR" }]),
        15..=16 => strings(&["PROJECT"]),
        17 => strings(&["VBA", "_VBA_PROJECT"]),
        _ => strings(&["VBA", "Nope"]),
    };
    match rng.below(4) {
        0 => {
            let bytes = stream_of(clean, &path).unwrap_or_default();
            Corruption::TruncateStream { keep: keep(rng, &bytes), path }
        }
        1 => Corruption::OffsetPastEnd { module: module(rng) },
        2 => Corruption::OversizedLength { path },
        _ => Corruption::CyclicChain { path },
    }
}

/// A P4 case: 0 to 8 modules (sources up to 12,000 bytes, so some streams have several chunks
/// and some pass the 4,096-byte mini stream cutoff) with one to three random corruptions.
#[must_use]
pub fn p4_case(seed: u64) -> P4Case {
    let mut rng = Rng::new(seed, P4_STREAM);
    let count = match rng.below(10) {
        0 => 0,
        1..=6 => rng.range(1, 3),
        _ => rng.range(4, 8),
    };
    let mods = modules(&mut rng, count, |r| match r.below(10) {
        0..=5 => r.range(0, 512),
        6..=8 => r.range(513, CHUNK),
        _ => r.range(CHUNK + 1, 12_000),
    });
    let mut fixture = Fixture::new(mods);
    let clean = fixture.to_streams();
    let n = match rng.below(20) {
        0..=11 => 1,
        12..=16 => 2,
        _ => 3,
    };
    for _ in 0..n {
        let c = corruption(&mut rng, &fixture, &clean);
        fixture.corrupt.push(c);
    }
    P4Case { fixture, clean }
}

fn same_path(a: &[String], b: &[&str]) -> bool {
    a.len() == b.len() && a.iter().zip(b).all(|(x, y)| x.eq_ignore_ascii_case(y))
}

/// What a case's corruptions actually damage.
#[derive(Debug, Default)]
struct Damage {
    /// `VBA/dir` is cut short or unreadable.
    dir: bool,
    /// `PROJECT` is cut short or unreadable.
    project: bool,
    /// Per fixture module: `None` when untouched, else whether anything other than a
    /// truncation hit it.
    modules: Vec<Option<bool>>,
}

fn damage(case: &P4Case) -> Damage {
    let f = &case.fixture;
    let mut d = Damage { modules: vec![None; f.modules.len()], ..Damage::default() };
    // OffsetPastEnd rewrites `dir` before truncation, which can change its compressed length.
    let dir_len = if f.corrupt.iter().any(|c| matches!(c, Corruption::OffsetPastEnd { .. })) {
        let rewritten =
            Fixture { corrupt: f.corrupt.iter().filter(|c| matches!(c, Corruption::OffsetPastEnd { .. })).cloned().collect(), ..f.clone() };
        stream_of(&rewritten.to_streams(), &strings(&["VBA", "dir"])).map_or(0, |b| b.len())
    } else {
        stream_of(&case.clean, &strings(&["VBA", "dir"])).map_or(0, |b| b.len())
    };
    for c in &f.corrupt {
        let (path, truncation) = match c {
            Corruption::OffsetPastEnd { module } => {
                if let Some(slot) = f.modules.iter().position(|m| m.name.eq_ignore_ascii_case(module)).and_then(|i| d.modules.get_mut(i)) {
                    *slot = Some(true);
                }
                continue;
            }
            Corruption::TruncateStream { path, keep } => {
                let len = if same_path(path, &["VBA", "dir"]) { dir_len } else { stream_of(&case.clean, path).map_or(0, |b| b.len()) };
                if *keep >= len {
                    continue;
                }
                (path, true)
            }
            Corruption::OversizedLength { path } | Corruption::CyclicChain { path } => (path, false),
        };
        if same_path(path, &["VBA", "dir"]) {
            d.dir = true;
        } else if same_path(path, &["PROJECT"]) {
            d.project = true;
        } else if let Some(i) = f.modules.iter().position(|m| same_path(path, &["VBA", &m.name])) {
            let other = d.modules.get(i).copied().flatten().unwrap_or(false);
            if let Some(slot) = d.modules.get_mut(i) {
                *slot = Some(other || !truncation);
            }
        }
    }
    d
}

/// Whether `got` is `want` cut after whole 4,096-byte chunks (what truncating a module stream
/// between chunks leaves: a valid, shorter container). Sources are one byte per character.
fn whole_chunk_prefix(want: &str, got: &str) -> bool {
    want.starts_with(got) && got.chars().count().is_multiple_of(CHUNK)
}

/// P4: the result stays within the default limits; if `dir` was damaged there is a `dir`
/// diagnostic; otherwise every module is listed by name, each damaged module has a diagnostic
/// naming it (and empty source) or, when only truncated, a whole-chunk prefix of its source,
/// every other module keeps its exact source, and kinds and diagnostics change only when
/// `PROJECT` was damaged.
///
/// # Errors
///
/// What broke the property.
pub fn check_corrupt(case: &P4Case, p: &VbaProject) -> Result<(), String> {
    check_limits(p, &Limits::default())?;
    let f = &case.fixture;
    let d = damage(case);
    if d.dir {
        let reported = p.diagnostics.iter().any(|x| x.module.is_none() && matches!(x.code, DiagCode::DirCorrupt | DiagCode::DirMissing));
        if !reported || !p.modules.is_empty() {
            return Err(format!("dir damaged, but {} modules and diagnostics {:?}", p.modules.len(), codes(p)));
        }
        return Ok(());
    }
    if p.modules.len() != f.modules.len() {
        return Err(format!("{} modules read, {} built; diagnostics {:?}", p.modules.len(), f.modules.len(), codes(p)));
    }
    for (i, (got, want)) in p.modules.iter().zip(&f.modules).enumerate() {
        if got.name != want.name {
            return Err(format!("module {i}: read {:?}, built {:?}", got.name, want.name));
        }
        if !d.project && got.kind != want.kind {
            return Err(format!("module {}: kind {:?}, built {:?}, with PROJECT intact", want.name, got.kind, want.kind));
        }
        let failure = got.diagnostics.iter().find(|x| x.code != DiagCode::KindFromModuleType);
        if let Some(x) = failure
            && x.module.as_deref() != Some(want.name.as_str())
        {
            return Err(format!("module {}: diagnostic {:?} names {:?}", want.name, x.code, x.module));
        }
        if !d.project && got.diagnostics.iter().any(|x| x.code == DiagCode::KindFromModuleType) {
            return Err(format!("module {}: kindFromModuleType with PROJECT intact", want.name));
        }
        match d.modules.get(i).copied().flatten() {
            Some(other) => match failure {
                Some(_) if !got.source.is_empty() => {
                    return Err(format!("module {}: a failure diagnostic but {} chars of source", want.name, got.source.chars().count()));
                }
                Some(_) => {}
                None if !other && whole_chunk_prefix(&want.source, &got.source) => {}
                None => {
                    return Err(format!(
                        "module {} was damaged but has no diagnostic naming it; source {}",
                        want.name,
                        first_difference(&got.source, &want.source)
                    ));
                }
            },
            None => {
                if let Some(x) = failure {
                    return Err(format!("untouched module {} has diagnostic {:?}", want.name, x.code));
                }
                if got.source != want.source {
                    return Err(format!("untouched module {}: source differs, {}", want.name, first_difference(&got.source, &want.source)));
                }
            }
        }
    }
    let allowed = |code: DiagCode| d.project && code == DiagCode::ProjectStreamMissing;
    if let Some(x) = p.diagnostics.iter().find(|x| !allowed(x.code)) {
        return Err(format!("unexpected project diagnostic {:?}: {}", x.code, x.message));
    }
    Ok(())
}

// ---------------------------------------------------------------- limits and fuzz inputs

/// The result stays within `limits`: at most `max_modules` modules and `max_total_source` bytes
/// of source in total (Requirement 5).
///
/// # Errors
///
/// Which limit was passed.
pub fn check_limits(p: &VbaProject, limits: &Limits) -> Result<(), String> {
    if p.modules.len() > limits.max_modules {
        return Err(format!("{} modules, more than the limit {}", p.modules.len(), limits.max_modules));
    }
    let total = p.modules.iter().try_fold(0usize, |t, m| t.checked_add(m.source.len()));
    match total {
        Some(t) if t <= limits.max_total_source => Ok(()),
        _ => Err(format!("{total:?} bytes of source, more than the limit {}", limits.max_total_source)),
    }
}

/// Limits to fuzz with: the defaults, or small ones that real inputs reach.
#[must_use]
pub fn fuzz_limits(seed: u64) -> [Limits; 2] {
    let mut rng = Rng::new(seed, FUZZ_STREAM ^ 0x4C);
    [Limits::default(), Limits { max_modules: rng.range(0, 4), max_total_source: rng.range(0, 2000), max_stream: rng.range(0, 9000) }]
}

/// A small valid fixture for fuzz inputs to start from.
fn fuzz_base(rng: &mut Rng) -> Fixture {
    let count = rng.range(0, 4);
    Fixture::new(modules(rng, count, |r| r.range(0, 600)))
}

/// `bytes` with one to six random edits: flipped, inserted or deleted bytes, a 32-bit field
/// overwritten with a large value, a cut, or a duplicated slice.
fn mutate(rng: &mut Rng, bytes: &mut Vec<u8>) {
    for _ in 0..rng.range(1, 6) {
        let at = rng.below(bytes.len().saturating_add(1));
        match rng.below(6) {
            0 => {
                let flip = rng.byte() | 1;
                if let Some(b) = bytes.get_mut(at) {
                    *b ^= flip;
                }
            }
            1 => bytes.insert(at.min(bytes.len()), rng.byte()),
            2 => {
                if at < bytes.len() {
                    bytes.remove(at);
                }
            }
            3 => {
                let value = match rng.below(4) {
                    0 => u32::MAX,
                    1 => 0x7FFF_FFFF,
                    2 => u32::try_from(rng.range(0, 70_000)).unwrap_or(0),
                    _ => u32::try_from(rng.next_u64() >> 32).unwrap_or(0),
                };
                for (i, b) in value.to_le_bytes().into_iter().enumerate() {
                    if let Some(slot) = bytes.get_mut(at.saturating_add(i)) {
                        *slot = b;
                    }
                }
            }
            4 => bytes.truncate(at),
            _ => {
                let end = at.saturating_add(rng.range(1, 64)).min(bytes.len());
                let piece: Vec<u8> = bytes.get(at..end).map(<[u8]>::to_vec).unwrap_or_default();
                let to = rng.below(bytes.len().saturating_add(1)).min(bytes.len());
                let tail = bytes.split_off(to);
                bytes.extend(piece);
                bytes.extend(tail);
            }
        }
    }
}

/// A fuzz input for `read_project` (seed mod 4): 0, random streams; 1, a valid project whose
/// decompressed `dir` stream is mutated and recompressed (reaching the record parser); 2, a
/// valid project whose compressed `dir` stream is mutated; 3, a valid project whose module and
/// `PROJECT` streams are mutated.
#[must_use]
pub fn fuzz_source(seed: u64) -> MemSource {
    let mut rng = Rng::new(seed, FUZZ_STREAM);
    let dir = ["VBA", "dir"];
    match seed % 4 {
        0 => {
            let mut src = MemSource::new();
            let mut d = Vec::new();
            let n = rng.range(0, 600);
            rng.fill(&mut d, n);
            if rng.chance(50)
                && let Some(first) = d.first_mut()
            {
                *first = 0x01;
            }
            src.insert(&dir, d);
            let mut text = Vec::new();
            for _ in 0..rng.range(0, 6) {
                text.extend_from_slice(
                    rng.pick(&["Module=", "Class=", "BaseClass=", "Document=", "[Workspace]", "ID="]).copied().unwrap_or("").as_bytes(),
                );
                let n = rng.range(0, 40);
                rng.fill(&mut text, n);
                text.extend_from_slice(b"\r\n");
            }
            src.insert(&["PROJECT"], text);
            for i in 0..rng.range(0, 3) {
                let mut m = vec![0x01];
                let n = rng.range(0, 300);
                rng.fill(&mut m, n);
                src.insert(&["VBA", &format!("Module{i}")], m);
            }
            src
        }
        1 => {
            let f = fuzz_base(&mut rng);
            let mut bytes = f.dir_fixture().to_bytes();
            mutate(&mut rng, &mut bytes);
            let mut src = f.to_streams();
            src.insert(&dir, ovba::compress(&bytes));
            src
        }
        2 => {
            let f = fuzz_base(&mut rng);
            let mut src = f.to_streams();
            let mut bytes = src.stream(&dir).ok().flatten().unwrap_or_default();
            mutate(&mut rng, &mut bytes);
            src.insert(&dir, bytes);
            src
        }
        _ => {
            let f = fuzz_base(&mut rng);
            let mut src = f.to_streams();
            let mut paths: Vec<Vec<String>> = f.modules.iter().map(|m| strings(&["VBA", &m.name])).collect();
            paths.push(strings(&["PROJECT"]));
            for path in paths {
                if rng.chance(60) {
                    let mut bytes = stream_of(&src, &path).unwrap_or_default();
                    mutate(&mut rng, &mut bytes);
                    let p: Vec<&str> = path.iter().map(String::as_str).collect();
                    src.insert(&p, bytes);
                }
            }
            src
        }
    }
}

/// A fuzz input for reading a whole `vbaProject.bin` (seed mod 4): 0, random bytes, half of
/// them after a compound file signature; 1, a valid file with bytes mutated anywhere; 2, a
/// valid file mutated in its first 2,048 bytes (header, FAT and directory); 3, a valid file cut
/// short or with random bytes appended, and 32-bit fields overwritten.
#[must_use]
pub fn fuzz_bin(seed: u64) -> Vec<u8> {
    let mut rng = Rng::new(seed, FUZZ_STREAM ^ 0xB1);
    if seed.is_multiple_of(4) {
        let mut out = if rng.chance(50) { CFB_SIGNATURE.to_vec() } else { Vec::new() };
        let n = rng.range(0, 4096);
        rng.fill(&mut out, n);
        return out;
    }
    let f = fuzz_base(&mut rng);
    let mut bin = f.to_vba_project_bin().unwrap_or_default();
    match seed % 4 {
        1 => mutate(&mut rng, &mut bin),
        2 => {
            let tail = bin.split_off(bin.len().min(2048));
            mutate(&mut rng, &mut bin);
            bin.extend(tail);
        }
        _ => {
            if rng.chance(50) {
                let cut = rng.below(bin.len().saturating_add(1));
                bin.truncate(cut);
            } else {
                let n = rng.range(1, 2048);
                rng.fill(&mut bin, n);
            }
            for _ in 0..rng.range(0, 4) {
                let at = rng.below(bin.len() / 4).saturating_mul(4);
                for (i, b) in u32::MAX.to_le_bytes().into_iter().enumerate() {
                    if let Some(slot) = bin.get_mut(at.saturating_add(i)) {
                        *slot = b;
                    }
                }
            }
        }
    }
    bin
}
