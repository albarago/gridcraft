//! Property tests for [`read_project`] over Fixture_Builder projects (design properties P3 and
//! P4) and a seeded fuzz, driven by the seeded generators of `testutil` so every failure replays.
//! The engine runs the same properties through the real compound file reader
//! (`vba_source::props`).
//!
//! - P3 (round trip): 0 to 64 modules, names of 1 to 31 identifier characters (unique
//!   case-insensitively), any kind, sources of 0 to 65,536 bytes, read back with the same names,
//!   kinds and sources and no diagnostics.
//! - P4 (corruption hooks): one to three random [`Corruption`](crate::fixture::Corruption)s
//!   never panic and give a diagnostic naming each damaged module (or a `dir` diagnostic),
//!   while untouched modules keep their exact source.
//! - Fuzz: random streams, and valid projects with mutated `dir` (compressed and decompressed),
//!   module and `PROJECT` streams, never panic and stay within the [`Limits`], both the defaults
//!   and small ones.
//!
//! Knobs: `GRIDCRAFT_FUZZ_SEEDS` (default 64) gives P3 that many fixtures, P4 2 and the fuzz 16
//! inputs per seed; `GRIDCRAFT_FUZZ_SEED=17` replays one case. Soak with
//! `GRIDCRAFT_FUZZ_SEEDS=300 cargo test --release -p gridcraft-vba --lib project::props::`.
use super::{Limits, read_project};
use crate::testutil::{
    MAX_P3_MODULES, MAX_P3_SOURCE, P4Case, check_corrupt, check_limits, check_round_trip, fuzz_limits, fuzz_source, guarded, p3_fixture, p4_case,
    seed_list,
};

/// P4 cases per seed.
const P4_PER_SEED: u64 = 2;
/// Fuzz inputs per seed.
const FUZZ_PER_SEED: u64 = 16;

fn replay(seed: u64, test: &str) -> String {
    format!("replay: GRIDCRAFT_FUZZ_SEED={seed} cargo test -p gridcraft-vba --lib project::props::{test}")
}

/// P3: building a fixture and reading its streams gives back the same modules.
///
/// **Validates: Requirements 10.3**
#[test]
fn p3_round_trip() {
    let (seeds, replaying) = seed_list(1);
    let (mut most_modules, mut longest) = (0, 0);
    for &seed in &seeds {
        let checked = guarded(|| {
            let f = p3_fixture(seed);
            let p = read_project(&f.to_streams(), &Limits::default());
            // Sources are one Windows-1252 byte per character.
            let longest = f.modules.iter().map(|m| m.source.chars().count()).max().unwrap_or(0);
            check_round_trip(&f, &p).map(|()| (f.modules.len(), longest))
        });
        match checked {
            Ok(Ok((count, len))) => {
                most_modules = most_modules.max(count);
                longest = longest.max(len);
            }
            Ok(Err(why)) | Err(why) => panic!("P3 failed for seed {seed}: {why}\n  {}", replay(seed, "p3_round_trip")),
        }
    }
    // The seed classes must reach both edges.
    if !replaying && seeds.len() >= 16 {
        assert_eq!(most_modules, MAX_P3_MODULES, "P3 never built {MAX_P3_MODULES} modules");
        assert_eq!(longest, MAX_P3_SOURCE, "P3 never built a {MAX_P3_SOURCE}-byte source");
    }
}

/// P4: corrupted fixtures give diagnostics, never panics, and leave other modules alone.
///
/// **Validates: Requirements 10.5**
#[test]
fn p4_corruptions_give_diagnostics() {
    let (seeds, replaying) = seed_list(P4_PER_SEED);
    let (mut dir_failures, mut module_failures) = (0, 0);
    for &seed in &seeds {
        let checked = guarded(|| {
            let case: P4Case = p4_case(seed);
            let p = read_project(&case.fixture.to_streams(), &Limits::default());
            check_corrupt(&case, &p).map(|()| p)
        });
        match checked {
            Ok(Ok(p)) => {
                dir_failures += usize::from(p.diagnostics.iter().any(|d| d.code == crate::DiagCode::DirCorrupt));
                module_failures += p.modules.iter().filter(|m| !m.diagnostics.is_empty()).count();
            }
            Ok(Err(why)) | Err(why) => panic!("P4 failed for seed {seed}: {why}\n  {}", replay(seed, "p4_corruptions_give_diagnostics")),
        }
    }
    if !replaying && seeds.len() >= 64 {
        assert!(
            dir_failures > 0 && module_failures > 0,
            "P4 reached {dir_failures} dir and {module_failures} module failures in {} cases",
            seeds.len()
        );
    }
}

/// Fuzz: arbitrary streams never panic and stay within the limits.
///
/// **Validates: Requirements 10.7**
#[test]
fn fuzz_read_project_never_panics() {
    let (seeds, _) = seed_list(FUZZ_PER_SEED);
    for &seed in &seeds {
        let checked = guarded(|| {
            let src = fuzz_source(seed);
            fuzz_limits(seed).iter().try_for_each(|limits| check_limits(&read_project(&src, limits), limits))
        });
        if let Ok(Err(why)) | Err(why) = checked {
            panic!("fuzz failed for seed {seed}: {why}\n  {}", replay(seed, "fuzz_read_project_never_panics"));
        }
    }
}
