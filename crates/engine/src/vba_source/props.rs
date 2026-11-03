//! The VBA reader's properties through the real compound file reader: P3 and P4 on
//! `Fixture::to_vba_project_bin` read by [`read_vba_project`], and a seeded fuzz over whole
//! `vbaProject.bin` files. The generators and checkers are `gridcraft_vba::testutil`'s, shared
//! with the `MemSource` properties in `gridcraft-vba` (`project::props`).
//!
//! Knobs: `GRIDCRAFT_FUZZ_SEEDS` (default 64) gives P3 a quarter as many fixtures (they are the
//! `MemSource` P3's first ones), P4 2 and the fuzz 8 inputs per seed; `GRIDCRAFT_FUZZ_SEED=17`
//! replays one case.
use gridcraft_vba::fixture::Corruption;
use gridcraft_vba::testutil::{
    check_corrupt, check_limits, check_round_trip, fuzz_bin, fuzz_limits, guarded, p3_fixture, p4_case, seed_list, seed_list_scaled,
};
use gridcraft_vba::{Limits, read_project};

use super::read_vba_project;

/// P3 runs one fixture per this many seeds: the bin path compresses the same sources again.
const P3_DIVISOR: u64 = 4;
const P4_PER_SEED: u64 = 2;
const FUZZ_PER_SEED: u64 = 8;

fn replay(seed: u64, test: &str) -> String {
    format!("replay: GRIDCRAFT_FUZZ_SEED={seed} cargo test -p gridcraft-engine --lib vba_source::props::{test}")
}

/// P3 through the compound file: the bin reads back the fixture's modules, exactly as its
/// `MemSource` does.
///
/// **Validates: Requirements 10.3**
#[test]
fn p3_bin_round_trip() {
    let (seeds, _) = seed_list_scaled(1, P3_DIVISOR);
    for &seed in &seeds {
        let checked = guarded(|| {
            let f = p3_fixture(seed);
            let bin = f.to_vba_project_bin().map_err(|e| format!("the fixture didn't build: {e}"))?;
            let p = read_vba_project(&bin, &Limits::default());
            check_round_trip(&f, &p)?;
            if p != read_project(&f.to_streams(), &Limits::default()) {
                return Err("the bin and the MemSource read differently".to_owned());
            }
            Ok(())
        });
        if let Ok(Err(why)) | Err(why) = checked {
            panic!("P3 (bin) failed for seed {seed}: {why}\n  {}", replay(seed, "p3_bin_round_trip"));
        }
    }
}

/// P4 through the compound file. A cyclic chain on a stream that fits in one sector can't be
/// built; those cases are skipped.
///
/// **Validates: Requirements 10.5**
#[test]
fn p4_bin_corruptions_give_diagnostics() {
    let (seeds, replaying) = seed_list(P4_PER_SEED);
    let mut built = 0;
    for &seed in &seeds {
        let checked = guarded(|| {
            let case = p4_case(seed);
            let bin = match case.fixture.to_vba_project_bin() {
                Ok(bin) => bin,
                Err(_) if case.fixture.corrupt.iter().any(|c| matches!(c, Corruption::CyclicChain { .. })) => return Ok(false),
                Err(e) => return Err(format!("the fixture didn't build: {e}")),
            };
            check_corrupt(&case, &read_vba_project(&bin, &Limits::default())).map(|()| true)
        });
        match checked {
            Ok(Ok(b)) => built += usize::from(b),
            Ok(Err(why)) | Err(why) => panic!("P4 (bin) failed for seed {seed}: {why}\n  {}", replay(seed, "p4_bin_corruptions_give_diagnostics")),
        }
    }
    if !replaying {
        assert!(built * 2 > seeds.len(), "only {built} of {} P4 cases could be built", seeds.len());
    }
}

/// Fuzz: arbitrary `vbaProject.bin` bytes never panic and stay within the limits.
///
/// **Validates: Requirements 10.7**
#[test]
fn fuzz_read_vba_project_never_panics() {
    let (seeds, _) = seed_list(FUZZ_PER_SEED);
    for &seed in &seeds {
        let checked = guarded(|| {
            let bin = fuzz_bin(seed);
            fuzz_limits(seed).iter().try_for_each(|limits| check_limits(&read_vba_project(&bin, limits), limits))
        });
        if let Ok(Err(why)) | Err(why) = checked {
            panic!("fuzz failed for seed {seed}: {why}\n  {}", replay(seed, "fuzz_read_vba_project_never_panics"));
        }
    }
}
