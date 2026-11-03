//! Fixture_Builder projects read through the real compound file reader (Req 5.1).
use gridcraft_vba::fixture::{Corruption, Fixture, FixtureModule};
use gridcraft_vba::{DiagCode, Limits, ModuleKind, StreamSource, VbaProject, read_project};
use gridcraft_xlsx::Compound;

use super::{CompoundSource, read_vba_project};

fn four() -> Vec<FixtureModule> {
    vec![
        FixtureModule::new("Module1", ModuleKind::Standard, "Attribute VB_Name = \"Module1\"\r\n' Café € “quoted” Ž\r\nSub A()\r\nEnd Sub\r\n"),
        FixtureModule::new("Class1", ModuleKind::Class, "Attribute VB_Name = \"Class1\"\r\n"),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, "Attribute VB_Name = \"ThisWorkbook\"\r\n"),
        FixtureModule::new("UserForm1", ModuleKind::Form, ""),
    ]
}

/// A source of `lines` distinct lines. 2,000 lines compress past the 4,096-byte mini stream
/// cutoff, so that module stream is stored in regular sectors.
fn long_source(lines: usize) -> String {
    (0..lines).map(|i| format!("Public V{i} As Long ' {}\r\n", i * 7919)).collect()
}

fn path(p: &[&str]) -> Vec<String> {
    p.iter().map(|s| (*s).to_owned()).collect()
}

fn bin(f: &Fixture) -> Vec<u8> {
    f.to_vba_project_bin().unwrap()
}

/// Reads the fixture's bin through [`Compound`] and the adapter.
fn read_bin(f: &Fixture) -> VbaProject {
    read_vba_project(&bin(f), &Limits::default())
}

fn read_mem(f: &Fixture) -> VbaProject {
    read_project(&f.to_streams(), &Limits::default())
}

fn summary(p: &VbaProject) -> Vec<(String, ModuleKind, String)> {
    p.modules.iter().map(|m| (m.name.clone(), m.kind, m.source.clone())).collect()
}

fn expected(f: &Fixture) -> Vec<(String, ModuleKind, String)> {
    f.modules.iter().map(|m| (m.name.clone(), m.kind, m.source.clone())).collect()
}

/// Every diagnostic code, project-level first, then each module's.
fn codes(p: &VbaProject) -> Vec<DiagCode> {
    p.diagnostics.iter().chain(p.modules.iter().flat_map(|m| &m.diagnostics)).map(|d| d.code).collect()
}

fn fixtures() -> Vec<(&'static str, Fixture)> {
    let mut with_big = four();
    with_big.push(FixtureModule::new("Big", ModuleKind::Standard, &long_source(2000)));
    vec![
        ("0 modules", Fixture::new(Vec::new())),
        ("1 module", Fixture::new(four().into_iter().take(1).collect())),
        ("4 modules", Fixture::new(four())),
        ("4 modules + big", Fixture::new(with_big)),
    ]
}

#[test]
fn fixtures_read_the_same_through_the_compound_reader() {
    for (label, f) in fixtures() {
        let p = read_bin(&f);
        assert_eq!(summary(&p), expected(&f), "{label}");
        assert_eq!(p, read_mem(&f), "{label}: the bin and the MemSource disagree");
        assert_eq!(p.name.as_deref(), Some("VBAProject"), "{label}");
        assert!(codes(&p).is_empty(), "{label}: {p:#?}");
    }
}

#[test]
fn the_big_module_uses_regular_sectors() {
    let f = Fixture::new(vec![FixtureModule::new("Big", ModuleKind::Standard, &long_source(2000))]);
    assert!(f.module_stream(&f.modules[0]).len() > 4096, "the big module must pass the mini stream cutoff");
    assert_eq!(summary(&read_bin(&f)), expected(&f));
}

#[test]
fn non_ascii_cp1252_source_survives_the_compound_reader() {
    let p = read_bin(&Fixture::new(four()));
    assert!(p.modules[0].source.contains("' Café € “quoted” Ž"), "{:?}", p.modules[0].source);
    assert_eq!(p.code_page, Some(1252));
}

#[test]
fn children_list_the_project_storages() {
    for (label, f) in fixtures() {
        let b = bin(&f);
        let src = CompoundSource(Compound::open(&b).unwrap());
        assert_eq!(src.children(&[]).unwrap(), ["PROJECT", "VBA"], "{label}");
        let mut want: Vec<String> = f.modules.iter().map(|m| m.name.clone()).collect();
        want.extend(["_VBA_PROJECT".to_owned(), "dir".to_owned()]);
        want.sort();
        assert_eq!(src.children(&["VBA"]).unwrap(), want, "{label}");
        // Names match case-insensitively, like the MemSource.
        assert_eq!(src.children(&["vba"]).unwrap(), want, "{label}");
        let mut mem = f.to_streams().children(&["VBA"]).unwrap();
        mem.sort();
        assert_eq!(mem, want, "{label}");
        assert_eq!(src.stream(&["VBA", "Nope"]).unwrap(), None, "{label}");
    }
}

#[test]
fn bin_corruptions_give_diagnostics_through_the_compound_reader() {
    let m = |name: &str| path(&["VBA", name]);
    let base = Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, &long_source(40)), four()[1].clone()]);
    let big = Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, &long_source(2000)), four()[1].clone()]);
    let half = base.module_stream(&base.modules[0]).len() / 2;
    // (fixture, corruption, the module the diagnostic lands on or None for the project, the code).
    let cases: Vec<(&Fixture, Corruption, Option<&str>, DiagCode)> = vec![
        (&base, Corruption::CyclicChain { path: m("Module1") }, Some("Module1"), DiagCode::ModuleCorrupt),
        (&big, Corruption::CyclicChain { path: m("Module1") }, Some("Module1"), DiagCode::ModuleCorrupt),
        (&base, Corruption::CyclicChain { path: m("dir") }, None, DiagCode::DirCorrupt),
        (&base, Corruption::CyclicChain { path: path(&["PROJECT"]) }, None, DiagCode::ProjectStreamMissing),
        (&base, Corruption::OversizedLength { path: m("Module1") }, Some("Module1"), DiagCode::ModuleCorrupt),
        (&big, Corruption::OversizedLength { path: m("Module1") }, Some("Module1"), DiagCode::ModuleCorrupt),
        (&base, Corruption::OversizedLength { path: m("dir") }, None, DiagCode::DirCorrupt),
        (&base, Corruption::OversizedLength { path: path(&["PROJECT"]) }, None, DiagCode::ProjectStreamMissing),
        (&base, Corruption::TruncateStream { path: m("Module1"), keep: half }, Some("Module1"), DiagCode::ModuleTruncated),
        (&base, Corruption::TruncateStream { path: m("Module1"), keep: 0 }, Some("Module1"), DiagCode::ModuleOffsetOutOfRange),
        (&base, Corruption::TruncateStream { path: m("dir"), keep: 40 }, None, DiagCode::DirCorrupt),
        (&base, Corruption::OffsetPastEnd { module: "Module1".to_owned() }, Some("Module1"), DiagCode::ModuleOffsetOutOfRange),
    ];
    for (f, corruption, module, code) in cases {
        let f = Fixture { corrupt: vec![corruption.clone()], ..f.clone() };
        let p = read_bin(&f);
        let got: Vec<DiagCode> = match module {
            Some(name) => p.modules.iter().find(|x| x.name == name).map(|x| x.diagnostics.iter().map(|d| d.code).collect()).unwrap_or_default(),
            None => p.diagnostics.iter().map(|d| d.code).collect(),
        };
        assert!(got.contains(&code), "{corruption:?}: wanted {code:?}, got {:?}", codes(&p));
        // A module-level failure leaves the other module readable.
        if module.is_some() {
            assert_eq!(p.modules[1].source, f.modules[1].source, "{corruption:?}");
        }
    }
}

/// xorshift64, so the blobs are the same on every run.
fn random_bytes(mut seed: u64, len: usize) -> Vec<u8> {
    (0..len)
        .map(|_| {
            seed ^= seed << 13;
            seed ^= seed >> 7;
            seed ^= seed << 17;
            seed.to_le_bytes()[0]
        })
        .collect()
}

#[test]
fn a_file_that_isnt_a_compound_file_gives_dir_corrupt() {
    let mut blobs = vec![Vec::new(), b"PK\x03\x04 not a compound file".to_vec()];
    blobs.extend([1u64, 2, 0x9E37_79B9_7F4A_7C15].iter().map(|s| random_bytes(*s, 4096)));
    // The real signature with nothing after it.
    blobs.push(bin(&Fixture::new(four()))[..8].to_vec());
    for b in blobs {
        assert!(Compound::open(&b).is_err(), "{} bytes opened", b.len());
        let p = read_vba_project(&b, &Limits::default());
        assert!(p.modules.is_empty() && p.name.is_none() && p.code_page.is_none(), "{p:#?}");
        assert_eq!(codes(&p), [DiagCode::DirCorrupt], "{} bytes", b.len());
    }
}

#[test]
fn a_truncated_bin_never_panics() {
    let (_, f) = fixtures().swap_remove(3);
    let full = bin(&f);
    for len in (0..full.len()).step_by(97).chain([511, 512, 513, full.len() - 1]) {
        let p = read_vba_project(&full[..len], &Limits::default());
        if len < 512 {
            assert_eq!(codes(&p), [DiagCode::DirCorrupt], "{len} bytes");
        }
    }
}
