//! Fixture_Builder projects read through the real compound file reader (Req 5.1).
use gridcraft_vba::fixture::{ALL_DIAG_CODES, Corruption, Fixture, FixtureModule, SignatureKind, diag_fixture};
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

// --- Signatures (Req 7.4-7.6) ---

const SIGNED: DiagCode = DiagCode::Signed;

#[test]
fn a_signature_stream_is_found_through_the_compound_file() {
    // Only root streams: no project, but the root listing still sees the signature.
    let cfb = gridcraft_xlsx::testing::compound_file(&[("\u{5}DigitalSignatureAgile", b"0\x82 never read"), ("PROJECT", b"ID=\"{}\"\r\n")]);
    let p = read_vba_project(&cfb, &Limits::default());
    assert!(p.signed);
    assert_eq!(codes(&p), [DiagCode::DirMissing, SIGNED]);
    assert!(!read_bin(&Fixture::new(four())).signed, "a fixture without signature streams is unsigned");
}

#[test]
fn the_package_signature_flag_marks_the_workbook_signed() {
    let mut wb = gridcraft_model::Workbook::new();
    assert_eq!(super::read_workbook_vba(&wb, &Limits::default()), None, "no project, nothing read");
    wb.vba_project = Some(bin(&Fixture::new(four())).into());
    let unsigned = super::read_workbook_vba(&wb, &Limits::default()).unwrap();
    assert!(!unsigned.signed);
    assert!(!codes(&unsigned).contains(&SIGNED));
    wb.vba_signature_part = true;
    let signed = super::read_workbook_vba(&wb, &Limits::default()).unwrap();
    assert!(signed.signed);
    assert_eq!(signed.diagnostics.iter().filter(|d| d.code == SIGNED).count(), 1);
    assert_eq!(summary(&signed), summary(&unsigned), "the modules are read the same");
    // A bin with a signature stream as well still gives one diagnostic.
    wb.vba_project = Some(gridcraft_xlsx::testing::compound_file(&[("\u{5}DigitalSignature", b"x")]).into());
    let both = super::read_workbook_vba(&wb, &Limits::default()).unwrap();
    assert_eq!(codes(&both), [DiagCode::DirMissing, SIGNED]);
}

/// The package with `add` parts added.
fn with_parts(bytes: &[u8], add: &[(&str, &[u8])]) -> Vec<u8> {
    use std::io::{Read, Write};
    let mut z = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut out = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for i in 0..z.len() {
        let mut f = z.by_index(i).unwrap();
        let name = f.name().to_string();
        let mut data = Vec::new();
        f.read_to_end(&mut data).unwrap();
        out.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
        out.write_all(&data).unwrap();
    }
    for (name, data) in add {
        out.start_file(*name, zip::write::SimpleFileOptions::default()).unwrap();
        out.write_all(data).unwrap();
    }
    out.finish().unwrap().into_inner()
}

#[test]
fn a_package_signature_part_is_reported_through_the_engine() {
    let vba = bin(&Fixture::new(four()));
    let mut wb = gridcraft_model::Workbook::new();
    wb.vba_project = Some(vba.clone().into());
    let xlsm = gridcraft_xlsx::write_xlsx_as(&wb, gridcraft_xlsx::XlsxKind::MacroWorkbook).unwrap();
    let rels = br#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/></Relationships>"#;
    let signed = with_parts(&xlsm, &[("xl/_rels/vbaProject.bin.rels", rels), ("xl/vbaProjectSignature.bin", b"0\x82 opaque")]);
    let (back, _) = gridcraft_xlsx::read_xlsx(&signed).unwrap();
    assert!(back.vba_signature_part);
    assert_eq!(back.vba_project.as_deref(), Some(&vba), "the bin is kept byte for byte");
    let p = super::read_workbook_vba(&back, &Limits::default()).unwrap();
    assert!(p.signed);
    assert_eq!(codes(&p), [SIGNED]);
    let (plain, _) = gridcraft_xlsx::read_xlsx(&xlsm).unwrap();
    assert!(!super::read_workbook_vba(&plain, &Limits::default()).unwrap().signed);
}

// --- One fixture per diagnostic, signatures and P_Code regions (Req 10.6) ---

/// The diagnostic codes on `module`, or on the project for `None`.
fn codes_on(p: &VbaProject, module: Option<&str>) -> Vec<DiagCode> {
    match module {
        Some(name) => p.modules.iter().find(|m| m.name == name).map(|m| m.diagnostics.iter().map(|d| d.code).collect()).unwrap_or_default(),
        None => p.diagnostics.iter().map(|d| d.code).collect(),
    }
}

/// The bin half of `every_diagnostic_code_has_a_fixture` in the vba crate: each code's fixture,
/// written as a `vbaProject.bin`, read through [`Compound`] and the adapter.
#[test]
fn every_diagnostic_code_has_a_fixture_through_the_compound_reader() {
    for code in ALL_DIAG_CODES {
        let d = diag_fixture(code);
        let p = read_vba_project(&bin(&d.fixture), &d.limits);
        assert!(codes_on(&p, d.module).contains(&code), "{code:?}: got {:?}", codes(&p));
    }
}

#[test]
fn signed_and_stomped_fixtures_read_the_same_through_the_compound_reader() {
    let stomped = diag_fixture(DiagCode::PossibleStomping).fixture;
    let signed = diag_fixture(DiagCode::Signed).fixture;
    let both = Fixture { signatures: SignatureKind::ALL.to_vec(), ..stomped.clone() };
    let cases = [
        (&stomped, vec![DiagCode::PossibleStomping]),
        (&signed, vec![SIGNED]),
        (&both, vec![SIGNED, DiagCode::PossibleStomping]),
        (&diag_fixture(DiagCode::CompiledCodeNotChecked).fixture, vec![DiagCode::CompiledCodeNotChecked]),
    ];
    for (f, want) in cases {
        let p = read_bin(f);
        assert_eq!(p, read_mem(f), "the bin and the MemSource disagree: {want:?}");
        assert_eq!(codes(&p), want);
        assert_eq!(p.signed, want.contains(&SIGNED));
        assert_eq!(summary(&p), expected(f));
    }
    // The stomped module's P_Code region is what MODULEOFFSET skips.
    let p = read_bin(&stomped);
    assert_eq!(p.modules[0].text_offset as usize, stomped.modules[0].cache_prefix.len());
    // Through the workbook too: a signature stream and a package part still give one diagnostic.
    let mut wb = gridcraft_model::Workbook::new();
    wb.vba_project = Some(bin(&both).into());
    wb.vba_signature_part = true;
    let p = super::read_workbook_vba(&wb, &Limits::default()).unwrap();
    assert_eq!(codes(&p), [SIGNED, DiagCode::PossibleStomping]);
}

#[test]
fn each_signature_stream_is_found_through_the_compound_reader() {
    for kind in SignatureKind::ALL {
        let f = Fixture { signatures: vec![kind], ..Fixture::new(four()) };
        let b = bin(&f);
        let src = CompoundSource(Compound::open(&b).unwrap());
        assert!(src.children(&[]).unwrap().iter().any(|n| n == kind.stream_name()), "{kind:?}");
        let p = read_vba_project(&b, &Limits::default());
        assert!(p.signed, "{kind:?}");
        assert_eq!(codes(&p), [SIGNED], "{kind:?}");
    }
}
