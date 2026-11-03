use std::collections::BTreeMap;

use super::cfb::{self, CfbError, Damage, StreamSpec};
use super::{
    ALL_DIAG_CODES, Corruption, DIAG_MODULE, Fixture, FixtureModule, MAX_PCODE, PCODE_PLACEHOLDER, SIGNATURE_PLACEHOLDER, SignatureKind,
    diag_fixture, encode, vba_project_header,
};
use crate::diag::DiagCode;
use crate::encoding::{CP_MAC_ROMAN, CP_UTF8, CP_WINDOWS_1252};
use crate::project::{Limits, MemSource, StreamSource, VbaProject, read_project};
use crate::project_stream::ModuleKind;

fn four() -> Vec<FixtureModule> {
    vec![
        FixtureModule::new("Module1", ModuleKind::Standard, "Attribute VB_Name = \"Module1\"\r\n' Café € “quoted” Ž\r\nSub A()\r\nEnd Sub\r\n"),
        FixtureModule::new("Class1", ModuleKind::Class, "Attribute VB_Name = \"Class1\"\r\n"),
        FixtureModule::new("ThisWorkbook", ModuleKind::Document, "Attribute VB_Name = \"ThisWorkbook\"\r\n"),
        FixtureModule::new("UserForm1", ModuleKind::Form, ""),
    ]
}

fn read(src: &MemSource) -> VbaProject {
    read_project(src, &Limits::default())
}

/// (name, kind, source) of every module.
fn summary(p: &VbaProject) -> Vec<(String, ModuleKind, String)> {
    p.modules.iter().map(|m| (m.name.clone(), m.kind, m.source.clone())).collect()
}

fn expected(f: &Fixture) -> Vec<(String, ModuleKind, String)> {
    f.modules.iter().map(|m| (m.name.clone(), m.kind, m.source.clone())).collect()
}

fn path(p: &[&str]) -> Vec<String> {
    p.iter().map(|s| (*s).to_owned()).collect()
}

/// A source long enough to span several mini sectors once compressed.
fn long_source(lines: usize) -> String {
    (0..lines).map(|i| format!("Public V{i} As Long ' {}\r\n", i * 7919)).collect()
}

#[test]
fn round_trips_zero_one_and_four_modules() {
    for count in [0, 1, 4] {
        let f = Fixture::new(four().into_iter().take(count).collect());
        let p = read(&f.to_streams());
        assert_eq!(summary(&p), expected(&f), "{count} modules");
        assert_eq!(p.name.as_deref(), Some("VBAProject"));
        assert_eq!(p.code_page, Some(CP_WINDOWS_1252));
        assert!(p.diagnostics.is_empty(), "{:?}", p.diagnostics);
        assert!(p.modules.iter().all(|m| m.diagnostics.is_empty() && m.text_offset == 0), "{p:#?}");
    }
}

#[test]
fn non_ascii_cp1252_source_is_stored_as_cp1252() {
    let f = Fixture::new(four());
    let stream = f.module_stream(&f.modules[0]);
    let decompressed = crate::ovba::decompress(&stream, 1 << 20).unwrap();
    // é, €, “, ”, Ž as single Windows-1252 bytes.
    for b in [0xE9u8, 0x80, 0x93, 0x94, 0x8E] {
        assert!(decompressed.contains(&b), "missing 0x{b:02X}");
    }
}

#[test]
fn round_trips_a_utf8_project() {
    let f = Fixture {
        code_page: CP_UTF8,
        project_name: "Projekt".to_owned(),
        modules: vec![
            FixtureModule::new("Modul1", ModuleKind::Standard, "' 日本語 ✓ ünïcode\r\nSub B()\r\nEnd Sub\r\n"),
            FixtureModule::new("Klasse", ModuleKind::Class, "' Ω\r\n"),
        ],
        ..Fixture::default()
    };
    let p = read(&f.to_streams());
    assert_eq!(summary(&p), expected(&f));
    assert_eq!((p.name.as_deref(), p.code_page), (Some("Projekt"), Some(CP_UTF8)));
    assert!(p.diagnostics.is_empty(), "{:?}", p.diagnostics);
}

#[test]
fn cache_prefix_moves_the_module_offset() {
    let mut module = FixtureModule::new("Module1", ModuleKind::Standard, "Sub C()\r\nEnd Sub\r\n");
    module.cache_prefix = vec![0xAA; 37];
    let f = Fixture::new(vec![module]);
    let p = read(&f.to_streams());
    assert_eq!(p.modules[0].text_offset, 37);
    assert_eq!(summary(&p), expected(&f));
    // A cache next to real code is reported as not checked (Req 7.3).
    let codes: Vec<DiagCode> = p.modules[0].diagnostics.iter().map(|d| d.code).collect();
    assert_eq!(codes, [DiagCode::CompiledCodeNotChecked]);
}

#[test]
fn encode_inverts_decode() {
    for cp in [CP_WINDOWS_1252, CP_MAC_ROMAN] {
        let all: Vec<u8> = (0..=u8::MAX).collect();
        let text = crate::encoding::decode(cp, &all).0;
        assert_eq!(encode(cp, &text), all, "code page {cp}");
    }
    assert_eq!(encode(CP_WINDOWS_1252, "a日b"), b"a?b");
    assert_eq!(encode(CP_UTF8, "日"), "日".as_bytes());
}

#[test]
fn project_stream_lists_every_kind() {
    let text = String::from_utf8(Fixture::new(four()).project_text()).unwrap();
    for line in
        ["Module=Module1\r\n", "Class=Class1\r\n", "Document=ThisWorkbook/&H00000000\r\n", "BaseClass=UserForm1\r\n", "Name=\"VBAProject\"\r\n"]
    {
        assert!(text.contains(line), "{line:?} missing from {text}");
    }
    assert!(text.starts_with("ID=\"{"));
}

/// The codes on the named module, or on the project when `module` is `None`.
fn codes_of(p: &VbaProject, module: Option<&str>) -> Vec<DiagCode> {
    match module {
        Some(name) => p.modules.iter().find(|m| m.name == name).map(|m| m.diagnostics.iter().map(|d| d.code).collect()).unwrap_or_default(),
        None => p.diagnostics.iter().map(|d| d.code).collect(),
    }
}

#[test]
fn each_corruption_gives_its_diagnostic() {
    let m = |name: &str| path(&["VBA", name]);
    let module_len = |f: &Fixture| f.module_stream(&f.modules[0]).len();
    let mut base = Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, &long_source(40)), four()[1].clone()]);
    base.modules[0].cache_prefix = vec![1, 2, 3];
    let cut = module_len(&base) / 2;
    let cases: Vec<(Corruption, Option<&str>, DiagCode)> = vec![
        (Corruption::TruncateStream { path: m("Module1"), keep: cut }, Some("Module1"), DiagCode::ModuleTruncated),
        (Corruption::TruncateStream { path: m("Module1"), keep: 2 }, Some("Module1"), DiagCode::ModuleOffsetOutOfRange),
        (Corruption::TruncateStream { path: m("dir"), keep: 40 }, None, DiagCode::DirCorrupt),
        (Corruption::OffsetPastEnd { module: "module1".to_owned() }, Some("Module1"), DiagCode::ModuleOffsetOutOfRange),
        (Corruption::OversizedLength { path: m("Module1") }, Some("Module1"), DiagCode::ModuleCorrupt),
        (Corruption::OversizedLength { path: path(&["PROJECT"]) }, None, DiagCode::ProjectStreamMissing),
        (Corruption::CyclicChain { path: m("Module1") }, Some("Module1"), DiagCode::ModuleCorrupt),
        (Corruption::CyclicChain { path: m("dir") }, None, DiagCode::DirCorrupt),
    ];
    for (corruption, module, code) in cases {
        let f = Fixture { corrupt: vec![corruption.clone()], ..base.clone() };
        let p = read(&f.to_streams());
        assert!(codes_of(&p, module).contains(&code), "{corruption:?}: wanted {code:?}, got {p:#?}");
        // The other module still reads.
        if module.is_some() {
            assert_eq!(p.modules[1].source, base.modules[1].source, "{corruption:?}");
        }
    }
}

#[test]
fn unknown_corruption_targets_are_ignored() {
    let f = Fixture {
        corrupt: vec![Corruption::TruncateStream { path: path(&["VBA", "Nope"]), keep: 0 }, Corruption::OffsetPastEnd { module: "Nope".to_owned() }],
        ..Fixture::new(four())
    };
    assert_eq!(f.to_streams(), Fixture::new(four()).to_streams());
}

// ---- The compound file ----

/// Reads every stream of a compound file by path, independently of the writer: follows the
/// DIFAT (header and chained sectors), the FAT, the mini FAT and the directory tree. Each chain
/// is read to its end, so a loop or a short chain is an error even where a lazier reader would
/// stop early.
fn read_cfb(b: &[u8]) -> Result<BTreeMap<Vec<String>, Vec<u8>>, String> {
    let u16_at = |at: usize| b.get(at..at + 2).map(|s| u16::from_le_bytes([s[0], s[1]])).ok_or("short");
    let u32_at = |buf: &[u8], at: usize| buf.get(at..at + 4).map(|s| u32::from_le_bytes([s[0], s[1], s[2], s[3]])).ok_or("short".to_owned());
    if b.get(..8) != Some(&cfb::SIGNATURE[..]) {
        return Err("signature".into());
    }
    assert_eq!((u16_at(24)?, u16_at(26)?, u16_at(28)?, u16_at(30)?, u16_at(32)?), (0x3E, 3, 0xFFFE, 9, 6));
    assert_eq!(u32_at(b, 56)?, 4096, "mini cutoff");
    let sector = |id: u32| b.get((id as usize + 1) * 512..(id as usize + 2) * 512).ok_or(format!("sector {id} out of range"));
    let mut fat_ids: Vec<u32> = (0..109).map(|i| u32_at(b, 76 + i * 4)).collect::<Result<_, _>>()?;
    let mut next = u32_at(b, 68)?;
    let mut difat_count = 0;
    while next < 0xFFFF_FFFA {
        let s = sector(next)?;
        fat_ids.extend((0..127).map(|i| u32_at(s, i * 4)).collect::<Result<Vec<_>, _>>()?);
        next = u32_at(s, 508)?;
        difat_count += 1;
    }
    assert_eq!(difat_count, u32_at(b, 72)?, "DIFAT sector count");
    fat_ids.retain(|&id| id < 0xFFFF_FFFA);
    assert_eq!(fat_ids.len(), u32_at(b, 44)? as usize, "FAT sector count");
    let mut fat = Vec::new();
    for id in fat_ids {
        let s = sector(id)?;
        fat.extend((0..128).map(|i| u32_at(s, i * 4)).collect::<Result<Vec<_>, _>>()?);
    }
    let follow = |table: &[u32], start: u32, unit: usize, data: &dyn Fn(u32) -> Result<Vec<u8>, String>| -> Result<Vec<u8>, String> {
        let mut out = Vec::new();
        let mut id = start;
        let mut seen = vec![false; table.len()];
        while id != 0xFFFF_FFFE {
            let slot = seen.get_mut(id as usize).ok_or("broken chain")?;
            if *slot {
                return Err("chain loops".into());
            }
            *slot = true;
            let bytes = data(id)?;
            assert_eq!(bytes.len(), unit);
            out.extend(bytes);
            id = table[id as usize];
        }
        Ok(out)
    };
    let regular = |id: u32| sector(id).map(<[u8]>::to_vec);
    let dir = follow(&fat, u32_at(b, 48)?, 512, &regular)?;
    let mini_fat_raw = match u32_at(b, 60)? {
        0xFFFF_FFFE => Vec::new(),
        start => follow(&fat, start, 512, &regular)?,
    };
    assert_eq!(mini_fat_raw.len() / 512, u32_at(b, 64)? as usize, "mini FAT sector count");
    let mini_fat: Vec<u32> = (0..mini_fat_raw.len() / 4).map(|i| u32_at(&mini_fat_raw, i * 4)).collect::<Result<_, _>>()?;
    let entry = |i: u32| dir.get(i as usize * 128..i as usize * 128 + 128).ok_or("entry out of range".to_owned());
    let root = entry(0)?;
    assert_eq!(root[66], 5);
    let container = match u32_at(root, 116)? {
        0xFFFF_FFFE => Vec::new(),
        start => follow(&fat, start, 512, &regular)?,
    };
    let mini = |id: u32| container.get(id as usize * 64..id as usize * 64 + 64).map(<[u8]>::to_vec).ok_or("mini sector out of range".to_owned());
    let mut out = BTreeMap::new();
    let mut stack = vec![(u32_at(root, 76)?, Vec::<String>::new())];
    let mut visited = 0;
    while let Some((i, prefix)) = stack.pop() {
        if i == 0xFFFF_FFFF {
            continue;
        }
        visited += 1;
        assert!(visited <= dir.len() / 128, "tree loops");
        let e = entry(i)?;
        let len = (u16::from_le_bytes([e[64], e[65]]) as usize).saturating_sub(2) / 2;
        let units: Vec<u16> = (0..len).map(|k| u16::from_le_bytes([e[k * 2], e[k * 2 + 1]])).collect();
        let name = String::from_utf16(&units).map_err(|e| e.to_string())?;
        assert_eq!(e[67], 1, "black");
        stack.push((u32_at(e, 68)?, prefix.clone()));
        stack.push((u32_at(e, 72)?, prefix.clone()));
        let mut path = prefix;
        path.push(name);
        match e[66] {
            1 => stack.push((u32_at(e, 76)?, path)),
            2 => {
                let size = u32_at(e, 120)? as usize;
                let start = u32_at(e, 116)?;
                let mut data = if size < 4096 { follow(&mini_fat, start, 64, &mini)? } else { follow(&fat, start, 512, &regular)? };
                let unit = if size < 4096 { 64 } else { 512 };
                if data.len() < size {
                    return Err(format!("{} shorter than declared", path.join("/")));
                }
                assert_eq!(data.len(), size.div_ceil(unit) * unit, "chain length of {}", path.join("/"));
                data.truncate(size);
                out.insert(path, data);
            }
            k => return Err(format!("entry kind {k}")),
        }
    }
    Ok(out)
}

/// Every stream of a [`MemSource`] by path, found by listing storages from the root.
fn mem_streams(src: &MemSource) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut out = BTreeMap::new();
    let mut stack = vec![Vec::<String>::new()];
    while let Some(prefix) = stack.pop() {
        let p: Vec<&str> = prefix.iter().map(String::as_str).collect();
        for name in src.children(&p).unwrap() {
            let mut child = prefix.clone();
            child.push(name);
            let c: Vec<&str> = child.iter().map(String::as_str).collect();
            match src.stream(&c).unwrap() {
                Some(bytes) => {
                    out.insert(child, bytes);
                }
                None => stack.push(child),
            }
        }
    }
    out
}

#[test]
fn the_bin_holds_the_same_streams_as_the_mem_source() {
    let mut big = FixtureModule::new("Big", ModuleKind::Class, &long_source(2000));
    big.cache_prefix = vec![7; 5000];
    let mut modules = four();
    modules.push(big);
    let truncated = Fixture {
        corrupt: vec![
            Corruption::TruncateStream { path: path(&["VBA", "Class1"]), keep: 10 },
            Corruption::OffsetPastEnd { module: "Module1".to_owned() },
        ],
        ..Fixture::new(modules.clone())
    };
    // Signature streams (U+0005 names) in the root, a P_Code region and a removed stream.
    let signed = Fixture {
        signatures: SignatureKind::ALL.to_vec(),
        corrupt: vec![Corruption::RemoveStream { path: path(&["VBA", "Class1"]) }],
        ..Fixture::new(vec![four()[0].clone().with_pcode(300), four()[1].clone()])
    };
    for f in [Fixture::default(), Fixture::new(four()), Fixture::new(modules), truncated, signed] {
        let bin = f.to_vba_project_bin().unwrap();
        assert_eq!(bin[..8], [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
        assert_eq!(bin.len() % 512, 0);
        let streams = read_cfb(&bin).unwrap();
        assert_eq!(streams, mem_streams(&f.to_streams()));
        assert_eq!(streams[&path(&["VBA", "_VBA_PROJECT"])], vba_project_header());
    }
}

#[test]
fn bin_corruptions_are_rejected_by_a_reader() {
    let base = Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, &long_source(40)), four()[1].clone()]);
    let big = Fixture::new(vec![FixtureModule::new("Module1", ModuleKind::Standard, &long_source(2000))]);
    let m = path(&["VBA", "Module1"]);
    for (f, corruption, want) in [
        (&base, Corruption::CyclicChain { path: m.clone() }, "chain loops"),
        (&big, Corruption::CyclicChain { path: m.clone() }, "chain loops"),
        (&base, Corruption::OversizedLength { path: m.clone() }, "shorter than declared"),
        (&base, Corruption::OversizedLength { path: path(&["VBA", "Class1"]) }, "shorter than declared"),
        (&big, Corruption::OversizedLength { path: m.clone() }, "shorter than declared"),
    ] {
        let damaged = Fixture { corrupt: vec![corruption.clone()], ..f.clone() };
        let err = read_cfb(&damaged.to_vba_project_bin().unwrap()).unwrap_err();
        assert!(err.contains(want), "{corruption:?}: {err}");
    }
}

#[test]
fn a_cycle_in_a_one_sector_stream_is_refused() {
    let f = Fixture { corrupt: vec![Corruption::CyclicChain { path: path(&["VBA", "_VBA_PROJECT"]) }], ..Fixture::default() };
    assert_eq!(f.to_vba_project_bin(), Err(CfbError::CycleNeedsTwoSectors("VBA/_VBA_PROJECT".to_owned())));
}

#[test]
fn bad_names_are_errors_not_panics() {
    for name in ["", "a/b", "x!", &"N".repeat(32)] {
        let f = Fixture::new(vec![FixtureModule::new(name, ModuleKind::Standard, "")]);
        assert!(matches!(f.to_vba_project_bin(), Err(CfbError::BadName(_))), "{name:?}");
    }
    let twice = Fixture::new(vec![FixtureModule::new("A", ModuleKind::Standard, ""), FixtureModule::new("a", ModuleKind::Class, "")]);
    assert!(matches!(twice.to_vba_project_bin(), Err(CfbError::Duplicate(_))));
    let clash = [
        StreamSpec { path: path(&["S"]), data: vec![1], damage: Damage::None },
        StreamSpec { path: path(&["S", "T"]), data: vec![2], damage: Damage::None },
    ];
    assert!(matches!(cfb::write(&clash), Err(CfbError::Duplicate(_))));
    assert_eq!(cfb::write(&[StreamSpec { path: vec![], data: vec![], damage: Damage::None }]), Err(CfbError::EmptyPath));
}

#[test]
fn large_files_use_difat_sectors() {
    // 9 MiB in regular sectors needs more than the header's 109 FAT sectors.
    let data: Vec<u8> = (0..9 * 1024 * 1024u32).map(|i| (i.wrapping_mul(2_654_435_761) >> 24) as u8).collect();
    let specs = [
        StreamSpec { path: path(&["Big"]), data: data.clone(), damage: Damage::None },
        StreamSpec { path: path(&["VBA", "small"]), data: vec![9; 100], damage: Damage::None },
    ];
    let bin = cfb::write(&specs).unwrap();
    assert!(u32::from_le_bytes(bin[72..76].try_into().unwrap()) >= 1, "DIFAT sectors");
    let streams = read_cfb(&bin).unwrap();
    assert_eq!(streams[&path(&["Big"])], data);
    assert_eq!(streams[&path(&["VBA", "small"])], vec![9; 100]);
}

#[test]
fn siblings_are_in_cfb_order() {
    // Shorter names first, then uppercased: the reader's walk doesn't care, but the order is
    // what [MS-CFB] §2.6.4 asks of a search tree.
    let specs: Vec<StreamSpec> =
        ["bb", "A", "ccc", "aa"].iter().map(|n| StreamSpec { path: path(&[n]), data: vec![], damage: Damage::None }).collect();
    let bin = cfb::write(&specs).unwrap();
    let dir = &bin[1024..];
    let name = |i: usize| {
        String::from_utf16(
            &dir[i * 128..i * 128 + 64].chunks(2).map(|c| u16::from_le_bytes([c[0], c[1]])).take_while(|&u| u != 0).collect::<Vec<_>>(),
        )
        .unwrap()
    };
    let link = |i: usize, at: usize| u32::from_le_bytes(dir[i * 128 + at..i * 128 + at + 4].try_into().unwrap());
    let mut order = Vec::new();
    let mut i = link(0, 76);
    while i != 0xFFFF_FFFF {
        order.push(name(i as usize));
        i = link(i as usize, 72);
    }
    assert_eq!(order, ["A", "aa", "bb", "ccc"]);
}

// ---- P_Code regions, signature streams and one fixture per diagnostic (Req 10.6) ----

#[test]
fn with_pcode_fills_a_deterministic_placeholder_region() {
    let m = || FixtureModule::new("Module1", ModuleKind::Standard, "Sub C()\r\nEnd Sub\r\n");
    for len in [0, 1, PCODE_PLACEHOLDER.len(), 1000] {
        let a = m().with_pcode(len);
        assert_eq!(a.cache_prefix.len(), len);
        assert_eq!(a, m().with_pcode(len), "the same length gives the same bytes");
        assert!(a.cache_prefix.iter().zip(PCODE_PLACEHOLDER.iter().cycle()).all(|(x, y)| x == y));
        // The region is what MODULEOFFSET skips; the source still reads.
        let f = Fixture::new(vec![a]);
        let p = read(&f.to_streams());
        assert_eq!(p.modules[0].text_offset as usize, len);
        assert_eq!(summary(&p), expected(&f));
    }
    assert_eq!(m().with_pcode(10).with_pcode(0).cache_prefix, Vec::<u8>::new());
    assert_eq!(m().with_cache(vec![1, 2, 3]).cache_prefix, [1, 2, 3]);
    assert_eq!(MAX_PCODE, crate::ovba::MAX_DECOMPRESSED);
}

#[test]
fn signature_kinds_name_the_signature_streams() {
    assert_eq!(SignatureKind::ALL.map(SignatureKind::stream_name), crate::project::SIGNATURE_STREAMS);
    assert!(SignatureKind::ALL.iter().all(|k| k.stream_name().starts_with('\u{5}')));
}

#[test]
fn signature_streams_are_root_streams_with_placeholder_bytes() {
    for kinds in [vec![SignatureKind::Legacy], vec![SignatureKind::Ex], vec![SignatureKind::Agile], SignatureKind::ALL.to_vec()] {
        // A kind listed twice is written once.
        let twice: Vec<SignatureKind> = kinds.iter().chain(&kinds).copied().collect();
        let f = Fixture { signatures: twice, ..Fixture::new(four()) };
        let mem = f.to_streams();
        let mut root = mem.children(&[]).unwrap();
        root.sort();
        let mut want: Vec<String> = kinds.iter().map(|k| k.stream_name().to_owned()).collect();
        want.extend(["PROJECT".to_owned(), "VBA".to_owned()]);
        want.sort();
        assert_eq!(root, want, "{kinds:?}");
        // The writer accepts the U+0005 names, and the bin holds the same bytes.
        let streams = read_cfb(&f.to_vba_project_bin().unwrap()).unwrap();
        assert_eq!(streams, mem_streams(&mem), "{kinds:?}");
        for k in &kinds {
            assert_eq!(streams[&path(&[k.stream_name()])], SIGNATURE_PLACEHOLDER, "{k:?}");
        }
        let p = read(&mem);
        assert!(p.signed, "{kinds:?}");
        assert_eq!(codes_of(&p, None), [DiagCode::Signed], "{kinds:?}");
        assert_eq!(summary(&p), expected(&f), "signatures don't change the modules");
    }
    let unsigned = read(&Fixture::new(four()).to_streams());
    assert!(!unsigned.signed && unsigned.diagnostics.is_empty());
}

#[test]
fn control_characters_are_valid_cfb_names() {
    // [MS-CFB] §2.6.1 forbids only / \ : ! in names.
    let specs: Vec<StreamSpec> = ["\u{5}DigitalSignature", "\u{1}CompObj", "a\u{7}b"]
        .iter()
        .map(|n| StreamSpec { path: path(&[n]), data: vec![1], damage: Damage::None })
        .collect();
    let streams = read_cfb(&cfb::write(&specs).unwrap()).unwrap();
    assert_eq!(streams.len(), 3);
    assert_eq!(streams[&path(&["\u{5}DigitalSignature"])], [1]);
}

#[test]
fn remove_stream_drops_it_from_both_outputs() {
    for target in [path(&["PROJECT"]), path(&["VBA", "dir"]), path(&["vba", "class1"])] {
        let f = Fixture { corrupt: vec![Corruption::RemoveStream { path: target.clone() }], ..Fixture::new(four()) };
        let mem = mem_streams(&f.to_streams());
        assert!(!mem.keys().any(|k| k.len() == target.len() && k.iter().zip(&target).all(|(a, b)| a.eq_ignore_ascii_case(b))), "{target:?}");
        assert_eq!(mem.len(), mem_streams(&Fixture::new(four()).to_streams()).len() - 1, "{target:?}");
        assert_eq!(read_cfb(&f.to_vba_project_bin().unwrap()).unwrap(), mem, "{target:?}");
    }
}

#[test]
fn pad_project_stream_passes_the_cap() {
    let f = Fixture { corrupt: vec![Corruption::PadProjectStream], ..Fixture::new(four()) };
    let text = mem_streams(&f.to_streams()).remove(&path(&["PROJECT"])).unwrap();
    assert!(text.len() > crate::project_stream::MAX_PROJECT_STREAM);
    let plain = Fixture::new(four()).project_text();
    assert!(text.starts_with(&plain[..plain.iter().position(|&b| b == b'\n').unwrap() + 1]), "the ID= line comes first");
    assert!(text.ends_with(&plain[plain.iter().position(|&b| b == b'\n').unwrap() + 1..]), "the rest follows the filler");
    assert!(crate::project_stream::parse_project_stream(&text).truncated);
}

/// One fixture per [`DiagCode`], read from the [`MemSource`]. The bin half, through the real
/// compound file reader, is in the engine (`vba_source/tests.rs`).
#[test]
fn every_diagnostic_code_has_a_fixture() {
    let mut ids: Vec<&str> = ALL_DIAG_CODES.iter().map(DiagCode::as_str).collect();
    ids.sort_unstable();
    ids.dedup();
    assert_eq!(ids.len(), ALL_DIAG_CODES.len(), "ALL_DIAG_CODES lists each code once");
    for code in ALL_DIAG_CODES {
        let d = diag_fixture(code);
        let p = read_project(&d.fixture.to_streams(), &d.limits);
        assert!(codes_of(&p, d.module).contains(&code), "{code:?}: got {p:#?}");
        if let Some(name) = d.module {
            assert_eq!(name, DIAG_MODULE);
            assert!(!codes_of(&p, None).contains(&code), "{code:?} is a module diagnostic");
        }
    }
}
