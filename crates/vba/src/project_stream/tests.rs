use super::{MAX_LINES, MAX_PROJECT_STREAM, ModuleKind, ProjectStreamInfo, parse_project_stream, resolve_kind};
use crate::dir::DirModuleType;

/// The [MS-OVBA] §3.1 "PROJECT Stream Example" (VBAPROJECTText), typed from the spec page.
const SPEC_EXAMPLE: &[u8] = b"ID=\"{917DED54-440B-4FD1-A5C1-74ACF261E600}\"\r\n\
Document=ThisWorkbook/&H00000000\r\n\
Document=Sheet1/&H00000000\r\n\
Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}\r\n\
BaseClass=UserForm1\r\n\
HelpFile=\"c:\\example path\\example.hlp\"\r\n\
Name=\"VBAProject\"\r\n\
HelpContextID=\"1\"\r\n\
Description=\"Example VBA Project Description\"\r\n\
VersionCompatible32=\"393222000\"\r\n\
CMG=\"0705D8E3D8EDDBF1DBF1DBF1DBF1\"\r\n\
DPB=\"0E0CD1ECDFF4E7F5E7F5E7\"\r\n\
GC=\"1517CAF1D6F9D7F9D706\"\r\n\
\r\n\
[Host Extender Info]\r\n\
&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000\r\n\
\r\n\
[Workspace]\r\n\
ThisWorkbook=23, 23, 911, 280, \r\n\
Sheet1=69, 69, 724, 317, C\r\n\
UserForm1=0, 0, 0, 0, C, 46, 46, 701, 294, Z\r\n";

fn entry(name: &[u8], kind: ModuleKind) -> (Vec<u8>, ModuleKind) {
    (name.to_vec(), kind)
}

fn spec_modules() -> Vec<(Vec<u8>, ModuleKind)> {
    vec![entry(b"ThisWorkbook", ModuleKind::Document), entry(b"Sheet1", ModuleKind::Document), entry(b"UserForm1", ModuleKind::Form)]
}

#[test]
fn spec_example_gives_documents_and_form() {
    let info = parse_project_stream(SPEC_EXAMPLE);
    assert_eq!(info, ProjectStreamInfo { modules: spec_modules(), truncated: false });
}

#[test]
fn all_four_kinds() {
    let text = b"ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n\
Document=ThisWorkbook/&H00000000\r\n\
Module=Module1\r\n\
Class=Class1\r\n\
BaseClass=UserForm1\r\n\
Name=\"VBAProject\"\r\n";
    let info = parse_project_stream(text);
    assert_eq!(
        info.modules,
        vec![
            entry(b"ThisWorkbook", ModuleKind::Document),
            entry(b"Module1", ModuleKind::Standard),
            entry(b"Class1", ModuleKind::Class),
            entry(b"UserForm1", ModuleKind::Form),
        ]
    );
}

#[test]
fn lf_endings_whitespace_and_key_case() {
    let text = b"module=Module1  \t\n  CLASS = Class1\r\nDocument=Sheet2 / &H00000000 \n\
BaseClass=Form1\r\nModule=\r\nnot a property\n=Orphan\nHelpFile=\n";
    let info = parse_project_stream(text);
    assert_eq!(
        info.modules,
        vec![
            entry(b"Module1", ModuleKind::Standard),
            entry(b"Class1", ModuleKind::Class),
            entry(b"Sheet2", ModuleKind::Document),
            entry(b"Form1", ModuleKind::Form),
        ]
    );
    // A Document= value without the version suffix keeps the whole name.
    let bare = parse_project_stream(b"Document=Sheet3\r\n");
    assert_eq!(bare.modules, vec![entry(b"Sheet3", ModuleKind::Document)]);
}

#[test]
fn properties_after_a_section_header_are_ignored() {
    let mut text = SPEC_EXAMPLE.to_vec();
    text.extend_from_slice(b"Module=Late\r\nClass=Later\r\n");
    assert_eq!(parse_project_stream(&text).modules, spec_modules());
    // Any section header stops reading, not just the two the spec names.
    let early = parse_project_stream(b"Module=A\r\n[Other]\r\nModule=B\r\n");
    assert_eq!(early.modules, vec![entry(b"A", ModuleKind::Standard)]);
}

#[test]
fn resolve_uses_project_entry_case_insensitively() {
    let info = parse_project_stream(SPEC_EXAMPLE);
    // The PROJECT kind wins over MODULETYPE, whichever way the dir names it.
    assert_eq!(resolve_kind(Some(&info), b"THISWORKBOOK", DirModuleType::Procedural), (ModuleKind::Document, false));
    assert_eq!(resolve_kind(Some(&info), b"userform1", DirModuleType::NonProcedural), (ModuleKind::Form, false));
}

#[test]
fn resolve_falls_back_to_moduletype() {
    let info = parse_project_stream(SPEC_EXAMPLE);
    // Module missing from PROJECT.
    assert_eq!(resolve_kind(Some(&info), b"Module9", DirModuleType::Procedural), (ModuleKind::Standard, true));
    assert_eq!(resolve_kind(Some(&info), b"Class9", DirModuleType::NonProcedural), (ModuleKind::Class, true));
    // PROJECT stream missing.
    assert_eq!(resolve_kind(None, b"ThisWorkbook", DirModuleType::Procedural), (ModuleKind::Standard, true));
    assert_eq!(resolve_kind(None, b"ThisWorkbook", DirModuleType::NonProcedural), (ModuleKind::Class, true));
}

#[test]
fn mbcs_names_match_byte_exact() {
    // cp1252 "Módulo": only ASCII letters fold, so 0xF3 does not match 0xD3.
    let info = parse_project_stream(b"Module=M\xF3dulo\r\n");
    assert_eq!(info.modules, vec![entry(b"M\xF3dulo", ModuleKind::Standard)]);
    assert_eq!(resolve_kind(Some(&info), b"m\xF3DULO", DirModuleType::NonProcedural), (ModuleKind::Standard, false));
    assert_eq!(resolve_kind(Some(&info), b"M\xD3dulo", DirModuleType::NonProcedural), (ModuleKind::Class, true));
}

#[test]
fn every_prefix_of_the_example_is_safe() {
    let full = spec_modules();
    for end in 0..=SPEC_EXAMPLE.len() {
        let Some(prefix) = SPEC_EXAMPLE.get(..end) else { continue };
        let info = parse_project_stream(prefix);
        assert!(!info.truncated, "prefix {end}");
        assert!(info.modules.len() <= full.len(), "prefix {end}");
        // Each entry is the expected one, or a cut-short name of the right kind.
        for (got, want) in info.modules.iter().zip(&full) {
            assert_eq!(got.1, want.1, "prefix {end}");
            assert!(want.0.starts_with(&got.0), "prefix {end}");
        }
    }
}

#[test]
fn garbage_never_panics() {
    // Seeded xorshift64 over an alphabet that reaches every branch.
    const ALPHABET: &[u8] = b"ModuleClassBaseDocument=/&H[]\r\n \t\x00\xFF\x82";
    let mut state: u64 = 0x9E37_79B9_7F4A_7C15;
    for round in 0..256usize {
        let len = round.wrapping_mul(7) % 400;
        let mut bytes = Vec::with_capacity(len);
        for _ in 0..len {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            let pick = usize::try_from(state % ALPHABET.len() as u64).unwrap_or(0);
            bytes.push(ALPHABET.get(pick).copied().unwrap_or(b'x'));
        }
        let info = parse_project_stream(&bytes);
        for (name, _) in &info.modules {
            assert!(!name.is_empty(), "round {round}");
            assert!(!name.contains(&b'\n') && !name.contains(&b'\r'), "round {round}");
        }
    }
    assert_eq!(parse_project_stream(b""), ProjectStreamInfo::default());
    assert_eq!(parse_project_stream(&[0u8; 4096]), ProjectStreamInfo::default());
}

#[test]
fn byte_cap_stops_reading_and_drops_the_cut_line() {
    let mut text = b"Module=First\r\n".to_vec();
    // Long filler lines, so the byte cap is hit well before the line cap.
    let push_filler = |text: &mut Vec<u8>, len: usize| {
        text.extend_from_slice(b"x=");
        text.resize(text.len() + len, b'a');
        text.extend_from_slice(b"\r\n");
    };
    let target = MAX_PROJECT_STREAM - 10;
    while text.len() + 2000 < target {
        push_filler(&mut text, 1000);
    }
    let gap = target - text.len();
    push_filler(&mut text, gap - 4);
    assert_eq!(text.len(), target);
    // This line straddles the cap: it must not come back as a shorter name.
    text.extend_from_slice(b"Module=StraddlesTheCap\r\nModule=Late\r\n");
    assert!(text.len() > MAX_PROJECT_STREAM);
    let info = parse_project_stream(&text);
    assert_eq!(info.modules, vec![entry(b"First", ModuleKind::Standard)]);
    assert!(info.truncated);
}

#[test]
fn line_cap_stops_reading() {
    let mut text = b"Module=First\r\n".to_vec();
    for _ in 1..MAX_LINES {
        text.extend_from_slice(b"x\n");
    }
    // Exactly MAX_LINES lines, then blank lines and a section: nothing was missed.
    let mut exact = text.clone();
    exact.extend_from_slice(b"\r\n\n[Workspace]\r\nModule=W\r\n");
    let info = parse_project_stream(&exact);
    assert_eq!(info.modules, vec![entry(b"First", ModuleKind::Standard)]);
    assert!(!info.truncated);
    // One property past the cap is not read and is reported.
    text.extend_from_slice(b"Module=Late\r\n");
    let info = parse_project_stream(&text);
    assert_eq!(info.modules, vec![entry(b"First", ModuleKind::Standard)]);
    assert!(info.truncated);
}

#[test]
fn kind_names_are_stable() {
    let names: Vec<&str> = [ModuleKind::Standard, ModuleKind::Class, ModuleKind::Document, ModuleKind::Form].iter().map(|k| k.as_str()).collect();
    assert_eq!(names, ["standard", "class", "document", "form"]);
}
