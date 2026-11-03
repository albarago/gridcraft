use super::{Limits, MemSource, StreamSource, VbaModule, VbaProject, read_project};
use crate::diag::DiagCode;
use crate::dir::fixture::{DirFixture, ModuleFixture, ReferenceFixture};
use crate::dir::{DirModuleType, ReferenceKind};
use crate::ovba::compress;
use crate::project_stream::ModuleKind;

const PROCEDURAL: DirModuleType = DirModuleType::Procedural;
const OTHER: DirModuleType = DirModuleType::NonProcedural;

/// A PROJECT stream listing the given properties (`("Module", "Module1")`, …).
fn project_text(props: &[(&str, &str)]) -> Vec<u8> {
    let mut text = String::from("ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n");
    for (key, value) in props {
        text.push_str(&format!("{key}={value}\r\n"));
    }
    text.push_str("Name=\"VBAProject\"\r\n\r\n[Workspace]\r\nModule1=0, 0, 0, 0, C\r\n");
    text.into_bytes()
}

/// A source with `VBA/dir` from `dir`, an optional PROJECT stream, and the given module streams.
fn source(dir: &DirFixture, project: Option<Vec<u8>>, modules: &[(&str, Vec<u8>)]) -> MemSource {
    let mut src = MemSource::new().with(&["VBA", "dir"], compress(&dir.to_bytes()));
    if let Some(project) = project {
        src.insert(&["PROJECT"], project);
    }
    for (stream, bytes) in modules {
        src.insert(&["VBA", stream], bytes.clone());
    }
    src
}

fn modules(specs: &[(&str, DirModuleType)]) -> Vec<ModuleFixture> {
    specs.iter().map(|&(n, t)| ModuleFixture::new(n, t)).collect()
}

fn codes(diags: &[crate::diag::Diag]) -> Vec<DiagCode> {
    diags.iter().map(|d| d.code).collect()
}

fn module<'a>(p: &'a VbaProject, name: &str) -> &'a VbaModule {
    p.modules.iter().find(|m| m.name == name).unwrap_or_else(|| panic!("no module {name} in {p:#?}"))
}

/// Module streams by stream name.
type Streams = Vec<(&'static str, Vec<u8>)>;

/// Three modules (standard, class, document) with Windows-1252 source.
fn happy() -> (DirFixture, Vec<u8>, Streams) {
    let dir = DirFixture {
        references: vec![ReferenceFixture::new("stdole", ReferenceKind::Registered)],
        modules: modules(&[("Module1", PROCEDURAL), ("Class1", OTHER), ("ThisWorkbook", OTHER)]),
        ..DirFixture::default()
    };
    let project = project_text(&[("Document", "ThisWorkbook/&H00000000"), ("Module", "Module1"), ("Class", "Class1")]);
    let streams = vec![
        ("Module1", compress(b"Attribute VB_Name = \"Module1\"\r\n' Caf\xE9 \x80 5\r\nSub A()\r\nEnd Sub\r\n")),
        ("Class1", compress(b"Attribute VB_Name = \"Class1\"\r\n")),
        ("ThisWorkbook", compress(b"Attribute VB_Name = \"ThisWorkbook\"\r\n")),
    ];
    (dir, project, streams)
}

#[test]
fn reads_three_modules_of_different_kinds() {
    let (dir, project, streams) = happy();
    let p = read_project(&source(&dir, Some(project), &streams), &Limits::default());
    assert_eq!(p.diagnostics, vec![]);
    assert_eq!(p.name.as_deref(), Some("VBAProject"));
    assert_eq!(p.code_page, Some(1252));
    assert!(!p.signed);
    assert_eq!(p.references.len(), 1);
    assert_eq!(p.references[0].name, "stdole");
    assert_eq!(p.references[0].kind, ReferenceKind::Registered);
    let names: Vec<&str> = p.modules.iter().map(|m| m.name.as_str()).collect();
    assert_eq!(names, ["Module1", "Class1", "ThisWorkbook"], "dir order is kept");
    let kinds: Vec<ModuleKind> = p.modules.iter().map(|m| m.kind).collect();
    assert_eq!(kinds, [ModuleKind::Standard, ModuleKind::Class, ModuleKind::Document]);
    let m1 = module(&p, "Module1");
    assert_eq!(m1.source, "Attribute VB_Name = \"Module1\"\r\n' Caf\u{e9} \u{20ac} 5\r\nSub A()\r\nEnd Sub\r\n");
    assert_eq!((m1.code_name.as_str(), m1.stream_name.as_str(), m1.text_offset), ("Module1", "Module1", 0));
    assert!(p.modules.iter().all(|m| m.diagnostics.is_empty()));
}

#[test]
fn source_starts_at_module_offset() {
    let mut dir = DirFixture { modules: modules(&[("Module1", PROCEDURAL)]), ..DirFixture::default() };
    dir.modules[0].text_offset = 7;
    let mut stream = vec![0xAB; 7];
    stream.extend(compress(b"Sub A()\r\nEnd Sub\r\n"));
    let p = read_project(&source(&dir, Some(project_text(&[("Module", "Module1")])), &[("Module1", stream)]), &Limits::default());
    assert_eq!(p.modules[0].source, "Sub A()\r\nEnd Sub\r\n");
    assert_eq!(p.modules[0].text_offset, 7);
    assert_eq!(p.modules[0].diagnostics, vec![]);
}

#[test]
fn dir_missing_gives_no_modules() {
    let p = read_project(&MemSource::new().with(&["PROJECT"], project_text(&[])), &Limits::default());
    assert_eq!(codes(&p.diagnostics), [DiagCode::DirMissing]);
    assert!(p.modules.is_empty() && p.name.is_none() && p.code_page.is_none());
}

#[test]
fn dir_bad_compression_parse_error_or_read_error_is_corrupt() {
    let bad_signature = MemSource::new().with(&["VBA", "dir"], vec![0x02, 0x00]);
    let bad_records = MemSource::new().with(&["VBA", "dir"], compress(b"not a dir stream at all"));
    let read_error = MemSource::new().failing(&["VBA", "dir"]);
    for src in [bad_signature, bad_records, read_error] {
        let p = read_project(&src, &Limits::default());
        assert_eq!(codes(&p.diagnostics), [DiagCode::DirCorrupt], "{p:#?}");
        assert!(p.modules.is_empty() && p.code_page.is_none());
        assert!(p.diagnostics[0].module.is_none());
    }
}

#[test]
fn project_stream_missing_falls_back_to_module_type() {
    let (dir, _, streams) = happy();
    let p = read_project(&source(&dir, None, &streams), &Limits::default());
    assert_eq!(codes(&p.diagnostics), [DiagCode::ProjectStreamMissing]);
    let kinds: Vec<ModuleKind> = p.modules.iter().map(|m| m.kind).collect();
    assert_eq!(kinds, [ModuleKind::Standard, ModuleKind::Class, ModuleKind::Class]);
    for m in &p.modules {
        assert_eq!(codes(&m.diagnostics), [DiagCode::KindFromModuleType]);
        assert_eq!(m.diagnostics[0].module.as_deref(), Some(m.name.as_str()));
        assert!(m.diagnostics[0].message.contains(&m.name));
    }
    assert!(module(&p, "Module1").source.contains("Sub A()"), "source is still read");
}

#[test]
fn project_stream_read_error_and_unlisted_module() {
    let (dir, _, streams) = happy();
    let p = read_project(&source(&dir, None, &streams).failing(&["PROJECT"]), &Limits::default());
    assert_eq!(codes(&p.diagnostics), [DiagCode::ProjectStreamMissing]);
    // A PROJECT stream that lists only two of the three modules.
    let partial = project_text(&[("Module", "Module1"), ("Class", "Class1")]);
    let p = read_project(&source(&dir, Some(partial), &streams), &Limits::default());
    assert_eq!(p.diagnostics, vec![]);
    assert_eq!(codes(&module(&p, "ThisWorkbook").diagnostics), [DiagCode::KindFromModuleType]);
    assert_eq!(module(&p, "Class1").diagnostics, vec![]);
}

#[test]
fn missing_module_stream_does_not_stop_the_others() {
    let (dir, project, mut streams) = happy();
    streams.retain(|(n, _)| *n != "Class1");
    let p = read_project(&source(&dir, Some(project), &streams), &Limits::default());
    assert_eq!(p.modules.len(), 3);
    let class1 = module(&p, "Class1");
    assert_eq!(class1.source, "");
    assert_eq!(codes(&class1.diagnostics), [DiagCode::ModuleStreamMissing]);
    assert!(class1.diagnostics[0].message.contains("\"Class1\""));
    assert!(module(&p, "Module1").source.contains("Sub A()"));
    assert!(module(&p, "ThisWorkbook").source.contains("ThisWorkbook"));
}

#[test]
fn offset_at_or_past_the_end_is_out_of_range() {
    for (offset, len) in [(100u32, 10usize), (10, 10)] {
        let mut dir = DirFixture { modules: modules(&[("Module1", PROCEDURAL)]), ..DirFixture::default() };
        dir.modules[0].text_offset = offset;
        let p = read_project(&source(&dir, Some(project_text(&[("Module", "Module1")])), &[("Module1", vec![1; len])]), &Limits::default());
        assert_eq!(p.modules[0].source, "");
        assert_eq!(codes(&p.modules[0].diagnostics), [DiagCode::ModuleOffsetOutOfRange], "offset {offset}");
    }
}

#[test]
fn damaged_and_cut_short_sources() {
    let text: Vec<u8> = (0..3000u32).map(|i| b'a' + u8::try_from(i % 26).unwrap_or(0)).collect();
    let full = compress(&text);
    let cases = [
        (vec![0x01, 0x00, 0x00, 0x00], DiagCode::ModuleCorrupt),      // bad chunk signature
        (vec![0x02, 0x00], DiagCode::ModuleCorrupt),                  // bad container signature
        (full[..full.len() - 5].to_vec(), DiagCode::ModuleTruncated), // chunk reaches past the end
        (full[..2].to_vec(), DiagCode::ModuleTruncated),              // header cut in half
    ];
    for (bytes, expected) in cases {
        let dir = DirFixture { modules: modules(&[("Module1", PROCEDURAL), ("Module2", PROCEDURAL)]), ..DirFixture::default() };
        let project = project_text(&[("Module", "Module1"), ("Module", "Module2")]);
        let src = source(&dir, Some(project), &[("Module1", bytes), ("Module2", compress(b"Sub B()\r\nEnd Sub\r\n"))]);
        let p = read_project(&src, &Limits::default());
        assert_eq!(p.modules[0].source, "");
        assert_eq!(codes(&p.modules[0].diagnostics), [expected]);
        assert_eq!(p.modules[1].source, "Sub B()\r\nEnd Sub\r\n");
    }
}

/// Found by P3: a source whose last chunk can't be compressed is stored as a raw chunk, which
/// always decompresses to 4,096 bytes. Its 0x00 padding must not show up as NUL characters.
#[test]
fn raw_last_chunk_padding_is_not_source() {
    // 4,096 compressible bytes, then 3,900 printable bytes with no repeats worth a copy token.
    let mut text: Vec<u8> = b"Sub A()\r\n".iter().copied().cycle().take(4096).collect();
    let mut x = 0x9E37_79B9_7F4A_7C15u64;
    text.extend((0..3900).map(|_| {
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        b' ' + u8::try_from(x % 95).unwrap_or(0)
    }));
    let stream = compress(&text);
    // The second chunk is raw: CompressedChunkFlag (bit 15 of its header) is 0.
    let second = 1 + 3 + usize::from(u16::from_le_bytes([stream[1], stream[2]]) & 0x0FFF);
    assert_eq!(u16::from_le_bytes([stream[second], stream[second + 1]]) & 0x8000, 0, "the last chunk must be raw");
    assert_eq!(crate::ovba::decompress(&stream, 1 << 20).map(|v| v.len()), Ok(8192), "the raw chunk is padded");
    let dir = DirFixture { modules: modules(&[("Module1", PROCEDURAL)]), ..DirFixture::default() };
    let p = read_project(&source(&dir, Some(project_text(&[("Module", "Module1")])), &[("Module1", stream)]), &Limits::default());
    assert_eq!(p.modules[0].source.as_bytes(), text.as_slice());
    assert_eq!(p.modules[0].diagnostics, vec![]);
}

#[test]
fn module_stream_read_error_is_corrupt() {
    let (dir, project, streams) = happy();
    let p = read_project(&source(&dir, Some(project), &streams).failing(&["VBA", "Class1"]), &Limits::default());
    assert_eq!(codes(&module(&p, "Class1").diagnostics), [DiagCode::ModuleCorrupt]);
    assert_eq!(module(&p, "Module1").diagnostics, vec![]);
}

#[test]
fn unsupported_code_page_is_reported_once() {
    let (mut dir, project, streams) = happy();
    dir.code_page = 1251;
    let p = read_project(&source(&dir, Some(project), &streams), &Limits::default());
    assert_eq!(codes(&p.diagnostics), [DiagCode::UnsupportedCodePage(1251)]);
    assert!(p.diagnostics[0].message.contains("1251"));
    assert_eq!(p.code_page, Some(1251));
    assert!(p.modules.iter().all(|m| m.diagnostics.is_empty()));
    assert!(module(&p, "Module1").source.contains("Caf\u{e9} \u{20ac}"), "decoded as Windows-1252");
}

#[test]
fn module_limit_keeps_the_first_modules() {
    let (dir, project, streams) = happy();
    let limits = Limits { max_modules: 2, ..Limits::default() };
    let p = read_project(&source(&dir, Some(project), &streams), &limits);
    assert_eq!(p.modules.iter().map(|m| m.name.as_str()).collect::<Vec<_>>(), ["Module1", "Class1"]);
    assert_eq!(codes(&p.diagnostics), [DiagCode::ModuleLimit]);
    assert!(p.diagnostics[0].message.contains('3') && p.diagnostics[0].message.contains('2'));
}

#[test]
fn source_limit_stops_reading() {
    let dir = DirFixture { modules: modules(&[("A", PROCEDURAL), ("B", PROCEDURAL), ("C", PROCEDURAL)]), ..DirFixture::default() };
    let project = project_text(&[("Module", "A"), ("Module", "B"), ("Module", "C")]);
    let ten = compress(b"0123456789");
    let src = source(&dir, Some(project), &[("A", ten.clone()), ("B", ten.clone()), ("C", ten)]);
    let p = read_project(&src, &Limits { max_total_source: 25, ..Limits::default() });
    assert_eq!(p.modules.iter().map(|m| m.name.as_str()).collect::<Vec<_>>(), ["A", "B"]);
    assert_eq!(codes(&p.diagnostics), [DiagCode::SourceLimit]);
    // Exactly at the limit is still fine.
    let p = read_project(&src, &Limits { max_total_source: 30, ..Limits::default() });
    assert_eq!((p.modules.len(), p.diagnostics.len()), (3, 0));
}

#[test]
fn unicode_name_is_preferred_for_display() {
    // The fixture writes MBCS fields as UTF-8 bytes, so in code page 1252 the MBCS name reads as
    // mojibake while MODULENAMEUNICODE and StreamNameUnicode read correctly.
    let mut dir = DirFixture { modules: modules(&[("M\u{f3}dulo", PROCEDURAL), ("N\u{e9}", PROCEDURAL)]), ..DirFixture::default() };
    dir.modules[1].name_unicode = false;
    let project = project_text(&[("Module", "M\u{f3}dulo"), ("Module", "N\u{e9}")]);
    let src = source(&dir, Some(project), &[("M\u{f3}dulo", compress(b"' a")), ("N\u{e9}", compress(b"' b"))]);
    let p = read_project(&src, &Limits::default());
    let (a, b) = (&p.modules[0], &p.modules[1]);
    assert_eq!((a.name.as_str(), a.code_name.as_str(), a.stream_name.as_str()), ("M\u{f3}dulo", "M\u{c3}\u{b3}dulo", "M\u{f3}dulo"));
    assert_eq!((b.name.as_str(), b.code_name.as_str()), ("N\u{c3}\u{a9}", "N\u{c3}\u{a9}"), "no MODULENAMEUNICODE: MBCS name");
    assert_eq!((a.source.as_str(), b.source.as_str()), ("' a", "' b"));
    assert!(a.diagnostics.is_empty() && b.diagnostics.is_empty(), "{p:#?}");
}

#[test]
fn empty_project_has_no_project_stream_diagnostic() {
    let p = read_project(&source(&DirFixture::default(), None, &[]), &Limits::default());
    assert_eq!(p, VbaProject { name: Some("VBAProject".to_owned()), code_page: Some(1252), ..VbaProject::default() });
}

#[test]
fn mem_source_lists_children_case_insensitively() {
    let (dir, project, streams) = happy();
    let src = source(&dir, Some(project), &streams).failing(&["Broken"]);
    assert_eq!(src.children(&[]).unwrap(), ["PROJECT", "VBA"]);
    assert_eq!(src.children(&["vba"]).unwrap(), ["Class1", "Module1", "ThisWorkbook", "dir"]);
    assert_eq!(src.children(&["VBA", "dir"]).unwrap(), Vec::<String>::new());
    assert!(src.stream(&["vba", "DIR"]).unwrap().is_some());
    assert!(src.stream(&["VBA", "nope"]).unwrap().is_none());
    assert!(src.children(&["Broken"]).is_err() && src.stream(&["Broken"]).is_err());
}
