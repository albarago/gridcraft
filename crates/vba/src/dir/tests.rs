use super::fixture::{DirFixture, ModuleFixture, ReferenceFixture, utf16};
use super::{DirError, DirModuleType, MAX_FIELD, MAX_MODULES, ReferenceKind, id, parse_dir};

/// PROJECTSYSKIND + PROJECTLCID + PROJECTLCIDINVOKE + PROJECTCODEPAGE, without PROJECTCOMPATVERSION.
const NAME_AT_WITHOUT_COMPAT: usize = 10 + 10 + 10 + 8;
/// The dir Terminator and its reserved field.
const DIR_TERMINATOR_LEN: usize = 6;
/// PROJECTMODULES Id, Size and Count, plus PROJECTCOOKIE.
const MODULES_HEADER_LEN: usize = 8 + 8;

fn module_names(names: &[&str]) -> Vec<ModuleFixture> {
    names.iter().map(|n| ModuleFixture::new(n, DirModuleType::Procedural)).collect()
}

fn put_u16(bytes: &mut [u8], at: usize, v: u16) {
    bytes[at..at + 2].copy_from_slice(&v.to_le_bytes());
}

fn put_u32(bytes: &mut [u8], at: usize, v: u32) {
    bytes[at..at + 4].copy_from_slice(&v.to_le_bytes());
}

/// A project using every optional record and every reference kind, with three modules.
fn rich() -> DirFixture {
    let mut modules = module_names(&["Module1", "Class1", "ThisWorkbook"]);
    modules[1].module_type = DirModuleType::NonProcedural;
    modules[1].read_only = true;
    modules[2].module_type = DirModuleType::NonProcedural;
    modules[2].private = true;
    modules[2].name_unicode = false;
    modules[2].stream_name = "ThisWorkbookStream".to_owned();
    modules[2].doc_string = "The workbook".to_owned();
    modules[2].text_offset = 0x0123_4567;
    DirFixture {
        compat_version: Some(2),
        doc_string: "Example VBA Project Description".to_owned(),
        help_file: "c:\\example path\\example.hlp".to_owned(),
        constants: Some("abc = 1".to_owned()),
        references: vec![
            ReferenceFixture::new("stdole", ReferenceKind::Registered),
            ReferenceFixture::new("VBAProject1", ReferenceKind::Project),
            ReferenceFixture::new("MSForms", ReferenceKind::Control),
            ReferenceFixture::new("Original", ReferenceKind::Original),
        ],
        modules,
        ..DirFixture::default()
    }
}

#[test]
fn zero_modules() {
    let info = parse_dir(&DirFixture::default().to_bytes(), MAX_MODULES).unwrap();
    assert_eq!(info.sys_kind, 1);
    assert_eq!(info.code_page, 1252);
    assert_eq!(info.name, b"VBAProject");
    assert_eq!(info.compat_version, None);
    assert!(info.constants.is_empty());
    assert!(info.references.is_empty());
    assert!(info.modules.is_empty());
    assert_eq!(info.truncated_modules, None);
}

#[test]
fn one_module() {
    let fixture = DirFixture { modules: module_names(&["Module1"]), ..DirFixture::default() };
    let info = parse_dir(&fixture.to_bytes(), MAX_MODULES).unwrap();
    assert_eq!(info.modules.len(), 1);
    let m = &info.modules[0];
    assert_eq!(m.name, b"Module1");
    assert_eq!(m.name_unicode, Some(utf16("Module1")));
    assert_eq!(m.stream_name, b"Module1");
    assert_eq!(m.stream_name_unicode, Some(utf16("Module1")));
    assert_eq!(m.text_offset, 0);
    assert_eq!(m.module_type, DirModuleType::Procedural);
    assert!(!m.read_only && !m.private);
}

#[test]
fn rich_project_reads_every_field() {
    let info = parse_dir(&rich().to_bytes(), MAX_MODULES).unwrap();
    assert_eq!(info.compat_version, Some(2));
    assert_eq!(info.constants, b"abc = 1");
    let refs: Vec<_> = info.references.iter().map(|r| (r.name.as_slice(), r.kind)).collect();
    assert_eq!(
        refs,
        [
            (&b"stdole"[..], ReferenceKind::Registered),
            (&b"VBAProject1"[..], ReferenceKind::Project),
            (&b"MSForms"[..], ReferenceKind::Control),
            (&b"Original"[..], ReferenceKind::Original),
        ]
    );
    let names: Vec<_> = info.modules.iter().map(|m| m.name.as_slice()).collect();
    assert_eq!(names, [&b"Module1"[..], b"Class1", b"ThisWorkbook"]);
    let [m1, class, wb] = &info.modules[..] else { panic!("three modules") };
    assert_eq!((m1.module_type, m1.read_only, m1.private), (DirModuleType::Procedural, false, false));
    assert_eq!((class.module_type, class.read_only, class.private), (DirModuleType::NonProcedural, true, false));
    assert_eq!((wb.module_type, wb.read_only, wb.private), (DirModuleType::NonProcedural, false, true));
    assert_eq!(wb.name_unicode, None);
    assert_eq!(wb.stream_name, b"ThisWorkbookStream");
    assert_eq!(wb.stream_name_unicode, Some(utf16("ThisWorkbookStream")));
    assert_eq!(wb.text_offset, 0x0123_4567);
}

#[test]
fn optional_records_in_every_combination() {
    for bits in 0..32u8 {
        let mut module = ModuleFixture::new("Module1", DirModuleType::Procedural);
        module.name_unicode = bits & 1 != 0;
        module.read_only = bits & 2 != 0;
        module.private = bits & 4 != 0;
        let fixture = DirFixture {
            compat_version: (bits & 8 != 0).then_some(7),
            constants: (bits & 16 != 0).then(|| "DEBUG = -1".to_owned()),
            modules: vec![module.clone()],
            ..DirFixture::default()
        };
        let info = parse_dir(&fixture.to_bytes(), MAX_MODULES).unwrap_or_else(|e| panic!("bits {bits:05b}: {e}"));
        assert_eq!(info.compat_version, fixture.compat_version, "bits {bits:05b}");
        assert_eq!(info.constants, fixture.constants.unwrap_or_default().as_bytes(), "bits {bits:05b}");
        let m = &info.modules[0];
        assert_eq!(m.name_unicode.is_some(), module.name_unicode, "bits {bits:05b}");
        assert_eq!((m.read_only, m.private), (module.read_only, module.private), "bits {bits:05b}");
    }
}

#[test]
fn references_with_and_without_names() {
    for kind in [ReferenceKind::Control, ReferenceKind::Original, ReferenceKind::Registered, ReferenceKind::Project] {
        let named = ReferenceFixture::new("Lib", kind);
        let unnamed = ReferenceFixture { name: None, ..named.clone() };
        let fixture = DirFixture { references: vec![unnamed, named], modules: module_names(&["M"]), ..DirFixture::default() };
        let info = parse_dir(&fixture.to_bytes(), MAX_MODULES).unwrap_or_else(|e| panic!("{kind:?}: {e}"));
        let refs: Vec<_> = info.references.iter().map(|r| (r.name.as_slice(), r.kind)).collect();
        assert_eq!(refs, [(&b""[..], kind), (&b"Lib"[..], kind)]);
        assert_eq!(info.modules.len(), 1);
    }
}

/// The writer's record sizes agree with the dir stream example in [MS-OVBA] §3.1.2: the
/// PROJECTINFORMATION example ends at 0x122, Reference[0] (stdole) at 0x1AE, and the ThisWorkbook
/// MODULE record is 0x94 bytes (0x4B2 to 0x546).
#[test]
fn layout_matches_spec_example() {
    let stdole = ReferenceFixture {
        libid: "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation".to_owned(),
        ..ReferenceFixture::new("stdole", ReferenceKind::Registered)
    };
    let mut fixture = DirFixture {
        code_page: 0x04E4,
        doc_string: "Example VBA Project Description".to_owned(),
        help_file: "c:\\example path\\example.hlp".to_owned(),
        constants: Some("abc = 1".to_owned()),
        ..DirFixture::default()
    };
    assert_eq!(fixture.to_bytes().len(), 0x122 + MODULES_HEADER_LEN + DIR_TERMINATOR_LEN);
    fixture.references.push(stdole);
    let without_module = fixture.to_bytes();
    assert_eq!(without_module.len(), 0x1AE + MODULES_HEADER_LEN + DIR_TERMINATOR_LEN);
    fixture.modules.push(ModuleFixture::new("ThisWorkbook", DirModuleType::NonProcedural));
    let with_module = fixture.to_bytes();
    assert_eq!(with_module.len() - without_module.len(), 0x94);
    // The example's PROJECTMODULES record starts at Id 0x000F; the module record follows it.
    let modules_at = 0x1AE;
    assert_eq!(with_module[modules_at..modules_at + 2], id::PROJECT_MODULES.to_le_bytes());
    assert_eq!(parse_dir(&with_module, MAX_MODULES).unwrap().modules[0].name, b"ThisWorkbook");
}

#[test]
fn every_prefix_is_truncated() {
    let bytes = rich().to_bytes();
    for n in 0..bytes.len() {
        match parse_dir(&bytes[..n], MAX_MODULES) {
            Err(DirError::Truncated { at }) => assert!(at <= n, "prefix {n}: at {at}"),
            other => panic!("prefix {n} of {}: {other:?}", bytes.len()),
        }
    }
    assert!(parse_dir(&bytes, MAX_MODULES).is_ok());
}

#[test]
fn bytes_after_the_terminator_are_ignored() {
    let mut bytes = rich().to_bytes();
    let expected = parse_dir(&bytes, MAX_MODULES).unwrap();
    bytes.extend_from_slice(&[0; 4096]);
    assert_eq!(parse_dir(&bytes, MAX_MODULES).unwrap(), expected);
}

#[test]
fn wrong_required_ids() {
    let fixture = DirFixture { modules: module_names(&["Module1"]), ..DirFixture::default() };
    let good = fixture.to_bytes();

    let mut bytes = good.clone();
    put_u16(&mut bytes, 0, 0x0099);
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::UnexpectedId { at: 0, expected: id::SYS_KIND, found: 0x0099 }));

    let mut bytes = good.clone();
    put_u16(&mut bytes, NAME_AT_WITHOUT_COMPAT, id::DOC_STRING);
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::UnexpectedId { at: NAME_AT_WITHOUT_COMPAT, expected: id::NAME, found: id::DOC_STRING }));

    let terminator_at = good.len() - DIR_TERMINATOR_LEN;
    let mut bytes = good.clone();
    put_u16(&mut bytes, terminator_at, 0x0011);
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::UnexpectedId { at: terminator_at, expected: id::DIR_TERMINATOR, found: 0x0011 }));

    // MODULE Terminator, then MODULETYPE, counted back from the dir Terminator.
    let module_terminator_at = terminator_at - 6;
    let mut bytes = good.clone();
    put_u16(&mut bytes, module_terminator_at, id::DIR_TERMINATOR);
    assert_eq!(
        parse_dir(&bytes, MAX_MODULES),
        Err(DirError::UnexpectedId { at: module_terminator_at, expected: id::MODULE_TERMINATOR, found: id::DIR_TERMINATOR })
    );
    let type_at = module_terminator_at - 6;
    let mut bytes = good;
    put_u16(&mut bytes, type_at, 0x0023);
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::UnknownRecord { at: type_at, found: 0x0023 }));
}

#[test]
fn unknown_reference_kind() {
    let base = DirFixture::default().to_bytes();
    let reference_at = base.len() - DIR_TERMINATOR_LEN - MODULES_HEADER_LEN;
    let unnamed = |kind| ReferenceFixture { name: None, ..ReferenceFixture::new("x", kind) };

    let mut bytes = DirFixture { references: vec![unnamed(ReferenceKind::Registered)], ..DirFixture::default() }.to_bytes();
    assert_eq!(bytes[reference_at..reference_at + 2], id::REFERENCE_REGISTERED.to_le_bytes());
    put_u16(&mut bytes, reference_at, 0x0099);
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::UnknownRecord { at: reference_at, found: 0x0099 }));

    // REFERENCEORIGINAL must wrap a REFERENCECONTROL: Id 2 + size 4 + libid, then 0x002F.
    let reference = unnamed(ReferenceKind::Original);
    let control_at = reference_at + 6 + reference.libid.len();
    let mut bytes = DirFixture { references: vec![reference], ..DirFixture::default() }.to_bytes();
    put_u16(&mut bytes, control_at, id::REFERENCE_REGISTERED);
    assert_eq!(
        parse_dir(&bytes, MAX_MODULES),
        Err(DirError::UnexpectedId { at: control_at, expected: id::REFERENCE_CONTROL, found: id::REFERENCE_REGISTERED })
    );
}

#[test]
fn bad_sizes() {
    let fixture = DirFixture { modules: module_names(&["Module1"]), ..DirFixture::default() };
    let good = fixture.to_bytes();
    let name_size_at = NAME_AT_WITHOUT_COMPAT + 2;

    // ProjectName above the spec's 128-byte cap, and a huge size.
    for size in [129, u32::MAX] {
        let mut bytes = good.clone();
        put_u32(&mut bytes, name_size_at, size);
        assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::BadSize { at: NAME_AT_WITHOUT_COMPAT, id: id::NAME }), "size {size}");
    }

    // PROJECTCODEPAGE must declare Size 2.
    let code_page_at = NAME_AT_WITHOUT_COMPAT - 8;
    let mut bytes = good.clone();
    put_u32(&mut bytes, code_page_at + 2, 4);
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::BadSize { at: code_page_at, id: id::CODE_PAGE }));

    // A module name above MAX_FIELD is BadSize; within the cap but past the end, Truncated.
    let module_at = DirFixture::default().to_bytes().len() - DIR_TERMINATOR_LEN;
    let mut bytes = good.clone();
    put_u32(&mut bytes, module_at + 2, u32::try_from(MAX_FIELD + 1).unwrap());
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::BadSize { at: module_at, id: id::MODULE_NAME }));
    let mut bytes = good.clone();
    put_u32(&mut bytes, module_at + 2, u32::try_from(MAX_FIELD).unwrap());
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::Truncated { at: module_at + 6 }));

    // MODULENAMEUNICODE must have an even size: "Module1" is 7 bytes, its UTF-16 copy 14.
    let unicode_at = module_at + 6 + 7;
    let mut bytes = good;
    assert_eq!(bytes[unicode_at..unicode_at + 2], id::MODULE_NAME_UNICODE.to_le_bytes());
    put_u32(&mut bytes, unicode_at + 2, 13);
    assert_eq!(parse_dir(&bytes, MAX_MODULES), Err(DirError::BadSize { at: unicode_at, id: id::MODULE_NAME_UNICODE }));
}

#[test]
fn module_count_above_the_limit() {
    let fixture = DirFixture { modules: module_names(&["A", "B", "C"]), ..DirFixture::default() };
    let bytes = fixture.to_bytes();
    let kept = |max| {
        let info = parse_dir(&bytes, max).unwrap();
        (info.modules.iter().map(|m| m.name.clone()).collect::<Vec<_>>(), info.truncated_modules)
    };
    assert_eq!(kept(3), (vec![b"A".to_vec(), b"B".to_vec(), b"C".to_vec()], None));
    assert_eq!(kept(2), (vec![b"A".to_vec(), b"B".to_vec()], Some(3)));
    assert_eq!(kept(0), (vec![], Some(3)));
}

#[test]
fn declared_count_above_the_records_is_truncated() {
    let mut bytes = DirFixture { modules: module_names(&["A"]), ..DirFixture::default() }.to_bytes();
    let count_at = DirFixture::default().to_bytes().len() - DIR_TERMINATOR_LEN - MODULES_HEADER_LEN + 6;
    put_u16(&mut bytes, count_at, u16::MAX);
    assert!(matches!(parse_dir(&bytes, MAX_MODULES), Err(DirError::UnexpectedId { expected: id::MODULE_NAME, .. })));
}

#[test]
fn error_messages_are_plain() {
    assert_eq!(DirError::Truncated { at: 5 }.to_string(), "the dir stream is cut short at byte 5");
    assert_eq!(
        DirError::UnexpectedId { at: 0, expected: 1, found: 0x99 }.to_string(),
        "the dir stream has record 0x0099 at byte 0 where record 0x0001 belongs"
    );
    assert_eq!(DirError::BadSize { at: 38, id: 4 }.to_string(), "record 0x0004 at byte 38 declares a size that isn't allowed");
}
