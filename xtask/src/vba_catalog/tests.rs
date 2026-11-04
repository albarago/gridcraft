//! Synthetic pages in the VBA-Docs *shape* (front matter, `## Syntax`, `## Parameters` table,
//! `## Return value`). The objects, members and sentences are made up; no VBA-Docs text is used.
use super::{Kind, NOTE_MAX_CHARS, Page, Param, SEED_NOTE, Summary, TRANCHE_1, build, parse_page};

const SHA: &str = "0123456789abcdef0123456789abcdef01234567";

const PROPERTY: &str = "---
title: Widget.Size property (Excel)
keywords: none
api_name:
- Excel.Widget.Size
ms.date: 01/01/2020
---

# Widget.Size property (Excel)

Zorbles the frobnicated widget quickly. Read/write **Long**.

## Syntax

_expression_.**Size**

_expression_ A placeholder sentence about the widget.

## Remarks

This remark sentence must never reach the catalog.
";

const METHOD: &str = "---
title: Widget.Spin method (Excel)
api_name:
- Excel.Widget.Spin
---

# Widget.Spin method (Excel)

Twirls the widget around twice.

## Syntax

_expression_.**Spin** (_Turns_, _Direction_)

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Turns_|Required| **Long**|How many turns the widget makes.|
| _Direction_|Optional| **[XlSpin](Excel.XlSpin.md)**|Which way it twirls.|

## Return value

A **Widget** object that has finished twirling.

## Example

```vb
Widget.Spin 3
```
";

const EVENT: &str = "---
title: Gadget.Wobble event (Excel)
api_name:
- Excel.Gadget.Wobble
---

# Gadget.Wobble event (Excel)

Happens whenever the gadget wobbles.

## Syntax

_expression_.**Wobble** (_Cancel_)

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set it to stop the wobble.|

## Return value

**Nothing**
";

const OVERVIEW: &str = "---
title: Widget object (Excel)
api_name:
- Excel.Widget
---

# Widget object (Excel)

A widget is a pretend thing.
";

fn page(name: &str, text: &str) -> Page {
    (name.to_string(), Ok(text.to_string()))
}

fn pages() -> Vec<Page> {
    vec![
        page("Excel.Widget.Spin.md", METHOD),
        page("Excel.Widget(object).md", OVERVIEW),
        page("Excel.Gadget.Wobble.md", EVENT),
        page("Excel.Widget.Size.md", PROPERTY),
    ]
}

fn run(existing: Option<&str>) -> (String, Summary) {
    build(&pages(), existing, SHA).unwrap()
}

/// The data rows (no comment lines).
fn rows(text: &str) -> Vec<&str> {
    text.lines().filter(|l| !l.starts_with('#')).collect()
}

#[test]
fn property_page() {
    let m = parse_page(PROPERTY).unwrap();
    assert_eq!((m.object.as_str(), m.member.as_str(), m.kind), ("Widget", "Size", Kind::Property));
    assert!(m.params.is_empty());
    assert_eq!(m.returns, "Long");
}

#[test]
fn method_page_with_required_and_optional_params() {
    let m = parse_page(METHOD).unwrap();
    assert_eq!(m.kind, Kind::Method);
    assert_eq!(
        m.params,
        vec![
            Param { name: "Turns".into(), optional: false, ty: "Long".into() },
            Param { name: "Direction".into(), optional: true, ty: "XlSpin".into() },
        ]
    );
    assert_eq!(m.returns, "Widget");
}

#[test]
fn event_page() {
    let m = parse_page(EVENT).unwrap();
    assert_eq!((m.object.as_str(), m.member.as_str(), m.kind), ("Gadget", "Wobble", Kind::Event));
    assert_eq!(m.params, vec![Param { name: "Cancel".into(), optional: false, ty: "Boolean".into() }]);
    assert_eq!(m.returns, "none");
}

#[test]
fn names_fall_back_to_api_name_and_params_to_the_syntax_line() {
    let text = "---\ntitle: Twist method (Excel)\napi_name:\n- Excel.Knob.Twist\n---\n# Twist method (Excel)\nTurns a knob.\n\
                ## Syntax\n_expression_.**Twist** (_Angle_, [_Speed_])\n";
    let m = parse_page(text).unwrap();
    assert_eq!((m.object.as_str(), m.member.as_str()), ("Knob", "Twist"));
    assert_eq!(
        m.params,
        vec![Param { name: "Angle".into(), optional: false, ty: String::new() }, Param { name: "Speed".into(), optional: true, ty: String::new() }]
    );
    assert_eq!(m.returns, "none");
}

#[test]
fn numbered_ranges_arrays_and_alternate_header_rows() {
    let text = "---\ntitle: Box.Pack method (Excel)\n---\n# Box.Pack method (Excel)\nPuts things in a box.\n\
                ## Parameters\n|Parameter|Required/Optional|Data type|Description|\n|:--|:--|:--|:--|\n\
                | _Items()_|Required| **String**|Things.|\n| _Arg2_ &ndash; _Arg4_|Optional| **Variant**|More things.|\n\
                ## Return value\nLong\n";
    let m = parse_page(text).unwrap();
    let got: Vec<(&str, bool, &str)> = m.params.iter().map(|p| (p.name.as_str(), p.optional, p.ty.as_str())).collect();
    assert_eq!(got, vec![("Items", false, "String()"), ("Arg2", true, "Variant"), ("Arg3", true, "Variant"), ("Arg4", true, "Variant")]);
    assert_eq!(m.returns, "Long");
    let reversed = text.replace("_Arg2_ &ndash; _Arg4_", "Arg4 - Arg2");
    assert!(parse_page(&reversed).unwrap_err().contains("bad parameter name"));
}

#[test]
fn bracketed_and_lower_case_types_and_void_returns() {
    let text = "---\ntitle: Box.Tie method (Excel)\n---\n# Box.Tie method (Excel)\nTies it.\n## Parameters\n\
                |Name|Required/Optional|Data type|Description|\n|:--|:--|:--|:--|\n| _Knots_|Required| **[INT]**|Knots.|\n\
                | _Style_|Optional| **xlKnotStyle**|Style.|\n| _Odd_|Optional| two words|Odd.|\n## Return value\n**VOID**\n";
    let m = parse_page(text).unwrap();
    let types: Vec<&str> = m.params.iter().map(|p| p.ty.as_str()).collect();
    assert_eq!(types, vec!["INT", "xlKnotStyle", "?"]);
    assert_eq!(m.returns, "none");
}

#[test]
fn return_type_from_the_summary_line() {
    let page = |summary: &str| format!("---\ntitle: Box.Lid property (Excel)\n---\n# Box.Lid property (Excel)\n{summary}\n## Syntax\n");
    let lid = |summary: &str| parse_page(&page(summary)).unwrap().returns;
    assert_eq!(lid("Used alone, this property returns a **[Lid](Excel.Lid(object).md)** object for the box. Read-only."), "Lid");
    assert_eq!(lid("Returns **True** if the box is shut."), "Boolean");
    assert_eq!(lid("Returns the thing on top. Read-only."), "?");
    assert_eq!(lid("Whether the lid creaks."), "?");
}

#[test]
fn overview_and_malformed_pages_are_skipped_with_a_reason() {
    assert_eq!(parse_page(OVERVIEW).unwrap_err(), "overview page (object)");
    assert_eq!(parse_page("# Widget.Size property (Excel)\n").unwrap_err(), "no front matter");
    assert_eq!(parse_page("---\ntitle: Widget.Size property (Excel)\n").unwrap_err(), "unterminated front matter");
    assert_eq!(parse_page("---\ntitle: Widget.Size (Excel)\n---\n").unwrap_err(), "no member kind in title");
    assert_eq!(parse_page("---\ntitle: Widget.Size gizmo (Excel)\n---\n").unwrap_err(), "unknown page kind `gizmo`");
    assert_eq!(parse_page("---\ntitle: Size property (Excel)\n---\n").unwrap_err(), "no object or member name");
    assert!(parse_page("---\ntitle: Size property (Excel Graph)\n---\n").unwrap_err().contains("Excel Graph"));
    let bad_cell = METHOD.replace("|Required|", "|Sometimes|");
    assert!(parse_page(&bad_cell).unwrap_err().contains("Required/Optional"));

    let mut all = pages();
    all.push(page("Excel.Broken.md", "no front matter here"));
    all.push(("Excel.Unreadable.md".into(), Err("not UTF-8".into())));
    let (text, s) = build(&all, None, SHA).unwrap();
    assert_eq!(s.pages_read, 6);
    assert_eq!(s.members, 3);
    assert_eq!(
        s.skipped,
        vec![
            ("Excel.Broken.md".to_string(), "no front matter".to_string()),
            ("Excel.Unreadable.md".to_string(), "not UTF-8".to_string()),
            ("Excel.Widget(object).md".to_string(), "overview page (object)".to_string()),
        ]
    );
    assert_eq!(rows(&text).len(), 3);
}

#[test]
fn output_is_sorted_with_header_and_new_members_default_to_unreviewed() {
    let (text, s) = run(None);
    let expected = format!(
        "# generated by cargo xtask vba-catalog from MicrosoftDocs/VBA-Docs {SHA} (CC-BY-4.0); names and signatures only\n\
         # object\tmember\tkind\tparams\treturns\tstatus\tnote\n\
         Gadget\tWobble\tevent\tCancel:req:Boolean\tnone\tunreviewed\t\n\
         Widget\tSize\tproperty\t\tLong\tunreviewed\t\n\
         Widget\tSpin\tmethod\tTurns:req:Long;Direction:opt:XlSpin\tWidget\tunreviewed\t\n"
    );
    assert_eq!(text, expected);
    assert_eq!((s.added, s.kept), (3, 0));
    assert!(s.removed.is_empty());
}

#[test]
fn merge_keeps_status_and_note_and_reports_removed_members() {
    let existing = "# old header\n\
                    widget\tsize\tproperty\t\tLong\tsupported\tdone in VBA-4.2\n\
                    Widget\tSpin\tmethod\t\tnone\tnotPlanned\tneeds a spinner\n\
                    Widget\tGone\tmethod\t\tnone\tplanned\tdropped upstream\n";
    let (text, s) = run(Some(existing));
    let r = rows(&text);
    assert_eq!(r.first().copied(), Some("Gadget\tWobble\tevent\tCancel:req:Boolean\tnone\tunreviewed\t"));
    assert_eq!(r.get(1).copied(), Some("Widget\tSize\tproperty\t\tLong\tsupported\tdone in VBA-4.2"));
    assert_eq!(r.get(2).copied(), Some("Widget\tSpin\tmethod\tTurns:req:Long;Direction:opt:XlSpin\tWidget\tnotPlanned\tneeds a spinner"));
    assert!(r.iter().all(|l| !l.contains("Gone")));
    assert_eq!((s.kept, s.added), (2, 1));
    assert_eq!(s.removed, vec!["Widget.Gone (method)".to_string()]);
}

#[test]
fn byte_stable_across_runs_and_input_order() {
    let (first, _) = run(None);
    let mut reversed = pages();
    reversed.reverse();
    let (second, _) = build(&reversed, None, SHA).unwrap();
    assert_eq!(first.as_bytes(), second.as_bytes());
    // Feeding the output back as the status file changes nothing.
    let (third, s) = run(Some(&first));
    assert_eq!(first.as_bytes(), third.as_bytes());
    assert_eq!((s.kept, s.added), (3, 0));
}

#[test]
fn note_longer_than_500_characters_is_an_error_naming_the_member() {
    let long = "é".repeat(NOTE_MAX_CHARS + 1);
    let existing = format!("Widget\tSize\tproperty\t\tLong\tplanned\t{long}\n");
    let err = build(&pages(), Some(&existing), SHA).unwrap_err();
    assert!(err.contains("Widget.Size (property)") && err.contains("501"), "{err}");
    // Exactly 500 is fine.
    let ok = format!("Widget\tSize\tproperty\t\tLong\tplanned\t{}\n", "é".repeat(NOTE_MAX_CHARS));
    assert!(build(&pages(), Some(&ok), SHA).is_ok());
}

#[test]
fn tranche_1_seed_plans_only_unreviewed_properties_and_methods() {
    // Made-up pages that happen to carry tranche-1 names.
    let member = |object: &str, name: &str, kind: &str| {
        let title = format!("{object}.{name} {kind} (Excel)");
        page(&format!("Excel.{object}.{name}.{kind}.md"), &format!("---\ntitle: {title}\n---\n# {title}\nA made-up line.\n"))
    };
    let mut all = pages();
    all.extend([
        member("Range", "Value", "property"),
        member("Range", "Copy", "method"),
        member("Font", "Bold", "property"),
        member("Interior", "Color", "property"),
        member("Workbook", "Activate", "method"),
        member("Workbook", "Activate", "event"),
    ]);
    let existing = "Range\tCopy\tmethod\t\tnone\tnotPlanned\tpeople decided\n\
                    Font\tBold\tproperty\t\t?\tunreviewed\tkeep this note\n\
                    Interior\tColor\tproperty\t\t?\tsupported\t\n";
    let (text, s) = build(&all, Some(existing), SHA).unwrap();
    let status =
        |prefix: &str| rows(&text).into_iter().find(|l| l.starts_with(prefix)).map(|l| l.splitn(6, '\t').nth(5).unwrap_or_default().to_string());
    assert_eq!(status("Range\tValue\tproperty"), Some(format!("planned\t{SEED_NOTE}")), "new member is seeded");
    assert_eq!(status("Workbook\tActivate\tmethod"), Some(format!("planned\t{SEED_NOTE}")));
    assert_eq!(status("Workbook\tActivate\tevent"), Some("unreviewed\t".into()), "events are never seeded");
    assert_eq!(status("Range\tCopy\tmethod"), Some("notPlanned\tpeople decided".into()), "a human status is kept");
    assert_eq!(status("Font\tBold\tproperty"), Some("planned\tkeep this note".into()), "an unreviewed row keeps its note");
    assert_eq!(status("Interior\tColor\tproperty"), Some("supported\t".into()));
    assert_eq!(status("Widget\tSize\tproperty"), Some("unreviewed\t".into()));
    assert_eq!(s.seeded, 3);
    let total: usize = TRANCHE_1.iter().map(|(_, m)| m.len()).sum();
    assert_eq!(s.seed_unmatched.len(), total - 5, "{:?}", s.seed_unmatched);
    assert!(s.seed_unmatched.contains(&"Range.Value2".to_string()) && !s.seed_unmatched.contains(&"Range.Value".to_string()));
    // Seeding is part of the fixed point: feeding the output back changes nothing.
    let (again, s2) = build(&all, Some(&text), SHA).unwrap();
    assert_eq!(text.as_bytes(), again.as_bytes());
    assert_eq!(s2.seeded, 0);
}

#[test]
fn tranche_1_names_are_identifiers_and_unique() {
    let mut seen = std::collections::BTreeSet::new();
    for (object, members) in TRANCHE_1 {
        for member in *members {
            assert!(super::is_ident(object) && super::is_ident(member), "{object}.{member}");
            assert!(seen.insert(format!("{object}.{member}").to_ascii_lowercase()), "{object}.{member} listed twice");
        }
    }
}

#[test]
fn bad_status_and_duplicate_rows_are_errors() {
    let bad = "Widget\tSize\tproperty\t\tLong\tnot planned\t\n";
    assert!(build(&pages(), Some(bad), SHA).unwrap_err().contains("unknown status"));
    let unreviewed = "Widget\tSize\tproperty\t\tLong\tunreviewed\t\n";
    assert!(build(&pages(), Some(unreviewed), SHA).is_ok());
    let dup = "Widget\tSize\tproperty\t\tLong\tplanned\t\nWIDGET\tSIZE\tproperty\t\tLong\tplanned\t\n";
    assert!(build(&pages(), Some(dup), SHA).unwrap_err().contains("listed twice"));
}

#[test]
fn no_prose_is_copied() {
    let (text, _) = run(None);
    for sentence in [
        "Zorbles the frobnicated widget quickly",
        "A placeholder sentence",
        "This remark sentence",
        "Twirls the widget",
        "How many turns",
        "Which way it twirls",
        "finished twirling",
        "Widget.Spin 3",
        "Happens whenever",
        "Set it to stop",
        "pretend thing",
    ] {
        assert!(!text.contains(sentence), "catalog contains page prose: {sentence}");
    }
    // Every data field is a name, a signature token or a status: no spaces anywhere in a row.
    assert!(rows(&text).iter().all(|l| !l.contains(' ')));
}
