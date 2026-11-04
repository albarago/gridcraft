//! VBA projects, read only: `vba.modules` lists the modules of the active workbook's VBA project,
//! `vba.source` returns one module's source and `vba.analyze` reports what the source uses. The
//! project is read from the kept `vbaProject.bin` on the first request ([`crate::vba_cache`]) and
//! never run; all three commands are queries, so they change nothing and never open a dialog.
use super::{CommandSpec, bad, cmd, has_doc};
use crate::{EngineError, Result, Session};
use gridcraft_vba::{CapSpan, CapabilityManifest, CompatibilityReport, Diag, Flag, ProcInfo, Reason, VbaModule, VbaProject, VerdictRule};
use serde_json::{Value as Json, json};
pub fn specs() -> Vec<CommandSpec> {
    vec![
        cmd!(
            query "vba.modules",
            "List VBA Modules",
            [],
            None,
            "{} → {project: {name, codePage, signed, diagnostics}, modules: [{name, kind: standard|class|document|form, codeName, lines, diagnostics}]} lists the workbook's VBA modules in project order, with what reading them found (each diagnostic is {code, severity: info|warning|error, module, message}). A workbook without a VBA project gives no modules. Nothing runs: VBA is read, never executed",
            has_doc,
            modules
        ),
        cmd!(
            query "vba.source",
            "VBA Module Source",
            [],
            None,
            "{module} → {module, source}: the full source of the VBA module with this name (matched ignoring case; a module's name is tried before its code name). Nothing runs: VBA is read, never executed",
            has_doc,
            source
        ),
        cmd!(
            query "vba.analyze",
            "Analyse VBA Compatibility",
            [],
            None,
            "{} → {verdict: noMacros|usesUnsupportedFeatures|analysisOnly|wouldRunInFuturePhase, reasons: [{rule, flag?, module?, line?, member?}], procedures: [{module, name, kind, params: [{name, type}], startLine, endLine, marks}], members: [{object, member, count, status}], flags: [{kind, module, startLine, endLine, progId?}], unparsed, unreadable, manifest: {categories: [{category, present, spans: [{module, startLine, startCol, endLine, endCol, construct}]}], dynamic: [{module, span, construct}], incomplete, incompleteReasons: [{kind, module?}]}}: what the workbook's VBA source uses and how ready it is for GridCraft. A workbook without a VBA project gives the verdict noMacros. Nothing runs: VBA is read, never executed",
            has_doc,
            analyze
        ),
    ]
}
/// The active workbook's VBA project, from the cache: the empty project when it has none.
fn project(s: &mut Session) -> Result<std::sync::Arc<VbaProject>> {
    let wb = s.doc()?.wb.clone();
    Ok(s.vba_cache.project(&wb))
}
fn diag_json(d: &Diag) -> Json {
    json!({"code": d.code.as_str(), "severity": d.code.severity().as_str(), "module": d.module, "message": d.message})
}
fn diags_json(ds: &[Diag]) -> Json {
    Json::Array(ds.iter().map(diag_json).collect())
}
/// The number of lines in `source`: each CRLF, lone LF and lone CR ends one line, a last line
/// without a terminator counts too, and empty source has none (Req 9.2).
pub(crate) fn line_count(source: &str) -> usize {
    let mut lines: usize = 0;
    let mut bytes = source.bytes().peekable();
    let mut last = None;
    while let Some(b) = bytes.next() {
        last = Some(b);
        match b {
            b'\r' => {
                lines = lines.saturating_add(1);
                if bytes.peek() == Some(&b'\n') {
                    bytes.next();
                    last = Some(b'\n');
                }
            }
            b'\n' => lines = lines.saturating_add(1),
            _ => {}
        }
    }
    match last {
        None | Some(b'\r' | b'\n') => lines,
        Some(_) => lines.saturating_add(1),
    }
}
fn modules(s: &mut Session, _: &Json) -> Result<Json> {
    let p = project(s)?;
    let modules: Vec<Json> = p
        .modules
        .iter()
        .map(|m| {
            json!({
                "name": m.name,
                "kind": m.kind.as_str(),
                "codeName": m.code_name,
                "lines": line_count(&m.source),
                "diagnostics": diags_json(&m.diagnostics),
            })
        })
        .collect();
    Ok(json!({
        "project": {"name": p.name, "codePage": p.code_page, "signed": p.signed, "diagnostics": diags_json(&p.diagnostics)},
        "modules": modules,
    }))
}
/// The module `wanted` names, ignoring case: the first whose display name matches, else the first
/// whose code name does. A name match wins over a code-name match on another module.
pub(crate) fn find_module<'a>(project: &'a VbaProject, wanted: &str) -> Option<&'a VbaModule> {
    let wanted = wanted.to_lowercase();
    project.modules.iter().find(|m| m.name.to_lowercase() == wanted).or_else(|| project.modules.iter().find(|m| m.code_name.to_lowercase() == wanted))
}
/// `vba.source`. A missing, non-text or empty `module` is a [`EngineError::BadParams`]; a name no
/// module has (also when there's no project) is an [`EngineError::Other`] reading
/// "unknown module: X".
fn source(s: &mut Session, p: &Json) -> Result<Json> {
    let wanted = match p.get("module") {
        Some(Json::String(m)) if !m.is_empty() => m.as_str(),
        Some(Json::String(_)) => return Err(bad("vba.source", "`module` is empty")),
        Some(_) => return Err(bad("vba.source", "`module` must be text")),
        None => return Err(bad("vba.source", "missing `module`")),
    };
    let project = project(s)?;
    let Some(m) = find_module(&project, wanted) else {
        return Err(EngineError::Other(format!("unknown module: {wanted}")));
    };
    Ok(json!({"module": m.name, "source": m.source}))
}
/// `vba.analyze`: the active workbook's compatibility report (Req 13.12, 13.13), from the cache.
fn analyze(s: &mut Session, _: &Json) -> Result<Json> {
    let wb = s.doc()?.wb.clone();
    let report = s.vba_cache.report(&wb);
    Ok(report_json(&report))
}
/// Inserts `key: value` into `obj` when `value` is `Some`; optional fields are left out, not null.
fn put<T: Into<Json>>(obj: &mut serde_json::Map<String, Json>, key: &str, value: Option<T>) {
    if let Some(v) = value {
        obj.insert(key.to_owned(), v.into());
    }
}
fn span_json(s: &CapSpan) -> Json {
    json!({"module": s.module, "startLine": s.start_line, "startCol": s.start_col, "endLine": s.end_line, "endCol": s.end_col, "construct": s.construct})
}
fn reason_json(r: &Reason) -> Json {
    let mut o = serde_json::Map::new();
    o.insert("rule".to_owned(), r.rule.as_str().into());
    if let VerdictRule::Flagged(kind) = r.rule {
        o.insert("flag".to_owned(), kind.as_str().into());
    }
    put(&mut o, "module", r.module.clone());
    put(&mut o, "line", r.line);
    put(&mut o, "member", r.member.clone());
    Json::Object(o)
}
fn proc_json(p: &ProcInfo) -> Json {
    let params: Vec<Json> = p.params.iter().map(|a| json!({"name": a.name, "type": a.ty})).collect();
    let marks: Vec<&str> = p.marks.iter().map(|m| m.as_str()).collect();
    json!({"module": p.module, "name": p.name, "kind": p.kind.as_str(), "params": params, "startLine": p.start_line, "endLine": p.end_line, "marks": marks})
}
fn flag_json(f: &Flag) -> Json {
    let mut o = serde_json::Map::new();
    o.insert("kind".to_owned(), f.kind.as_str().into());
    o.insert("module".to_owned(), f.module.clone().into());
    o.insert("startLine".to_owned(), f.start_line.into());
    o.insert("endLine".to_owned(), f.end_line.into());
    put(&mut o, "progId", f.prog_id.as_ref().map(|p| p.as_str().to_owned()));
    Json::Object(o)
}
fn manifest_json(m: &CapabilityManifest) -> Json {
    let categories: Vec<Json> = m
        .categories
        .iter()
        .map(|c| json!({"category": c.category.as_str(), "present": c.present, "spans": c.spans.iter().map(span_json).collect::<Vec<_>>()}))
        .collect();
    let dynamic: Vec<Json> = m
        .dynamic
        .iter()
        .map(|d| {
            let s = &d.span;
            json!({"module": d.module, "span": {"startLine": s.start_line, "startCol": s.start_col, "endLine": s.end_line, "endCol": s.end_col}, "construct": d.construct})
        })
        .collect();
    let reasons: Vec<Json> = m
        .incomplete_reasons
        .iter()
        .map(|r| {
            let mut o = serde_json::Map::new();
            o.insert("kind".to_owned(), r.as_str().into());
            put(&mut o, "module", r.module().map(str::to_owned));
            Json::Object(o)
        })
        .collect();
    json!({"categories": categories, "dynamic": dynamic, "incomplete": m.incomplete, "incompleteReasons": reasons})
}
/// The JSON of `vba.analyze`. Every enum is written with its stable `as_str` id; optional fields
/// (`flag`, `module`, `line`, `member` of a reason, `progId` of a flag, `module` of an
/// incomplete reason) are left out when there's nothing to say.
pub(crate) fn report_json(r: &CompatibilityReport) -> Json {
    let members: Vec<Json> =
        r.members.iter().map(|m| json!({"object": m.object, "member": m.member, "count": m.count, "status": m.status.as_str()})).collect();
    json!({
        "verdict": r.verdict.as_str(),
        "reasons": r.reasons.iter().map(reason_json).collect::<Vec<_>>(),
        "procedures": r.procedures.iter().map(proc_json).collect::<Vec<_>>(),
        "members": members,
        "flags": r.flags.iter().map(flag_json).collect::<Vec<_>>(),
        "unparsed": r.unparsed,
        "unreadable": r.unreadable,
        "manifest": manifest_json(&r.manifest),
    })
}
#[cfg(test)]
mod props;
#[cfg(test)]
mod tests;
