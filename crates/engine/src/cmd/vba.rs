//! VBA projects, read only: `vba.modules` lists the modules of the active workbook's VBA project
//! and `vba.source` returns one module's source. The project is read from the kept
//! `vbaProject.bin` on the first request ([`crate::vba_cache`]) and never run; both commands are
//! queries, so they change nothing and never open a dialog.
use super::{CommandSpec, bad, cmd, has_doc};
use crate::{EngineError, Result, Session};
use gridcraft_vba::{Diag, VbaModule, VbaProject};
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
#[cfg(test)]
mod props;
#[cfg(test)]
mod tests;
