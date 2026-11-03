//! View Code (Macros dialog › View Code): a workbook's VBA project, read only. The modules are
//! listed on the left, grouped standard, class, document, form and sorted by name ignoring case;
//! the selected module's source shows on the right in the code editor, read only, with VBA
//! colours. What reading the project found (signatures, stomping, parse failures) shows in a
//! banner above the source. Everything comes from the `vba.modules` and `vba.source` queries:
//! nothing here changes the workbook, and VBA never runs.
//!
//! From the keyboard the list is one Tab stop (the selected module): Up, Down, Home and End move
//! the selection across the groups, and Tab and Shift+Tab move between the list and the source.
//! A screen reader hears each module as "name, kind", the group titles as headings and the source
//! as "Source of" the module.

use egui::{Key, Ui, vec2};
use serde_json::{Value as Json, json};

use crate::SheetApp;
use crate::code_editor::{self, FieldOptions, Lang};
use crate::dialogs::Dialog;
use crate::theme::{self, Tokens};

/// The dialog's name in [`Dialog`] and `ui.dialog.*`.
pub const DIALOG: &str = "vbaCode";

/// The module kinds `vba.modules` reports, in the order View Code groups them, with each
/// group's heading.
pub const GROUPS: [(&str, &str); 4] = [("standard", "Modules"), ("class", "Class Modules"), ("document", "Document Modules"), ("form", "Forms")];

/// The sentence the window keeps under the project's name (Req 8.8).
pub const KEPT: &str = "GridCraft keeps the VBA project when you save but does not run it.";

/// What reading the project or a module found.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Diagnostic {
    /// `info`, `warning` or `error`.
    pub severity: String,
    pub code: String,
    pub message: String,
}

/// A module in the list.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Module {
    pub name: String,
    /// One of the [`GROUPS`] kinds.
    pub kind: String,
    pub lines: u64,
    pub diagnostics: Vec<Diagnostic>,
}

/// A `vba.modules` result, read for the window.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Listing {
    pub name: String,
    pub signed: bool,
    pub diagnostics: Vec<Diagnostic>,
    /// The non-empty groups in [`GROUPS`] order: (kind, heading, modules sorted by name).
    pub groups: Vec<(&'static str, &'static str, Vec<Module>)>,
}

impl Listing {
    /// Every module, in the order the list shows them.
    pub fn modules(&self) -> impl Iterator<Item = &Module> {
        self.groups.iter().flat_map(|(_, _, m)| m.iter())
    }

    /// The module named `name` (exactly, as listed).
    pub fn module(&self, name: &str) -> Option<&Module> {
        self.modules().find(|m| m.name == name)
    }

    /// The module selected when the window opens: the first of the first non-empty group.
    pub fn first(&self) -> Option<&Module> {
        self.modules().next()
    }
}

fn diagnostics(v: &Json) -> Vec<Diagnostic> {
    let text = |d: &Json, k: &str| d.get(k).and_then(Json::as_str).unwrap_or("").to_string();
    v.as_array()
        .into_iter()
        .flatten()
        .map(|d| Diagnostic { severity: text(d, "severity"), code: text(d, "code"), message: text(d, "message") })
        .collect()
}

/// Reads a `vba.modules` result: the project's name, signature and diagnostics, and its modules
/// grouped by kind (standard, class, document, form; a kind the list doesn't know goes with the
/// standard modules) and sorted by name ignoring case. Groups without modules are left out.
pub fn listing(v: &Json) -> Listing {
    let project = &v["project"];
    let mut groups: Vec<(&'static str, &'static str, Vec<Module>)> = GROUPS.iter().map(|(k, h)| (*k, *h, Vec::new())).collect();
    for m in v["modules"].as_array().into_iter().flatten() {
        let Some(name) = m.get("name").and_then(Json::as_str) else { continue };
        let kind = m.get("kind").and_then(Json::as_str).unwrap_or("standard");
        let at = GROUPS.iter().position(|(k, _)| *k == kind).unwrap_or(0);
        if let Some((k, _, list)) = groups.get_mut(at) {
            list.push(Module {
                name: name.to_string(),
                kind: (*k).to_string(),
                lines: m.get("lines").and_then(Json::as_u64).unwrap_or(0),
                diagnostics: diagnostics(&m["diagnostics"]),
            });
        }
    }
    for (_, _, list) in &mut groups {
        // Ignoring case first; the exact name keeps the order stable when two differ in case only.
        list.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()).then_with(|| a.name.cmp(&b.name)));
    }
    groups.retain(|(_, _, list)| !list.is_empty());
    Listing {
        name: project.get("name").and_then(Json::as_str).unwrap_or("").to_string(),
        signed: project.get("signed").and_then(Json::as_bool).unwrap_or(false),
        diagnostics: diagnostics(&project["diagnostics"]),
        groups,
    }
}

/// Where Up, Down, Home and End move the selection in a list of `len` modules from `at`: one
/// up or down, or to the first or last, stopping at the ends. Headings aren't in the list, so the
/// selection runs on from one group into the next.
pub fn step(len: usize, at: usize, key: Key) -> usize {
    let last = len.saturating_sub(1);
    let to = match key {
        Key::ArrowUp => at.saturating_sub(1),
        Key::ArrowDown => at.saturating_add(1),
        Key::Home => 0,
        Key::End => last,
        _ => at,
    };
    to.min(last)
}
/// What a module of `kind` is called in its screen-reader label (English; see [`item_label`]).
pub fn kind_word(kind: &str) -> &'static str {
    match kind {
        "class" => "class module",
        "document" => "document module",
        "form" => "form",
        _ => "module",
    }
}
/// A module's screen-reader label: its name and kind, as in "Sheet1, document module".
pub fn item_label(lang: crate::i18n::Language, name: &str, kind: &str) -> String {
    format!("{name}, {}", lang.tr(kind_word(kind)))
}
/// Whether the active workbook carries a VBA project: `view.macros` lists it as a `vba` item, as
/// the Macros dialog shows it.
pub fn has_project(app: &mut SheetApp) -> bool {
    app.run("view.macros", json!({}))
        .ok()
        .and_then(|v| v.as_array().map(|items| items.iter().any(|it| it.get("kind").and_then(Json::as_str) == Some("vba"))))
        .unwrap_or(false)
}

/// Opens View Code on the active workbook's VBA project, its first module selected. Nothing when
/// the workbook has no project (the reason is shown).
pub fn open(app: &mut SheetApp, _params: &Json) -> Option<Dialog> {
    let title = "View Code";
    if !has_project(app) {
        app.message = Some((title.into(), "This workbook has no VBA project.".into()));
        return None;
    }
    let result = match app.run("vba.modules", json!({})) {
        Ok(v) => v,
        Err(e) => {
            app.message = Some((title.into(), crate::clean_error(&e)));
            return None;
        }
    };
    let first = listing(&result).first().map(|m| m.name.clone()).unwrap_or_default();
    let doc = app.session.active().map_or(0, |d| d.uid);
    let mut d = Dialog::custom(DIALOG, title, json!({"doc": doc, "module": first}));
    d.result = Some(result);
    Some(d)
}

/// The source of the selected module, read once per selection with `vba.source` and kept in the
/// dialog (`source`, for `sourceOf`); a failure is kept as `sourceError`.
fn load_source(app: &mut SheetApp, d: &mut Dialog, module: &str) {
    if d.values.get("sourceOf").and_then(Json::as_str) == Some(module) {
        return;
    }
    d.values.insert("sourceOf".into(), json!(module));
    match app.run("vba.source", json!({"module": module})) {
        Ok(v) => {
            d.values.insert("source".into(), json!(v["source"].as_str().unwrap_or("")));
            d.values.remove("sourceError");
        }
        Err(e) => {
            d.values.insert("source".into(), json!(""));
            d.values.insert("sourceError".into(), json!(crate::clean_error(&e)));
        }
    }
}

/// The banner's colour for a severity, from the theme.
fn severity_color(t: &Tokens, severity: &str) -> egui::Color32 {
    match severity {
        "error" => t.danger,
        "warning" => t.notice_warn_border,
        _ => t.notice_info_border,
    }
}

/// The diagnostics banner: one wrapped line each, its severity in its colour.
fn banner(app: &SheetApp, ui: &mut Ui, items: &[&Diagnostic]) {
    if items.is_empty() {
        return;
    }
    let t = Tokens::get(ui.ctx());
    let lang = app.ui.language;
    let warn = items.iter().any(|d| d.severity == "warning" || d.severity == "error");
    let (bg, border) = if warn { (t.notice_warn_bg, t.notice_warn_border) } else { (t.notice_info_bg, t.notice_info_border) };
    egui::Frame::new().fill(bg).stroke(egui::Stroke::new(1.0, border)).corner_radius(4.0).inner_margin(6.0).show(ui, |ui| {
        ui.set_width(ui.available_width());
        egui::ScrollArea::vertical().id_salt("vba_banner").max_height(96.0).show(ui, |ui| {
            for d in items {
                let label = match d.severity.as_str() {
                    "error" => lang.tr("Error"),
                    "warning" => lang.tr("Warning"),
                    _ => lang.tr("Info"),
                };
                let mut job = egui::text::LayoutJob::default();
                let font = theme::ui_font(12.5);
                // The border colours are too pale for text on the banner's fill: the word takes the
                // text colour, an error the danger colour.
                let word = if d.severity == "error" { t.danger } else { t.text };
                job.append(&format!("{label}: "), 0.0, egui::TextFormat::simple(font.clone(), word));
                job.append(&d.message, 0.0, egui::TextFormat::simple(font, t.text));
                ui.add(egui::Label::new(job).wrap());
            }
        });
    });
    ui.add_space(6.0);
}

/// The View Code window's body: the project line, the diagnostics banner, the module list and
/// the selected module's source, read only.
pub fn view(app: &mut SheetApp, ui: &mut Ui, d: &mut Dialog, open: &mut bool) {
    // The window shows the workbook it opened on: another workbook becoming active closes it.
    let active = app.session.active().map(|doc| doc.uid);
    if d.values.get("doc").and_then(Json::as_u64).is_some_and(|doc| Some(doc) != active) {
        *open = false;
        return;
    }
    let doc = active.unwrap_or(0);
    let t = Tokens::get(ui.ctx());
    let lang = app.ui.language;
    let list = d.result.as_ref().map(listing).unwrap_or_default();
    let mut selected = d.values.get("module").and_then(Json::as_str).unwrap_or("").to_string();
    if list.module(&selected).is_none() {
        selected = list.first().map(|m| m.name.clone()).unwrap_or_default();
    }

    // The project's name, whether it is signed, and what GridCraft does with it.
    ui.horizontal_wrapped(|ui| {
        let name = if list.name.is_empty() { lang.tr("VBA project").to_string() } else { list.name.clone() };
        ui.label(egui::RichText::new(name).strong());
        if list.signed {
            ui.label(egui::RichText::new(lang.tr("signed (not verified)")).color(t.text_dim));
        }
    });
    ui.label(egui::RichText::new(lang.tr(KEPT)).color(t.text_dim));
    ui.add_space(6.0);

    let module = list.module(&selected).cloned();
    let mut items: Vec<&Diagnostic> = list.diagnostics.iter().collect();
    if let Some(m) = &module {
        items.extend(m.diagnostics.iter());
    }
    banner(app, ui, &items);

    if let Some(m) = &module {
        load_source(app, d, &m.name);
    }
    let height = (ui.available_height() - 8.0).max(240.0);
    ui.horizontal_top(|ui| {
        // The modules, by group. The list is one Tab stop: only the selected module takes the
        // focus, and Up, Down, Home and End move the selection while it has it.
        egui::Frame::NONE.fill(t.input_bg).stroke(egui::Stroke::new(1.0, t.input_border)).inner_margin(2.0).show(ui, |ui| {
            // Top to bottom, though the row around it runs left to right.
            ui.vertical(|ui| {
                ui.set_width(200.0);
                ui.set_min_height(height);
                egui::ScrollArea::vertical().id_salt("vba_modules").max_height(height).auto_shrink([false, false]).show(ui, |ui| {
                    if list.groups.is_empty() {
                        ui.label(egui::RichText::new(lang.tr("This project has no modules.")).italics().color(t.text_dim));
                    }
                    let mut items: Vec<(String, egui::Response)> = Vec::new();
                    for (kind, heading, modules) in &list.groups {
                        let h = ui.label(egui::RichText::new(lang.tr(heading)).font(theme::ui_font(11.5)).strong().color(t.text_dim));
                        ui.ctx().accesskit_node_builder(h.id, |node| node.set_role(egui::Role::Heading));
                        for m in modules {
                            let on = m.name == selected;
                            // Only the selected module is a Tab stop, so Tab leaves the list.
                            let sense = if on { egui::Sense::click() } else { egui::Sense::CLICK };
                            let r = ui.add(egui::Button::selectable(on, m.name.as_str()).sense(sense).min_size(vec2(194.0, 20.0)));
                            let label = item_label(lang, &m.name, kind);
                            r.widget_info(|| egui::WidgetInfo::selected(egui::Role::Button, true, on, &label));
                            if r.clicked() {
                                selected = m.name.clone();
                                r.request_focus();
                            }
                            items.push((m.name.clone(), r));
                        }
                        ui.add_space(4.0);
                    }
                    // The keyboard focus as egui keeps it, whether or not the native window has the
                    // focus (a window that doesn't gets no keys anyway).
                    let focused = |r: &egui::Response| r.ctx.memory(|mem| mem.has_focus(r.id));
                    if !d.focus_initialized {
                        // The window opens with the selected module focused, so the arrows work at once.
                        // A new window's first pass is an invisible sizing pass that ignores a focus
                        // request, so this asks until the focus lands.
                        match items.iter().find(|(n, _)| *n == selected) {
                            Some((_, r)) if !focused(r) => r.request_focus(),
                            _ => d.focus_initialized = true,
                        }
                    }
                    let Some(at) = items.iter().position(|(_, r)| focused(r)) else { return };
                    let key = ui.input_mut(|i| {
                        [Key::ArrowUp, Key::ArrowDown, Key::Home, Key::End].into_iter().find(|k| i.consume_key(egui::Modifiers::NONE, *k))
                    });
                    // The list moves the selection itself: egui must not also move the focus by position.
                    ui.memory_mut(|mem| {
                        if let Some((_, r)) = items.get(at) {
                            mem.set_focus_lock_filter(r.id, egui::EventFilter { vertical_arrows: true, ..Default::default() });
                        }
                        if key.is_some() {
                            mem.move_focus(egui::FocusDirection::None);
                        }
                    });
                    let Some(key) = key else { return };
                    let to = step(items.len(), at, key);
                    if to != at
                        && let Some((name, r)) = items.get(to)
                    {
                        selected = name.clone();
                        r.request_focus();
                        r.scroll_to_me(None);
                        // The banner and the source follow on the next pass.
                        ui.ctx().request_repaint();
                    }
                });
            });
        });
        // The selected module's source.
        ui.vertical(|ui| {
            let Some(m) = &module else { return };
            if let Some(e) = d.values.get("sourceError").and_then(Json::as_str) {
                ui.label(egui::RichText::new(e).color(t.danger));
                return;
            }
            let empty = d.values.get("source").and_then(Json::as_str).is_none_or(str::is_empty);
            if empty {
                egui::Frame::new().fill(t.notice_info_bg).stroke(egui::Stroke::new(1.0, t.notice_info_border)).inner_margin(8.0).show(ui, |ui| {
                    ui.set_width(ui.available_width());
                    ui.label(egui::RichText::new(lang.tr("This module has no stored source.")).color(t.text));
                    for diag in &m.diagnostics {
                        ui.add(egui::Label::new(egui::RichText::new(&diag.message).color(severity_color(&t, &diag.severity))).wrap());
                    }
                });
                return;
            }
            let name = format!("{} {}", lang.tr("Source of"), m.name);
            let options = FieldOptions { max_height: height - 10.0, read_only: true, name: &name, ..FieldOptions::new(Lang::Vba) };
            egui::Frame::new().fill(t.input_bg).stroke(egui::Stroke::new(1.0, t.input_border)).corner_radius(4.0).inner_margin(4.0).show(ui, |ui| {
                ui.set_min_height(height - 10.0);
                // The cached text is handed over as is: the field is read only, so it stays unchanged.
                if let Some(Json::String(source)) = d.values.get_mut("source") {
                    // One field per workbook and module, so each keeps its own scroll position.
                    code_editor::code_field(ui, egui::Id::unique("gridcraft.vba_source").with(doc).with(m.name.as_str()), source, &options);
                }
            });
        });
    });
    d.values.insert("module".into(), json!(selected));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn modules_are_grouped_by_kind_and_sorted_ignoring_case() {
        let v = json!({
            "project": {"name": "VBAProject", "codePage": 1252, "signed": true, "diagnostics": [{"code": "signed", "severity": "info", "module": null, "message": "Signed"}]},
            "modules": [
                {"name": "UserForm1", "kind": "form", "codeName": "UserForm1", "lines": 3, "diagnostics": []},
                {"name": "zeta", "kind": "standard", "codeName": "zeta", "lines": 1, "diagnostics": []},
                {"name": "Sheet1", "kind": "document", "codeName": "Sheet1", "lines": 0, "diagnostics": [{"code": "stomped", "severity": "warning", "module": "Sheet1", "message": "m"}]},
                {"name": "Alpha", "kind": "standard", "codeName": "Alpha", "lines": 9, "diagnostics": []},
                {"name": "beta", "kind": "standard", "codeName": "beta", "lines": 2, "diagnostics": []},
                {"name": "ThisWorkbook", "kind": "document", "codeName": "ThisWorkbook", "lines": 0, "diagnostics": []},
            ]
        });
        let l = listing(&v);
        let order: Vec<(&str, Vec<&str>)> = l.groups.iter().map(|(k, _, m)| (*k, m.iter().map(|m| m.name.as_str()).collect())).collect();
        assert_eq!(
            order,
            vec![("standard", vec!["Alpha", "beta", "zeta"]), ("document", vec!["Sheet1", "ThisWorkbook"]), ("form", vec!["UserForm1"])]
        );
        assert_eq!(l.first().map(|m| m.name.as_str()), Some("Alpha"));
        assert!(l.signed);
        assert_eq!(l.name, "VBAProject");
        assert_eq!(l.diagnostics.len(), 1);
        assert_eq!(l.module("Sheet1").and_then(|m| m.diagnostics.first()).map(|d| d.severity.as_str()), Some("warning"));
    }

    #[test]
    fn the_first_module_comes_from_the_first_group_with_modules() {
        let v = json!({"project": {"name": "", "signed": false, "diagnostics": []}, "modules": [
            {"name": "Form", "kind": "form", "lines": 1, "diagnostics": []},
            {"name": "Klass", "kind": "class", "lines": 1, "diagnostics": []},
        ]});
        let l = listing(&v);
        assert_eq!(l.groups.iter().map(|g| g.0).collect::<Vec<_>>(), vec!["class", "form"]);
        assert_eq!(l.first().map(|m| m.name.as_str()), Some("Klass"));
        assert_eq!(listing(&json!({})), Listing::default());
    }
    #[test]
    fn arrows_move_through_every_module_and_stop_at_the_ends() {
        assert_eq!(step(5, 0, Key::ArrowUp), 0);
        assert_eq!(step(5, 0, Key::ArrowDown), 1);
        // The keys walk modules only, so Down from a group's last module reaches the next group.
        assert_eq!(step(5, 2, Key::ArrowDown), 3);
        assert_eq!(step(5, 4, Key::ArrowDown), 4);
        assert_eq!(step(5, 3, Key::Home), 0);
        assert_eq!(step(5, 1, Key::End), 4);
        assert_eq!(step(5, 2, Key::Tab), 2);
        assert_eq!(step(0, 0, Key::ArrowDown), 0);
        assert_eq!(step(1, 0, Key::ArrowDown), 0);
        assert_eq!(step(3, usize::MAX, Key::ArrowDown), 2);
    }
    #[test]
    fn module_labels_name_the_module_and_its_kind() {
        let en = crate::i18n::Language::En;
        assert_eq!(item_label(en, "Module1", "standard"), "Module1, module");
        assert_eq!(item_label(en, "Sheet1", "document"), "Sheet1, document module");
        assert_eq!(item_label(en, "Klass", "class"), "Klass, class module");
        assert_eq!(item_label(en, "UserForm1", "form"), "UserForm1, form");
    }
}
