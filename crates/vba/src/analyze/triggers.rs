//! The Trigger_Table (Requirement 14.2): which VBA constructs set which Capability_Manifest
//! category, and which constructs make the manifest incomplete (14.4). The analyzer is generic
//! over these rows; adding a construct is adding a row.
//!
//! # Matching rules
//!
//! Names and ProgIDs compare ASCII case-insensitively.
//!
//! - [`Rule::Call`] with [`Callee::Vba`]: a VBA library function or statement, called bare
//!   (`Kill "a"`, `x = Dir("*.txt")`) or qualified as `VBA.`, `Interaction.`, `FileSystem.`,
//!   `VBA.Interaction.` or `VBA.FileSystem.` (the analyzer doesn't check which library module a
//!   function really lives in). A bare name with no arguments counts too (`x = Dir`).
//! - [`Rule::Call`] with [`Callee::On`]: a member on exactly one of the listed object paths
//!   (`Workbooks.Open`, `Application.Workbooks.Open`), or the same member inside a `With` on
//!   one of them.
//! - [`Rule::Call`] with [`Callee::Method`]: a member with that name on any object, also
//!   `With`-relative (`wb.SaveAs`, `UserForm1.Show`, `.PutInClipboard`). Not a bare name.
//! - [`Rule::ProgId`]: `CreateObject`/`GetObject` (as the 13.8 flags find them) with a literal
//!   ProgID equal to the pattern or starting with the pattern and a `.` (`MSXML2` matches
//!   `MSXML2.XMLHTTP.6.0`, `WScript.Shell` doesn't match `WScript.ShellX`).
//! - [`Rule::OpenMode`]: an `Open` statement whose mode (the first word after `For`) is the
//!   word; an `Open` without `For` is `Random`, VBA's default.
//! - [`Rule::OpenAccess`]: an `Open` statement whose `Access` clause holds the word
//!   (`Access Write`, `Access Read Write`).
//! - [`Rule::File`]: a file statement (`Print #`, `Get #`, …).
//! - [`Rule::Declare`]: a `Declare Sub|Function … Lib` statement.
//! - [`Rule::Type`]: a variable declared `As` the type (`As New` included), or `New` of it.
//! - [`Rule::Mark`]: a procedure with that [`Mark`] (Requirement 13.4, 13.5); the occurrence
//!   is the whole procedure. Procedures the parser recovered from count too.
//!
//! An occurrence is the call (statement or expression, arguments included), the statement, the
//! declaration or the procedure that matched.
//!
//! # Not covered
//!
//! - `Name old As new` isn't recognised by the parser, so the module is unparsed and the
//!   manifest incomplete (14.5) rather than showing file write.
//! - The `[A1]` evaluate shorthand always holds literal text, so it is never dynamic.
//! - `DataObject` created through `CreateObject("new:{…}")` isn't matched.
//! - Late-bound objects stored in variables: only the creation (`CreateObject`) is seen, plus the
//!   method rows (`.RegRead`, `.SaveAs`) on whatever object they are called on.
use Callee::{Method, On, Vba};
use Capability::{AutoExec, Clipboard, EventHandlers, FileRead, FileWrite, Network, Process, Registry, UiDialogs};
use Rule::{Call, Declare, File, OpenAccess, OpenMode, ProgId, Type};

use super::{Capability, Mark};
use crate::ast::FileOp;

/// One row: a construct that sets `category`.
#[derive(Debug)]
pub(super) struct Trigger {
    /// The category it sets.
    pub category: Capability,
    /// The construct's name in the manifest (`Open For Input`, `Workbooks.Open`).
    pub construct: &'static str,
    /// How it is matched.
    pub rule: Rule,
}

/// How a row matches; see the module docs.
#[derive(Debug)]
pub(super) enum Rule {
    /// A call or bare use of a name.
    Call(Callee),
    /// `CreateObject`/`GetObject` with this literal ProgID or ProgID prefix.
    ProgId(&'static str),
    /// `Open … For <word>`.
    OpenMode(&'static str),
    /// `Open … Access <… word …>`.
    OpenAccess(&'static str),
    /// A file statement.
    File(FileOp),
    /// `Declare`.
    Declare,
    /// `As [New] <type>` or `New <type>`.
    Type(&'static str),
    /// A procedure mark.
    Mark(Mark),
}

/// Which names a call row matches.
#[derive(Debug)]
pub(super) enum Callee {
    /// A VBA library name, bare or library-qualified.
    Vba(&'static str),
    /// A member on one of these object paths.
    On(&'static str, &'static [&'static [&'static str]]),
    /// A member on any object.
    Method(&'static str),
}

/// The qualifiers [`Callee::Vba`] accepts.
pub(super) const VBA_QUALIFIERS: &[&[&str]] = &[&["VBA"], &["Interaction"], &["FileSystem"], &["VBA", "Interaction"], &["VBA", "FileSystem"]];
/// `Application` as written in code.
pub(super) const APPLICATION: &[&[&str]] = &[&["Application"], &["Excel", "Application"]];
/// The `Workbooks` collection as written in code.
const WORKBOOKS: &[&[&str]] = &[&["Workbooks"], &["Application", "Workbooks"], &["Excel", "Application", "Workbooks"]];

const fn row(category: Capability, construct: &'static str, rule: Rule) -> Trigger {
    Trigger { category, construct, rule }
}

/// The Trigger_Table.
pub(super) static TRIGGERS: &[Trigger] = &[
    // File read.
    row(FileRead, "Open For Input", OpenMode("Input")),
    row(FileRead, "Open For Binary", OpenMode("Binary")),
    row(FileRead, "Open For Random", OpenMode("Random")),
    row(FileRead, "Input #", File(FileOp::Input)),
    row(FileRead, "Line Input #", File(FileOp::LineInput)),
    row(FileRead, "Get #", File(FileOp::Get)),
    row(FileRead, "Dir", Call(Vba("Dir"))),
    row(FileRead, "FileLen", Call(Vba("FileLen"))),
    row(FileRead, "FileDateTime", Call(Vba("FileDateTime"))),
    row(FileRead, "GetAttr", Call(Vba("GetAttr"))),
    row(FileRead, "Workbooks.Open", Call(On("Open", WORKBOOKS))),
    row(FileRead, "Workbooks.OpenText", Call(On("OpenText", WORKBOOKS))),
    row(FileRead, "Workbooks.OpenXML", Call(On("OpenXML", WORKBOOKS))),
    row(FileRead, "CreateObject(\"Scripting.FileSystemObject\")", ProgId("Scripting.FileSystemObject")),
    // File write.
    row(FileWrite, "Open For Output", OpenMode("Output")),
    row(FileWrite, "Open For Append", OpenMode("Append")),
    row(FileWrite, "Open Access Write", OpenAccess("Write")),
    row(FileWrite, "Print #", File(FileOp::Print)),
    row(FileWrite, "Write #", File(FileOp::Write)),
    row(FileWrite, "Put #", File(FileOp::Put)),
    row(FileWrite, "Kill", Call(Vba("Kill"))),
    row(FileWrite, "MkDir", Call(Vba("MkDir"))),
    row(FileWrite, "RmDir", Call(Vba("RmDir"))),
    row(FileWrite, "FileCopy", Call(Vba("FileCopy"))),
    row(FileWrite, "SetAttr", Call(Vba("SetAttr"))),
    row(FileWrite, "SaveAs", Call(Method("SaveAs"))),
    row(FileWrite, "SaveCopyAs", Call(Method("SaveCopyAs"))),
    row(FileWrite, "CreateObject(\"Scripting.FileSystemObject\")", ProgId("Scripting.FileSystemObject")),
    // Network.
    row(Network, "CreateObject(\"MSXML2.*\")", ProgId("MSXML2")),
    row(Network, "CreateObject(\"WinHttp.*\")", ProgId("WinHttp")),
    row(Network, "CreateObject(\"Microsoft.XMLHTTP\")", ProgId("Microsoft.XMLHTTP")),
    row(Network, "CreateObject(\"InternetExplorer.Application\")", ProgId("InternetExplorer.Application")),
    // Process.
    row(Process, "Shell", Call(Vba("Shell"))),
    row(Process, "CreateObject(\"WScript.Shell\")", ProgId("WScript.Shell")),
    row(Process, "CreateObject(\"Shell.Application\")", ProgId("Shell.Application")),
    row(Process, "Declare", Declare),
    // Registry.
    row(Registry, "GetSetting", Call(Vba("GetSetting"))),
    row(Registry, "GetAllSettings", Call(Vba("GetAllSettings"))),
    row(Registry, "SaveSetting", Call(Vba("SaveSetting"))),
    row(Registry, "DeleteSetting", Call(Vba("DeleteSetting"))),
    row(Registry, "RegRead", Call(Method("RegRead"))),
    row(Registry, "RegWrite", Call(Method("RegWrite"))),
    row(Registry, "RegDelete", Call(Method("RegDelete"))),
    // Clipboard.
    row(Clipboard, "DataObject", Type("DataObject")),
    row(Clipboard, "DataObject", Type("MSForms.DataObject")),
    row(Clipboard, "PutInClipboard", Call(Method("PutInClipboard"))),
    row(Clipboard, "GetFromClipboard", Call(Method("GetFromClipboard"))),
    // UI dialogs.
    row(UiDialogs, "MsgBox", Call(Vba("MsgBox"))),
    row(UiDialogs, "InputBox", Call(Vba("InputBox"))),
    row(UiDialogs, "Application.InputBox", Call(On("InputBox", APPLICATION))),
    row(UiDialogs, "Show", Call(Method("Show"))),
    row(UiDialogs, "Application.GetOpenFilename", Call(On("GetOpenFilename", APPLICATION))),
    row(UiDialogs, "Application.GetSaveAsFilename", Call(On("GetSaveAsFilename", APPLICATION))),
    row(UiDialogs, "Application.FileDialog", Call(On("FileDialog", APPLICATION))),
    // Event handlers and auto-exec.
    row(EventHandlers, "event handler", Rule::Mark(Mark::EventHandler)),
    row(AutoExec, "auto-exec", Rule::Mark(Mark::AutoExec)),
];

/// One dynamic-dispatch row (14.4): a call the analyzer can't resolve when `arg` isn't a literal.
#[derive(Debug)]
pub(super) struct Dynamic {
    /// The construct's name in the manifest.
    pub construct: &'static str,
    /// What is called.
    pub callee: Callee,
    /// The argument that names what runs: its parameter name and positional index. `None` when
    /// any call is dynamic.
    pub arg: Option<(&'static str, usize)>,
}

/// The dynamic-dispatch rows. `CreateObject`/`GetObject` with a non-literal ProgID are dynamic
/// too; the analyzer takes them from the 13.8 flags. A missing argument counts as non-literal.
/// The bare `Run` and `Evaluate` are `Application`'s global members.
pub(super) static DYNAMIC: &[Dynamic] = &[
    Dynamic { construct: "CallByName", callee: Vba("CallByName"), arg: None },
    Dynamic { construct: "Application.Run", callee: On("Run", APPLICATION), arg: Some(("Macro", 0)) },
    Dynamic { construct: "Application.Run", callee: Vba("Run"), arg: Some(("Macro", 0)) },
    Dynamic { construct: "Evaluate", callee: Vba("Evaluate"), arg: Some(("Name", 0)) },
    Dynamic { construct: "Evaluate", callee: Method("Evaluate"), arg: Some(("Name", 0)) },
];
