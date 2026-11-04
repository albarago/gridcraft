//! The keyword table: every [MS-VBAL] §3.3.5.2 reserved identifier, plus the contextual words the
//! syntactic grammar quotes but doesn't reserve (`Property`, `Step`, `Explicit`, …).
//!
//! One static table, kept sorted by ASCII-lowercased bytes so lookup is a case-insensitive binary
//! search (VBA identifiers are case-insensitive, §3.3.5). `tests::keyword_table_is_sorted` checks
//! the order.
//!
//! [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/

use std::cmp::Ordering;

/// Generates [`Keyword`], its spelling and reserved flag, and the lookup table from one list of
/// `Variant => "Spelling", reserved;` rows.
macro_rules! keywords {
    ($($variant:ident => $text:literal, $reserved:literal;)*) => {
        /// A keyword: a reserved identifier from [MS-VBAL] §3.3.5.2, or a contextual word
        /// ([`Keyword::is_reserved`] is `false`) that the grammar quotes but that is still a valid
        /// identifier, so a parser must accept it wherever an identifier may appear.
        ///
        /// [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/
        #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
        pub enum Keyword {
            $(
                #[doc = concat!("`", $text, "`")]
                $variant,
            )*
        }

        impl Keyword {
            /// The spelling used by [MS-VBAL] (contextual words and literal identifiers use their
            /// usual VBA capitalisation).
            ///
            /// [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/
            #[must_use]
            pub const fn as_str(self) -> &'static str {
                match self {
                    $(Keyword::$variant => $text,)*
                }
            }

            /// `true` for an [MS-VBAL] §3.3.5.2 reserved identifier, which can't be used as an
            /// IDENTIFIER; `false` for a contextual word.
            ///
            /// [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/
            #[must_use]
            pub const fn is_reserved(self) -> bool {
                match self {
                    $(Keyword::$variant => $reserved,)*
                }
            }
        }

        /// Every keyword with its spelling, sorted by ASCII-lowercased bytes.
        pub(super) const KEYWORDS: &[(&str, Keyword)] = &[$(($text, Keyword::$variant),)*];
    };
}

keywords! {
    Abs => "Abs", true;
    Access => "Access", false;
    AddressOf => "AddressOf", true;
    Alias => "Alias", false;
    And => "And", true;
    Any => "Any", true;
    Append => "Append", false;
    Array => "Array", true;
    As => "As", true;
    Attribute => "Attribute", true;
    Base => "Base", false;
    Binary => "Binary", false;
    Boolean => "Boolean", true;
    ByRef => "ByRef", true;
    Byte => "Byte", true;
    ByVal => "ByVal", true;
    Call => "Call", true;
    Case => "Case", true;
    CBool => "CBool", true;
    CByte => "CByte", true;
    CCur => "CCur", true;
    CDate => "CDate", true;
    CDbl => "CDbl", true;
    CDec => "CDec", true;
    CDecl => "CDecl", true;
    CInt => "CInt", true;
    Circle => "Circle", true;
    CLng => "CLng", true;
    CLngLng => "CLngLng", true;
    CLngPtr => "CLngPtr", true;
    Close => "Close", true;
    Compare => "Compare", false;
    Const => "Const", true;
    CSng => "CSng", true;
    CStr => "CStr", true;
    Currency => "Currency", true;
    CVar => "CVar", true;
    CVErr => "CVErr", true;
    Date => "Date", true;
    Debug => "Debug", true;
    Decimal => "Decimal", true;
    Declare => "Declare", true;
    DefBool => "DefBool", true;
    DefByte => "DefByte", true;
    DefCur => "DefCur", true;
    DefDate => "DefDate", true;
    DefDbl => "DefDbl", true;
    DefDec => "DefDec", true;
    DefInt => "DefInt", true;
    DefLng => "DefLng", true;
    DefLngLng => "DefLngLng", true;
    DefLngPtr => "DefLngPtr", true;
    DefObj => "DefObj", true;
    DefSng => "DefSng", true;
    DefStr => "DefStr", true;
    DefVar => "DefVar", true;
    Dim => "Dim", true;
    Do => "Do", true;
    DoEvents => "DoEvents", true;
    Double => "Double", true;
    Each => "Each", true;
    Else => "Else", true;
    ElseIf => "ElseIf", true;
    Empty => "Empty", true;
    End => "End", true;
    EndIf => "EndIf", true;
    Enum => "Enum", true;
    Eqv => "Eqv", true;
    Erase => "Erase", true;
    Error => "Error", false;
    Event => "Event", true;
    Exit => "Exit", true;
    Explicit => "Explicit", false;
    False => "False", true;
    Fix => "Fix", true;
    For => "For", true;
    Friend => "Friend", true;
    Function => "Function", true;
    Get => "Get", true;
    Global => "Global", true;
    GoSub => "GoSub", true;
    GoTo => "GoTo", true;
    If => "If", true;
    Imp => "Imp", true;
    Implements => "Implements", true;
    In => "In", true;
    Input => "Input", true;
    InputB => "InputB", true;
    Int => "Int", true;
    Integer => "Integer", true;
    Is => "Is", true;
    LBound => "LBound", true;
    Len => "Len", true;
    LenB => "LenB", true;
    Let => "Let", true;
    Lib => "Lib", false;
    Like => "Like", true;
    Line => "Line", false;
    LineInput => "LINEINPUT", true;
    Lock => "Lock", true;
    Long => "Long", true;
    LongLong => "LongLong", true;
    LongPtr => "LongPtr", true;
    Loop => "Loop", true;
    LSet => "LSet", true;
    Me => "Me", true;
    Mid => "Mid", false;
    MidB => "MidB", false;
    Mod => "Mod", true;
    Module => "Module", false;
    New => "New", true;
    Next => "Next", true;
    Not => "Not", true;
    Nothing => "Nothing", true;
    Null => "Null", true;
    Object => "Object", false;
    On => "On", true;
    Open => "Open", true;
    Option => "Option", true;
    Optional => "Optional", true;
    Or => "Or", true;
    Output => "Output", false;
    ParamArray => "ParamArray", true;
    Preserve => "Preserve", true;
    Print => "Print", true;
    Private => "Private", true;
    Property => "Property", false;
    PSet => "PSet", true;
    PtrSafe => "PtrSafe", false;
    Public => "Public", true;
    Put => "Put", true;
    RaiseEvent => "RaiseEvent", true;
    Random => "Random", false;
    Read => "Read", false;
    ReDim => "ReDim", true;
    Rem => "Rem", true;
    Reset => "Reset", false;
    Resume => "Resume", true;
    Return => "Return", true;
    RSet => "RSet", true;
    Scale => "Scale", true;
    Seek => "Seek", true;
    Select => "Select", true;
    Set => "Set", true;
    Sgn => "Sgn", true;
    Shared => "Shared", true;
    Single => "Single", true;
    Spc => "Spc", true;
    Static => "Static", true;
    Step => "Step", false;
    Stop => "Stop", true;
    String => "String", true;
    Sub => "Sub", true;
    Tab => "Tab", true;
    Text => "Text", false;
    Then => "Then", true;
    To => "To", true;
    True => "True", true;
    Type => "Type", true;
    TypeOf => "TypeOf", true;
    UBound => "UBound", true;
    Unlock => "Unlock", true;
    Until => "Until", true;
    Variant => "Variant", true;
    VbBase => "VB_Base", true;
    VbControl => "VB_Control", true;
    VbCreatable => "VB_Creatable", true;
    VbCustomizable => "VB_Customizable", true;
    VbDescription => "VB_Description", true;
    VbExposed => "VB_Exposed", true;
    VbExtKey => "VB_Ext_KEY", true;
    VbGlobalNameSpace => "VB_GlobalNameSpace", true;
    VbHelpId => "VB_HelpID", true;
    VbInvokeFunc => "VB_Invoke_Func", true;
    VbInvokeProperty => "VB_Invoke_Property", true;
    VbInvokePropertyPut => "VB_Invoke_PropertyPut", true;
    VbInvokePropertyPutRef => "VB_Invoke_PropertyPutRef", true;
    VbMemberFlags => "VB_MemberFlags", true;
    VbName => "VB_Name", true;
    VbPredeclaredId => "VB_PredeclaredId", true;
    VbProcData => "VB_ProcData", true;
    VbTemplateDerived => "VB_TemplateDerived", true;
    VbUserMemId => "VB_UserMemId", true;
    VbVarDescription => "VB_VarDescription", true;
    VbVarHelpId => "VB_VarHelpID", true;
    VbVarMemberFlags => "VB_VarMemberFlags", true;
    VbVarProcData => "VB_VarProcData", true;
    VbVarUserMemId => "VB_VarUserMemId", true;
    Wend => "Wend", true;
    While => "While", true;
    Width => "Width", false;
    With => "With", true;
    WithEvents => "WithEvents", true;
    Write => "Write", true;
    Xor => "Xor", true;
}

impl Keyword {
    /// Looks `word` up case-insensitively (`DIM`, `dim` and `Dim` all give [`Keyword::Dim`]).
    /// Returns `None` for anything that isn't exactly one keyword.
    #[must_use]
    pub fn lookup(word: &str) -> Option<Self> {
        let index = KEYWORDS.binary_search_by(|(text, _)| cmp_ignore_ascii_case(text, word)).ok()?;
        KEYWORDS.get(index).map(|&(_, keyword)| keyword)
    }

    /// Every keyword, in table order.
    pub fn all() -> impl Iterator<Item = Self> {
        KEYWORDS.iter().map(|&(_, keyword)| keyword)
    }

    /// Statements whose next `#` is a file number (`Print #1, x`, `Open f For Input As #1`),
    /// never the start of a date literal.
    pub(super) const fn takes_file_number(self) -> bool {
        matches!(
            self,
            Keyword::Print
                | Keyword::Write
                | Keyword::Input
                | Keyword::LineInput
                | Keyword::Get
                | Keyword::Put
                | Keyword::Seek
                | Keyword::Close
                | Keyword::Lock
                | Keyword::Unlock
                | Keyword::As
                | Keyword::Width
        )
    }
}

/// Compares two words byte by byte after ASCII-lowercasing both.
pub(super) fn cmp_ignore_ascii_case(a: &str, b: &str) -> Ordering {
    a.bytes().map(|byte| byte.to_ascii_lowercase()).cmp(b.bytes().map(|byte| byte.to_ascii_lowercase()))
}
