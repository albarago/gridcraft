//! Text decoding for the code pages a VBA project names in its `PROJECTCODEPAGE` record
//! ([MS-OVBA] §2.3.4.2.1.4), plus UTF-16LE for the `*UNICODE` records.
//!
//! Phase 1 decodes three code pages with built-in tables and no new dependency:
//!
//! - **1252 (Windows-1252).** Bytes 0x00–0x7F and 0xA0–0xFF map to the code point with the same
//!   value. 0x80–0x9F use the Windows-1252 table. The five positions that table leaves undefined
//!   (0x81, 0x8D, 0x8F, 0x90, 0x9D) map to the C1 control with the same value, as the WHATWG
//!   Encoding Standard's "windows-1252" decoder does. Decoding never fails.
//! - **65001 (UTF-8).** Invalid sequences become U+FFFD (`String::from_utf8_lossy`).
//! - **10000 (Mac OS Roman).** 0x00–0x7F are ASCII. 0x80–0xFF follow the Unicode consortium's
//!   `MAPPINGS/VENDORS/APPLE/ROMAN.TXT` (the post-1998 table): 0xDB is U+20AC EURO SIGN (not the
//!   older U+00A4 CURRENCY SIGN) and 0xF0, the Apple logo, is the private-use U+F8FF.
//!
//! **Deviation from Requirement 5 criterion 5.** The requirement proposes a larger set (including
//! the East Asian multibyte code pages 932, 936, 949 and 950). Phase 1 decodes any other code page
//! as Windows-1252 and reports it, so the caller can attach an `UnsupportedCodePage` diagnostic.
//! Covering the rest means either more hand-typed tables or adding `encoding_rs`; adding that
//! dependency is a maintainers' decision for later (follow-up VBA-1.2b).
//!
//! [MS-OVBA]: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/

/// Windows-1252 (Western European).
pub const CP_WINDOWS_1252: u16 = 1252;
/// UTF-8.
pub const CP_UTF8: u16 = 65001;
/// Mac OS Roman.
pub const CP_MAC_ROMAN: u16 = 10000;

/// Windows-1252 bytes 0x80–0x9F. Undefined positions are the same-valued C1 control (WHATWG).
static CP1252_80_9F: [char; 32] = [
    '\u{20AC}', '\u{0081}', '\u{201A}', '\u{0192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}', // 80-87
    '\u{02C6}', '\u{2030}', '\u{0160}', '\u{2039}', '\u{0152}', '\u{008D}', '\u{017D}', '\u{008F}', // 88-8F
    '\u{0090}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}', // 90-97
    '\u{02DC}', '\u{2122}', '\u{0161}', '\u{203A}', '\u{0153}', '\u{009D}', '\u{017E}', '\u{0178}', // 98-9F
];

/// Mac OS Roman bytes 0x80–0xFF, typed from Unicode's `VENDORS/APPLE/ROMAN.TXT`.
static MAC_ROMAN_80_FF: [char; 128] = [
    '\u{00C4}', '\u{00C5}', '\u{00C7}', '\u{00C9}', '\u{00D1}', '\u{00D6}', '\u{00DC}', '\u{00E1}', // 80-87
    '\u{00E0}', '\u{00E2}', '\u{00E4}', '\u{00E3}', '\u{00E5}', '\u{00E7}', '\u{00E9}', '\u{00E8}', // 88-8F
    '\u{00EA}', '\u{00EB}', '\u{00ED}', '\u{00EC}', '\u{00EE}', '\u{00EF}', '\u{00F1}', '\u{00F3}', // 90-97
    '\u{00F2}', '\u{00F4}', '\u{00F6}', '\u{00F5}', '\u{00FA}', '\u{00F9}', '\u{00FB}', '\u{00FC}', // 98-9F
    '\u{2020}', '\u{00B0}', '\u{00A2}', '\u{00A3}', '\u{00A7}', '\u{2022}', '\u{00B6}', '\u{00DF}', // A0-A7
    '\u{00AE}', '\u{00A9}', '\u{2122}', '\u{00B4}', '\u{00A8}', '\u{2260}', '\u{00C6}', '\u{00D8}', // A8-AF
    '\u{221E}', '\u{00B1}', '\u{2264}', '\u{2265}', '\u{00A5}', '\u{00B5}', '\u{2202}', '\u{2211}', // B0-B7
    '\u{220F}', '\u{03C0}', '\u{222B}', '\u{00AA}', '\u{00BA}', '\u{03A9}', '\u{00E6}', '\u{00F8}', // B8-BF
    '\u{00BF}', '\u{00A1}', '\u{00AC}', '\u{221A}', '\u{0192}', '\u{2248}', '\u{2206}', '\u{00AB}', // C0-C7
    '\u{00BB}', '\u{2026}', '\u{00A0}', '\u{00C0}', '\u{00C3}', '\u{00D5}', '\u{0152}', '\u{0153}', // C8-CF
    '\u{2013}', '\u{2014}', '\u{201C}', '\u{201D}', '\u{2018}', '\u{2019}', '\u{00F7}', '\u{25CA}', // D0-D7
    '\u{00FF}', '\u{0178}', '\u{2044}', '\u{20AC}', '\u{2039}', '\u{203A}', '\u{FB01}', '\u{FB02}', // D8-DF
    '\u{2021}', '\u{00B7}', '\u{201A}', '\u{201E}', '\u{2030}', '\u{00C2}', '\u{00CA}', '\u{00C1}', // E0-E7
    '\u{00CB}', '\u{00C8}', '\u{00CD}', '\u{00CE}', '\u{00CF}', '\u{00CC}', '\u{00D3}', '\u{00D4}', // E8-EF
    '\u{F8FF}', '\u{00D2}', '\u{00DA}', '\u{00DB}', '\u{00D9}', '\u{0131}', '\u{02C6}', '\u{02DC}', // F0-F7
    '\u{00AF}', '\u{02D8}', '\u{02D9}', '\u{02DA}', '\u{00B8}', '\u{02DD}', '\u{02DB}', '\u{02C7}', // F8-FF
];

/// Whether [`decode`] has a real decoder for `code_page` (rather than the Windows-1252 fallback).
pub fn is_supported(code_page: u16) -> bool {
    matches!(code_page, CP_WINDOWS_1252 | CP_UTF8 | CP_MAC_ROMAN)
}

/// Decode MBCS bytes in `code_page`.
///
/// Returns the text, plus `Some(code_page)` when the code page is unsupported and the bytes were
/// decoded as Windows-1252 instead (the caller turns that into an `UnsupportedCodePage`
/// diagnostic). Never fails and never panics; the only allocation is the output string.
pub fn decode(code_page: u16, bytes: &[u8]) -> (String, Option<u16>) {
    match code_page {
        CP_UTF8 => (String::from_utf8_lossy(bytes).into_owned(), None),
        CP_MAC_ROMAN => (decode_single_byte(bytes, mac_roman_char), None),
        CP_WINDOWS_1252 => (decode_single_byte(bytes, cp1252_char), None),
        other => (decode_single_byte(bytes, cp1252_char), Some(other)),
    }
}

/// Decode UTF-16LE (`MODULENAMEUNICODE` and the other `*UNICODE` records).
///
/// Unpaired surrogates become U+FFFD. An odd trailing byte is ignored.
pub fn decode_utf16le(bytes: &[u8]) -> String {
    let units = bytes.as_chunks::<2>().0.iter().map(|&pair| u16::from_le_bytes(pair));
    char::decode_utf16(units).map(|unit| unit.unwrap_or(char::REPLACEMENT_CHARACTER)).collect()
}

fn decode_single_byte(bytes: &[u8], map: fn(u8) -> char) -> String {
    let mut out = String::with_capacity(bytes.len());
    out.extend(bytes.iter().map(|&b| map(b)));
    out
}

fn cp1252_char(b: u8) -> char {
    match b {
        0x80..=0x9F => CP1252_80_9F.get(usize::from(b & 0x7F)).copied().unwrap_or(char::REPLACEMENT_CHARACTER),
        _ => char::from(b),
    }
}

fn mac_roman_char(b: u8) -> char {
    if b < 0x80 { char::from(b) } else { MAC_ROMAN_80_FF.get(usize::from(b & 0x7F)).copied().unwrap_or(char::REPLACEMENT_CHARACTER) }
}

#[cfg(test)]
mod tests;
