use super::{CP_MAC_ROMAN, CP_UTF8, CP_WINDOWS_1252, decode, decode_utf16le, is_supported};

/// Windows-1252 0x80–0x9F as published (WHATWG index-windows-1252 / Unicode `bestfit1252`),
/// with the five undefined bytes as their same-valued C1 controls.
const CP1252_EXPECTED_80_9F: [(u8, u32); 32] = [
    (0x80, 0x20AC), // EURO SIGN
    (0x81, 0x0081), // undefined -> C1
    (0x82, 0x201A), // SINGLE LOW-9 QUOTATION MARK
    (0x83, 0x0192), // LATIN SMALL LETTER F WITH HOOK
    (0x84, 0x201E), // DOUBLE LOW-9 QUOTATION MARK
    (0x85, 0x2026), // HORIZONTAL ELLIPSIS
    (0x86, 0x2020), // DAGGER
    (0x87, 0x2021), // DOUBLE DAGGER
    (0x88, 0x02C6), // MODIFIER LETTER CIRCUMFLEX ACCENT
    (0x89, 0x2030), // PER MILLE SIGN
    (0x8A, 0x0160), // LATIN CAPITAL LETTER S WITH CARON
    (0x8B, 0x2039), // SINGLE LEFT-POINTING ANGLE QUOTATION MARK
    (0x8C, 0x0152), // LATIN CAPITAL LIGATURE OE
    (0x8D, 0x008D), // undefined -> C1
    (0x8E, 0x017D), // LATIN CAPITAL LETTER Z WITH CARON
    (0x8F, 0x008F), // undefined -> C1
    (0x90, 0x0090), // undefined -> C1
    (0x91, 0x2018), // LEFT SINGLE QUOTATION MARK
    (0x92, 0x2019), // RIGHT SINGLE QUOTATION MARK
    (0x93, 0x201C), // LEFT DOUBLE QUOTATION MARK
    (0x94, 0x201D), // RIGHT DOUBLE QUOTATION MARK
    (0x95, 0x2022), // BULLET
    (0x96, 0x2013), // EN DASH
    (0x97, 0x2014), // EM DASH
    (0x98, 0x02DC), // SMALL TILDE
    (0x99, 0x2122), // TRADE MARK SIGN
    (0x9A, 0x0161), // LATIN SMALL LETTER S WITH CARON
    (0x9B, 0x203A), // SINGLE RIGHT-POINTING ANGLE QUOTATION MARK
    (0x9C, 0x0153), // LATIN SMALL LIGATURE OE
    (0x9D, 0x009D), // undefined -> C1
    (0x9E, 0x017E), // LATIN SMALL LETTER Z WITH CARON
    (0x9F, 0x0178), // LATIN CAPITAL LETTER Y WITH DIAERESIS
];

/// Mac OS Roman spot checks typed from Unicode's `VENDORS/APPLE/ROMAN.TXT`.
const MAC_ROMAN_SPOT: [(u8, u32); 32] = [
    (0x80, 0x00C4), // A WITH DIAERESIS
    (0x87, 0x00E1), // a WITH ACUTE
    (0x8E, 0x00E9), // e WITH ACUTE
    (0x96, 0x00F1), // n WITH TILDE
    (0x9F, 0x00FC), // u WITH DIAERESIS
    (0xA0, 0x2020), // DAGGER
    (0xA5, 0x2022), // BULLET
    (0xA7, 0x00DF), // SHARP S
    (0xA9, 0x00A9), // COPYRIGHT SIGN
    (0xAA, 0x2122), // TRADE MARK SIGN
    (0xAD, 0x2260), // NOT EQUAL TO
    (0xB0, 0x221E), // INFINITY
    (0xB5, 0x00B5), // MICRO SIGN
    (0xB9, 0x03C0), // GREEK SMALL LETTER PI
    (0xBD, 0x03A9), // GREEK CAPITAL LETTER OMEGA
    (0xC0, 0x00BF), // INVERTED QUESTION MARK
    (0xC3, 0x221A), // SQUARE ROOT
    (0xC6, 0x2206), // INCREMENT
    (0xC9, 0x2026), // HORIZONTAL ELLIPSIS
    (0xCA, 0x00A0), // NO-BREAK SPACE
    (0xCE, 0x0152), // CAPITAL LIGATURE OE
    (0xD0, 0x2013), // EN DASH
    (0xD7, 0x25CA), // LOZENGE
    (0xD8, 0x00FF), // y WITH DIAERESIS
    (0xDA, 0x2044), // FRACTION SLASH
    (0xDB, 0x20AC), // EURO SIGN (post-1998 table)
    (0xDE, 0xFB01), // LIGATURE FI
    (0xE1, 0x00B7), // MIDDLE DOT
    (0xF0, 0xF8FF), // Apple logo (private use)
    (0xF5, 0x0131), // DOTLESS I
    (0xFA, 0x02D9), // DOT ABOVE
    (0xFF, 0x02C7), // CARON
];

fn single(cp: u16, b: u8) -> Vec<char> {
    let (text, unsupported) = decode(cp, &[b]);
    assert_eq!(unsupported, None);
    text.chars().collect()
}

fn ch(cp: u32) -> char {
    char::from_u32(cp).unwrap_or(char::REPLACEMENT_CHARACTER)
}

#[test]
fn supported_set_is_exactly_three() {
    assert!(is_supported(CP_WINDOWS_1252));
    assert!(is_supported(CP_UTF8));
    assert!(is_supported(CP_MAC_ROMAN));
    for cp in [0u16, 437, 850, 932, 936, 949, 950, 1250, 1251, 1200, 65000, 65535] {
        assert!(!is_supported(cp), "{cp}");
    }
}

#[test]
fn cp1252_every_byte() {
    for b in 0u8..=0xFF {
        let expected = match b {
            0x80..=0x9F => CP1252_EXPECTED_80_9F.iter().find(|(byte, _)| *byte == b).map(|&(_, cp)| ch(cp)).unwrap_or(char::REPLACEMENT_CHARACTER),
            _ => char::from(b),
        };
        assert_eq!(single(CP_WINDOWS_1252, b), vec![expected], "byte {b:#04X}");
    }
}

#[test]
fn cp1252_non_ascii_spot_checks() {
    // 0xA0–0xFF are Latin-1: spot-check that they are the identity.
    for (b, cp) in [(0xA0u8, 0xA0u32), (0xA9, 0xA9), (0xC4, 0xC4), (0xE9, 0xE9), (0xF1, 0xF1), (0xFF, 0xFF)] {
        assert_eq!(single(CP_WINDOWS_1252, b), vec![ch(cp)]);
    }
    let (text, unsupported) = decode(CP_WINDOWS_1252, b"\x93Caf\xE9\x94 \x80 5 \x96 \x99");
    assert_eq!(unsupported, None);
    assert_eq!(text, "\u{201C}Caf\u{E9}\u{201D} \u{20AC} 5 \u{2013} \u{2122}");
}

#[test]
fn mac_roman_every_byte() {
    let mut seen = Vec::with_capacity(128);
    for b in 0u8..=0xFF {
        let chars = single(CP_MAC_ROMAN, b);
        assert_eq!(chars.len(), 1, "byte {b:#04X}");
        let c = chars.first().copied().unwrap_or(char::REPLACEMENT_CHARACTER);
        if b < 0x80 {
            assert_eq!(c, char::from(b), "ASCII byte {b:#04X}");
        } else {
            // The high half is a bijection onto non-ASCII, non-replacement code points.
            assert!(!c.is_ascii() && c != char::REPLACEMENT_CHARACTER, "byte {b:#04X} -> {c:?}");
            assert!(!seen.contains(&c), "byte {b:#04X} repeats {c:?}");
            seen.push(c);
        }
    }
    assert_eq!(seen.len(), 128);
    for (b, cp) in MAC_ROMAN_SPOT {
        assert_eq!(single(CP_MAC_ROMAN, b), vec![ch(cp)], "byte {b:#04X}");
    }
}

#[test]
fn mac_roman_text() {
    let (text, unsupported) = decode(CP_MAC_ROMAN, b"Caf\x8E \xDB5 \xA9\xF0");
    assert_eq!(unsupported, None);
    assert_eq!(text, "Caf\u{E9} \u{20AC}5 \u{A9}\u{F8FF}");
}

#[test]
fn utf8_valid_invalid_and_lossy() {
    assert_eq!(decode(CP_UTF8, "Sub Café_€ 😀".as_bytes()), ("Sub Café_€ 😀".to_owned(), None));
    assert_eq!(decode(CP_UTF8, b""), (String::new(), None));
    // A lone continuation byte and a truncated sequence each become U+FFFD.
    assert_eq!(decode(CP_UTF8, b"a\x80b"), ("a\u{FFFD}b".to_owned(), None));
    assert_eq!(decode(CP_UTF8, b"x\xE2\x82"), ("x\u{FFFD}".to_owned(), None));
    // An overlong encoding of '/' is invalid.
    let (text, unsupported) = decode(CP_UTF8, b"\xC0\xAF");
    assert_eq!(unsupported, None);
    assert!(text.chars().all(|c| c == '\u{FFFD}') && !text.is_empty());
}

#[test]
fn unsupported_code_pages_fall_back_to_cp1252() {
    let bytes = b"Mod\x80\x81\xE9";
    let (expected, none) = decode(CP_WINDOWS_1252, bytes);
    assert_eq!(none, None);
    for cp in [932u16, 1251, 0, 65535] {
        assert_eq!(decode(cp, bytes), (expected.clone(), Some(cp)), "code page {cp}");
    }
    assert_eq!(decode(932, b""), (String::new(), Some(932)));
}

#[test]
fn utf16le_cases() {
    // BMP: "Aé€".
    assert_eq!(decode_utf16le(&[0x41, 0x00, 0xE9, 0x00, 0xAC, 0x20]), "A\u{E9}\u{20AC}");
    // Surrogate pair: U+1F600 = D83D DE00.
    assert_eq!(decode_utf16le(&[0x3D, 0xD8, 0x00, 0xDE]), "\u{1F600}");
    // Lone high surrogate, then lone low surrogate.
    assert_eq!(decode_utf16le(&[0x3D, 0xD8, 0x41, 0x00]), "\u{FFFD}A");
    assert_eq!(decode_utf16le(&[0x00, 0xDE]), "\u{FFFD}");
    // Odd length: the trailing byte is ignored.
    assert_eq!(decode_utf16le(&[0x41, 0x00, 0x42]), "A");
    assert_eq!(decode_utf16le(&[0x42]), "");
    assert_eq!(decode_utf16le(&[]), "");
}
