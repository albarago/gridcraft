//! Date-literal recognition ([MS-VBAL] §3.3.3).
//!
//! The grammar here is `date-or-time` and its parts. The surrounding `#` delimiters and the
//! optional WSC just inside them are handled by the lexer (`"#" *WSC date-or-time *WSC "#"`).
//! The ABNF is ambiguous (a `date-separator` may be plain whitespace, so `1/2 12:00` could be a
//! three-part date `1/2 12` or a date and a time), so [`date_or_time`] returns every way the text
//! can be read and the lexer picks the longest one that is followed by the closing `#`.
//!
//! Only the shape is checked. Whether the numbers form a real calendar date (the static semantics
//! of §3.3.3) is for a later phase.
//!
//! [MS-VBAL]: https://learn.microsoft.com/en-us/openspecs/microsoft_general_purpose_programming_languages/ms-vbal/

use super::{byte_at, skip_wsc};

/// `English-month-name / English-month-abbreviation`, longest first so `june` wins over `jun`.
const MONTH_NAMES: [&str; 23] = [
    "september",
    "february",
    "november",
    "december",
    "january",
    "october",
    "august",
    "march",
    "april",
    "june",
    "july",
    "may",
    "jan",
    "feb",
    "mar",
    "apr",
    "jun",
    "jul",
    "aug",
    "sep",
    "oct",
    "nov",
    "dec",
];

/// `ampm`'s words, longest first.
const AMPM: [&str; 4] = ["am", "pm", "a", "p"];

/// The most readings [`date_or_time`] can return: two date-value lengths, each with or without a
/// time, and a time alone.
const MAX_READINGS: usize = 5;

/// The end offsets of each way `src[pos..]` starts with a `date-or-time`.
#[derive(Debug, Default)]
pub(super) struct Readings {
    ends: [usize; MAX_READINGS],
    len: usize,
}

impl Readings {
    fn push(&mut self, end: usize) {
        if let Some(slot) = self.ends.get_mut(self.len) {
            *slot = end;
            self.len = self.len.saturating_add(1);
        }
    }

    /// The end offsets found, in no particular order.
    pub(super) fn ends(&self) -> &[usize] {
        self.ends.get(..self.len).unwrap_or(&[])
    }
}

/// `date-or-time = (date-value 1*WSC time-value) / date-value / time-value`, starting at `pos`.
pub(super) fn date_or_time(src: &str, pos: usize) -> Readings {
    let mut readings = Readings::default();
    for date_end in date_values(src, pos).into_iter().flatten() {
        readings.push(date_end);
        let time_start = skip_wsc(src, date_end);
        if time_start > date_end
            && let Some(time_end) = time_value(src, time_start)
        {
            readings.push(time_end);
        }
    }
    if let Some(time_end) = time_value(src, pos) {
        readings.push(time_end);
    }
    readings
}

/// `date-value = left-date-value date-separator middle-date-value [date-separator
/// right-date-value]`: the end after the middle part, and the end after the right part if there
/// is one.
fn date_values(src: &str, pos: usize) -> [Option<usize>; 2] {
    let Some(middle_end) = date_part(src, pos).and_then(|left| date_separator(src, left)).and_then(|middle| date_part(src, middle)) else {
        return [None, None];
    };
    let right_end = date_separator(src, middle_end).and_then(|right| date_part(src, right));
    [Some(middle_end), right_end]
}

/// `left-date-value / middle-date-value / right-date-value = decimal-literal / month-name`.
fn date_part(src: &str, pos: usize) -> Option<usize> {
    decimal(src, pos).or_else(|| month_name(src, pos))
}

/// `date-separator = 1*WSC / (*WSC ("/" / "-" / ",") *WSC)`.
fn date_separator(src: &str, pos: usize) -> Option<usize> {
    let after_space = skip_wsc(src, pos);
    match byte_at(src, after_space) {
        Some(b'/' | b'-' | b',') => Some(skip_wsc(src, after_space.saturating_add(1))),
        _ if after_space > pos => Some(after_space),
        _ => None,
    }
}

/// `time-value = (hour-value ampm) / (hour-value time-separator minute-value [time-separator
/// second-value] [ampm])`, the longest match.
fn time_value(src: &str, pos: usize) -> Option<usize> {
    let hour_end = decimal(src, pos)?;
    let Some(minute_end) = time_separator(src, hour_end).and_then(|minute| decimal(src, minute)) else {
        return ampm(src, hour_end);
    };
    let end = time_separator(src, minute_end).and_then(|second| decimal(src, second)).unwrap_or(minute_end);
    Some(ampm(src, end).unwrap_or(end))
}

/// `time-separator = *WSC (":" / ".") *WSC`.
fn time_separator(src: &str, pos: usize) -> Option<usize> {
    let after_space = skip_wsc(src, pos);
    match byte_at(src, after_space) {
        Some(b':' | b'.') => Some(skip_wsc(src, after_space.saturating_add(1))),
        _ => None,
    }
}

/// `ampm = *WSC ("am" / "pm" / "a" / "p")`, case-insensitive like every ABNF quoted string.
fn ampm(src: &str, pos: usize) -> Option<usize> {
    let start = skip_wsc(src, pos);
    word_at(src, start, &AMPM)
}

/// `month-name`, case-insensitive.
fn month_name(src: &str, pos: usize) -> Option<usize> {
    word_at(src, pos, &MONTH_NAMES)
}

/// The end of the first of `words` found at `pos`, compared ASCII case-insensitively.
fn word_at(src: &str, pos: usize, words: &[&str]) -> Option<usize> {
    let rest = src.as_bytes().get(pos..)?;
    words
        .iter()
        .find(|word| rest.get(..word.len()).is_some_and(|head| head.eq_ignore_ascii_case(word.as_bytes())))
        .map(|word| pos.saturating_add(word.len()))
}

/// `decimal-literal = 1*decimal-digit`.
fn decimal(src: &str, pos: usize) -> Option<usize> {
    let end = super::digits_end(src, pos, |byte| byte.is_ascii_digit());
    (end > pos).then_some(end)
}

/// `true` when a date-or-time reading can only be meant as a date: it uses `/`, `-` or `:`, or a
/// month name or am/pm marker (the only letters the grammar allows). A `#` followed by such text
/// and no closing `#` is an unterminated date; anything less date-like (`#1, 2`) stays a `#`
/// token, because `#` is also a file-number prefix and a type suffix.
pub(super) fn looks_like_date(text: &str) -> bool {
    text.bytes().any(|byte| matches!(byte, b'/' | b'-' | b':') || byte.is_ascii_alphabetic())
}
