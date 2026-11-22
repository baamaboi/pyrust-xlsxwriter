//! Converting calamine cell values to Python objects.
//!
//! | calamine                  | Python                                          |
//! |---------------------------|-------------------------------------------------|
//! | `Empty`                   | `None`                                          |
//! | `Int`                     | `int`                                           |
//! | `Float`                   | `float`, or `int` when the value is integral    |
//! | `String`                  | `str`                                           |
//! | `Bool`                    | `bool`                                          |
//! | `Error`                   | the error text, e.g. `"#DIV/0!"`                |
//! | `DateTime` (date format)  | `datetime`, `time` (serial < 1) or `timedelta`  |
//! | `DateTimeIso` (ods, `t="d"`) | `datetime` / `time` (the text if unparsable) |
//! | `DurationIso` (ods)       | `timedelta` (the text if unparsable)            |
//!
//! xlsx stores every number as text without a type, so (like openpyxl, which
//! returns `int` for values written without a decimal point, as Excel does)
//! integral floats with a magnitude below 2**53 are returned as `int`.

use calamine::{Data, ExcelDateTime};
use chrono::{DateTime, NaiveDate, NaiveDateTime, NaiveTime, TimeDelta};
use pyo3::prelude::*;
use pyo3::types::{PyBool, PyFloat, PyString};

const MS_PER_DAY: f64 = 86_400_000.0;
/// Integral floats below this magnitude are returned as `int`.
const MAX_EXACT_INT: f64 = 9_007_199_254_740_992.0;

/// Python's `datetime` supports years 1..=9999.
fn python_year(year: i32) -> bool {
    (1..=9999).contains(&year)
}

fn none(py: Python<'_>) -> Bound<'_, PyAny> {
    py.None().into_bound(py)
}

pub fn float_to_py(py: Python<'_>, value: f64) -> Bound<'_, PyAny> {
    if value.fract() == 0.0 && value.abs() < MAX_EXACT_INT {
        // `fract()` is NaN for NaN and infinities, so those stay floats.
        (value as i64)
            .into_pyobject(py)
            .map(|v| v.into_any())
            .unwrap_or_else(|e| match e {})
    } else {
        PyFloat::new(py, value).into_any()
    }
}

/// Excel serial date to a naive datetime (`None` if out of Python's range).
pub fn serial_to_datetime(value: f64, is_1904: bool) -> Option<NaiveDateTime> {
    if !value.is_finite() {
        return None;
    }
    let (epoch, mut days) = if is_1904 {
        (NaiveDate::from_ymd_opt(1904, 1, 1)?, value)
    } else {
        (NaiveDate::from_ymd_opt(1899, 12, 30)?, value)
    };
    // Excel's fictitious 1900-02-29 (serial 60): serials before it are one day off.
    if !is_1904 && value > 0.0 && value < 60.0 {
        days += 1.0;
    }
    let ms = (days * MS_PER_DAY).round();
    // About 27,000 years either way: well outside Python's datetime range.
    if ms.abs() > 1e18 {
        return None;
    }
    let delta = TimeDelta::try_milliseconds(ms as i64)?;
    let dt = epoch.and_hms_opt(0, 0, 0)?.checked_add_signed(delta)?;
    use chrono::Datelike;
    python_year(dt.year()).then_some(dt)
}

/// A serial time of day (0 <= value < 1) to a time, if it does not round to a full day.
fn serial_to_time(value: f64) -> Option<NaiveTime> {
    let ms = (value * MS_PER_DAY).round() as u64;
    if ms >= MS_PER_DAY as u64 {
        return None;
    }
    NaiveTime::from_num_seconds_from_midnight_opt(
        (ms / 1000) as u32,
        ((ms % 1000) * 1_000_000) as u32,
    )
}

fn serial_to_duration(value: f64) -> Option<TimeDelta> {
    let ms = (value * MS_PER_DAY).round();
    // Python's timedelta is limited to 999,999,999 days.
    if !ms.is_finite() || ms.abs() >= 999_999_999.0 * MS_PER_DAY {
        return None;
    }
    TimeDelta::try_milliseconds(ms as i64)
}

/// Convert a date-formatted number, like openpyxl: durations (`[h]:mm`
/// formats) become `timedelta`, serials in `[0, 1)` become `time`, others
/// `datetime`. Values outside Python's ranges stay floats.
pub fn excel_datetime_to_py<'py>(
    py: Python<'py>,
    value: &ExcelDateTime,
    is_1904: bool,
) -> Bound<'py, PyAny> {
    let serial = value.as_f64();
    let converted = if value.is_duration() {
        serial_to_duration(serial).and_then(|d| d.into_pyobject(py).ok().map(Bound::into_any))
    } else if (0.0..1.0).contains(&serial) {
        match serial_to_time(serial) {
            Some(t) => t.into_pyobject(py).ok().map(Bound::into_any),
            None => serial_to_datetime(serial, is_1904)
                .and_then(|d| d.into_pyobject(py).ok().map(Bound::into_any)),
        }
    } else {
        serial_to_datetime(serial, is_1904)
            .and_then(|d| d.into_pyobject(py).ok().map(Bound::into_any))
    };
    converted.unwrap_or_else(|| PyFloat::new(py, serial).into_any())
}

/// An ISO 8601 date, datetime or time (ods `date-value`, xlsx `t="d"`).
pub enum IsoValue {
    DateTime(NaiveDateTime),
    Time(NaiveTime),
}

pub fn parse_iso_datetime(text: &str) -> Option<IsoValue> {
    let text = text.trim();
    for format in [
        "%Y-%m-%dT%H:%M:%S%.f",
        "%Y-%m-%dT%H:%M",
        "%Y-%m-%d %H:%M:%S%.f",
    ] {
        if let Ok(dt) = NaiveDateTime::parse_from_str(text, format) {
            return Some(IsoValue::DateTime(dt));
        }
    }
    if let Ok(dt) = DateTime::parse_from_rfc3339(text) {
        return Some(IsoValue::DateTime(dt.naive_local()));
    }
    if let Ok(dt) = text.parse::<DateTime<chrono::Utc>>() {
        return Some(IsoValue::DateTime(dt.naive_utc()));
    }
    if let Ok(d) = NaiveDate::parse_from_str(text, "%Y-%m-%d") {
        return d.and_hms_opt(0, 0, 0).map(IsoValue::DateTime);
    }
    if let Ok(t) = NaiveTime::parse_from_str(text, "%H:%M:%S%.f") {
        return Some(IsoValue::Time(t));
    }
    None
}

/// Parse an ISO 8601 duration such as `PT12H30M15S`, `-P1DT2H` or `PT0.5S`.
/// Years and months are rejected (their length is not fixed).
pub fn parse_iso_duration(text: &str) -> Option<TimeDelta> {
    let text = text.trim();
    let (negative, rest) = match text.strip_prefix('-') {
        Some(rest) => (true, rest),
        None => (false, text.strip_prefix('+').unwrap_or(text)),
    };
    let rest = rest.strip_prefix('P').or_else(|| rest.strip_prefix('p'))?;
    let mut seconds = 0.0f64;
    let mut number = String::new();
    let mut in_time = false;
    let mut any = false;
    for ch in rest.chars() {
        match ch {
            '0'..='9' | '.' => number.push(ch),
            ',' => number.push('.'),
            'T' | 't' => {
                if !number.is_empty() {
                    return None;
                }
                in_time = true;
            }
            _ => {
                let n: f64 = number.parse().ok()?;
                number.clear();
                any = true;
                seconds += n * match (ch.to_ascii_uppercase(), in_time) {
                    ('W', false) => 7.0 * 86_400.0,
                    ('D', false) => 86_400.0,
                    ('H', true) => 3_600.0,
                    ('M', true) => 60.0,
                    ('S', true) => 1.0,
                    _ => return None,
                };
            }
        }
    }
    if !number.is_empty() || !any {
        return None;
    }
    let ms = (seconds * 1000.0).round();
    if !ms.is_finite() || ms >= 999_999_999.0 * MS_PER_DAY {
        return None;
    }
    let delta = TimeDelta::try_milliseconds(ms as i64)?;
    Some(if negative { -delta } else { delta })
}

fn iso_datetime_to_py<'py>(py: Python<'py>, text: &str) -> Bound<'py, PyAny> {
    let converted = match parse_iso_datetime(text) {
        Some(IsoValue::DateTime(dt)) => {
            use chrono::Datelike;
            if python_year(dt.year()) {
                dt.into_pyobject(py).ok().map(Bound::into_any)
            } else {
                None
            }
        }
        Some(IsoValue::Time(t)) => t.into_pyobject(py).ok().map(Bound::into_any),
        None => None,
    };
    converted.unwrap_or_else(|| PyString::new(py, text).into_any())
}

fn iso_duration_to_py<'py>(py: Python<'py>, text: &str) -> Bound<'py, PyAny> {
    parse_iso_duration(text)
        .and_then(|d| d.into_pyobject(py).ok().map(Bound::into_any))
        .unwrap_or_else(|| PyString::new(py, text).into_any())
}

/// Convert one cell value to a Python object (never fails: values that
/// cannot be represented fall back to a float or the original text).
pub fn data_to_py<'py>(py: Python<'py>, value: &Data, is_1904: bool) -> Bound<'py, PyAny> {
    match value {
        Data::Empty => none(py),
        Data::Float(f) => float_to_py(py, *f),
        Data::String(s) => PyString::new(py, s).into_any(),
        Data::Int(i) => i
            .into_pyobject(py)
            .map(|v| v.into_any())
            .unwrap_or_else(|e| match e {}),
        Data::Bool(b) => PyBool::new(py, *b).to_owned().into_any(),
        Data::DateTime(dt) => excel_datetime_to_py(py, dt, is_1904),
        Data::Error(e) => PyString::new(py, &e.to_string()).into_any(),
        Data::DateTimeIso(s) => iso_datetime_to_py(py, s),
        Data::DurationIso(s) => iso_duration_to_py(py, s),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn dt(y: i32, m: u32, d: u32, h: u32, mi: u32, s: u32) -> NaiveDateTime {
        NaiveDate::from_ymd_opt(y, m, d)
            .unwrap()
            .and_hms_opt(h, mi, s)
            .unwrap()
    }

    #[test]
    fn serial_dates_1900() {
        assert_eq!(
            serial_to_datetime(1.0, false),
            Some(dt(1900, 1, 1, 0, 0, 0))
        );
        assert_eq!(
            serial_to_datetime(59.0, false),
            Some(dt(1900, 2, 28, 0, 0, 0))
        );
        assert_eq!(
            serial_to_datetime(61.0, false),
            Some(dt(1900, 3, 1, 0, 0, 0))
        );
        assert_eq!(
            serial_to_datetime(45292.5, false),
            Some(dt(2024, 1, 1, 12, 0, 0))
        );
        assert_eq!(serial_to_datetime(f64::NAN, false), None);
        assert_eq!(serial_to_datetime(1e300, false), None);
        assert_eq!(serial_to_datetime(-1e300, false), None);
        assert_eq!(serial_to_datetime(3e6, false), None);
    }

    #[test]
    fn serial_dates_1904() {
        assert_eq!(serial_to_datetime(0.0, true), Some(dt(1904, 1, 1, 0, 0, 0)));
        assert_eq!(
            serial_to_datetime(43830.0, true),
            Some(dt(2024, 1, 1, 0, 0, 0))
        );
    }

    #[test]
    fn serial_times_and_durations() {
        assert_eq!(serial_to_time(0.5), NaiveTime::from_hms_opt(12, 0, 0));
        assert_eq!(serial_to_time(0.999_999_999_9), None);
        assert_eq!(serial_to_duration(1.5), TimeDelta::try_hours(36));
        assert_eq!(serial_to_duration(f64::INFINITY), None);
    }

    #[test]
    fn iso_values() {
        assert!(
            matches!(parse_iso_datetime("2024-01-02"), Some(IsoValue::DateTime(d)) if d == dt(2024, 1, 2, 0, 0, 0))
        );
        assert!(
            matches!(parse_iso_datetime("2024-01-02T03:04:05"), Some(IsoValue::DateTime(d)) if d == dt(2024, 1, 2, 3, 4, 5))
        );
        assert!(
            matches!(parse_iso_datetime("2024-01-02T03:04:05Z"), Some(IsoValue::DateTime(d)) if d == dt(2024, 1, 2, 3, 4, 5))
        );
        assert!(matches!(
            parse_iso_datetime("13:30:00"),
            Some(IsoValue::Time(_))
        ));
        assert!(parse_iso_datetime("nope").is_none());
        assert_eq!(
            parse_iso_duration("PT12H30M15S"),
            TimeDelta::try_seconds(12 * 3600 + 30 * 60 + 15)
        );
        assert_eq!(parse_iso_duration("-P1DT2H"), TimeDelta::try_hours(-26));
        assert_eq!(
            parse_iso_duration("PT0.5S"),
            TimeDelta::try_milliseconds(500)
        );
        assert_eq!(parse_iso_duration("P1Y"), None);
        assert_eq!(parse_iso_duration("PT"), None);
        assert_eq!(parse_iso_duration("12:00"), None);
    }
}
